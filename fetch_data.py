import gspread
from oauth2client.service_account import ServiceAccountCredentials
import json
import os
from datetime import datetime
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import requests

def extract_sheet_id_from_url(url):
    """Extract Google Sheet ID from sharing URL"""
    if '/spreadsheets/d/' in url:
        return url.split('/spreadsheets/d/')[1].split('/')[0]
    return None

def setup_google_services(credentials_json):
    """Set up Google Sheets and Drive services"""
    scope = [
        'https://spreadsheets.google.com/feeds',
        'https://www.googleapis.com/auth/drive',
        'https://www.googleapis.com/auth/spreadsheets'
    ]
    
    creds_dict = json.loads(credentials_json)
    creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    
    # Set up both gspread and direct API access
    gspread_client = gspread.authorize(creds)
    sheets_service = build('sheets', 'v4', credentials=creds)
    
    return gspread_client, sheets_service

def get_sheet_data_with_formatting(sheets_service, spreadsheet_id, sheet_name):
    """Get sheet data including formatting information"""
    try:
        # First, get the sheet ID for the specific tab
        spreadsheet = sheets_service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        
        sheet_id = None
        for sheet in spreadsheet['sheets']:
            if sheet['properties']['title'] == sheet_name:
                sheet_id = sheet['properties']['sheetId']
                break
        
        if sheet_id is None:
            raise ValueError(f"Sheet '{sheet_name}' not found")
        
        # Get values
        values_result = sheets_service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=f'{sheet_name}!A:Z'
        ).execute()
        
        values = values_result.get('values', [])
        
        # Get formatting data
        formatting_result = sheets_service.spreadsheets().get(
            spreadsheetId=spreadsheet_id,
            ranges=[f'{sheet_name}!A:Z'],
            includeGridData=True
        ).execute()
        
        return values, formatting_result
        
    except HttpError as e:
        print(f"HTTP Error: {e}")
        return None, None
    except Exception as e:
        print(f"Error getting sheet data: {e}")
        return None, None

def check_cell_strikethrough(formatting_data, sheet_name, row_index, col_index):
    """Check if a specific cell has strikethrough formatting"""
    try:
        sheets = formatting_data.get('sheets', [])
        
        for sheet in sheets:
            if sheet['properties']['title'] == sheet_name:
                grid_data = sheet.get('data', [])
                
                if grid_data and len(grid_data) > 0:
                    row_data = grid_data[0].get('rowData', [])
                    
                    if row_index < len(row_data):
                        row = row_data[row_index]
                        values = row.get('values', [])
                        
                        if col_index < len(values):
                            cell = values[col_index]
                            
                            # Check for strikethrough in effective format
                            effective_format = cell.get('effectiveFormat', {})
                            text_format = effective_format.get('textFormat', {})
                            
                            if text_format.get('strikethrough', False):
                                return True
                            
                            # Also check userEnteredFormat
                            user_format = cell.get('userEnteredFormat', {})
                            user_text_format = user_format.get('textFormat', {})
                            
                            if user_text_format.get('strikethrough', False):
                                return True
                            
                            # Check for inline formatting in textFormatRuns
                            text_format_runs = cell.get('textFormatRuns', [])
                            for run in text_format_runs:
                                run_format = run.get('format', {})
                                if run_format.get('strikethrough', False):
                                    return True
                
                break
        
        return False
        
    except Exception as e:
        print(f"Error checking strikethrough for cell ({row_index}, {col_index}): {e}")
        return False

def analyze_text_with_manual_indicators(text):
    """Check for manual strikethrough indicators as fallback"""
    if not text or not isinstance(text, str):
        return False
    
    text_upper = text.upper().strip()
    
    # Manual indicators that suggest completion
    completion_indicators = [
        '[DONE]', '[COMPLETED]', '[FINISHED]', '[COMPLETE]',
        '✓', '✗', 'DONE:', 'COMPLETED:', 'FINISHED:',
        '(DONE)', '(COMPLETED)', '(FINISHED)',
        '- DONE', '- COMPLETED', '- FINISHED'
    ]
    
    for indicator in completion_indicators:
        if indicator in text_upper:
            return True
    
    # Check for markdown-style strikethrough
    if text.count('~~') >= 2:
        return True
    
    return False

def send_discord_notification(webhook_url, result):
    """Send notification to Discord using webhook with support for long todo lists"""
    
    if not webhook_url:
        print("No Discord webhook URL provided, skipping notification")
        return False
    
    try:
        # Determine embed color based on status
        if result['status'] != 'success':
            color = 0xFF0000  # Red for errors
        elif result['has_new_questions']:
            color = 0xFF9900  # Orange for new questions
        elif result['todo_count'] == 0:
            color = 0x00FF00  # Green for all done
        else:
            color = 0x0099FF  # Blue for normal status
        
        embeds = []
        
        # Main embed
        main_embed = {
            "title": "📋 Group 4 Questions Update",
            "color": color,
            "timestamp": result['timestamp'],
            "footer": {
                "text": "Questions Parser Bot"
            }
        }
        
        if result['status'] == 'success':
            # Status field
            status_emoji = "🆕" if result['has_new_questions'] else "✅" if result['todo_count'] == 0 else "📊"
            main_embed["description"] = f"{status_emoji} **Status Update**"
            
            # Summary fields
            main_embed["fields"] = [
                {
                    "name": "📊 Summary",
                    "value": f"**Total Questions:** {result['total_questions']}\n**✅ Done:** {result['done_count']}\n**📝 Todo:** {result['todo_count']}",
                    "inline": True
                }
            ]
            
            # Add alert for new questions
            if result['has_new_questions']:
                main_embed["fields"].insert(0, {
                    "name": "🚨 Alert",
                    "value": f"**{result['todo_count']} question(s) need attention!**",
                    "inline": False
                })
            
            # Add recently completed with trimming (as before)
            if result['done_questions']:
                done_text = ""
                for i, q in enumerate(result['done_questions'][-3:], 1):  # Show last 3 completed
                    trimmed_text = q['text'][:60] + ('...' if len(q['text']) > 60 else '')
                    done_text += f"~~{trimmed_text}~~\n"
                
                main_embed["fields"].append({
                    "name": f"✅ Recently Completed",
                    "value": done_text if done_text else "None",
                    "inline": False
                })
        
        else:
            # Error status
            main_embed["description"] = f"❌ **Error occurred**"
            main_embed["fields"] = [
                {
                    "name": "Error Message",
                    "value": result.get('message', 'Unknown error'),
                    "inline": False
                }
            ]
        
        embeds.append(main_embed)
        
        # Add todo questions in separate embeds if needed - NO TRIMMING
        if result['status'] == 'success' and result['todo_questions']:
            todo_text = ""
            questions_in_current_embed = 0
            embed_count = 1
            
            for i, q in enumerate(result['todo_questions'], 1):
                question_line = f"{i}. {q['text']}\n"
                
                # Check if adding this question would exceed Discord's field limit
                if len(todo_text + question_line) > 1000:  # Leave buffer for Discord limits
                    # Create new embed for previous questions
                    todo_embed = {
                        "color": color,
                        "fields": [
                            {
                                "name": f"📝 Todo Questions ({embed_count})",
                                "value": todo_text,
                                "inline": False
                            }
                        ]
                    }
                    embeds.append(todo_embed)
                    
                    # Start new embed
                    todo_text = question_line
                    embed_count += 1
                else:
                    todo_text += question_line
                
                questions_in_current_embed += 1
            
            # Add remaining questions
            if todo_text:
                todo_embed = {
                    "color": color,
                    "fields": [
                        {
                            "name": f"📝 Todo Questions ({embed_count})" if embed_count > 1 else f"📝 Todo Questions ({result['todo_count']})",
                            "value": todo_text,
                            "inline": False
                        }
                    ]
                }
                embeds.append(todo_embed)
        
        # Prepare webhook payload
        payload = {
            "embeds": embeds
        }
        
        # Add mention for urgent cases (optional)
        if result.get('has_new_questions') and result.get('todo_count', 0) > 3:
            payload["content"] = "@here Multiple questions need attention!"
        
        # Send to Discord
        response = requests.post(webhook_url, json=payload)
        response.raise_for_status()
        
        print("✅ Discord notification sent successfully")
        print(f"📊 Sent {len(embeds)} embed(s) to Discord")
        return True
        
    except requests.exceptions.RequestException as e:
        print(f"❌ Failed to send Discord notification: {e}")
        return False
    except Exception as e:
        print(f"❌ Error creating Discord notification: {e}")
        return False



def send_simple_discord_message(webhook_url, message):
    """Send a simple text message to Discord"""
    try:
        payload = {"content": message}
        response = requests.post(webhook_url, json=payload)
        response.raise_for_status()
        return True
    except Exception as e:
        print(f"Failed to send simple Discord message: {e}")
        return False

def parse_group4_questions(sheet_url, credentials_json):
    """Parse the Questions tab and extract Group 4 column data with formatting"""
    
    if not credentials_json:
        return {
            "status": "error",
            "message": "Google credentials required for formatting detection",
            "timestamp": datetime.now().isoformat(),
            "done_questions": [],
            "todo_questions": [],
            "total_questions": 0,
            "has_new_questions": False
        }
    
    try:
        gspread_client, sheets_service = setup_google_services(credentials_json)
        sheet_id = extract_sheet_id_from_url(sheet_url)
        
        # Get data with formatting
        values, formatting_data = get_sheet_data_with_formatting(
            sheets_service, sheet_id, "Questions"
        )
        
        if not values:
            return {
                "status": "error",
                "message": "No data found in Questions sheet",
                "timestamp": datetime.now().isoformat(),
                "done_questions": [],
                "todo_questions": [],
                "total_questions": 0,
                "has_new_questions": False
            }
        
    except Exception as e:
        print(f"Error accessing sheet: {e}")
        return {
            "status": "error",
            "message": f"Error accessing sheet: {str(e)}",
            "timestamp": datetime.now().isoformat(),
            "done_questions": [],
            "todo_questions": [],
            "total_questions": 0,
            "has_new_questions": False
        }
    
    # Find "Group 4" column
    header_row = values[0] if values else []
    group4_column_index = None
    
    for i, header in enumerate(header_row):
        if header and "Group4" in str(header):
            group4_column_index = i
            break
    
    if group4_column_index is None:
        print("Group 4 column not found")
        return {
            "status": "error",
            "message": "Group 4 column not found",
            "timestamp": datetime.now().isoformat(),
            "done_questions": [],
            "todo_questions": [],
            "total_questions": 0,
            "has_new_questions": False
        }
    
    # Extract Group 4 column data (skip header row)
    group4_data = []
    done_questions = []
    todo_questions = []
    
    for row_index in range(1, len(values)):
        row = values[row_index]
        
        if group4_column_index < len(row):
            cell_content = row[group4_column_index]
            
            if cell_content and str(cell_content).strip():
                
                # Check for strikethrough formatting
                has_strikethrough = check_cell_strikethrough(
                    formatting_data, "Questions", row_index, group4_column_index
                )
                
                # Fallback to manual indicators if no formatting detected
                if not has_strikethrough:
                    has_strikethrough = analyze_text_with_manual_indicators(cell_content)
                
                print(f"  Strikethrough detected: {has_strikethrough}")
                
                question_data = {
                    "row_number": row_index + 1,
                    "text": str(cell_content).strip(),
                    "is_crossed_out": has_strikethrough,
                    "formatting_source": "api_formatting" if check_cell_strikethrough(
                        formatting_data, "Questions", row_index, group4_column_index
                    ) else "manual_indicators"
                }
                
                group4_data.append(question_data)
                
                if has_strikethrough:
                    done_questions.append(question_data)
                else:
                    todo_questions.append(question_data)
    
    # Determine if there are new questions
    has_new_questions = len(todo_questions) > 0
    
    result = {
        "status": "success",
        "timestamp": datetime.now().isoformat(),
        "sheet_url": sheet_url,
        "group4_column_index": group4_column_index + 1,  # 1-indexed for user display
        "total_questions": len(group4_data),
        "done_count": len(done_questions),
        "todo_count": len(todo_questions),
        "has_new_questions": has_new_questions,
        "done_questions": done_questions,
        "todo_questions": todo_questions,
        "all_questions": group4_data
    }
    
    return result

def main():
    sheet_url = os.getenv('SHEET_URL')
    credentials_json = os.getenv('GOOGLE_CREDENTIALS')
    discord_webhook = os.getenv('DISCORD_WEBHOOK_URL')

    if not sheet_url:
        raise ValueError("SHEET_URL environment variable not set")
    
    if not credentials_json:
        raise ValueError("GOOGLE_CREDENTIALS environment variable not set for formatting detection")
    
    print(f"Processing sheet...")

    # Parse the questions
    result = parse_group4_questions(sheet_url, credentials_json)
    
    # Create output directory
    os.makedirs('output', exist_ok=True)
    
    # Save detailed results
    with open('output/group4_questions.json', 'w') as f:
        json.dump(result, f, indent=2, ensure_ascii=False)
    
    # Save summary
    summary = {
        "status": result["status"],
        "timestamp": result["timestamp"],
        "has_new_questions": result["has_new_questions"],
        "todo_count": result["todo_count"],
        "done_count": result["done_count"],
        "total_questions": result["total_questions"]
    }
    
    with open('output/summary.json', 'w') as f:
        json.dump(summary, f, indent=2)

    # Send Discord notification
    if discord_webhook:
        notification_sent = send_discord_notification(discord_webhook, result)
        
        # Save notification status
        summary["discord_notification_sent"] = notification_sent
        with open('output/summary.json', 'w') as f:
            json.dump(summary, f, indent=2)
    
    # Print summary
    print("\n" + "="*50)
    print("SUMMARY")
    print("="*50)
    print(f"Status: {result['status']}")
    
    if result['status'] == 'success':
        print("Successfully processed questions!")
    else:
        print(f"Error: {result.get('message', 'Unknown error')}")
    
    print("\nResults saved to output/ directory")

if __name__ == "__main__":
    main()
