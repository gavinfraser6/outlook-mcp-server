import datetime
import os
import win32com.client
from typing import List, Optional, Dict, Any
from mcp.server.fastmcp import FastMCP, Context
import json
import re # Import regex for parsing actions

# Initialize FastMCP server
# Now enhanced with bilingual (English/Afrikaans) email analysis capabilities.
mcp = FastMCP("outlook-assistant")

# Constants
MAX_DAYS = 180
ACTIONABLE_EMAIL_MAX_DAYS = 60

# Email cache for storing retrieved emails by number
email_cache = {}

# Helper functions
def connect_to_outlook():
    """Connect to Outlook application using COM"""
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        return outlook, namespace
    except Exception as e:
        raise Exception(f"Failed to connect to Outlook: {str(e)}")

def get_manager_name(namespace) -> Optional[str]:
    """
    Tries to get the name of the current user's manager from Outlook/Exchange.
    Returns the manager's name or None if it cannot be determined.
    """
    try:
        currentUser = namespace.CurrentUser
        if currentUser:
            exchangeUser = currentUser.GetExchangeUser()
            if exchangeUser:
                manager = exchangeUser.GetManager()
                if manager:
                    return manager.Name
    except Exception as e:
        print(f"Warning: Could not determine manager's name: {str(e)}")
        # This can fail if not on an Exchange account, which is fine.
        return None
    return None

def get_my_email_address(namespace) -> Optional[str]:
    """
    Gets the primary SMTP email address of the current user.
    """
    try:
        if namespace.CurrentUser:
            # First, try the most reliable method for Exchange accounts
            exchange_user = namespace.CurrentUser.AddressEntry.GetExchangeUser()
            if exchange_user and exchange_user.PrimarySmtpAddress:
                return exchange_user.PrimarySmtpAddress.lower()
        # Fallback for non-Exchange or other setups
        if namespace.Accounts:
            for account in namespace.Accounts:
                if account.SmtpAddress:
                    return account.SmtpAddress.lower()
    except Exception as e:
        print(f"Warning: Could not determine user's email address: {str(e)}")
    return None


def get_folder_by_name(namespace, folder_name: str):
    """Get a specific Outlook folder by name"""
    try:
        # First check inbox subfolder
        inbox = namespace.GetDefaultFolder(6)  # 6 is the index for inbox folder
        
        # Check inbox subfolders first (most common)
        for folder in inbox.Folders:
            if folder.Name.lower() == folder_name.lower():
                return folder
                
        # Then check all folders at root level
        for folder in namespace.Folders:
            if folder.Name.lower() == folder_name.lower():
                return folder
            
            # Also check subfolders
            for subfolder in folder.Folders:
                if subfolder.Name.lower() == folder_name.lower():
                    return subfolder
                    
        # If not found
        return None
    except Exception as e:
        raise Exception(f"Failed to access folder {folder_name}: {str(e)}")

def format_email(mail_item) -> Dict[str, Any]:
    """Format an Outlook mail item into a structured dictionary"""
    try:
        # Extract recipients
        recipients = []
        if mail_item.Recipients:
            for i in range(1, mail_item.Recipients.Count + 1):
                recipient = mail_item.Recipients(i)
                try:
                    recipients.append(f"{recipient.Name} <{recipient.Address}>")
                except:
                    recipients.append(f"{recipient.Name}")
        
        is_sent_item = False
        try:
            # A simple way to check if it's a sent item is to see if its parent folder is the 'Sent Items' folder
            if mail_item.Parent.EntryID == mail_item.Application.GetNamespace("MAPI").GetDefaultFolder(5).EntryID:
                is_sent_item = True
        except:
            pass
            
        # Format the email data
        email_data = {
            "id": mail_item.EntryID,
            "conversation_id": mail_item.ConversationID if hasattr(mail_item, 'ConversationID') else None,
            "subject": mail_item.Subject,
            "sender": mail_item.SenderName,
            "sender_email": mail_item.SenderEmailAddress,
            "received_time": mail_item.ReceivedTime.strftime("%Y-%m-%d") if hasattr(mail_item, 'ReceivedTime') and mail_item.ReceivedTime else None,
            "sent_time": mail_item.SentOn.strftime("%Y-%m-%d %H:%M:%S") if hasattr(mail_item, 'SentOn') and mail_item.SentOn else None,
            "is_sent_item": is_sent_item,
            "recipients": recipients,
            "body": mail_item.Body,
            "has_attachments": mail_item.Attachments.Count > 0,
            "attachment_count": mail_item.Attachments.Count if hasattr(mail_item, 'Attachments') else 0,
            "unread": mail_item.UnRead if hasattr(mail_item, 'UnRead') else False,
            "importance": mail_item.Importance if hasattr(mail_item, 'Importance') else 1, # 0=Low, 1=Normal, 2=High
            "categories": mail_item.Categories if hasattr(mail_item, 'Categories') else "",
            "flagged": mail_item.FlagStatus if hasattr(mail_item, 'FlagStatus') else 0 # 0=olNoFlag, 1=olMarkComplete, 2=olFlagged, 3=olFollowUp, 4=olForward, 5=olReply, 6=olUnflagged
        }
        return email_data
    except Exception as e:
        raise Exception(f"Failed to format email: {str(e)}")

def get_todays_appointments(namespace) -> List[Dict[str, Any]]:
    """
    Fetches and formats today's calendar appointments from Outlook.

    Args:
        namespace: The active Outlook MAPI namespace.

    Returns:
        A list of dictionaries, where each dictionary represents an appointment for today.
    """
    appointments_list = []
    try:
        calendar = namespace.GetDefaultFolder(9) # 9 is the index for the Calendar folder
        items = calendar.Items
        
        # This is crucial to include recurring appointments in the search
        items.IncludeRecurrences = True
        items.Sort("[Start]")

        # Define the time range for today (from midnight to midnight)
        today_start = datetime.datetime.now().replace(hour=0, minute=0, second=0)
        today_end = today_start + datetime.timedelta(days=1)
        
        # Format dates for the Outlook filter string (e.g., '10/26/2023 12:00 AM')
        start_str = today_start.strftime('%m/%d/%Y %H:%M %p')
        end_str = today_end.strftime('%m/%d/%Y %H:%M %p')
        
        # The filter finds items that start before the end of today AND end after the start of today.
        # This correctly captures all-day events and events that may span across midnight.
        restriction = f"[Start] < '{end_str}' AND [End] > '{start_str}'"
        
        restricted_items = items.Restrict(restriction)

        for item in restricted_items:
            appointments_list.append({
                "subject": item.Subject,
                "start": item.Start, # Keep as datetime for sorting
                "end": item.End,
                "location": item.Location if item.Location else "Not specified",
                "all_day": item.AllDayEvent
            })
    except Exception as e:
        print(f"Warning: Could not retrieve calendar appointments: {e}")
    
    # Sort by start time before returning
    appointments_list.sort(key=lambda x: x["start"])
    return appointments_list

def get_todays_tasks(namespace) -> List[Dict[str, Any]]:
    """Fetches incomplete Outlook tasks due on or before today."""
    tasks_list = []
    try:
        tasks_folder = namespace.GetDefaultFolder(13) # 13 is Tasks folder
        items = tasks_folder.Items
        items.Sort("[DueDate]")
        items.IncludeRecurrences = True # Important for recurring tasks

        today_str = datetime.date.today().strftime('%m/%d/%Y')

        # Filter for tasks that are not complete and are due on or before today (includes overdue)
        restriction = f"[Complete] = false AND [DueDate] <= '{today_str}'"
        
        restricted_items = items.Restrict(restriction)

        for item in restricted_items:
            # Check if it's actually a task item (olTask = 48)
            if hasattr(item, 'Class') and item.Class == 48:
                tasks_list.append({
                    "subject": item.Subject,
                    "due_date": item.DueDate.strftime('%Y-%m-%d') if hasattr(item, 'DueDate') and item.DueDate else "No due date"
                })
    except Exception as e:
        print(f"Warning: Could not retrieve Outlook tasks: {e}")

    return tasks_list

def clear_email_cache():
    """Clear the email cache"""
    global email_cache
    email_cache = {}

def get_emails_from_folder(folder, days: int, search_term: Optional[str] = None):
    """Get emails from a folder with optional search filter"""
    emails_list = []
    
    # Calculate the date threshold
    now = datetime.datetime.now()
    threshold_date = now - datetime.timedelta(days=days)
    
    try:
        # Set up filtering
        folder_items = folder.Items
        folder_items.Sort("[ReceivedTime]", True)  # Sort by received time, newest first
        
        # If we have a search term, apply it
        if search_term:
            # Handle OR operators in search term
            search_terms = [term.strip() for term in search_term.split(" OR ")]
            
            # Try to create a filter for subject, sender name or body
            try:
                # Build SQL filter with OR conditions for each search term
                sql_conditions = []
                for term in search_terms:
                    sql_conditions.append(f"\"urn:schemas:httpmail:subject\" LIKE '%{term}%'")
                    sql_conditions.append(f"\"urn:schemas:httpmail:fromname\" LIKE '%{term}%'")
                    sql_conditions.append(f"\"urn:schemas:httpmail:textdescription\" LIKE '%{term}%'")
                
                filter_term = f"@SQL=" + " OR ".join(sql_conditions)
                folder_items = folder_items.Restrict(filter_term)
            except:
                # If filtering fails, we'll do manual filtering later
                pass
        
        # Process emails
        count = 0
        for item in folder_items:
            try:
                if hasattr(item, 'ReceivedTime') and item.ReceivedTime:
                    # Convert to naive datetime for comparison
                    received_time = item.ReceivedTime.replace(tzinfo=None)
                    
                    # Skip emails older than our threshold
                    if received_time < threshold_date:
                        continue
                    
                    # Manual search filter if needed
                    if search_term and folder_items == folder.Items:  # If we didn't apply filter earlier
                        # Handle OR operators in search term for manual filtering
                        search_terms = [term.strip().lower() for term in search_term.split(" OR ")]
                        
                        # Check if any of the search terms match
                        found_match = False
                        for term in search_terms:
                            if (term in item.Subject.lower() or 
                                term in item.SenderName.lower() or 
                                term in item.Body.lower()):
                                found_match = True
                                break
                        
                        if not found_match:
                            continue
                    
                    # Format and add the email
                    email_data = format_email(item)
                    emails_list.append(email_data)
                    count += 1
            except Exception as e:
                print(f"Warning: Error processing email: {str(e)}")
                continue
                
    except Exception as e:
        print(f"Error retrieving emails: {str(e)}")
        
    return emails_list

@mcp.tool()
def create_outlook_task(subject: str, due_date_str: str, reminder_time_str: Optional[str] = None) -> str:
    """
    Creates a new task in Outlook's To-Do list.

    Args:
        subject: The subject or name of the task (e.g., "Follow up on project proposal").
        due_date_str: The date the task is due. Can be a string like "tomorrow", "next Friday", or "2024-10-31".
        reminder_time_str: Optional. The time for a reminder (e.g., "9:00 AM"). If provided, a reminder will be set.

    Returns:
        A confirmation message indicating success or failure.
    """
    print(f"Tool: create_outlook_task called with subject='{subject}', due_date='{due_date_str}', reminder_time='{reminder_time_str}'")
    try:
        outlook, _ = connect_to_outlook()
        task = outlook.CreateItem(3) # 3 represents olTaskItem
        task.Subject = subject
        task.DueDate = due_date_str # win32com is smart enough to parse "tomorrow", "next monday", etc.

        if reminder_time_str:
            task.ReminderSet = True
            # The COM object can often parse combined date/time strings correctly.
            task.ReminderTime = f"{due_date_str} {reminder_time_str}"

        task.Save()
        # To provide better feedback, let's get the actual parsed due date.
        actual_due_date = task.DueDate.strftime('%A, %B %d, %Y')
        msg = f"Success: Task '{subject}' created, due on {actual_due_date}."
        if task.ReminderSet:
            actual_reminder_time = task.ReminderTime.strftime('%I:%M %p').lstrip('0')
            msg += f" A reminder is set for {actual_reminder_time}."
        return msg
    except Exception as e:
        return f"Error creating Outlook task: {str(e)}"

@mcp.tool()
def get_outlook_tasks(due: str = 'today') -> str:
    """
    Retrieves incomplete tasks from Outlook based on their due date.

    Args:
        due: A filter for which tasks to retrieve. Can be 'today', 'tomorrow', 'this week', or 'all'. Defaults to 'today'.

    Returns:
        A JSON string representing a list of due tasks, or a message if none are found.
    """
    print(f"Tool: get_outlook_tasks called with due='{due}'")
    due = due.lower().strip()
    if due not in ['today', 'tomorrow', 'this week', 'all']:
        return "Error: Invalid 'due' parameter. Must be one of 'today', 'tomorrow', 'this week', or 'all'."

    try:
        _, namespace = connect_to_outlook()
        tasks_folder = namespace.GetDefaultFolder(13) # 13 is Tasks folder
        items = tasks_folder.Items
        items.Sort("[DueDate]")
        items.IncludeRecurrences = True

        # Build filter string
        today = datetime.date.today()
        # Use a format Outlook understands well in filters
        today_str = today.strftime('%m/%d/%Y')

        if due == 'today':
            # Includes overdue tasks
            restriction = f"[Complete] = false AND [DueDate] <= '{today_str}'"
        elif due == 'tomorrow':
            tomorrow = today + datetime.timedelta(days=1)
            tomorrow_str = tomorrow.strftime('%m/%d/%Y')
            restriction = f"[Complete] = false AND [DueDate] = '{tomorrow_str}'"
        elif due == 'this week':
            start_of_week = today - datetime.timedelta(days=today.weekday())
            end_of_week = start_of_week + datetime.timedelta(days=6)
            start_str = start_of_week.strftime('%m/%d/%Y')
            end_str = end_of_week.strftime('%m/%d/%Y')
            restriction = f"[Complete] = false AND [DueDate] >= '{start_str}' AND [DueDate] <= '{end_str}'"
        else: # 'all'
            restriction = "[Complete] = false"

        due_tasks = []
        restricted_items = items.Restrict(restriction)

        for item in restricted_items:
            # olTask = 48
            if hasattr(item, 'Class') and item.Class == 48:
                 due_tasks.append({
                    "subject": item.Subject,
                    "due_date": item.DueDate.strftime('%Y-%m-%d') if hasattr(item, 'DueDate') and item.DueDate else "No due date",
                    "reminder_set": item.ReminderSet
                })

        if not due_tasks:
            return json.dumps({"message": f"No incomplete tasks found for the '{due}' category."})

        return json.dumps(due_tasks, indent=2)

    except Exception as e:
        return f"Error retrieving Outlook tasks: {str(e)}"

@mcp.tool()
def mark_task_complete(task_subject: str) -> str:
    """
    Finds an incomplete task by its exact subject and marks it as complete.

    Args:
        task_subject: The exact subject of the task to mark as complete.

    Returns:
        A confirmation message.
    """
    print(f"Tool: mark_task_complete called for subject='{task_subject}'")
    try:
        _, namespace = connect_to_outlook()
        tasks_folder = namespace.GetDefaultFolder(13) # 13 is Tasks folder
        
        # Using a simple filter is generally safer than a raw SQL-style one
        restriction = f"[Subject] = '{task_subject}' AND [Complete] = False"
        tasks = tasks_folder.Items.Restrict(restriction)
        
        if tasks.Count == 0:
            return f"Error: No active task with the subject '{task_subject}' was found."
        
        if tasks.Count > 1:
            print(f"Warning: Found {tasks.Count} active tasks with the same subject. Completing the first one found.")

        # Get the first task from the filtered results
        task_to_complete = tasks.Item(1)
        task_to_complete.MarkComplete() # This is the dedicated method for a TaskItem

        return f"Success: Task '{task_subject}' has been marked as complete."
    except Exception as e:
        return f"Error marking task as complete: {str(e)}"

@mcp.tool()
def generate_morning_briefing(days_to_scan: int = 3, follow_up_days: int = 2) -> str:
    """
    **OPTIMIZED FOR AI ANALYSIS & NOW INCLUDES OUTLOOK TASKS**
    Gathers all data for a comprehensive morning briefing.

    This tool acts as a data aggregator. It fetches calendar appointments for today, analyzes
    recent email conversations, and **retrieves any due Outlook tasks**. It then returns a structured
    JSON object containing all this raw data.

    Your task is to analyze the returned JSON to construct a user-friendly briefing. Identify:
    1.  **Today's Tasks**: Check the `todays_reminders` section first (this contains due tasks).
    2.  **Today's Schedule**: Review the `todays_calendar` section.
    3.  **Email Priorities**: Analyze `conversation_threads` for urgent items, replies needed, and follow-ups.
    4.  Synthesize these points into a clear, actionable summary for the user.

    Args:
        days_to_scan: How many days back to scan for relevant email threads (default 3, max 14).
        follow_up_days: The number of days to wait before an item might be considered "awaiting reply" (default 2).

    Returns:
        A JSON string containing calendar data, a list of active conversation threads, and today's due tasks.
    """
    print(f"Tool: generate_morning_briefing (AI-driven) called with days_to_scan={days_to_scan}, follow_up_days={follow_up_days}")
    if not isinstance(days_to_scan, int) or not 1 <= days_to_scan <= 14:
        return "Error: 'days_to_scan' must be an integer between 1 and 14."
    if not isinstance(follow_up_days, int) or follow_up_days < 1:
        return "Error: 'follow_up_days' must be a positive integer."
    
    try:
        # 1. Data Collection
        _, namespace = connect_to_outlook()
        my_email = get_my_email_address(namespace)
        if not my_email:
            return "Error: Could not determine your email address. Cannot analyze conversations."

        manager_name = get_manager_name(namespace)
        inbox = namespace.GetDefaultFolder(6)
        sent_folder = namespace.GetDefaultFolder(5)

        # -- Calendar Data --
        todays_appointments = get_todays_appointments(namespace)
        # Convert datetime objects to strings for clean JSON serialization
        for app in todays_appointments:
            app['start'] = app['start'].strftime('%I:%M %p').lstrip('0')
            app['end'] = app['end'].strftime('%I:%M %p').lstrip('0')

        # -- Email Data --
        start_date = datetime.datetime.now() - datetime.timedelta(days=days_to_scan)
        start_date_str = start_date.strftime('%m/%d/%Y %H:%M %p')
        
        inbox_items = inbox.Items.Restrict(f"[ReceivedTime] >= '{start_date_str}'")
        sent_items = sent_folder.Items.Restrict(f"[SentOn] >= '{start_date_str}'")
        all_items = list(inbox_items) + list(sent_items)
        
        conversations = {}
        for item in all_items:
            try:
                conv_id = item.ConversationID
                if conv_id not in conversations: conversations[conv_id] = []
                conversations[conv_id].append(item)
            except Exception: continue

        def get_item_datetime(item):
            dt = getattr(item, 'ReceivedTime', None) or getattr(item, 'SentOn', None)
            return dt.replace(tzinfo=None) if dt else datetime.datetime.min

        for conv_id in conversations:
            conversations[conv_id].sort(key=get_item_datetime)

        # 2. Data Processing for AI
        clear_email_cache()
        email_number = 1
        threads_for_ai = []

        for conv_id, thread in conversations.items():
            if not thread: continue
            
            last_email_item = thread[-1]
            last_email_data = format_email(last_email_item)
            
            # Cache the full data for other tools
            email_cache[email_number] = last_email_data

            is_from_me = my_email in last_email_data.get('sender_email', '').lower() if last_email_data.get('sender_email') else False
            
            time_since_last_email = datetime.datetime.now() - get_item_datetime(last_email_item)
            
            thread_status = {
                "email_number": email_number,
                "subject": last_email_data.get('subject'),
                "last_email_from": "me" if is_from_me else last_email_data.get('sender'),
                "last_email_timestamp": get_item_datetime(last_email_item).strftime('%Y-%m-%d %H:%M'),
                "is_last_email_unread": last_email_data.get('unread', False) and not is_from_me,
                "is_from_manager": manager_name and manager_name.lower() in last_email_data.get('sender', '').lower(),
                "contains_question_in_body": '?' in last_email_data.get('body', ''),
                "days_since_last_email": time_since_last_email.days
            }
            
            # Add context for the AI to decide if a follow-up is needed
            if is_from_me and time_since_last_email.days >= follow_up_days:
                thread_status["follow_up_suggestion"] = f"Awaiting reply for {time_since_last_email.days} days."
            
            threads_for_ai.append(thread_status)
            email_number += 1

        # -- Fetch Today's Outlook Tasks --
        todays_tasks = get_todays_tasks(namespace)
        
        # 3. Construct Final JSON Payload
        briefing_data = {
            "briefing_metadata": {
                "date": datetime.date.today().strftime('%A, %B %d, %Y'),
                "user_email": my_email,
                "manager_name": manager_name or "Not found"
            },
            "todays_reminders": todays_tasks, # Key is "reminders" for consistent AI interpretation, value is today's tasks
            "todays_calendar": todays_appointments,
            "conversation_threads": sorted(
                threads_for_ai,
                key=lambda x: (not x['is_last_email_unread'], x['last_email_timestamp']),
                reverse=True
            ),
             "analysis_instructions": "Review reminders (which are Outlook tasks), calendar, and conversation_threads to create a prioritized morning briefing for the user."
        }

        return json.dumps(briefing_data, indent=2)

    except Exception as e:
        error_payload = {
            "status": "error",
            "message": f"An error occurred while gathering briefing data: {str(e)}"
        }
        return json.dumps(error_payload, indent=2)

# MCP Tools
@mcp.tool()
def prioritize_inbox(days: int = 1, max_emails_to_scan: int = 25) -> str:
    """
    **OPTIMIZED FOR AI ANALYSIS**
    Fetches recent emails from the inbox for AI-powered prioritization.

    This tool does NOT rank emails itself. Instead, it gathers raw email data (sender, subject, body snippet)
    and returns it as a JSON string. YOU, the AI assistant, must then analyze this data to identify
    which emails are most important based on their content, context, and tone.

    Use this tool when the user asks "what are my most important emails?", "triage my inbox", or
    "what needs my attention?". You should then process the returned JSON to provide a
    human-readable summary to the user, explaining WHY each email is a priority.

    Args:
        days: Number of days to look back for emails (default 1, max 31). A smaller number is better to keep the data manageable.
        max_emails_to_scan: The maximum number of emails to retrieve for analysis (default 25).

    Returns:
        A JSON string representing a list of emails. Each email is a dictionary with keys:
        'email_number', 'sender', 'subject', 'body_snippet', 'received_time', 'importance', and 'is_from_manager'.
        The 'email_number' can be used with other tools like `get_email_by_number`.
        Returns an error message if it fails.
    """
    print(f"Tool: prioritize_inbox (AI-driven) called with days={days}, max_emails_to_scan={max_emails_to_scan}")

    # Parameter validation
    if not isinstance(days, int) or not 1 <= days <= 31:
        return "Error: 'days' must be an integer between 1 and 31 for AI analysis."
    if not isinstance(max_emails_to_scan, int) or not 5 <= max_emails_to_scan <= 50:
        return "Error: 'max_emails_to_scan' must be between 5 and 50."

    try:
        _, namespace = connect_to_outlook()
        manager_name = get_manager_name(namespace)
        inbox = namespace.GetDefaultFolder(6)
        # We still use the robust get_emails_from_folder helper
        emails = get_emails_from_folder(inbox, days)

        if not emails:
            return json.dumps({"status": "success", "message": f"No emails found in the Inbox from the last {days} {'day' if days == 1 else 'days'}."})

        emails_for_ai = []
        # Limit the list to avoid excessive token usage for the AI
        emails_to_process = emails[:max_emails_to_scan]
        
        clear_email_cache() # Clear cache before populating it for this session

        for i, email_data in enumerate(emails_to_process, 1):
            # Cache the FULL original data, so get_email_by_number still works
            email_cache[i] = email_data
            
            is_from_manager_flag = False
            if manager_name and manager_name.lower() in email_data.get('sender', '').lower():
                is_from_manager_flag = True

            # Create a dictionary with just the data the AI needs for analysis
            emails_for_ai.append({
                "email_number": i, # Crucial for follow-up actions like 'get_email_by_number(5)'
                "sender": email_data.get('sender'),
                "subject": email_data.get('subject'),
                "body_snippet": (email_data.get('body', '') or "").strip()[:500] + '...', # Truncate body to save tokens
                "received_time": email_data.get('received_time'),
                "importance": "High" if email_data.get('importance') == 2 else "Normal",
                "is_from_manager": is_from_manager_flag
            })
        
        if not emails_for_ai:
             return json.dumps({"status": "success", "message": "No suitable emails were found for analysis."})

        # Return the data as a JSON string for the AI to parse and analyze
        return json.dumps(emails_for_ai, indent=2)

    except Exception as e:
        return f"Error fetching emails for AI analysis: {str(e)}"

@mcp.tool()
def inbox_load_and_mood_estimator(days_to_scan: int = 30) -> str:
    """
    **OPTIMIZED FOR AI ANALYSIS**
    Calculates key metrics about the inbox for AI-powered 'load' or 'stress' analysis.

    This tool does NOT calculate a 'load score' or 'mood'. Instead, it provides raw metrics
    like the count of urgent emails, flagged items, and average response delay.
    YOU, the AI assistant, must interpret these metrics to assess the overall state
    of the user's inbox and provide a qualitative summary (e.g., 'calm', 'busy', 'overloaded')
    and suggest a course of action.

    Args:
        days_to_scan: How many days back to scan for emails to calculate metrics (max 60).

    Returns:
        A JSON string containing key performance indicators for the inbox, such as:
        'unread_urgent_count', 'flagged_threads_count', 'average_response_delay_hours',
        and 'total_active_conversations'.
    """
    print(f"Tool: inbox_load_and_mood_estimator (AI-driven) called with days_to_scan={days_to_scan}")
    if not isinstance(days_to_scan, int) or not 1 <= days_to_scan <= ACTIONABLE_EMAIL_MAX_DAYS:
        return f"Error: 'days_to_scan' must be an integer between 1 and {ACTIONABLE_EMAIL_MAX_DAYS}."

    try:
        _, namespace = connect_to_outlook()
        my_email = get_my_email_address(namespace)
        if not my_email:
            return "Error: Could not determine your email address. Cannot estimate inbox load."

        inbox = namespace.GetDefaultFolder(6)
        sent_folder = namespace.GetDefaultFolder(5)
        
        start_date = datetime.datetime.now() - datetime.timedelta(days=days_to_scan)
        start_date_str = start_date.strftime('%m/%d/%Y %H:%M %p')

        inbox_items = inbox.Items.Restrict(f"[ReceivedTime] >= '{start_date_str}'")
        sent_items = sent_folder.Items.Restrict(f"[SentOn] >= '{start_date_str}'")
        all_items = list(inbox_items) + list(sent_items)

        # Bilingual keywords for analysis
        URGENT_KEYWORDS = [
            # English
            "urgent", "action required", "asap", "deadline", "critical",
            # Afrikaans
            "dringend", "aksie vereis", "sgm", "sperdatum", "krities", "belangrik", "spoedig", "gou", "NB"
        ]

        unread_urgent_count = 0
        flagged_threads_count = 0
        total_response_delay_seconds = 0
        replied_to_count = 0
        
        conversations = {}
        for item in all_items:
            try:
                conv_id = item.ConversationID
                if conv_id not in conversations:
                    conversations[conv_id] = []
                conversations[conv_id].append(item)
            except Exception as e:
                print(f"Warning: Could not get ConversationID for an item, skipping: {e}")
                continue

        # Sort threads to analyze chronologically
        def get_item_datetime(item):
            dt = getattr(item, 'ReceivedTime', None) or getattr(item, 'SentOn', None)
            return dt.replace(tzinfo=None) if dt else datetime.datetime.min

        for conv_id, thread in conversations.items():
            if not thread: continue
            thread.sort(key=get_item_datetime)

            last_email = thread[-1]
            last_sender_addr = getattr(last_email, 'SenderEmailAddress', '').lower()
            
            # Metric 1: Unread Urgent Count
            if getattr(last_email, 'UnRead', False) and last_sender_addr != my_email:
                subject_lower = getattr(last_email, 'Subject', '').lower()
                if any(kw in subject_lower for kw in URGENT_KEYWORDS):
                    unread_urgent_count += 1
            
            # Metric 2: Flagged Threads Count
            if any(getattr(item, 'FlagStatus', 0) == 2 for item in thread): # olFlagged
                flagged_threads_count += 1

            # Metric 3: Average Response Delay
            for i in range(len(thread) - 1):
                current_item = thread[i]
                next_item = thread[i+1]
                
                # Check for a reply pattern: their email -> my email
                if getattr(current_item, 'SenderEmailAddress', '').lower() != my_email and getattr(next_item, 'SenderEmailAddress', '').lower() == my_email:
                    time_in = get_item_datetime(current_item)
                    time_out = get_item_datetime(next_item)
                    if time_out > time_in:
                        delay = (time_out - time_in).total_seconds()
                        total_response_delay_seconds += delay
                        replied_to_count += 1

        # Finalize Metrics
        avg_response_delay_hours = (total_response_delay_seconds / replied_to_count / 3600) if replied_to_count > 0 else 0
        
        # Construct the JSON payload for the AI
        result = {
            "analysis_metadata": {
                "scan_period_days": days_to_scan,
                "timestamp": datetime.datetime.now().isoformat()
            },
            "inbox_metrics": {
                "unread_urgent_count": unread_urgent_count,
                "flagged_threads_count": flagged_threads_count,
                "average_response_delay_hours": round(avg_response_delay_hours, 2),
                "total_active_conversations": len(conversations)
            },
            "ai_instructions": "Analyze these metrics to assess the user's inbox load and provide a qualitative summary and recommendations."
        }
        
        return json.dumps(result, indent=2)

    except Exception as e:
        error_payload = {
            "status": "error",
            "message": f"An error occurred while calculating inbox metrics: {str(e)}"
        }
        return json.dumps(error_payload, indent=2)

@mcp.tool()
def list_folders() -> str:
    """
    List all available mail folders in Outlook
    
    Returns:
        A list of available mail folders
    """
    print("Tool: list_folders called")
    try:
        # Connect to Outlook
        _, namespace = connect_to_outlook()
        
        result = "Available mail folders:\n\n"
        
        # List all root folders and their subfolders
        for folder in namespace.Folders:
            result += f"- {folder.Name}\n"
            
            # List subfolders
            for subfolder in folder.Folders:
                result += f"  - {subfolder.Name}\n"
                
                # List subfolders (one more level)
                try:
                    for subsubfolder in subfolder.Folders:
                        result += f"    - {subsubfolder.Name}\n"
                except:
                    pass
        
        return result
    except Exception as e:
        return f"Error listing mail folders: {str(e)}"

@mcp.tool()
def list_recent_emails(days: int = 7, folder_name: Optional[str] = None) -> str:
    """
    List email titles from the specified number of days
    
    Args:
        days: Number of days to look back for emails (max 30)
        folder_name: Name of the folder to check (if not specified, checks the Inbox)
        
    Returns:
        Numbered list of email titles with sender information
    """
    print(f"Tool: list_recent_emails called with days={days}, folder_name={folder_name}")
    if not isinstance(days, int) or days < 1 or days > MAX_DAYS:
        return f"Error: 'days' must be an integer between 1 and {MAX_DAYS}"
    
    try:
        # Connect to Outlook
        _, namespace = connect_to_outlook()
        
        # Get the appropriate folder
        if folder_name:
            folder = get_folder_by_name(namespace, folder_name)
            if not folder:
                return f"Error: Folder '{folder_name}' not found"
        else:
            folder = namespace.GetDefaultFolder(6)  # Default inbox
        
        # Clear previous cache
        clear_email_cache()
        
        # Get emails from folder
        emails = get_emails_from_folder(folder, days)
        
        # Format the output and cache emails
        folder_display = f"'{folder_name}'" if folder_name else "Inbox"
        if not emails:
            return f"No emails found in {folder_display} from the last {days} days."
        
        result = f"Found {len(emails)} emails in {folder_display} from the last {days} days:\n\n"
        
        # Cache emails and build result
        for i, email in enumerate(emails, 1):
            # Store in cache
            email_cache[i] = email
            
            # Format for display
            result += f"Email #{i}\n"
            result += f"Subject: {email['subject']}\n"
            result += f"From: {email['sender']} <{email['sender_email']}>\n"
            result += f"Received: {email['received_time']}\n"
            result += f"Read Status: {'Read' if not email['unread'] else 'Unread'}\n"
            result += f"Has Attachments: {'Yes' if email['has_attachments'] else 'No'}\n\n"
        
        result += "To view the full content of an email, use the get_email_by_number tool with the email number."
        return result
    
    except Exception as e:
        return f"Error retrieving email titles: {str(e)}"

@mcp.tool()
def search_emails(search_term: str, days: int = 7, folder_name: Optional[str] = None) -> str:
    """
    Search emails by contact name or keyword within a time period
    
    Args:
        search_term: Name or keyword to search for
        days: Number of days to look back (max 30)
        folder_name: Name of the folder to search (if not specified, searches the Inbox)
        
    Returns:
        Numbered list of matching email titles
    """
    print(f"Tool: search_emails called with search_term='{search_term}', days={days}, folder_name={folder_name}")
    if not search_term:
        return "Error: Please provide a search term"
        
    if not isinstance(days, int) or days < 1 or days > MAX_DAYS:
        return f"Error: 'days' must be an integer between 1 and {MAX_DAYS}"
    
    try:
        # Connect to Outlook
        _, namespace = connect_to_outlook()
        
        # Get the appropriate folder
        if folder_name:
            folder = get_folder_by_name(namespace, folder_name)
            if not folder:
                return f"Error: Folder '{folder_name}' not found"
        else:
            folder = namespace.GetDefaultFolder(6)  # Default inbox
        
        # Clear previous cache
        clear_email_cache()
        
        # Get emails matching search term
        emails = get_emails_from_folder(folder, days, search_term)
        
        # Format the output and cache emails
        folder_display = f"'{folder_name}'" if folder_name else "Inbox"
        if not emails:
            return f"No emails matching '{search_term}' found in {folder_display} from the last {days} days."
        
        result = f"Found {len(emails)} emails matching '{search_term}' in {folder_display} from the last {days} days:\n\n"
        
        # Cache emails and build result
        for i, email in enumerate(emails, 1):
            # Store in cache
            email_cache[i] = email
            
            # Format for display
            result += f"Email #{i}\n"
            result += f"Subject: {email['subject']}\n"
            result += f"From: {email['sender']} <{email['sender_email']}>\n"
            result += f"Received: {email['received_time']}\n"
            result += f"Read Status: {'Read' if not email['unread'] else 'Unread'}\n"
            result += f"Has Attachments: {'Yes' if email['has_attachments'] else 'No'}\n\n"
        
        result += "To view the full content of an email, use the get_email_by_number tool with the email number."
        return result
    
    except Exception as e:
        return f"Error searching emails: {str(e)}"

@mcp.tool()
def count_unread_emails(folder_name: Optional[str] = None) -> str:
    """
    Counts the number of unread emails in a specified folder.
    
    Args:
        folder_name: Name of the folder to check (if not specified, checks the Inbox).
        
    Returns:
        A string stating the number of unread emails.
    """
    print(f"Tool: count_unread_emails called with folder_name={folder_name}")
    try:
        # Connect to Outlook
        _, namespace = connect_to_outlook()

        # Get the appropriate folder
        if folder_name:
            folder_display_name = f"'{folder_name}'"
            folder = get_folder_by_name(namespace, folder_name)
            if not folder:
                return f"Error: Folder {folder_display_name} not found"
        else:
            folder_display_name = "Inbox"
            folder = namespace.GetDefaultFolder(6)  # Default inbox

        # Filter for unread emails using the Restrict method for efficiency
        unread_filter = "[UnRead] = True"
        unread_emails = folder.Items.Restrict(unread_filter)
        count = unread_emails.Count

        return f"You have {count} unread emails in your {folder_display_name}."

    except Exception as e:
        return f"Error counting unread emails: {str(e)}"

@mcp.tool()
def get_email_by_number(email_number: int) -> str:
    """
    Get detailed content of a specific email by its number from the last listing or prioritization.
    
    Args:
        email_number: The number of the email from the list results (e.g., Email #1, Priority Email #1)
        
    Returns:
        Full details of the specified email
    """
    print(f"Tool: get_email_by_number called with email_number={email_number}")
    try:
        if not email_cache:
            return "Error: No emails have been listed yet. Please use list_recent_emails, search_emails, or prioritize_inbox first."
        
        if email_number not in email_cache:
            return f"Error: Email #{email_number} not found in the current listing. Please run a listing tool again."
        
        email_data = email_cache[email_number]
        
        # Connect to Outlook to get the full email content
        _, namespace = connect_to_outlook()
        
        # Retrieve the specific email
        email = namespace.GetItemFromID(email_data["id"])
        if not email:
            return f"Error: Email #{email_number} could not be retrieved from Outlook."
        
        # Format the output
        result = f"Email #{email_number} Details:\n\n"
        result += f"Subject: {email_data['subject']}\n"
        if email_data.get('is_sent_item'):
            result += f"To: {', '.join(email_data['recipients'])}\n"
            result += f"Sent: {email_data['sent_time']}\n"
        else:
            result += f"From: {email_data['sender']} <{email_data['sender_email']}>\n"
            result += f"Received: {email_data['received_time']}\n"
            result += f"Recipients: {', '.join(email_data['recipients'])}\n"

        result += f"Has Attachments: {'Yes' if email_data['has_attachments'] else 'No'}\n"
        
        if email_data['has_attachments']:
            result += "Attachments:\n"
            for i in range(1, email.Attachments.Count + 1):
                attachment = email.Attachments(i)
                result += f"  - {attachment.FileName}\n"
        
        result += "\nBody:\n"
        result += email_data['body']
        
        if not email_data.get('is_sent_item'):
            result += "\n\nTo reply to this email, use the reply_to_email_by_number tool with this email number."
        
        return result
    
    except Exception as e:
        return f"Error retrieving email details: {str(e)}"

@mcp.tool()
def reply_to_email_by_number(email_number: int, reply_text: str) -> str:
    """
    Reply to a specific email by its number from the last listing or prioritization.
    
    Args:
        email_number: The number of the email from the list results
        reply_text: The text content for the reply
        
    Returns:
        Status message indicating success or failure
    """
    print(f"Tool: reply_to_email_by_number called with email_number={email_number}, reply_text='{reply_text}'")
    try:
        if not email_cache:
            return "Error: No emails have been listed yet. Please use list_recent_emails, search_emails, or prioritize_inbox first."
        
        if email_number not in email_cache:
            return f"Error: Email #{email_number} not found in the current listing."
        
        email_data = email_cache[email_number]
        if email_data.get('is_sent_item'):
            return f"Error: Email #{email_number} is a sent item. You cannot reply to it."

        email_id = email_data["id"]
        
        # Connect to Outlook
        outlook, namespace = connect_to_outlook()
        
        # Retrieve the specific email
        email = namespace.GetItemFromID(email_id)
        if not email:
            return f"Error: Email #{email_number} could not be retrieved from Outlook."
        
        # Create reply
        reply = email.Reply()
        reply.Body = reply_text
        
        # Send the reply
        reply.Send()
        
        return f"Reply sent successfully to: {email.SenderName} <{email.SenderEmailAddress}>"
    
    except Exception as e:
        return f"Error replying to email: {str(e)}"

@mcp.tool()
def move_email_by_number(email_number: int, destination_folder_name: str) -> str:
    """
    Moves a specific email from its current location to another folder.

    This function is used to organize your inbox by filing emails into appropriate
    folders. You must first have a list of emails from 'list_recent_emails' or
    'search_emails'.

    Args:
        email_number: The number of the email from the list results.
        destination_folder_name: The exact name of the folder you want to move the email to.
        Use the `list_folders` tool to see valid folder names.

    Returns:
        A confirmation message indicating success or failure.
    """
    print(f"Tool: move_email_by_number called with email_number={email_number}, destination_folder_name='{destination_folder_name}'")
    # Step 1: Input Validation and Pre-Checks
    if not email_cache:
        return "Error: No emails have been listed yet. Please use list_recent_emails or search_emails first."
    if email_number not in email_cache:
        return f"Error: Email #{email_number} not found in the current listing."
    if not destination_folder_name:
        return "Error: You must provide a destination folder name."

    try:
        # Step 2: Retrieving Outlook Objects
        _, namespace = connect_to_outlook()
        
        email_data = email_cache[email_number]
        email_id = email_data["id"]
        
        email_to_move = namespace.GetItemFromID(email_id)
        destination_folder = get_folder_by_name(namespace, destination_folder_name)
        
        # Step 3: Post-Retrieval Validation
        if not email_to_move:
            return f"Error: Email #{email_number} could no longer be found in Outlook. It may have been moved or deleted."
        if not destination_folder:
            return f"Error: Destination folder '{destination_folder_name}' could not be found. Use the list_folders tool to see available folders."
            
        # Step 4: The Core Action - Performing the Move
        email_to_move.Move(destination_folder)
        
        # Step 5: Confirmation and Cleanup
        email_subject = email_data['subject']
        del email_cache[email_number]
        
        return f"Success: Email #{email_number}, '{email_subject}', has been moved to the '{destination_folder_name}' folder."

    # Step 6: Final Error Handling
    except Exception as e:
        return f"Error moving email: {str(e)}"

@mcp.tool()
def compose_email(recipient_email: str, subject: str, body: str, cc_email: Optional[str] = None) -> str:
    """
    Compose and send a new email
    
    Args:
        recipient_email: Email address of the recipient
        subject: Subject line of the email
        body: Main content of the email
        cc_email: Email address for CC (optional)
        
    Returns:
        Status message indicating success or failure
    """
    print(f"Tool: compose_email called with recipient_email='{recipient_email}', subject='{subject}', body length={len(body)}, cc_email={cc_email}")
    try:
        # Connect to Outlook
        outlook, _ = connect_to_outlook()
        
        # Create a new email
        mail = outlook.CreateItem(0)  # 0 is the value for a mail item
        mail.Subject = subject
        mail.To = recipient_email
        
        if cc_email:
            mail.CC = cc_email
        
        # Add signature to the body
        mail.Body = body
        
        # Send the email
        mail.Send()
        
        return f"Email sent successfully to: {recipient_email}"
    
    except Exception as e:
        return f"Error sending email: {str(e)}"

# Run the server
if __name__ == "__main__":
    print("Starting Outlook MCP Server...")
    print("Connecting to Outlook...")

    try:
        # Test Outlook connection
        outlook, namespace = connect_to_outlook()
        inbox = namespace.GetDefaultFolder(6)  # 6 is inbox
        print(f"Successfully connected to Outlook. Inbox has {inbox.Items.Count} items.")
        
        # Run the MCP server
        print("Starting MCP server. Press Ctrl+C to stop.")
        mcp.run()
    except Exception as e:
        print(f"Error starting server: {str(e)}")