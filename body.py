import os 
import subprocess
import webbrowser
import platform
import time
from datetime import datetime, timedelta, date
import pytz
import re
import cohere
import psutil
from collections import defaultdict
from pathlib import Path
import json

from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request

# --- Scopes Required for Calendar and Tasks ---
SCOPES = [
    'https://www.googleapis.com/auth/calendar',
    'https://www.googleapis.com/auth/tasks'
]

# --- Set up Cohere cCient ---
co = cohere.Client("x4EoaJk7qGWmA24hTCQyOsRLDmySQdM7W5XBqD9F")  # Replace with your real key
PROCESSED_EVENTS = set()
# --- Memory for Project Profiles---
APP_PROFILE_FILE = Path("project_profiles.json")
def load_profiles():
    if APP_PROFILE_FILE.exists():
        with open(APP_PROFILE_FILE, "r") as f:
            return json.load(f)
    return {}

def save_profiles(profiles):
    with open(APP_PROFILE_FILE, "w") as f:
        json.dump(profiles, f, indent=2)

PROJECT_PROFILES = load_profiles()

# --- Memory for Project Files  ---
MEMORY_FILE = Path("project_memory.json")

def load_memory():
    if MEMORY_FILE.exists():
        with open(MEMORY_FILE, "r") as f:
            return json.load(f)
    return {}

def save_memory(memory):
    with open(MEMORY_FILE, "w") as f:
        json.dump(memory, f, indent=2)

PROJECT_MEMORY = load_memory()

def remember_resources(project, urls=None, pdfs=None):
    if project not in PROJECT_MEMORY:
        PROJECT_MEMORY[project] = {"urls": [], "pdfs": []}
    if urls:
        for url in urls:
            if url not in PROJECT_MEMORY[project]["urls"]:
                PROJECT_MEMORY[project]["urls"].append(url)
    if pdfs:
        for pdf in pdfs:
            if pdf not in PROJECT_MEMORY[project]["pdfs"]:
                PROJECT_MEMORY[project]["pdfs"].append(pdf)
    save_memory(PROJECT_MEMORY)

def reopen_project_resources(project):
    resources = PROJECT_MEMORY.get(project, {})
    urls = resources.get("urls", [])
    pdfs = resources.get("pdfs", [])

    for url in urls:
        webbrowser.open(url)
    for pdf_path in pdfs:
        if os.path.exists(pdf_path):
            subprocess.Popen([pdf_path], shell=True)

    if urls or pdfs:
        print(f"Restored resources for project '{project}'")

# --- Known App Executable Names ---
APP_TARGETS = {
    "photoshop": "Photoshop.exe",   
    "visual studio code": "Code.exe",
    "word": "WINWORD.EXE",
    "powerpoint": "POWERPNT.EXE",
    "excel": "EXCEL.EXE",
    "chrome": "chrome.exe",
    "firefox": "firefox.exe",
    "notepad": "notepad.exe",
    "calculator": "calc.exe",
    "zoom": "Zoom.exe",
    "slack": "slack.exe",
    "obs studio": "obs64.exe",
    "fusion 360": "Fusion360.exe",
    "onenote": "ONENOTE.EXE",
    "discord": "Discord.exe",
    "matlab": "matlab.exe"  
}

# --- Known Install URLs ---
INSTALL_URLS = {
    "microsoft teams": "https://www.microsoft.com/en/microsoft-teams/download-app",
    "photoshop": "https://www.adobe.com/products/photoshop/free-trial-download.html",
    "visual studio code": "https://code.visualstudio.com/",
    "word": "https://www.microsoft.com/en-us/microsoft-365/word",
    "powerpoint": "https://www.microsoft.com/en-us/microsoft-365/powerpoint",
    "excel": "https://www.microsoft.com/en-us/microsoft-365/excel",
    "zoom": "https://zoom.us/download",
    "slack": "https://slack.com/downloads/windows",
    "fusion 360": "https://www.autodesk.com/products/fusion-360/download",
    "onenote": "https://www.onenote.com/download",
    "discord": "https://discord.com/download",
    "matlab": "https://www.mathworks.com/downloads/",
    "obs studio": "https://obsproject.com/download"
}
# --- Known install URLs for fallback ---
PROJECT_KEYWORDS = {
    "ai research": ["ml", "neural", "cohere", "inference", "ai research"],
    "3d modeling": ["cad", "3d model", "render", "bracket"],
    "game dev": ["unity", "game", "devlog", "sprites"],
    "web design": ["html", "css", "figma", "design", "landing page"],
    "course notes": ["lecture", "notes", "course", "onenote"]
}
RECENT_PROJECT_TASKS = defaultdict(list)

# --- Find Installed Apps ---
def find_apps(app_targets):
    search_dirs = [
        os.environ.get("ProgramFiles", "C:\\Program Files"),
        os.environ.get("ProgramFiles(x86)", "C:\\Program Files (x86)"),
        os.environ.get("LOCALAPPDATA", ""),
        os.environ.get("APPDATA", ""),
    ]

    found_apps = {}

    for root_dir in search_dirs:
        for root, dirs, files in os.walk(root_dir):
            # Skip any Teams Installer directories
            if "teams installer" in root.lower():
                continue

            for file in files:
                file_lower = file.lower()
                for app_name, exe_name in app_targets.items():
                    if exe_name.lower() == file_lower and app_name not in found_apps:
                        full_path = os.path.join(root, file)
                        found_apps[app_name] = full_path
    return found_apps

APP_PATHS = find_apps(APP_TARGETS)

# --- Speech to Text Input ---
def get_user_input(prompt_text, fallback_text=""):
    try:
        from speech_recognition import Recognizer, Microphone
        import speech_recognition as sr

        recognizer = Recognizer()
        with Microphone() as source:
            print(prompt_text + " (say something)")
            audio = recognizer.listen(source)
            response = recognizer.recognize_google(audio)
            print(f"You said: {response}")
            return response.lower().strip()
    except Exception as e:
        print(f"Speech failed ({e}), falling back to text input.")
        return input(prompt_text + " ").lower().strip()

# --- Auth Helper ---
def get_credentials():
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('creds.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    return creds

# --- Inference Function ---
def infer_app_to_launch(title):
    prompt = f"""
You are an intelligent system assistant that suggests which Windows application to launch based on the title of a calendar event.

Match the task to the best-known desktop app a user might expect to use.
If it's a meeting but the platform isn't specified, return 'unspecified meeting'.

Examples:
- "Fix login bug" → Visual Studio Code
- "Weekly sync (Google Meet)" → Google Meet
- "Edit marketing video" → Adobe Premiere Pro
- "Design flyer" → Photoshop
- "Write summary report" → Word
- "Join client call (Teams)" → Microsoft Teams
- "Host webinar" → Zoom
- "Build project presentation" → PowerPoint
- "Organize data sheet" → Excel
- "Review pull request" → Visual Studio Code
- "Stream live session" → OBS Studio
- "Team chat" → Slack
- "Weekly meeting" → unspecified meeting

- "Design mechanical bracket" → Fusion 360
- "Edit 3D model for CNC" → Fusion 360
- "Assemble gear components (Fusion)" → Fusion 360
- "CAD modeling session" → Fusion 360

- "Meeting notes (OneNote)" → OneNote
- "Plan lecture notes" → OneNote
- "Write research ideas" → OneNote
- "Class notebook session" → OneNote

- "Team voice chat" → Discord
- "Late-night code collab" → Discord
- "Project sync call (Discord)" → Discord
- "Gaming chat with teammates" → Discord

- "Daily standup" → Slack
- "Team update on Slack" → Slack
- "Reply to Slack threads" → Slack
- "Catch up on #dev channel" → Slack

- "Simulate control system" → MATLAB
- "Run script for data analysis" → MATLAB
- "MATLAB debugging session" → MATLAB
- "Calculate eigenvalues (math)" → MATLAB


Now, based on the event title: "{title}", respond with the name of the app only.
"""
    try:
        response = co.generate(
            model="command-r-plus",
            prompt=prompt,
            max_tokens=10,
            temperature=0
        )
        result = response.generations[0].text.strip().lower()
        result = re.sub(r'[^a-zA-Z0-9 ]', '', result)
        if not result:
            raise ValueError("Blank response")
        print(f"Cohere inferred: {result}")
        return result
    except Exception as e:
        print(f"Cohere failed ({e}), using fallback")
        return "visual studio code"
    
# --- Task and Project Recognition ---
def classify_project(text):
    text = text.lower()
    for project, keywords in PROJECT_KEYWORDS.items():
        if any(keyword in text for keyword in keywords):
            return project
    return "misc"

def store_task_by_project(task_title, notes):
    project = classify_project(task_title + " " + notes)
    RECENT_PROJECT_TASKS[project].append(task_title)

# --- Summarize Previous Work ---
def summarize_last_project_session(project):
    log_path = "logged_tasks.json"
    if not os.path.exists(log_path):
        print(f"No log file found.")
        return

    with open(log_path, "r") as f:
        data = json.load(f)

    sessions = data.get(project, [])
    if not sessions:
        print(f"No previous session found for project '{project}'.")
        return

    last_session = sessions.pop()  # Get and remove last session
    print(f"\nLast session summary for project '{project}' at {last_session['timestamp']}:")
    for task in last_session["tasks"]:
        print(f"  - {task}")

    # Save the updated log
    if sessions:
        data[project] = sessions
    else:
        del data[project]  # remove key if no sessions left

    with open(log_path, "w") as f:
        json.dump(data, f, indent=2)

# --- Add Apps to Project ---
def prompt_to_add_apps(project_name):
    print(f"\nAdd apps for project '{project_name}'")
    print("Type app names one by one (e.g., 'matlab', 'notepad', etc.).")
    print("Type 'done' when finished.\n")

    if project_name not in PROJECT_PROFILES:
        PROJECT_PROFILES[project_name] = []

    while True:
        app = get_user_input("App to add: ")
        if app == "done":
            break
        if app not in PROJECT_PROFILES[project_name]:
            PROJECT_PROFILES[project_name].append(app)
            print(f"Added '{app}' to {project_name}.")
        else:
            print(f"'{app}' is already in the list.")
    save_profiles(PROJECT_PROFILES)
# --- Open Project and Apps ---
def launch_project_environment(project_name):
    summarize_last_project_session(project_name)
    apps = PROJECT_PROFILES.get(project_name, [])
    if not apps:
        print(f"No apps defined for project '{project_name}'")
        prompt_to_add_apps(project_name)
        apps = PROJECT_PROFILES.get(project_name, [])
    print(f"Launching environment for project: {project_name}")
    for app in apps:
        if app not in APP_PATHS:
            print(f"App path for '{app}' not found. Skipping.")
            continue
        launch_app_by_name(app)
    
    reopen_project_resources(project_name)

def is_teams_running():
    for proc in psutil.process_iter(['name']):
        if proc.info['name'] and 'teams' in proc.info['name'].lower():
            return True
    return False

def launch_teams():
    update_exe = os.path.join(os.environ.get("LOCALAPPDATA", ""), "Microsoft", "Teams", "Update.exe")

    if os.path.exists(update_exe):
        print("Attempting to launch Microsoft Teams via Update.exe...")
        subprocess.Popen([update_exe, "--processStart", "Teams.exe"], shell=True)
        time.sleep(5)  # wait a bit for Teams to start

        if is_teams_running():
            print("Teams launched via Update.exe.")
            return
        else:
            print("Update.exe did not launch Teams. Falling back to Microsoft Store version.")

    # Try UWP version
    print("Attempting to launch Microsoft Teams via Microsoft Store...")
    try:
        subprocess.Popen('explorer shell:AppsFolder\\MSTeams_8wekyb3d8bbwe!MSTeams', shell=True)
        print("Microsoft Teams (UWP) launch attempted.")
    except Exception as e:
        print(f"Failed to launch Teams (UWP): {e}")

# --- App Launcher ---
def launch_app_by_name(app_name):
    if platform.system() != "Windows":
        print("App launching only supported on Windows.")
        return

    app_name = app_name.lower()

    # Handle unspecified meeting platform
    if app_name == "unspecified meeting":
        choice = get_user_input("A meeting was detected but no platform was specified. Choose platform (teams/meet/zoom): ")
        if choice in ["teams", "microsoft teams"]:
            app_name = "microsoft teams"
        elif choice in ["meet", "google meet"]:
            app_name = "google meet"
        elif choice == "zoom":
            app_name = "zoom"
        else:
            print("Unknown platform. Skipping launch.")
            return

    # Handle browser-based launch
    if app_name == "google meet":
        webbrowser.open("https://meet.google.com")
        return

    # Special handling for Microsoft Teams (desktop + fallback to UWP)
    if app_name == "microsoft teams":
        launch_teams()
        return

    # Try launching from discovered executable paths
    if app_name in APP_PATHS:
        print(f"Launching {app_name} via scanned path...")
        try:
            subprocess.Popen([APP_PATHS[app_name]], shell=True)
            return
        except Exception as e:
            print(f"Failed to launch {app_name} from path: {e}")

    # Prompt to install if not found
    if app_name in INSTALL_URLS:
        answer = get_user_input(f"{app_name} is not installed. Would you like to install it now? (yes/no): ")
        if answer in ["yes", "y"]:
            webbrowser.open(INSTALL_URLS[app_name])
        else:
            print("Skipping install.")
        return

    try:
        print(f"Fallback: Trying 'start {app_name}'")
        subprocess.Popen(f'start {app_name}', shell=True)
    except Exception as e:
        print(f"Could not launch app '{app_name}': {e}")

print("Scanned app paths:")
for app, path in APP_PATHS.items():
    print(f"  {app}: {path}")


# --- Scheduler ---
def schedule_app_launch(app_name, start_time):
    if platform.system() != "Windows":
        print("App launching only supported on Windows.")
        return

    launch_time = start_time - timedelta(minutes=5)
    now = datetime.now(start_time.tzinfo)
    delay = (launch_time - now).total_seconds()

    if delay <= 0:
        print(f"Event is too close or already passed. Launching {app_name} now.")
        launch_app_by_name(app_name)
        return
    print(f"Waiting {int(delay // 60)} minutes and {int(delay % 60)} seconds to launch {app_name}...")
    time.sleep(delay)
    launch_app_by_name(app_name)

# --- Event + Task Creation + App Launcher ---
def create_event_and_launch(title, description="", start_time=None, duration_minutes=60):
    tz = pytz.timezone("Europe/Rome")
    if start_time is None:
        start_time = datetime.now(tz) + timedelta(hours=1)
    end_time = start_time + timedelta(minutes=duration_minutes)
    due_time = (start_time + timedelta(minutes=30)).isoformat()

    creds = get_credentials()
    calendar_service = build('calendar', 'v3', credentials=creds)
    tasks_service = build('tasks', 'v1', credentials=creds)

    # Create calendar event
    event = {
        'summary': title,
        'description': description,
        'start': {'dateTime': start_time.isoformat(), 'timeZone': 'Europe/Rome'},
        'end': {'dateTime': end_time.isoformat(), 'timeZone': 'Europe/Rome'},
    }
    created_event = calendar_service.events().insert(calendarId='primary', body=event).execute()
    print("Event created:", created_event.get('htmlLink'))

    # Create corresponding task
    task = {
        'title': title,
        'notes': f'{description}\n\nLinked to event: {created_event.get("htmlLink")}',
        'due': due_time
    }
    created_task = tasks_service.tasks().insert(tasklist='@default', body=task).execute()
    print("Task created:", created_task['title'])

    # If event is for today, schedule the app launch
    if start_time.date() == date.today():
        project = classify_project(title + " " + description)
        if project != "misc":
            launch_project_environment(project)
        else:
            app = infer_app_to_launch(title)
            schedule_app_launch(app, start_time)

# --- Fetch Today's Events ---
def handle_todays_events():
    creds = get_credentials()
    calendar_service = build('calendar', 'v3', credentials=creds)

    tz = pytz.timezone("Europe/Rome")
    today = date.today()
    start_of_day = tz.localize(datetime.combine(today, datetime.min.time())).isoformat()
    end_of_day = tz.localize(datetime.combine(today, datetime.max.time())).isoformat()

    events_result = calendar_service.events().list(
        calendarId='primary',
        timeMin=start_of_day,
        timeMax=end_of_day,
        singleEvents=True,
        orderBy='startTime'
    ).execute()

    events = events_result.get('items', [])
    print(f"Found {len(events)} events for today.")

    for event in events:
        event_title = event.get('summary', '')
        event_description = event.get('description', '')
        start_str = event['start'].get('dateTime')
        event_id = event.get('id')

        if not event_id or not start_str:
            continue

        # Use ID or ID+start time as a unique identifier
        key = f"{event_id}_{start_str}"
        if key in PROCESSED_EVENTS:
            continue  # Skip if already handled

        start_time = datetime.fromisoformat(start_str.replace('Z', '+00:00'))
        if start_time > datetime.now(start_time.tzinfo):
            project = classify_project(event_title + " " + event_description)
            if project != "misc":
                print(f"Project detected: {project}. Launching environment...")
                launch_project_environment(project)
            else:
                app = infer_app_to_launch(event_title)
                print(f"No project match. Inferred app: {app}")
                schedule_app_launch(app, start_time)
            PROCESSED_EVENTS.add(key)

# --- Check Completed Tasks ---

LOG_FILE = "logged_tasks.json"

def log_completed_tasks_for_today():
    creds = get_credentials()
    service = build('tasks', 'v1', credentials=creds)
    tasks_result = service.tasks().list(tasklist='@default').execute()
    tasks = tasks_result.get('items', [])

    session_time = datetime.now().strftime("%Y-%m-%d %H:%M")
    today = date.today()
    log_path = "logged_tasks.json"

    if os.path.exists(log_path):
        with open(log_path, "r") as f:
            full_log = json.load(f)
    else:
        full_log = {}

    session_tasks = defaultdict(list)

    for task in tasks:
        if task.get('status') != 'completed':
            continue

        completed_str = task.get('completed')
        if not completed_str:
            continue

        completed_dt = datetime.fromisoformat(completed_str.replace('Z', '+00:00'))
        if completed_dt.date() != today:
            continue

        title = task.get('title', 'Untitled Task')
        project = classify_project(title)
        session_tasks[project].append(title)

    for project, titles in session_tasks.items():
        if project not in full_log:
            full_log[project] = []

        full_log[project].append({
            "timestamp": session_time,
            "tasks": titles
        })
        print(f"Logged {len(titles)} tasks under project '{project}'")

    with open(log_path, "w") as f:
        json.dump(full_log, f, indent=2)

    # Save new log
    with open(LOG_FILE, "w") as f:
        json.dump(new_log, f, indent=2)



# --- Example usage ---
if __name__ == '__main__':
    while True:
        handle_todays_events()
        time.sleep(300)
        
    create_event_and_launch(
        title="Train transformer model for benchmark dataset",
        description="Use Visual Studio Code and MATLAB to test performance on the synthetic classification dataset. Review related paper in PDF and open Cohere API docs."
    )

    # Example: creating a new event for today should still trigger its app launch
    log_completed_tasks_for_today() 
