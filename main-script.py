import tkinter as tk
from datetime import datetime, timedelta, date, timezone
import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

events_boardroom = ()

current_date = datetime.now()
shortened_days = []
for i in range(0, 7):
    next_day = current_date + timedelta(days=i)
    next_day_short = next_day.strftime("%a")
    shortened_days.append(next_day_short)

# Create the main window
root = tk.Tk()
root.title("Meeting Room Calendars")

# Get the screen resolution
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()

# Calculate the window size (full screen)
window_width = screen_width
window_height = screen_height

# Set the window size and position
root.geometry(f"{window_width}x{window_height}+0+0")

# Create a frame to hold the calendars
calendar_frame = tk.Frame(root)
calendar_frame.pack(fill="both", expand=True)

# Define the days of the week
days = ["", shortened_days[0], shortened_days[1], shortened_days[2], shortened_days[3], shortened_days[4], shortened_days[5], shortened_days[6]]

# Get the current day and calculate the next 7 days
today = datetime.now().date()
next_7_days = [today + timedelta(days=i) for i in range(7)]

# Function to create a calendar
def create_calendar(title, frame, row, column):
    # Create a dictionary to store the cell labels
    cell_labels = {}

    # Add the title
    title_label = tk.Label(frame, text=title, font=("Arial", 14, "bold"))
    title_label.grid(row=0, column=0, columnspan=7, pady=3)

    # Add the day labels
    day_labels = [tk.Label(frame, text=day, font=("Arial", 10, "bold"), width=10) for day in days]
    for i, label in enumerate(day_labels):
        label.grid(row=1, column=i)

    for row in range(8, 19):  # Range from 8 (8 AM) to 19 (7 PM)
        for sub_row in range(2):  # Two sub-rows for each hour
            hour = row if row < 13 else row - 12  # Convert 24-hour to 12-hour format
            am_pm = "AM" if row < 12 else "PM"
            if sub_row == 0:
                time_label = tk.Label(frame, text=f"{hour}:00 {am_pm}", font=("Arial", 10), width=12)
                time_label.grid(row=(row-8)*2+2, column=0, sticky="n")
            for col in range(1, 8):
                cell = tk.Label(frame, relief=tk.SUNKEN, width=10, height=1)
                cell.grid(row=(row-8)*2+sub_row+2, column=col, sticky="nsew")
                cell_labels[(row-8)*2+sub_row, col] = cell  # Store the cell label in the dictionary

    # Configure row and column weights
    for row in range(2, 26):
        frame.rowconfigure(row, weight=1)
    for col in range(1, 8):
        frame.columnconfigure(col, weight=1)

    return frame, cell_labels

# Add sample meetings to the small meeting room calendar
def add_small_room_meetings(cell_labels,events_small):
    clear_calendar(cell_labels)
    for event in events_small:
      start_str = event["start"].get("dateTime", event["start"].get("date"))
      end_str = event["end"].get("dateTime", event["end"].get("date"))
      start_datetime = datetime.fromisoformat(start_str)
      end_datetime = datetime.fromisoformat(end_str)
      date = start_datetime.date()
      difference = end_datetime - start_datetime
      slot_total = int(difference / timedelta(minutes=30))
      start_time = start_datetime.time()
      start_hour = start_time.hour
      start_minute = start_time.minute
      start_slot = (start_hour - 8) * 2 + (1 if start_minute >= 30 else 0)
      purpose = event["summary"]
      days_till=(date - today_date).days
      for num in range(slot_total):
        cell_labels[(start_slot+num,days_till+1)].config(text=purpose,bg="lightgreen")

# Add sample meetings to the boardroom calendar
def add_boardroom_meetings(cell_labels,events_boardroom):
    clear_calendar(cell_labels)
    for event in events_boardroom:
      start_str = event["start"].get("dateTime", event["start"].get("date"))
      end_str = event["end"].get("dateTime", event["end"].get("date"))
      start_datetime = datetime.fromisoformat(start_str)
      end_datetime = datetime.fromisoformat(end_str)
      date = start_datetime.date()
      difference = end_datetime - start_datetime
      slot_total = int(difference / timedelta(minutes=30))
      start_time = start_datetime.time()
      start_hour = start_time.hour
      start_minute = start_time.minute
      start_slot = (start_hour - 8) * 2 + (1 if start_minute >= 30 else 0)
      purpose = event["summary"]
      days_till=(date - today_date).days
      for num in range(slot_total):
        if start_slot+num < 0 or start_slot+num > 21:
           print("Out of range")
        else:
          cell_labels[(start_slot+num,days_till+1)].config(text=purpose,bg="lightgreen")

# Create the calendars
calendar_width = window_width // 2  # 50% of the screen width
calendar_height = window_height

small_room_frame = tk.LabelFrame(calendar_frame, text="Small meeting room", width=calendar_width, height=calendar_height)
small_room_frame.grid(row=0, column=0, sticky="nsew")
small_room_frame, small_room_cell_labels = create_calendar("Small meeting room", small_room_frame, 0, 0)

boardroom_frame = tk.LabelFrame(calendar_frame, text="Boardroom", width=calendar_width, height=calendar_height)
boardroom_frame.grid(row=0, column=1, sticky="nsew")
boardroom_frame, boardroom_cell_labels = create_calendar("Boardroom", boardroom_frame, 0, 1)


# If modifying these scopes, delete the file token.json.
SCOPES = ["https://www.googleapis.com/auth/calendar.readonly"]

today_date = date.today()

def main():
  """Shows basic usage of the Google Calendar API.
  Prints the start and name of the next 10 events on the user's calendar.
  """
  creds = None
  # The file token.json stores the user's access and refresh tokens, and is
  # created automatically when the authorization flow completes for the first
  # time.
  if os.path.exists("token.json"):
    creds = Credentials.from_authorized_user_file("token.json", SCOPES)
  # If there are no (valid) credentials available, let the user log in.
  if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
      creds.refresh(Request())
    else:
      flow = InstalledAppFlow.from_client_secrets_file(
          "credentials.json", SCOPES
      )
      creds = flow.run_local_server(port=0)
    # Save the credentials for the next run
    with open("token.json", "w") as token:
      token.write(creds.to_json())

  try:
    service = build("calendar", "v3", credentials=creds)

    # Call the Calendar API
    now = datetime.now(timezone.utc).isoformat()  # 'Z' indicates UTC time
    sevenDaysDate = datetime.now(timezone.utc) + timedelta(days=6)
    sevenDaysDate = sevenDaysDate.isoformat()
    events_result_small = (
        service.events()
        .list(
            calendarId="dd9e04b947e71ac0067ce3aabdfd7d0061a06bb8a6bab33c5b78432b13794955@group.calendar.google.com", #Small meeting room
            timeMin=now,
            timeMax=sevenDaysDate,
            singleEvents=True,
            orderBy="startTime",
        )
        .execute()
    )
    events_result_boardroom = (
        service.events()
        .list(
            calendarId="24ef60bc2555b748d4443501c952fccc520214d667049a32397d4a2ccf0b8a96@group.calendar.google.com", #Boardroom
            timeMin=now,
            timeMax=sevenDaysDate,
            singleEvents=True,
            orderBy="startTime",
        )
        .execute()
    )
    events_boardroom = events_result_boardroom.get("items", [])
    events_small = events_result_small.get("items", [])

    if not events_small:
      print("No upcoming small room events found.")
    else:
       add_small_room_meetings(small_room_cell_labels,events_small)
    
    if not events_boardroom:
      print("No upcoming boardroom events found.")
      return
    else:
       add_boardroom_meetings(boardroom_cell_labels,events_boardroom)
        
  except HttpError as error:
    print(f"An error occurred: {error}")


def clear_calendar(cell_labels):
    for cell in cell_labels.values():
        cell.config(text="", bg="white")

main()


def run_main():
    print("Rerunning")
    main()
    root.after(100000, run_main)


root.after(0, run_main)


# Configure row and column weights for the calendar frame
calendar_frame.rowconfigure(0, weight=1)
calendar_frame.columnconfigure(0, weight=1)
calendar_frame.columnconfigure(1, weight=1)


# Start the main event loop
root.mainloop()