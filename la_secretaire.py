import os
import math
import json
import email
import base64
import imaplib
import requests
import webbrowser
import tkinter as tk
from turtle import st
from wake_listener import WakeWordAgent, get_porcupine_key
from dateparser.search import search_dates
from datetime import datetime, date
from tkcalendar import DateEntry
from tkinter import ttk, messagebox, scrolledtext
from email.header import decode_header
from dotenv import load_dotenv
from voice_utils import speak, listen
from dateutil.parser import parse as parse_datetime
from calendar_utils import get_calendar_service, list_upcoming_events, create_event


PART1 = "c2stb3ItdjEtZGEzMGFjNmNkYzk3YjllNjc3"
PART2 = "ZWNjZjZhZjI1NjljMzkwMDdiODUwOGNhN"
PART3 = "2M0ZGNlMGFhY2Q2NWIyZDBhYjZiNg=="

def get_api_key():
    full = PART1 + PART2 + PART3
    return base64.b64decode(full).decode()


CRED_FILE = "user_credentials.json"

def load_saved_credentials():
    if os.path.exists(CRED_FILE):
        with open(CRED_FILE, "r") as f:
            return json.load(f)
    return None

def save_credentials(email, password):
    with open(CRED_FILE, "w") as f:
        json.dump({"email": email, "password": password}, f)

# Load environment variables
load_dotenv()

# Connect to Gmail
def connect_to_gmail(username, app_password):
    imap = imaplib.IMAP4_SSL("imap.gmail.com")
    imap.login(username, app_password)
    return imap

# Fetch last N emails
def fetch_emails(imap, n=5):
    imap.select("inbox")
    status, messages = imap.search(None, "ALL")
    email_ids = messages[0].split()[-n:]

    emails = []
    for eid in email_ids[::-1]:
        _, msg_data = imap.fetch(eid, "(RFC822)")
        msg = email.message_from_bytes(msg_data[0][1])
        subject, _ = decode_header(msg["Subject"])[0]
        if isinstance(subject, bytes):
            subject = subject.decode(errors="ignore")
        from_ = msg.get("From")
        date_ = msg.get("Date")

        body = ""
        if msg.is_multipart():
            for part in msg.walk():
                if part.get_content_type() == "text/plain":
                    body = part.get_payload(decode=True).decode(errors="ignore")
                    break
        else:
            body = msg.get_payload(decode=True).decode(errors="ignore")

        emails.append({
            "subject": subject,
            "from": from_,
            "date": date_,
            "body": body[:3000]  # Limit size
        })

    return emails

def extract_email_index(command):
    command = command.lower()

    # Shortcut if user says 'selected'
    if "selected email" in command or "selected mail" in command:
        return "selected"

    # Simple mapping for common spoken indices
    word_to_index = {
        "first": 0, "1st": 0, "one": 0,
        "second": 1, "2nd": 1, "two": 1,
        "third": 2, "3rd": 2, "three": 2,
        "fourth": 3, "4th": 3, "four": 3,
        "fifth": 4, "5th": 4, "five": 4,
        "sixth": 5, "6th": 5, "six": 5,
        "seventh": 6, "7th": 6, "seven": 6,
        "eighth": 7, "8th": 7, "eight": 7,
        "ninth": 8, "9th": 8, "nine": 8,
        "tenth": 9, "10th": 9, "ten": 9
    }

    for word, index in word_to_index.items():
        if word in command:
            print(f"📌 Extracted index from word '{word}': {index}")
            return index

    print("📌 No index or 'selected' keyword found.")
    return None


def handle_voice_triggered_command(command):
    
    global app
    command = command.lower()
    print("🎙️ Voice command received:", command)

    email_idx = extract_email_index(command)
    print("🔢 Extracted email index:", email_idx)

    # Utility to run any function safely in Tkinter GUI thread
    def run_on_gui_thread(func):
        app.root.after(0, func)

    if ("summarize" in command or "summarise" in command):
        def summarize():
            if email_idx == "selected":
                app.summarize_selected()
            elif email_idx is not None:
                app.email_listbox.selection_clear(0, tk.END)
                app.email_listbox.selection_set(email_idx)
                app.email_listbox.activate(email_idx)
                app.summarize_selected()
            else:
                app.show_output("Please specify which email to summarize.")
        run_on_gui_thread(summarize)

    elif ("reply" in command or "respond" in command):
        def reply():
            if email_idx == "selected":
                app.reply_selected()
            elif email_idx is not None:
                app.email_listbox.selection_clear(0, tk.END)
                app.email_listbox.selection_set(email_idx)
                app.email_listbox.activate(email_idx)
                app.reply_selected()
            else:
                app.show_output("Please specify which email to reply to.")
        run_on_gui_thread(reply)

    # === CALENDAR COMMANDS ===
    elif "create event" in command or "add event" in command:
        run_on_gui_thread(lambda: app.handle_voice_calendar_command(command))

    elif "view events" in command or "show calendar" in command:
        run_on_gui_thread(app.show_calendar)

    # === OUTPUT CLEAR ===
    elif "clear output" in command:
        run_on_gui_thread(app.clear_output_area)

    # === DEFAULT RESPONSE ===
    else:
        print("❓ Unrecognized voice command:", command)
        run_on_gui_thread(lambda: app.show_output("❓ Sorry, I didn't understand that command."))
        speak("Sorry, I didn't catch that.")


# Summarize Email using OpenAI
def summarize_email(body):
    prompt = f"Summarize this email:\n\n{body}"
    try:
        response = requests.post(
            "https://openrouter.ai/api/v1/chat/completions",
            headers={
                "Authorization": f"Bearer {get_api_key()}",
                "HTTP-Referer": "https://youropenaiapp.com",
            },
            json={
                "model": "meta-llama/llama-3-70b-instruct",
                "messages": [{"role": "user", "content": prompt}],
                "max_tokens": 150
            },
            timeout=20  # Optional timeout
        )

        data = response.json()

        if "choices" in data:
            return data["choices"][0]["message"]["content"].strip()
        elif "error" in data:
            return f"❌ Error: {data['error']}"
        else:
            return "❌ Unexpected API response format."

    except Exception as e:
        return f"❌ Exception occurred: {str(e)}"


# Generate Reply using OpenAI
def generate_reply(body):
    prompt = f"Generate a polite and professional reply to the following email:\n\n{body}"
    try:
        response = requests.post(
            "https://openrouter.ai/api/v1/chat/completions",
            headers={
                "Authorization": f"Bearer {get_api_key()}",
                "HTTP-Referer": "https://youropenaiapp.com",
            },
            json={
                "model": "meta-llama/llama-3-70b-instruct",
                "messages": [{"role": "user", "content": prompt}],
                "max_tokens": 200
            },
            timeout=20
        )

        data = response.json()

        if "choices" in data:
            return data["choices"][0]["message"]["content"].strip()
        elif "error" in data:
            return f"❌ Error: {data['error']}"
        else:
            return "❌ Unexpected API response format."

    except Exception as e:
        return f"❌ Exception occurred: {str(e)}"


# GUI App
class EmailAgentApp:
    def __init__(self, root):
        self.root = root
        try:
            self.root.iconbitmap('logo.ico')  # Optional: Add logo.ico if available
        except:
            print("No icon found, continuing without it.")

        self.emails = []
        self.setup_ui()
        saved = load_saved_credentials()
        if saved:
            self.email_user = saved["email"]
            self.email_pass = saved["password"]
            self.login()
        else:
            self.prompt_login()
        

    def setup_ui(self):
        self.root.title("La Secrétaire – Personal Productivity Assistant")
        self.root.geometry("1100x750")
        self.root.configure(bg="#e8edf3")

        # ====== Styling ======
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("TButton", font=("Segoe UI", 10, "bold"), padding=6)
        style.configure("TLabel", font=("Segoe UI", 10), background="#e8edf3")
        style.configure("TLabelFrame", font=("Segoe UI", 12, "bold"))
        style.configure("TLabelframe.Label", background="#e8edf3")

        # ====== Title ======
        ttk.Label(self.root, text="La Secrétaire", font=("Segoe UI", 24, "bold"), background="#e8edf3").pack(pady=15)

        # ====== Email Section ======
        email_container = ttk.Frame(self.root)  # NEW wrapper frame
        email_container.pack(pady=10, fill="x")

        email_frame = ttk.LabelFrame(email_container, text="📧 Email Actions", padding=15)
        email_frame.pack(anchor="center")  # Center the inner frame

        self.email_listbox = tk.Listbox(email_frame, font=("Segoe UI", 10), height=7, selectbackground="#007acc")
        self.email_listbox.grid(row=0, column=0, columnspan=2, pady=(5, 10), padx=5, sticky="nsew")
        self.email_listbox.bind("<<ListboxSelect>>", self.display_selected_email)

        self.body_box = scrolledtext.ScrolledText(email_frame, font=("Segoe UI", 10), width=95, height=10, wrap=tk.WORD)
        self.body_box.grid(row=1, column=0, columnspan=2, padx=5, pady=(0, 10))

        email_btn_frame = ttk.Frame(email_frame)
        email_btn_frame.grid(row=2, column=0, columnspan=2, pady=5)

        ttk.Button(email_btn_frame, text="🧠 Summarize Email", command=self.summarize_selected).pack(side="left", padx=10)
        ttk.Button(email_btn_frame, text="✉️ Generate Reply", command=self.reply_selected).pack(side="left", padx=10)

        # ====== Calendar Section ======
        calendar_frame = ttk.LabelFrame(self.root, text="📅 Calendar Actions", padding=15)
        calendar_frame.pack(fill="x", padx=20, pady=10)

        cal_btns = ttk.Frame(calendar_frame)
        cal_btns.pack()
        ttk.Button(cal_btns, text="📂 View Events", command=self.show_calendar).pack(side="left", padx=10)
        ttk.Button(cal_btns, text="➕ Add Event", command=self.add_calendar_event).pack(side="left", padx=10)

        # ====== Voice Assistant Section ======
        voice_frame = ttk.LabelFrame(self.root, text="🎤 Voice Assistant", padding=15)
        voice_frame.pack(fill="x", padx=20, pady=10)
        ttk.Button(voice_frame, text="🎙️ Voice Mode", command=self.handle_voice_mode).pack(padx=10)

        # ====== Assistant Output Section ======
        output_frame = ttk.LabelFrame(self.root, text="🧠 Assistant Output", padding=15)
        output_frame.pack(fill="both", expand=True, padx=20, pady=10)

        self.output_box = scrolledtext.ScrolledText(output_frame, font=("Segoe UI", 10), height=8, wrap=tk.WORD)
        self.output_box.pack(fill="both", expand=True, pady=(0, 10))
        self.output_box.config(state="disabled")

        ttk.Button(output_frame, text="🧹 Clear Output", command=lambda: self.show_output("")).pack(pady=5)

        # Reply buttons area
        self.output_button_frame = ttk.Frame(output_frame)
        self.output_button_frame.pack(fill="x")

        self.left_button_area = ttk.Frame(self.output_button_frame)
        self.left_button_area.pack(side="left", padx=10)

        self.right_button_area = ttk.Frame(self.output_button_frame)
        self.right_button_area.pack(side="right", padx=10)

        # ====== Status Bar ======
        self.status_label = ttk.Label(self.root, text="Status: Ready", anchor="w",
                                    font=("Segoe UI", 9, "italic"), background="#e8edf3")
        self.status_label.pack(fill="x", padx=15, pady=(0, 10))

        # === Listening Orb ===
        self.orb_canvas = tk.Canvas(self.root, width=150, height=150, bg="#e8edf3", highlightthickness=0)
        self.orb_canvas.place(relx=0.5, rely=0.1, anchor="n")  # center top
        self.orb_canvas.place_forget()  # ✅ this hides it initially
        self.orb_visible = False

    def prompt_login(self):
        login_window = tk.Toplevel(self.root)
        login_window.title("Gmail Login")
        login_window.geometry("350x200")
        login_window.grab_set()

        tk.Label(login_window, text="Gmail Address:").pack(pady=(10, 0))
        email_entry = tk.Entry(login_window, width=40)
        email_entry.pack()

        tk.Label(login_window, text="App Password:").pack(pady=(10, 0))
        pass_entry = tk.Entry(login_window, show="*", width=40)
        pass_entry.pack()

        remember_var = tk.BooleanVar()
        tk.Checkbutton(login_window, text="Remember Me", variable=remember_var).pack(pady=5)

        def submit():
            self.email_user = email_entry.get().strip()
            self.email_pass = pass_entry.get().strip()
            if remember_var.get():
                save_credentials(self.email_user, self.email_pass)
            login_window.destroy()
            self.login()  # ✅ login only after user has entered credentials

        tk.Button(login_window, text="Login", command=submit).pack(pady=10)


    def login(self):
        try:
            self.imap = connect_to_gmail(self.email_user, self.email_pass)
            self.emails = fetch_emails(self.imap)
            for idx, mail in enumerate(self.emails):
                self.email_listbox.insert(tk.END, f"{idx+1}. {mail['subject']} — {mail['from']} ({mail['date']})")
            self.status_label.config(text="Status: Email login successful and emails loaded.")
        except Exception as e:
            messagebox.showerror("Login Failed", str(e))
            self.root.destroy()


    def display_selected_email(self, event):
        idx = self.email_listbox.curselection()
        if idx:
            mail = self.emails[idx[0]]
            self.body_box.delete("1.0", tk.END)
            self.body_box.insert(tk.END, f"Subject: {mail['subject']}\nFrom: {mail['from']}\nDate: {mail['date']}\n\n{mail['body']}")

    def summarize_selected(self):
        idx = self.email_listbox.curselection()
        if not idx:
            idx = (self.email_listbox.index(tk.ACTIVE),)
        if idx:
            mail = self.emails[idx[0]]
            summary = summarize_email(mail["body"])
            self.show_output(f"📄 Summary for: {mail['subject']}\n\n{summary}")
            print("📍 summarize_selected_email triggered")
            print("Selected index:", idx[0])


    def reply_selected(self):
        idx = self.email_listbox.curselection()
        if not idx:
            idx = (self.email_listbox.index(tk.ACTIVE),)

        if idx:
            mail = self.emails[idx[0]]
            reply = generate_reply(mail["body"])
            self.show_output(f"✉️ Suggested reply for: {mail['subject']}\n\n{reply}")
            print("📍 reply_selected triggered")
            print("Selected index:", idx[0])

            # Clear previous buttons
            for widget in self.left_button_area.winfo_children():
                widget.destroy()
            for widget in self.right_button_area.winfo_children():
                widget.destroy()

            # Add buttons to the left
            ttk.Button(self.left_button_area, text="📋 Copy to Clipboard",
                       command=lambda: self.copy_to_clipboard(reply)).pack(side="left", padx=5)

            ttk.Button(self.left_button_area, text="📨 Open in Gmail",
                       command=lambda: self.open_gmail_compose(mail, reply)).pack(side="left", padx=5)

            ttk.Button(self.right_button_area, text="🧹 Clear Output",
                       command=lambda: self.show_output("")).pack()


    def copy_to_clipboard(self, text):
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        self.status_label.config(text="✅ Reply copied to clipboard.")

    def open_gmail_compose(self, mail, reply):
        to_email = mail['from'].split()[-1].strip("<>")
        subject = f"Re: {mail['subject']}"
        body_encoded = reply.replace("\n", "%0A")

        mailto_url = f"https://mail.google.com/mail/?view=cm&fs=1&to={to_email}&su={subject}&body={body_encoded}"
        webbrowser.open(mailto_url)
        self.status_label.config(text="📨 Gmail compose window opened.")

    def handle_voice_mode(self):
        command = listen().lower()

        if "summarise" in command:
            self.summarize_selected()
            speak("Here is the summary.")
        elif "reply" in command:
            self.reply_selected()
            speak("Here is the suggested reply.")
        elif "create event" in command:
            self.handle_voice_calendar_command(command)
        else:
            speak("Sorry, I didn’t understand. Try saying summarize, reply, or create event.")


    def show_calendar(self):
        # Create popup window
        timeline_window = tk.Toplevel(self.root)
        timeline_window.title("📅 Upcoming Events Timeline")
        timeline_window.geometry("750x400")

        # ===== Treeview Styling =====
        style = ttk.Style()
        style.theme_use("clam")  # cleaner theme

        # Configure table appearance
        style.configure("Treeview.Heading", font=("Helvetica", 10, "bold"))
        style.configure("Treeview", rowheight=25, font=("Helvetica", 10), fieldbackground="#ffffff")
        style.map("Treeview", background=[("selected", "#ececec")])

        # Custom style name
        style.layout("Custom.Treeview", style.layout("Treeview"))
        style.configure("Custom.Treeview", background="#ffffff", foreground="#000000")

        # ===== Timeline Treeview =====
        tree = ttk.Treeview(
            timeline_window,
            columns=("Date", "Start", "End", "Title"),
            show="headings",
            style="Custom.Treeview"
        )
        tree.pack(fill="both", expand=True)

        # Set column headers
        tree.heading("Date", text="Date")
        tree.heading("Start", text="Start Time")
        tree.heading("End", text="End Time")
        tree.heading("Title", text="Title")

        tree.column("Date", width=100)
        tree.column("Start", width=100)
        tree.column("End", width=100)
        tree.column("Title", width=400)

        # Configure row colors
        tree.tag_configure('evenrow', background='#f8f8f8')
        tree.tag_configure('oddrow', background='#ffffff')
        tree.tag_configure('todayrow', background='#d7f9e9')  # green for today

        # ===== Populate Events =====
        try:
            service = get_calendar_service()
            now = datetime.utcnow().isoformat() + 'Z'

            results = service.events().list(calendarId='primary', timeMin=now,
                                            maxResults=20, singleEvents=True,
                                            orderBy='startTime').execute()
            events = results.get('items', [])
            today_str = date.today().strftime("%Y-%m-%d")

            if not events:
                tree.insert("", "end", values=("No events", "", "", ""))
            else:
                for idx, event in enumerate(events):
                    start = event["start"].get("dateTime", event["start"].get("date"))
                    end = event["end"].get("dateTime", event["end"].get("date"))

                    # Try parsing ISO format
                    try:
                        start_dt = parse_datetime(start)
                        end_dt = parse_datetime(end)
                        date_str = start_dt.strftime("%Y-%m-%d")
                        start_time = start_dt.strftime("%I:%M %p")
                        end_time = end_dt.strftime("%I:%M %p")
                    except:
                        date_str = start.split("T")[0]
                        start_time = start.split("T")[1][:5] if "T" in start else ""
                        end_time = end.split("T")[1][:5] if "T" in end else ""

                    title = event.get("summary", "No Title")
                    tag = 'todayrow' if date_str == today_str else ('evenrow' if idx % 2 == 0 else 'oddrow')
                    tree.insert("", "end", values=(date_str, start_time, end_time, title), tags=(tag,))
        except Exception as e:
            tree.insert("", "end", values=("Error fetching events:", str(e), "", ""))
            
    def add_calendar_event(self):
        event_window = tk.Toplevel(self.root)
        event_window.title("➕ Add Calendar Event")

        # === Event Title ===
        tk.Label(event_window, text="Event Title:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        title_entry = ttk.Entry(event_window, width=40)
        title_entry.grid(row=0, column=1, padx=5, pady=5)

        # === Date Picker ===
        tk.Label(event_window, text="Select Date:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        date_picker = DateEntry(event_window, width=18, background="darkblue", foreground="white", date_pattern="yyyy-mm-dd")
        date_picker.grid(row=1, column=1, padx=5, pady=5, sticky="w")

        # === Start Time Dropdown ===
        tk.Label(event_window, text="Start Time (HH:MM):").grid(row=2, column=0, padx=5, pady=5, sticky="w")
        start_hour = ttk.Combobox(event_window, values=[f"{i:02}" for i in range(0, 24)], width=5)
        start_minute = ttk.Combobox(event_window, values=[f"{i:02}" for i in range(0, 60, 5)], width=5)
        start_hour.grid(row=2, column=1, sticky="w", padx=(5,0))
        start_minute.grid(row=2, column=1, padx=(60,0))

        # === End Time Dropdown ===
        tk.Label(event_window, text="End Time (HH:MM):").grid(row=3, column=0, padx=5, pady=5, sticky="w")
        end_hour = ttk.Combobox(event_window, values=[f"{i:02}" for i in range(0, 24)], width=5)
        end_minute = ttk.Combobox(event_window, values=[f"{i:02}" for i in range(0, 60, 5)], width=5)
        end_hour.grid(row=3, column=1, sticky="w", padx=(5,0))
        end_minute.grid(row=3, column=1, padx=(60,0))
        
        # === Save Event Button ===
        def save_event():
            title = title_entry.get()
            date = date_picker.get()
            try:
                start = f"{date}T{start_hour.get()}:{start_minute.get()}:00"
                end = f"{date}T{end_hour.get()}:{end_minute.get()}:00"
                link = create_event(title, start, end)
                self.show_output(f"✅ Event created!\nTitle: {title}\nStart: {start}\nEnd: {end}\n\n📎 {link}")
                event_window.destroy()
            except Exception as e:
                messagebox.showerror("Error", f"Something went wrong.\n\n{e}")

        ttk.Button(event_window, text="➕ Add Event", command=save_event).grid(row=4, column=0, columnspan=2, pady=10)
        
    def handle_voice_calendar_command(self, command):
        import re
        import dateparser
        from dateparser.search import search_dates
        from datetime import datetime, timedelta
        from dateutil.relativedelta import relativedelta, MO, TU, WE, TH, FR, SA, SU

        command = command.lower()
        print("📢 Voice command received:", command)
        
        # 👇 Enhanced natural date extractor that handles "next Monday" etc.
        def get_datetime_from_command(command):
            base = datetime.now()
            weekday_map = {
                "monday": MO,
                "tuesday": TU,
                "wednesday": WE,
                "thursday": TH,
                "friday": FR,
                "saturday": SA,
                "sunday": SU,
            }

            match = re.search(r"(next|this|coming)?\s*(monday|tuesday|wednesday|thursday|friday|saturday|sunday)", command)
            if match:
                prefix = match.group(1)
                day = match.group(2)
                delta = 0 if prefix in ["this", None] else 1
                weekday = weekday_map[day]
                return base + relativedelta(weekday=weekday(delta))

            # Fallback to normal date parsing
            results = search_dates(command, settings={"RELATIVE_BASE": base})
            if results:
                return results[0][1]

            return None

        if "create event" in command:
            # 🔍 Extract date and time
            event_date = get_datetime_from_command(command)
            if not event_date:
                self.show_output("❌ Sorry, I couldn't understand the date.")
                speak("Sorry, I couldn't understand the date.")
                return

            print("🛠️ Final Parsed Datetime:", event_date.strftime("%A, %d %B %Y"))

            # 🕗 Time parsing
            time_match = re.search(r'at\s*(\d{1,2})(?::(\d{2}))?\s*(am|pm)?', command)
            hour = 9  # Default time
            minute = 0
            meridian = None

            if time_match:
                hour = int(time_match.group(1))
                minute = int(time_match.group(2) or 0)
                meridian = time_match.group(3)

                current_hour = datetime.now().hour
                if meridian == "pm" and hour < 12:
                    hour += 12
                elif meridian == "am" and hour == 12:
                    hour = 0
                elif not meridian:
                    if current_hour >= 12 and hour < 12:
                        hour += 12
            else:
                speak("No time mentioned, so I will set it for 9 AM by default.")

            # 📝 Title extraction
            title_match = re.search(r'(called|about|named|titled)\s+(.+)', command)
            summary = title_match.group(2).strip() if title_match else "Untitled Event"

            # ⏱ Final event times
            start_dt = event_date.replace(hour=hour, minute=minute, second=0)
            end_dt = start_dt + timedelta(hours=1)
            readable_time = start_dt.strftime("%A, %B %d at %I:%M %p")

            # ✅ Confirmation
            confirm_text = f"Should I create an event titled '{summary}' on {readable_time}? To confirm say, 'go ahead'."
            self.show_output(confirm_text)
            speak(confirm_text)

            user_reply = listen().lower()
            print("🗣️ User replied:", user_reply)

            if any(x in user_reply for x in ["yes", "yeah", "sure", "go ahead", "do it", "yup"]):
                link = create_event(summary, start_dt.isoformat(), end_dt.isoformat())
                result = f"✅ Event Created!\nTitle: {summary}\nTime: {readable_time}\n📎 {link}"
                self.show_output(result)
                speak("Event created successfully.")
            else:
                self.show_output("❌ Event creation canceled.")
                speak("Okay, I’ve cancelled the event.")     
                
    def show_listening_orb(self):
        self.orb_visible = True
        self.orb_canvas.place(relx=0.5, rely=0.1, anchor="n")
        self.orb_pulse_frame = 0  # NEW: for smooth animation
        self.animate_siri_orb()

    def hide_listening_orb(self):
        self.orb_visible = False
        self.orb_canvas.delete("all")
        self.orb_canvas.place_forget()


    def animate_siri_orb(self):
        if not self.orb_visible:
            return

        self.orb_canvas.delete("all")
        center_x, center_y = 75, 75
        frame = self.orb_pulse_frame
        self.orb_pulse_frame += 1

        # ==== COLORS (Siri gradient palette) ====
        siri_colors = [
            "#00cfff",  # cyan
            "#886fff",  # indigo
            "#ff5bd7",  # magenta
            "#00ffcc",  # mint
        ]

        # ==== PULSE WAVES ====
        for i in range(len(siri_colors)):
            radius = 20 + 10 * math.sin((frame + i * 15) * 0.15)
            self.orb_canvas.create_oval(
                center_x - radius,
                center_y - radius,
                center_x + radius,
                center_y + radius,
                outline=siri_colors[i],
                width=3
            )

        # ==== CENTER GLOW ====
        self.orb_canvas.create_oval(
            center_x - 12, center_y - 12,
            center_x + 12, center_y + 12,
            fill="#ffffff", outline=""
        )

        self.root.after(50, self.animate_siri_orb)

                   
    def show_output(self, text):
        self.output_box.config(state="normal")
        self.output_box.delete("1.0", tk.END)
        self.output_box.insert(tk.END, text)
        self.output_box.config(state="disabled")
        

# Run the App
if __name__ == "__main__":
    root = tk.Tk()
    app = EmailAgentApp(root)
    
    wake_word_path = os.path.join("wake_words", "Hey-Secretary_en_windows_v3_0_0.ppn")
    
    # Start background wake-word listener
    wake_agent = WakeWordAgent(
    callback=handle_voice_triggered_command,
    app=app,
    access_key=get_porcupine_key()  # ✅ pass the access key
    )
    wake_agent.start()
    
    root.mainloop()
