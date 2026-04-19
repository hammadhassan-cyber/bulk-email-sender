import tkinter as tk               
from tkinter import ttk           
from tkinter import messagebox
from tkinter import filedialog  
import smtplib                     
import ssl, csv, json, os                                       
from datetime import datetime      
from email.mime.text import MIMEText 
from email.mime.multipart import MIMEMultipart

#   EMAIL TEMPLATES
TEMPLATES = {
    "Welcome": {
        "subject": "Welcome, {name}!",
        "body": "Hello {name},\n\nWelcome to {company}!\nWe are happy to have you.\n\nBest regards,\nThe Team"
    },
    "Newsletter": {
        "subject": "Monthly News from {company}",
        "body": "Dear {name},\n\nHere is this month's news from {company}.\n\nThank you for reading!\n\nBest,\n{company}"
    },
    "Invitation": {
        "subject": "You are invited, {name}!",
        "body": "Hi {name},\n\nYou are invited to our special event!\nWe hope to see you there.\n\nRegards,\n{company}"
    },
    "Custom": {
        "subject": "",
        "body": ""
    }
}
# ============================================================

def is_valid_email(email):
    return "@" in email and "." in email.split("@")[-1]

def save_to_history(recipient, subject, status):    
    # Load existing history (or start with empty list)
    history = []
    if os.path.exists("history.json"):
        with open("history.json", "r") as f:
            history = json.load(f)
    
    # Add new record
    record = {
        "to":        recipient,
        "subject":   subject,
        "status":    status,
        "time":      datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    }
    history.append(record)
    
    # Save back to file
    with open("history.json", "w") as f:
        json.dump(history, f, indent=2)

def send_one_email(sender_email, sender_password, recipient, subject, body): 
    # Build the email message
    msg = MIMEMultipart()
    msg["From"]    = sender_email
    msg["To"]      = recipient
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "plain"))  # plain, normal text (not HTML)
    
    try:
        # Connect to Gmail's server on port 587
        context = ssl.create_default_context()  
        
        with smtplib.SMTP("smtp.gmail.com", 587) as server:
            server.starttls(context=context)         # start encryption
            server.login(sender_email, sender_password)  # log in
            server.sendmail(sender_email, recipient, msg.as_string())
        
        return True, ""  
    
    except smtplib.SMTPAuthenticationError:
        return False, "Wrong email or password. Use Gmail App Password!"
    
    except smtplib.SMTPException as error:
        return False, f"Email error: {str(error)}"
    
    except Exception as error:
        return False, f"Something went wrong: {str(error)}"


# ============================================================
#  BUILDING THE MAIN WINDOW (GUI)

class EmailApp:
    def __init__(self, window):
        self.window = window
        self.window.title("Bulk Email Sender")
        self.window.geometry("650x550")
        self.window.configure(bg="#f0f4ff")
        
        self.contacts = []
        
        # Build the user interface
        self.build_ui()
    
    def build_ui(self):        
        title = tk.Label(self.window, text="Bulk Email Sender",font=("Arial", 18, "bold"),
                bg="#f0f4ff", fg="#2c3e8c")
        title.pack(pady=10)
        
        self.tabs = ttk.Notebook(self.window)
        self.tabs.pack(fill="both", expand=True, padx=15, pady=5)
        
        # Create each tab
        self.build_tab_send()      
        self.build_tab_contacts()  
        self.build_tab_template()  
        self.build_tab_history() 
    
    # ----------------------------------------------------------
    # SEND EMAIL
    
    def build_tab_send(self):
        
        tab = tk.Frame(self.tabs, bg="white", padx=15, pady=10)
        self.tabs.add(tab, text="  Send Email  ")
        
        def add_field(label_text, show_char=None):
            row = tk.Frame(tab, bg="white")
            row.pack(fill="x", pady=4)
            tk.Label(row, text=label_text, width=18, anchor="w",
                     bg="white", font=("Arial", 10)).pack(side="left")
            entry = tk.Entry(row, font=("Arial", 10), relief="solid", bd=1, show=show_char or "")
            entry.pack(side="left", fill="x", expand=True)
            return entry
        
        # --- Input fields ---
        tk.Label(tab, text="Your Gmail Credentials",
                 font=("Arial", 11, "bold"), bg="white", fg="#2c3e8c").pack(anchor="w", pady=(0, 3))
        
        self.entry_email = add_field("Your Gmail:")
        self.entry_password = add_field("App Password:", show_char="*")
        
        tk.Label(tab, text="Email Content",font=("Arial", 11, "bold"), bg="white",
                 fg="#2c3e8c").pack(anchor="w", pady=(0, 3))
        
        self.entry_subject = add_field("Subject:")
        
        # Recipient field
        recip_row = tk.Frame(tab, bg="white")
        recip_row.pack(fill="x", pady=4)
        tk.Label(recip_row, text="To (emails):", width=18, anchor="w",
                 bg="white", font=("Arial", 10)).pack(side="left")
        self.entry_to = tk.Entry(recip_row, font=("Arial", 10), relief="solid", bd=1)
        self.entry_to.pack(side="left", fill="x", expand=True)
        
        tk.Label(tab, text="Message:", bg="white", font=("Arial", 10)).pack(anchor="w", pady=(6, 2))
        
        self.text_body = tk.Text(tab, height=6, font=("Arial", 10), relief="solid", bd=1, wrap="word")
        self.text_body.pack(fill="x")
        
        # --- SEND button ---
        send_btn = tk.Button(tab,vtext="SEND EMAIL(S)", command=self.send_emails, bg="#2c3e8c", fg="white", font=("Arial", 12, "bold"),
            relief="flat", pady=8, cursor="hand2")
        send_btn.pack(fill="x")
        
        self.label_status = tk.Label(
            tab, text="", bg="white",
            font=("Arial", 9), fg="#2c3e8c"
        )
        self.label_status.pack(pady=5)
    
    # ----------------------------------------------------------
    #  TAB 2: LOAD CONTACTS FROM CSV
    # ----------------------------------------------------------
    
    def build_tab_contacts(self):        
        tab = tk.Frame(self.tabs, bg="white", padx=15, pady=15)
        self.tabs.add(tab, text="   Contacts  ")
        
        tk.Label(tab, text="Load Contacts from CSV File",
                 font=("Arial", 13, "bold"), bg="white", fg="#2c3e8c").pack(anchor="w")
        
        tk.Label(tab,
            text=(
                "\nYour CSV file needs these columns:\n\n"
                "   name,email,company\n"
                "   Ali Hassan,ali@gmail.com,HK Academy\n"
                "   Sara Khan,sara@gmail.com,Tech Corp\n"
            ),
            font=("Courier New", 10), bg="#f8f9ff", fg="#333", justify="left", relief="solid", bd=1,
              padx=10, pady=8).pack(fill="x", pady=10)
        
        # Load CSV button
        def load_csv():
            filepath = filedialog.askopenfilename(title="Choose your CSV file", filetypes=[("CSV files", "*.csv")])
            if not filepath:
                return   # user cancelled
            
            self.contacts = []
            
            with open(filepath, newline="", encoding="utf-8-sig") as f:
                reader = csv.DictReader(f)
                for row in reader:
                    contact = {k.lower().strip(): v.strip() for k, v in row.items()}
                    if contact.get("email"):   # only add if has email
                        self.contacts.append(contact)
            
            # Show loaded contacts in listbox
            listbox.delete(0, "end")
            for c in self.contacts:
                listbox.insert("end", f"  {c.get('name','')}  |  {c['email']}  |  {c.get('company','')}")
            
            count_label.config(text=f" Loaded {len(self.contacts)} contacts")
            
            emails = ", ".join(c["email"] for c in self.contacts)
            self.entry_to.delete(0, "end")
            self.entry_to.insert(0, emails)
        
        tk.Button(tab, text="Browse & Load CSV", command=load_csv, bg="#2c3e8c", fg="white", font=("Arial", 11, "bold"),
            relief="flat", pady=8, cursor="hand2").pack(fill="x")
        
        count_label = tk.Label(tab, text="No contacts loaded yet", bg="white", fg="gray", font=("Arial", 9))
        count_label.pack(pady=5)
        
        frame = tk.Frame(tab, bg="white")
        frame.pack(fill="both", expand=True, pady=5)
        
        scrollbar = tk.Scrollbar(frame)
        scrollbar.pack(side="right", fill="y")
        
        listbox = tk.Listbox(frame, font=("Arial", 9), yscrollcommand=scrollbar.set, relief="solid", bd=1)
        listbox.pack(fill="both", expand=True)
        scrollbar.config(command=listbox.yview)

    # ----------------------------------------------------------
    #  TAB 3: CHOOSE TEMPLATE
    # ----------------------------------------------------------
    
    def build_tab_template(self):        
        tab = tk.Frame(self.tabs, bg="white", padx=15, pady=15)
        self.tabs.add(tab, text="   Templates  ")
        
        tk.Label(tab, text="Choose an Email Template",
                 font=("Arial", 13, "bold"), bg="white", fg="#2c3e8c").pack(anchor="w", pady=(0, 10))
        
        row = tk.Frame(tab, bg="white")
        row.pack(fill="x", pady=5)
        
        tk.Label(row, text="Template:", bg="white", font=("Arial", 10)).pack(side="left")
        
        template_choice = tk.StringVar(value="Welcome")
        dropdown = ttk.Combobox(row, textvariable=template_choice, values=list(TEMPLATES.keys()), state="readonly", width=20)
        dropdown.pack(side="left", padx=10)
        
        # Preview areas
        tk.Label(tab, text="Subject Preview:", bg="white", font=("Arial", 10, "bold")).pack(anchor="w", pady=(10, 2))
        
        subject_preview = tk.Entry(tab, font=("Arial", 10), relief="solid", bd=1)
        subject_preview.pack(fill="x")
        
        tk.Label(tab, text="Body Preview:", bg="white", font=("Arial", 10, "bold")).pack(anchor="w", pady=(8, 2))
        
        body_preview = tk.Text(tab, height=7, font=("Arial", 10), relief="solid", bd=1, wrap="word")
        body_preview.pack(fill="x")
        
        def show_preview(event=None):
            name = template_choice.get()
            tmpl = TEMPLATES.get(name, {})
            
            subject_preview.delete(0, "end")
            subject_preview.insert(0, tmpl.get("subject", ""))
            
            body_preview.delete("1.0", "end")
            body_preview.insert("1.0", tmpl.get("body", ""))
        
        dropdown.bind("<<ComboboxSelected>>", show_preview)
        show_preview()  # show default on load
        
        def apply_template():
            name = template_choice.get()
            tmpl = TEMPLATES.get(name, {})
            
            self.entry_subject.delete(0, "end")
            self.entry_subject.insert(0, tmpl.get("subject", ""))
            
            self.text_body.delete("1.0", "end")
            self.text_body.insert("1.0", tmpl.get("body", ""))
            
            messagebox.showinfo("Done", f"Template '{name}' applied!\nGo to Send Email tab.")
        
        tk.Button(tab, text="Apply This Template", command=apply_template, bg="#27ae60", fg="white", font=("Arial", 11, "bold"),
            relief="flat", pady=8, cursor="hand2").pack(fill="x", pady=10)
    
    # ----------------------------------------------------------
    # VIEW HISTORY
    
    def build_tab_history(self):        
        tab = tk.Frame(self.tabs, bg="white", padx=15, pady=15)
        self.tabs.add(tab, text="   History  ")
        
        tk.Label(tab, text="Sent Email History",
                 font=("Arial", 13, "bold"), bg="white", fg="#2c3e8c").pack(anchor="w")
        
        # Table with columns
        columns = ("Time", "To", "Subject", "Status")
        
        frame = tk.Frame(tab)
        frame.pack(fill="both", expand=True, pady=10)
        
        scrollbar = ttk.Scrollbar(frame, orient="vertical")
        scrollbar.pack(side="right", fill="y")
        
        self.history_table = ttk.Treeview(frame, columns=columns, show="headings", yscrollcommand=scrollbar.set)
        scrollbar.config(command=self.history_table.yview)
        
        # Set column widths
        self.history_table.column("Time",    width=140)
        self.history_table.column("To",      width=180)
        self.history_table.column("Subject", width=170)
        self.history_table.column("Status",  width=70)
        
        for col in columns:
            self.history_table.heading(col, text=col)
        
        self.history_table.pack(fill="both", expand=True)
        
        # Color sent rows green, failed rows red
        self.history_table.tag_configure("sent",   foreground="green")
        self.history_table.tag_configure("failed", foreground="red")
        
        btn_row = tk.Frame(tab, bg="white")
        btn_row.pack(fill="x")
        
        tk.Button(btn_row, text="Refresh", command=self.load_history, bg="#2c3e8c", fg="white",
            font=("Arial", 9), relief="flat", padx=12, pady=5, cursor="hand2").pack(side="left", padx=(0, 5))
        
        tk.Button(btn_row, text="Clear All", command=self.clear_history, bg="#c0392b", fg="white", font=("Arial", 9), relief="flat",
            padx=12, pady=5, cursor="hand2").pack(side="left")
        
        # Load history on startup
        self.load_history()
    
    def send_emails(self):
        # Called when SEND button is clicked.        
        sender   = self.entry_email.get().strip()
        password = self.entry_password.get().strip()
        to_field = self.entry_to.get().strip()
        subject  = self.entry_subject.get().strip()
        body     = self.text_body.get("1.0", "end").strip()
        
        if not sender or not password or not to_field or not subject or not body:
            messagebox.showwarning("Missing Info", "Please fill in ALL fields before sending!")
            return
        
        recipients = [e.strip() for e in to_field.replace("\n", ",").split(",") if e.strip()]
        
        sent_count   = 0
        failed_count = 0
        
        for email in recipients:
            # Check email looks valid
            if not is_valid_email(email):
                self.label_status.config(
                    text=f" Skipped invalid email: {email}", fg="orange")
                save_to_history(email, subject, "Invalid")
                failed_count += 1
                continue
            
            # Personalize subject and body
            contact_data = {"name": email.split("@")[0], "company": ""}
            for c in self.contacts:
                if c.get("email", "").lower() == email.lower():
                    contact_data = c
                    break
            
            # Replace {name} and {company} with real values
            personal_subject = subject.replace("{name}", contact_data.get("name", ""))
            personal_subject = personal_subject.replace("{company}", contact_data.get("company", ""))
            
            personal_body = body.replace("{name}", contact_data.get("name", ""))
            personal_body = personal_body.replace("{company}", contact_data.get("company", ""))
            
            # Show status while sending
            self.label_status.config(text=f"Sending to {email}...", fg="#2c3e8c")
            self.window.update()   # refresh the window so user sees it
            
            # --- Actually send the email ---
            success, error_msg = send_one_email(
                sender, password, email,
                personal_subject, personal_body
            )
            
            if success:
                sent_count += 1
                save_to_history(email, personal_subject, "Sent ")
            else:
                failed_count += 1
                save_to_history(email, personal_subject, "Failed ")
                messagebox.showerror("Send Failed",
                    f"Could not send to {email}\n\nReason: {error_msg}")
        
        # --- Show final result ---
        result = f" Done!  Sent: {sent_count}  |  Failed: {failed_count}"
        self.label_status.config(text=result, fg="green")
        
        # Refresh history tab
        self.load_history()    
        messagebox.showinfo("Complete", result)
        
    def load_history(self):
        # Load and show history from history.json file.
        for row in self.history_table.get_children():
            self.history_table.delete(row)
        
        # Load from file
        if not os.path.exists("history.json"):
            return   
        
        with open("history.json", "r") as f:
            history = json.load(f)
    
        for record in reversed(history):
            status = record.get("status", "")
            tag = "sent" if "Sent" in status else "failed"
            
            self.history_table.insert("", "end",
                values=(record.get("time", ""), record.get("to", ""), record.get("subject", ""), status), tags=(tag,))
        
    def clear_history(self):
        # Delete all history records.
        answer = messagebox.askyesno("Confirm", "Delete all history?")
        if answer:
            with open("history.json", "w") as f:
                json.dump([], f)   # save empty list
            self.load_history()    # refresh the table

# ============================================================

if __name__ == "__main__":
    window = tk.Tk()
    app = EmailApp(window)
    window.mainloop()
