import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import pandas as pd
import webbrowser
import urllib.parse
import time
import threading
import re
import os
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart


class MessageSenderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Message Sender (SMS, WhatsApp & Email)")
        self.root.geometry("750x750")
        self.root.resizable(False, False)
        self.df = None
        self.email_template = None

        # --- GUI Layout ---
        self.title_label = tk.Label(root, text="Excel Message Sender Pro",
                                    font=("Arial", 16, "bold"))
        self.title_label.pack(pady=5)

        # 1. File Selection
        self.file_frame = tk.Frame(root)
        self.file_frame.pack(pady=5)
        self.file_label = tk.Label(self.file_frame, text="No Excel file selected", 
                                   fg="red")
        self.file_label.pack(side=tk.LEFT, padx=5)
        self.load_btn = tk.Button(self.file_frame, text="1. Load Excel File",
                                  command=self.load_file, bg="#4CAF50", 
                                  fg="white")
        self.load_btn.pack(side=tk.LEFT)

        # 2. Email Configuration
        self.email_frame = tk.LabelFrame(root, text="Email Configuration (For Email Sending)", padx=10, pady=5)
        self.email_frame.pack(pady=5, fill="x", padx=10)

        # SMTP Selection
        tk.Label(self.email_frame, text="Provider:").grid(row=0, column=0, sticky="w")
        self.smtp_var = tk.StringVar(value="Gmail")
        self.smtp_dropdown = tk.OptionMenu(self.email_frame, self.smtp_var, "Gmail", "Outlook")
        self.smtp_dropdown.grid(row=0, column=1, padx=5, sticky="w")

        # Sender Email
        tk.Label(self.email_frame, text="Sender Email:").grid(row=0, column=2, sticky="w")
        self.sender_email_entry = tk.Entry(self.email_frame, width=30)
        self.sender_email_entry.grid(row=0, column=3, padx=5)

        # Sender Password
        tk.Label(self.email_frame, text="App Password:").grid(row=1, column=0, sticky="w")
        self.sender_pass_entry = tk.Entry(self.email_frame, width=20, show="*")
        self.sender_pass_entry.grid(row=1, column=1, padx=5, sticky="w")

        # Load Template
        self.template_label = tk.Label(self.email_frame, text="No template loaded", fg="red")
        self.template_label.grid(row=1, column=2, sticky="w")
        self.load_template_btn = tk.Button(self.email_frame, text="Load .txt Template", command=self.load_template, bg="#FF9800", fg="white")
        self.load_template_btn.grid(row=1, column=3, padx=5)

        # 3. Log Area
        self.log_text = scrolledtext.ScrolledText(root, width=85, height=15, state=tk.DISABLED, font=("Consolas", 9))
        self.log_text.pack(pady=10)

        # 4. Action Buttons
        self.btn_frame = tk.Frame(root)
        self.btn_frame.pack(pady=5)

        self.sms_btn = tk.Button(self.btn_frame, text="Send SMS", command=self.start_sms_thread, bg="#2196F3", fg="white", width=15, height=2, state=tk.DISABLED)
        self.sms_btn.grid(row=0, column=0, padx=5)

        self.wa_btn = tk.Button(self.btn_frame, text="Send WhatsApp", command=self.start_wa_thread, bg="#25D366", fg="white", width=15, height=2, state=tk.DISABLED)
        self.wa_btn.grid(row=0, column=1, padx=5)

        self.email_btn = tk.Button(self.btn_frame, text="Send Emails", command=self.start_email_thread, bg="#9C27B0", fg="white", width=15, height=2, state=tk.DISABLED)
        self.email_btn.grid(row=0, column=2, padx=5)

        self.exit_btn = tk.Button(self.btn_frame, text="Exit Program", command=self.exit_app, bg="#f44336", fg="white", width=15, height=2)
        self.exit_btn.grid(row=1, column=0, columnspan=3, pady=10)

    # --- Helper Functions ---
    def log(self, message):
        """Thread-safe logging to the GUI"""
        def _update():
            self.log_text.config(state=tk.NORMAL)
            self.log_text.insert(tk.END, message + "\n")
            self.log_text.see(tk.END)
            self.log_text.config(state=tk.DISABLED)
        self.root.after(0, _update)

    def load_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
        if not file_path: return

        try:
            self.df = pd.read_excel(file_path)
            # Clean column names: lowercase, strip spaces, replace spaces with underscores for easy templating
            self.df.columns = [str(col).strip().lower().replace(" ", "_") for col in self.df.columns]
            
            self.file_label.config(text=f"Loaded: {os.path.basename(file_path)} ({len(self.df)} rows)", fg="green")
            self.sms_btn.config(state=tk.NORMAL)
            self.wa_btn.config(state=tk.NORMAL)
            self.email_btn.config(state=tk.NORMAL)
            
            self.log(f"Excel loaded. Available fields for template: {', '.join(self.df.columns)}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read file:\n{str(e)}")

    def load_template(self):
        file_path = filedialog.askopenfilename(filetypes=[("Text Files", "*.txt")])
        if not file_path: return
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                self.email_template = f.read()
            self.template_label.config(text=f"Loaded: {os.path.basename(file_path)}", fg="green")
            self.log("Email template loaded successfully.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read template:\n{str(e)}")

    def clean_phone(self, phone):
        return re.sub(r'\D', '', str(phone))

    # --- SMS Logic ---
    def send_sms_process(self):
        active_df = self.df[self.df.get('status', pd.Series(['active']*len(self.df))).str.lower() == 'active'] if 'status' in self.df.columns else self.df
        self.log(f"--- Starting SMS Process for {len(active_df)} contacts ---")
        for index, row in active_df.iterrows():
            phone = str(row.get('phone', row.get('telefono', '')))
            message = str(row.get('message', 'Test message'))
            sms_url = f"sms:{phone}?&body={urllib.parse.quote(message)}"
            self.log(f"Opening SMS for: {phone}")
            webbrowser.open(sms_url)
            time.sleep(1.5)
        self.log("--- SMS Process Complete! ---")
        self.root.after(0, lambda: self.sms_btn.config(state=tk.NORMAL))

    # --- WhatsApp Logic ---
    def send_whatsapp_process(self):
        active_df = self.df[self.df.get('status', pd.Series(['active']*len(self.df))).str.lower() == 'active'] if 'status' in self.df.columns else self.df
        self.log(f"--- Starting WhatsApp Process for {len(active_df)} contacts ---")
        for index, row in active_df.iterrows():
            phone = self.clean_phone(row.get('phone', row.get('telefono', '')))
            message = str(row.get('message', 'Test message'))
            wa_url = f"https://wa.me/{phone}?text={urllib.parse.quote(message)}"
            self.log(f"Opening WhatsApp for: {phone}")
            webbrowser.open(wa_url)
            time.sleep(3)
        self.log("--- WhatsApp Process Complete! ---")
        self.root.after(0, lambda: self.wa_btn.config(state=tk.NORMAL))

    # --- EMAIL Logic ---
    def send_email_process(self):
        sender_email = self.sender_email_entry.get().strip()
        sender_pass = self.sender_pass_entry.get().strip()
        provider = self.smtp_var.get()

        if not sender_email or not sender_pass:
            messagebox.showerror("Error", "Please enter Sender Email and App Password.")
            self.root.after(0, lambda: self.email_btn.config(state=tk.NORMAL))
            return
        if not self.email_template:
            messagebox.showerror("Error", "Please load a .txt message template first.")
            self.root.after(0, lambda: self.email_btn.config(state=tk.NORMAL))
            return

        # Determine SMTP Server
        if provider == "Gmail":
            smtp_server, smtp_port = "smtp.gmail.com", 587
        else:
            smtp_server, smtp_port = "smtp-mail.outlook.com", 587

        # Filter active rows (if 'status' column exists, otherwise send to all)
        if 'status' in self.df.columns:
            active_df = self.df[self.df['status'].astype(str).str.lower() == 'active']
        else:
            active_df = self.df

        self.log(f"--- Connecting to {provider} SMTP Server... ---")
        try:
            server = smtplib.SMTP(smtp_server, smtp_port)
            server.starttls()
            server.login(sender_email, sender_pass)
            self.log("Login successful! Starting to send emails...")
        except Exception as e:
            self.log(f"SMTP LOGIN FAILED: {str(e)}")
            messagebox.showerror("Login Failed", f"Could not connect to {provider}.\n\nIf using Gmail, ensure you are using a 16-digit 'App Password', NOT your regular password.")
            self.root.after(0, lambda: self.email_btn.config(state=tk.NORMAL))
            return

        success_count = 0
        for index, row in active_df.iterrows():
            # Convert row to dictionary, fill NaNs with empty strings
            row_dict = {k: str(v).strip() if pd.notna(v) else "" for k, v in row.items()}
            
            target_email = row_dict.get('email', '').strip()
            if not target_email or "@" not in target_email:
                self.log(f"SKIPPED (No valid email): Row {index}")
                continue

            # Dynamically replace placeholders in the template
            # Using a custom safe substitution to prevent crashes on missing keys
            try:
                # Simple regex replacement for {key}
                import re
                def replace_func(match):
                    key = match.group(1).strip().lower().replace(" ", "_")
                    return row_dict.get(key, f"[{key} not found]")
                
                final_message = re.sub(r'\{(.*?)\}', replace_func, self.email_template)
            except Exception as e:
                self.log(f"Template error for {target_email}: {e}")
                continue

            # Construct Email
            msg = MIMEMultipart()
            msg['From'] = sender_email
            msg['To'] = target_email
            msg['Subject'] = f"Message for {row_dict.get('nombre1', 'Jurado')}"
            msg.attach(MIMEText(final_message, 'plain', 'utf-8'))

            try:
                server.sendmail(sender_email, target_email, msg.as_string())
                self.log(f"SUCCESS: Sent to {target_email}")
                success_count += 1
                time.sleep(1) # Prevent rate-limiting by SMTP server
            except Exception as e:
                self.log(f"FAILED: {target_email} - {str(e)}")

        server.quit()
        self.log(f"--- Email Process Complete! {success_count} emails sent successfully. ---")
        self.root.after(0, lambda: self.email_btn.config(state=tk.NORMAL))

    # --- Threading Wrappers ---
    def start_sms_thread(self):
        self.sms_btn.config(state=tk.DISABLED)
        threading.Thread(target=self.send_sms_process, daemon=True).start()

    def start_wa_thread(self):
        self.wa_btn.config(state=tk.DISABLED)
        threading.Thread(target=self.send_whatsapp_process, daemon=True).start()

    def start_email_thread(self):
        self.email_btn.config(state=tk.DISABLED)
        threading.Thread(target=self.send_email_process, daemon=True).start()

    def exit_app(self):
        if messagebox.askyesno("Exit", "Are you sure you want to exit?"):
            self.root.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = MessageSenderApp(root)
    root.mainloop()