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

#  dksy ppty xoht sodm
# =============================================================================
#  POPUP WINDOW: Email Configuration (only appears when "Send Emails" clicked)
# =============================================================================     
class EmailConfigPopup:
    def __init__(self, parent, app):
        self.app = app
        self.win = tk.Toplevel(parent)
        self.win.title("📧  Email Configuration")
        self.win.geometry("560x420")
        self.win.resizable(False, False)
        self.win.configure(bg="#2C2C2C")
        self.win.grab_set()  # Modal: blocks main window until closed

        # --- Title ---
        tk.Label(
            self.win, text="Email Sending Configuration",
            font=("Arial", 14, "bold"), fg="#FFFFFF", bg="#2C2C2C"
        ).pack(pady=(15, 10))

        # --- Form Frame ---
        form = tk.Frame(self.win, bg="#2C2C2C")
        form.pack(padx=20, fill="x")

        # Row 0: Provider
        tk.Label(form, text="Email Provider:", fg="#CCCCCC", bg="#2C2C2C",
                 font=("Arial", 10)).grid(row=0, column=0, sticky="w", pady=6)
        self.provider_var = tk.StringVar(value="Gmail")
        self.provider_dd = tk.OptionMenu(form, self.provider_var, "Gmail", "Outlook")
        self.provider_dd.config(width=18, font=("Arial", 10))
        self.provider_dd.grid(row=0, column=1, padx=10, pady=6, sticky="w")

        # Row 1: Sender Email
        tk.Label(form, text="Sender Email:", fg="#CCCCCC", bg="#2C2C2C",
                 font=("Arial", 10)).grid(row=1, column=0, sticky="w", pady=6)
        self.email_entry = tk.Entry(form, width=35, font=("Arial", 10))
        self.email_entry.grid(row=1, column=1, padx=10, pady=6, sticky="w")

        # Row 2: App Password
        tk.Label(form, text="App Password:", fg="#CCCCCC", bg="#2C2C2C",
                 font=("Arial", 10)).grid(row=2, column=0, sticky="w", pady=6)
        self.pass_entry = tk.Entry(form, width=25, font=("Arial", 10), show="•")
        self.pass_entry.grid(row=2, column=1, padx=10, pady=6, sticky="w")

        # Row 3: Subject Line
        tk.Label(form, text="Email Subject:", fg="#CCCCCC", bg="#2C2C2C",
                 font=("Arial", 10)).grid(row=3, column=0, sticky="w", pady=6)
        self.subject_entry = tk.Entry(form, width=35, font=("Arial", 10))
        self.subject_entry.insert(0, "Notificación de Jurado de Votación")
        self.subject_entry.grid(row=3, column=1, padx=10, pady=6, sticky="w")

        # Row 4: Template
        tk.Label(form, text="Message Template:", fg="#CCCCCC", bg="#2C2C2C",
                 font=("Arial", 10)).grid(row=4, column=0, sticky="w", pady=6)
        self.template_label = tk.Label(form, text="No template loaded", fg="#FF6B6B",
                                       bg="#2C2C2C", font=("Arial", 9))
        self.template_label.grid(row=4, column=1, padx=10, pady=6, sticky="w")

        tk.Button(
            form, text="Load .txt Template", command=self.load_template,
            bg="#FF9800", fg="#FFFFFF", font=("Arial", 9, "bold"),
            activebackground="#E68900", activeforeground="#FFFFFF",
            relief="raised", bd=2, cursor="hand2"
        ).grid(row=5, column=1, padx=10, sticky="w")

        # --- Info Label ---
        tk.Label(
            self.win,
            text="⚠ Gmail requires a 16-digit App Password (not your regular password).\n"
                 "Go to: Google Account → Security → 2-Step Verification → App Passwords",
            fg="#FFD54F", bg="#2C2C2C", font=("Arial", 8), justify="left"
        ).pack(pady=(10, 5), padx=20, anchor="w")

        # --- Action Buttons ---
        btn_frame = tk.Frame(self.win, bg="#2C2C2C")
        btn_frame.pack(pady=(10, 15))

        tk.Button(
            btn_frame, text="📤  SEND EMAILS NOW", command=self.start_sending,
            bg="#9C27B0", fg="#FFFFFF", font=("Arial", 12, "bold"),
            activebackground="#7B1FA2", activeforeground="#FFFFFF",
            width=22, height=2, relief="raised", bd=3, cursor="hand2"
        ).grid(row=0, column=0, padx=10)

        tk.Button(
            btn_frame, text="Cancel", command=self.win.destroy,
            bg="#757575", fg="#FFFFFF", font=("Arial", 11, "bold"),
            activebackground="#616161", activeforeground="#FFFFFF",
            width=12, height=2, relief="raised", bd=3, cursor="hand2"
        ).grid(row=0, column=1, padx=10)

        self.email_template = None

    def load_template(self):
        path = filedialog.askopenfilename(
            filetypes=[("Text Files", "*.txt")], parent=self.win
        )
        if not path:
            return
        try:
            with open(path, "r", encoding="utf-8") as f:
                self.email_template = f.read()
            self.template_label.config(text=f"✅ {os.path.basename(path)}", fg="#66BB6A")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read template:\n{e}", parent=self.win)

    def start_sending(self):
        sender_email = self.email_entry.get().strip()
        sender_pass = self.pass_entry.get().strip()
        provider = self.provider_var.get()
        subject = self.subject_entry.get().strip()

        if not sender_email or not sender_pass:
            messagebox.showerror("Missing Info", "Please enter your email and app password.",
                                 parent=self.win)
            return
        if not self.email_template:
            messagebox.showerror("Missing Template", "Please load a .txt message template first.",
                                 parent=self.win)
            return
        if self.app.df is None:
            messagebox.showerror("No Data", "Please load an Excel file first.", parent=self.win)
            return

        # Close popup and start sending in background thread
        self.win.destroy()
        threading.Thread(
            target=self.app.send_email_process,
            args=(sender_email, sender_pass, provider, subject, self.email_template),
            daemon=True
        ).start()


# =============================================================================
#  MAIN APPLICATION WINDOW
# =============================================================================
class MessageSenderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("📨 Excel Message Sender Pro")
        self.root.geometry("780x620")
        self.root.configure(bg="#1E1E1E")

        self.df = None

        # ── Shared button style constants ──
        self.BTN_FONT = ("Arial", 12, "bold")
        self.BTN_FG = "#FFFFFF"
        self.BTN_ACTIVE_FG = "#FFFFFF"
        self.BTN_BD = 3
        self.BTN_RELIEF = "raised"
        self.BTN_CURSOR = "hand2"

        # ── Title ──
        tk.Label(
            root, text="📨  Excel Message Sender Pro",
            font=("Arial", 18, "bold"), fg="#FFFFFF", bg="#1E1E1E"
        ).pack(pady=(15, 5))

        tk.Label(
            root, text="Send SMS, WhatsApp & Email messages from your Excel list",
            font=("Arial", 10), fg="#AAAAAA", bg="#1E1E1E"
        ).pack(pady=(0, 10))

        # ── File Selection Bar ──
        file_frame = tk.Frame(root, bg="#1E1E1E")
        file_frame.pack(pady=5)

        self.file_label = tk.Label(
            file_frame, text="❌  No Excel file loaded",
            fg="#FF6B6B", bg="#1E1E1E", font=("Arial", 10, "bold")
        )
        self.file_label.pack(side=tk.LEFT, padx=10)

        tk.Button(
            file_frame, text="📂  Load Excel File", command=self.load_file,
            bg="#4CAF50", fg=self.BTN_FG, font=self.BTN_FONT,
            activebackground="#388E3C", activeforeground=self.BTN_ACTIVE_FG,
            bd=self.BTN_BD, relief=self.BTN_RELIEF, cursor=self.BTN_CURSOR,
            width=20, height=1
        ).pack(side=tk.LEFT, padx=5)

        # ── Log Area ──
        self.log_text = scrolledtext.ScrolledText(
            root, width=90, height=16, state=tk.DISABLED,
            font=("Consolas", 9), bg="#0D1117", fg="#C9D1D9",
            insertbackground="#FFFFFF", selectbackground="#264F78"
        )
        self.log_text.pack(pady=10, padx=15)

        # ── Action Buttons Row ──
        btn_frame = tk.Frame(root, bg="#1E1E1E")
        btn_frame.pack(pady=15)

        # SMS Button
        self.sms_btn = tk.Button(
            btn_frame, text="📱  Send SMS", command=self.start_sms_thread,
            bg="#1565C0", fg=self.BTN_FG, font=self.BTN_FONT,
            activebackground="#0D47A1", activeforeground=self.BTN_ACTIVE_FG,
            width=20, height=2, bd=self.BTN_BD, relief=self.BTN_RELIEF,
            cursor=self.BTN_CURSOR, state=tk.DISABLED
        )
        self.sms_btn.grid(row=0, column=0, padx=8)

        # WhatsApp Button
        self.wa_btn = tk.Button(
            btn_frame, text="💬  Send WhatsApp", command=self.start_wa_thread,
            bg="#2E7D32", fg=self.BTN_FG, font=self.BTN_FONT,
            activebackground="#1B5E20", activeforeground=self.BTN_ACTIVE_FG,
            width=20, height=2, bd=self.BTN_BD, relief=self.BTN_RELIEF,
            cursor=self.BTN_CURSOR, state=tk.DISABLED
        )
        self.wa_btn.grid(row=0, column=1, padx=8)

        # Email Button (opens popup)
        self.email_btn = tk.Button(
            btn_frame, text="📧  Send Emails", command=self.open_email_popup,
            bg="#7B1FA2", fg=self.BTN_FG, font=self.BTN_FONT,
            activebackground="#4A148C", activeforeground=self.BTN_ACTIVE_FG,
            width=20, height=2, bd=self.BTN_BD, relief=self.BTN_RELIEF,
            cursor=self.BTN_CURSOR, state=tk.DISABLED
        )
        self.email_btn.grid(row=0, column=2, padx=8)

        # Exit Button
        tk.Button(
            btn_frame, text="🚪  Exit Program", command=self.exit_app,
            bg="#C62828", fg=self.BTN_FG, font=self.BTN_FONT,
            activebackground="#B71C1C", activeforeground=self.BTN_ACTIVE_FG,
            width=20, height=2, bd=self.BTN_BD, relief=self.BTN_RELIEF,
            cursor=self.BTN_CURSOR
        ).grid(row=1, column=0, columnspan=3, pady=(15, 0))

    # ──────────────────────────────────────────────
    #  HELPERS
    # ──────────────────────────────────────────────
    def log(self, message):
        def _update():
            self.log_text.config(state=tk.NORMAL)
            self.log_text.insert(tk.END, message + "\n")
            self.log_text.see(tk.END)
            self.log_text.config(state=tk.DISABLED)
        self.root.after(0, _update)

    def clean_phone(self, phone):
        return re.sub(r'\D', '', str(phone))

    # ──────────────────────────────────────────────
    #  LOAD & CLEAN EXCEL FILE
    # ──────────────────────────────────────────────
    def load_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
        if not path:
            return

        try:
            # Read raw
            raw = pd.read_excel(path, header=None)
            self.log(f"Raw file loaded: {raw.shape[0]} rows × {raw.shape[1]} columns")

            # --- Smart header detection ---
            # Find rows that look like headers (contain "Phone" or "Name")
            header_indices = []
            for idx, row in raw.iterrows():
                row_str = " ".join(str(v) for v in row.values if pd.notna(v))
                if "phone" in row_str and "Name" in row_str:
                    header_indices.append(idx)

            if not header_indices:
                messagebox.showerror("Error", "Could not find header rows with phone/Name.")
                return

            # Use the first detected header row
            main_header_idx = header_indices[0]
            headers = [str(v).strip() for v in raw.iloc[main_header_idx].values]

            # Handle duplicate column names by appending suffixes
            seen = {}
            clean_headers = []
            for h in headers:
                if h == "nan" or h.strip() == "":
                    h = f"empty_{len(seen)}"
                if h in seen:
                    seen[h] += 1
                    clean_headers.append(f"{h}_{seen[h]}")
                else:
                    seen[h] = 0
                    clean_headers.append(h)

            # Extract data rows (skip all header rows, keep only data)
            data_rows = []
            for idx, row in raw.iterrows():
                if idx in header_indices:
                    continue
                # Skip completely empty rows
                non_null = row.dropna()
                if len(non_null) == 0:
                    continue
                # Skip rows that are just mesa numbers (only 1-2 values)
                vals = [str(v).strip() for v in row.values if pd.notna(v) and str(v).strip()]
                if len(vals) < 4:
                    continue
                # Skip rows that start with "mailto:" (email dump rows)
                first_val = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
                if first_val.startswith("mailto:"):
                    continue
                data_rows.append(row.values)

            df = pd.DataFrame(data_rows, columns=clean_headers)

            # Normalize column names: lowercase, replace spaces with underscores
            df.columns = [c.lower().replace(" ", "_") for c in df.columns]

            # Drop rows where key fields are all empty
            key_cols = [c for c in df.columns if c.startswith("Name") or c.startswith("phone")]
            if key_cols:
                df = df.dropna(subset=key_cols, how="all")

            # Replace NaN with empty string for template safety
            df = df.fillna("")

            self.df = df
            self.file_label.config(
                text=f"✅  Loaded: {os.path.basename(path)}  ({len(df)} data rows)",
                fg="#66BB6A"
            )
            self.sms_btn.config(state=tk.NORMAL)
            self.wa_btn.config(state=tk.NORMAL)
            self.email_btn.config(state=tk.NORMAL)

            self.log(f"Cleaned data: {len(df)} rows with actual juror/person data.")
            self.log(f"Available template fields: {', '.join(sorted(df.columns))}")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to read file:\n{str(e)}")

    # ──────────────────────────────────────────────
    #  OPEN EMAIL POPUP
    # ──────────────────────────────────────────────
    def open_email_popup(self):
        if self.df is None:
            messagebox.showwarning("No Data", "Please load an Excel file first.")
            return
        EmailConfigPopup(self.root, self)

    # ──────────────────────────────────────────────
    #  SMS PROCESS
    # ──────────────────────────────────────────────
    def send_sms_process(self):
        df = self.df
        # Use TELEFONO column (first occurrence)
        phone_col = next((c for c in df.columns if c == "phone"), None)
        name_col = next((c for c in df.columns if c == "Name"), None)

        if not phone_col:
            self.log("ERROR: No 'TELEFONO' column found in the file.")
            self.root.after(0, lambda: self.sms_btn.config(state=tk.NORMAL))
            return

        # Filter: only rows with actual phone numbers
        active = df[df[phone_col].astype(str).str.strip() != ""]
        self.log(f"--- Starting SMS for {len(active)} contacts with phone numbers ---")

        count = 0
        for _, row in active.iterrows():
            phone = self.clean_phone(row[phone_col])
            name = row[name_col] if name_col else "Contact"
            if not phone:
                continue

            # Build a default message from available fields
            msg = f"Hola {name}, le recordamos su asignación como jurado de votación."
            sms_url = f"sms:+1{phone}?&body={urllib.parse.quote(msg)}"

            self.log(f"  📱 SMS → {name} ({phone})")
            webbrowser.open(sms_url)
            time.sleep(1.5)
            count += 1

        self.log(f"--- SMS Complete: {count} messages opened. ---")
        self.root.after(0, lambda: self.sms_btn.config(state=tk.NORMAL))

    # ──────────────────────────────────────────────
    #  WHATSAPP PROCESS
    # ──────────────────────────────────────────────
    def send_whatsapp_process(self):
        df = self.df
        phone_col = next((c for c in df.columns if c == "phone"), None)
        name_col = next((c for c in df.columns if c == "Name"), None)

        if not phone_col:
            self.log("ERROR: No 'TELEFONO' column found.")
            self.root.after(0, lambda: self.wa_btn.config(state=tk.NORMAL))
            return

        active = df[df[phone_col].astype(str).str.strip() != ""]
        self.log(f"--- Starting WhatsApp for {len(active)} contacts ---")

        count = 0
        for _, row in active.iterrows():
            phone = self.clean_phone(row[phone_col])
            name = row[name_col] if name_col else "Contact"
            if not phone or len(phone) < 7:
                continue

            # Build message from row data
            mesa = row.get("mesa", "")
            puesto = row.get("puesto", "")
            dia = row.get("dia", "")
            msg = (f"Hola {name}, le recordamos que ha sido asignado como jurado "
                   f"de votación.\n📍 Puesto: {puesto}\n📋 Mesa: {mesa}\n📅 Día: {dia}")

            wa_url = f"https://wa.me/{phone}?text={urllib.parse.quote(msg)}"
            self.log(f"  💬 WhatsApp → {name} ({phone})")
            webbrowser.open(wa_url)
            time.sleep(3)
            count += 1

        self.log(f"--- WhatsApp Complete: {count} tabs opened. Click Send in each tab. ---")
        self.root.after(0, lambda: self.wa_btn.config(state=tk.NORMAL))

    # ──────────────────────────────────────────────
    #  EMAIL PROCESS (called from popup thread)
    # ──────────────────────────────────────────────
    def send_email_process(self, sender_email, sender_pass, provider, subject, template):
        # SMTP settings
        if provider == "Gmail":
            smtp_server, smtp_port = "smtp.gmail.com", 587
        else:
            smtp_server, smtp_port = "smtp-mail.outlook.com", 587

        # Find email column
        df = self.df
        email_col = next((c for c in df.columns if c == "email"), None)
        if not email_col:
            self.log("ERROR: No 'EMAIL' column found.")
            self.root.after(0, lambda: self.email_btn.config(state=tk.NORMAL))
            return

        active = df[df[email_col].astype(str).str.contains("@", na=False)]
        self.log(f"--- Connecting to {provider} SMTP ({smtp_server}:{smtp_port})... ---")

        try:
            server = smtplib.SMTP(smtp_server, smtp_port)
            server.starttls()
            server.login(sender_email, sender_pass)
            self.log("✅  Login successful! Sending emails...")
        except Exception as e:
            self.log(f"❌  SMTP LOGIN FAILED: {e}")
            if provider == "Gmail":
                self.log("   → Make sure you're using a 16-digit App Password, not your regular password.")
            self.root.after(0, lambda: self.email_btn.config(state=tk.NORMAL))
            return

        success = 0
        for _, row in active.iterrows():
            target = str(row[email_col]).strip()
            if not target or "@" not in target:
                continue

            # Build row dictionary for template substitution
            row_dict = {k: str(v).strip() for k, v in row.items() if v != ""}

            # Safe placeholder replacement
            def replace_placeholder(match):
                key = match.group(1).strip().lower().replace(" ", "_")
                return row_dict.get(key, f"[{key}]")

            body = re.sub(r'\{(.*?)\}', replace_placeholder, template)

            msg = MIMEMultipart()
            msg["From"] = sender_email
            msg["To"] = target
            msg["Subject"] = subject
            msg.attach(MIMEText(body, "plain", "utf-8"))

            try:
                server.sendmail(sender_email, target, msg.as_string())
                self.log(f"  ✅  Sent → {target}")
                success += 1
                time.sleep(1)  # Rate limit protection
            except Exception as e:
                self.log(f"  ❌  Failed → {target}: {e}")

        server.quit()
        self.log(f"--- Email Complete: {success}/{len(active)} sent successfully. ---")
        self.root.after(0, lambda: self.email_btn.config(state=tk.NORMAL))

    # ──────────────────────────────────────────────
    #  THREAD WRAPPERS
    # ──────────────────────────────────────────────
    # ──────────────────────────────────────────────
        #  OPEN SMS
        # ──────────────────────────────────────────────
    def start_sms_thread(self):
        self.sms_btn.config(state=tk.DISABLED)
        threading.Thread(target=self.send_sms_process, daemon=True).start()
    # ──────────────────────────────────────────────
    #  OPEN WHATSAPP
    # ──────────────────────────────────────────────

    def start_wa_thread(self):
        self.wa_btn.config(state=tk.DISABLED)
        threading.Thread(target=self.send_whatsapp_process, daemon=True).start()
    # ──────────────────────────────────────────────
    #  OPEN EMAIL POPUP
    # ──────────────────────────────────────────────
    def open_email_popup(self):
        if self.df is None:
            messagebox.showwarning("No Data", "Please load an Excel file first.")
            return
        EmailConfigPopup(self.root, self)       

    # ──────────────────────────────────────────────
    #  EXIT
    # ──────────────────────────────────────────────
    def exit_app(self):
        if messagebox.askyesno("Exit", "Are you sure you want to exit?"):
            self.root.destroy()


# =============================================================================
#  RUN
# =============================================================================
if __name__ == "__main__":
    root = tk.Tk()
    app = MessageSenderApp(root)
    root.mainloop()