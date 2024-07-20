import tkinter as tk
from tkinter import ttk, messagebox
import csv
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import time
import os
import socket

class EmailSenderApp:
    def __init__(self, master):
        self.master = master
        master.title("Automatic Email Sender")
        master.geometry("600x400")

        self.create_widgets()
        self.csv_file = "email_list.csv"
        self.last_modified_time = self.get_csv_modified_time()
        self.master.after(5000, self.check_csv_updates)

    def create_widgets(self):
        # Email content input
        ttk.Label(self.master, text="Email Content:").pack(pady=5)
        self.email_content = tk.Text(self.master, height=10)
        self.email_content.pack(pady=5, padx=10, fill=tk.BOTH, expand=True)

        # SMTP settings
        settings_frame = ttk.LabelFrame(self.master, text="SMTP Settings")
        settings_frame.pack(pady=10, padx=10, fill=tk.X)

        ttk.Label(settings_frame, text="SMTP Server:").grid(row=0, column=0, sticky="w", padx=5, pady=2)
        self.smtp_server = ttk.Entry(settings_frame)
        self.smtp_server.grid(row=0, column=1, padx=5, pady=2)

        ttk.Label(settings_frame, text="Port:").grid(row=1, column=0, sticky="w", padx=5, pady=2)
        self.smtp_port = ttk.Entry(settings_frame)
        self.smtp_port.grid(row=1, column=1, padx=5, pady=2)

        ttk.Label(settings_frame, text="Username:").grid(row=2, column=0, sticky="w", padx=5, pady=2)
        self.username = ttk.Entry(settings_frame)
        self.username.grid(row=2, column=1, padx=5, pady=2)

        ttk.Label(settings_frame, text="Password:").grid(row=3, column=0, sticky="w", padx=5, pady=2)
        self.password = ttk.Entry(settings_frame, show="*")
        self.password.grid(row=3, column=1, padx=5, pady=2)

        # Start button
        self.start_button = ttk.Button(self.master, text="Start Sending Emails", command=self.start_sending_emails)
        self.start_button.pack(pady=10)

        # Status label
        self.status_label = ttk.Label(self.master, text="")
        self.status_label.pack(pady=5)

    def start_sending_emails(self):
        if not self.validate_inputs():
            return
        self.status_label.config(text="Sending emails...")
        self.send_emails()

    def validate_inputs(self):
        if not self.email_content.get("1.0", tk.END).strip():
            messagebox.showerror("Error", "Email content cannot be empty.")
            return False
        if not all([self.smtp_server.get(), self.smtp_port.get(), self.username.get(), self.password.get()]):
            messagebox.showerror("Error", "All SMTP settings must be filled.")
            return False
        try:
            int(self.smtp_port.get())
        except ValueError:
            messagebox.showerror("Error", "Port must be a number.")
            return False
        return True

    def send_emails(self):
        content = self.email_content.get("1.0", tk.END)
        smtp_server = self.smtp_server.get()
        smtp_port = int(self.smtp_port.get())
        username = self.username.get()
        password = self.password.get()

        try:
            with smtplib.SMTP(smtp_server, smtp_port, timeout=10) as server:
                server.starttls()
                server.login(username, password)

                if not os.path.exists(self.csv_file):
                    raise FileNotFoundError(f"CSV file '{self.csv_file}' not found.")

                with open(self.csv_file, 'r', encoding='utf-8-sig') as file:
                    reader = csv.reader(file)
                    for row in reader:
                        if not row:
                            continue
                        email = row[0]
                        if not self.is_valid_email(email):
                            self.log_error(f"Invalid email address: {email}")
                            continue
                        try:
                            msg = MIMEMultipart()
                            msg['From'] = username
                            msg['To'] = email
                            msg['Subject'] = "Automatic Email"
                            msg.attach(MIMEText(content, 'plain', 'utf-8'))
                            server.send_message(msg)
                            time.sleep(1)  # Delay to avoid overwhelming the SMTP server
                        except smtplib.SMTPException as e:
                            self.log_error(f"Failed to send email to {email}: {str(e)}")

            self.status_label.config(text="Emails sent successfully!")
        except FileNotFoundError as e:
            self.show_error(f"File Error: {str(e)}")
        except smtplib.SMTPAuthenticationError:
            self.show_error("SMTP Authentication Error: Invalid username or password.")
        except socket.gaierror:
            self.show_error("Network Error: Unable to connect to the SMTP server. Check your internet connection and server address.")
        except socket.timeout:
            self.show_error("Timeout Error: The connection to the SMTP server timed out.")
        except smtplib.SMTPException as e:
            self.show_error(f"SMTP Error: {str(e)}")
        except UnicodeDecodeError as e:
            self.show_error(f"Encoding Error: Unable to read the CSV file. Try saving it as UTF-8. Error: {str(e)}")
        except Exception as e:
            self.show_error(f"Unexpected Error: {str(e)}")

    def is_valid_email(self, email):
        # Basic email validation
        return '@' in email and '.' in email.split('@')[1]

    def log_error(self, message):
        try:
            with open("error_log.txt", "a", encoding='utf-8') as log_file:
                log_file.write(f"{time.ctime()}: {message}\n")
        except Exception as e:
            print(f"Failed to write to error log: {str(e)}")

    def show_error(self, message):
        self.status_label.config(text="Error occurred. Check error log.")
        messagebox.showerror("Error", message)
        self.log_error(message)

    def get_csv_modified_time(self):
        try:
            return os.path.getmtime(self.csv_file)
        except FileNotFoundError:
            self.show_error(f"CSV file '{self.csv_file}' not found.")
            return None

    def check_csv_updates(self):
        current_modified_time = self.get_csv_modified_time()
        if current_modified_time and current_modified_time > self.last_modified_time:
            self.last_modified_time = current_modified_time
            self.send_emails()
        self.master.after(5000, self.check_csv_updates)

if __name__ == "__main__":
    root = tk.Tk()
    app = EmailSenderApp(root)
    root.mainloop()