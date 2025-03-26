import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import smtplib
import csv
import os
import threading
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication


class EmailSenderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Email Campaign Manager")
        self.root.geometry("800x700")
        self.root.configure(bg="#f0f0f0")
        self.root.resizable(True, True)
        
        # Set application icon
        # self.root.iconbitmap("email_icon.ico")  # Uncomment and add your icon if available
        
        # Variables
        self.csv_file_path = tk.StringVar()
        self.attachment_path = tk.StringVar()
        self.sender_email = tk.StringVar()
        self.sender_password = tk.StringVar()
        self.smtp_server = tk.StringVar(value="smtp.gmail.com")
        self.smtp_port = tk.IntVar(value=587)
        self.company_name = tk.StringVar()
        self.sender_name = tk.StringVar()
        self.sender_position = tk.StringVar()
        self.email_subject = tk.StringVar(value="Important Document for Your Review")
        
        # Default values
        self.sender_email.set("")
        self.sender_password.set("")
        self.company_name.set("Your Company Name")
        self.sender_name.set("Your Name")
        self.sender_position.set("Your Position")
        
        # Status variables
        self.total_emails = 0
        self.sent_emails = 0
        self.sending_in_progress = False
        
        self.create_widgets()
        
    def create_widgets(self):
        # Main frame
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Style configuration
        style = ttk.Style()
        style.configure("TFrame", background="#f0f0f0")
        style.configure("TLabel", background="#f0f0f0", font=("Arial", 10))
        style.configure("TButton", font=("Arial", 10, "bold"))
        style.configure("Header.TLabel", font=("Arial", 14, "bold"), background="#f0f0f0")
        style.configure("Subheader.TLabel", font=("Arial", 12, "bold"), background="#f0f0f0")
        
        # Title
        ttk.Label(main_frame, text="Email Campaign Manager", style="Header.TLabel").pack(pady=(0, 20))
        
        # Create notebook for tabs
        notebook = ttk.Notebook(main_frame)
        notebook.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # Setup tab
        setup_frame = ttk.Frame(notebook, padding=10)
        notebook.add(setup_frame, text="Setup")
        
        # Email Content tab
        content_frame = ttk.Frame(notebook, padding=10)
        notebook.add(content_frame, text="Email Content")
        
        # Preview tab
        preview_frame = ttk.Frame(notebook, padding=10)
        notebook.add(preview_frame, text="Preview")
        
        # Setup tab content
        self.setup_tab(setup_frame)
        
        # Email Content tab
        self.content_tab(content_frame)
        
        # Preview tab
        self.preview_tab(preview_frame)
        
        # Status bar
        status_frame = ttk.Frame(main_frame, padding=(0, 10, 0, 0))
        status_frame.pack(fill=tk.X, side=tk.BOTTOM)
        
        self.status_label = ttk.Label(status_frame, text="Ready")
        self.status_label.pack(side=tk.LEFT)
        
        self.progress = ttk.Progressbar(status_frame, orient=tk.HORIZONTAL, length=200, mode='determinate')
        self.progress.pack(side=tk.RIGHT, padx=(10, 0))
        
        # Send button
        send_button = ttk.Button(main_frame, text="Send Emails", command=self.start_email_sending, style="TButton")
        send_button.pack(pady=15)
        
    def setup_tab(self, parent):
        # SMTP Settings
        smtp_frame = ttk.LabelFrame(parent, text="SMTP Settings", padding=10)
        smtp_frame.pack(fill=tk.X, pady=10)
        
        # Sender Email
        ttk.Label(smtp_frame, text="Sender Email:").grid(row=0, column=0, sticky=tk.W, pady=5)
        ttk.Entry(smtp_frame, textvariable=self.sender_email, width=40).grid(row=0, column=1, sticky=tk.W, padx=5)
        
        # Password
        ttk.Label(smtp_frame, text="App Password:").grid(row=1, column=0, sticky=tk.W, pady=5)
        password_entry = ttk.Entry(smtp_frame, textvariable=self.sender_password, width=40, show="*")
        password_entry.grid(row=1, column=1, sticky=tk.W, padx=5)
        
        # SMTP Server
        ttk.Label(smtp_frame, text="SMTP Server:").grid(row=2, column=0, sticky=tk.W, pady=5)
        ttk.Entry(smtp_frame, textvariable=self.smtp_server, width=40).grid(row=2, column=1, sticky=tk.W, padx=5)
        
        # SMTP Port
        ttk.Label(smtp_frame, text="SMTP Port:").grid(row=3, column=0, sticky=tk.W, pady=5)
        ttk.Entry(smtp_frame, textvariable=self.smtp_port, width=40).grid(row=3, column=1, sticky=tk.W, padx=5)
        
        # Company Information
        company_frame = ttk.LabelFrame(parent, text="Company Information", padding=10)
        company_frame.pack(fill=tk.X, pady=10)
        
        # Company Name
        ttk.Label(company_frame, text="Company Name:").grid(row=0, column=0, sticky=tk.W, pady=5)
        ttk.Entry(company_frame, textvariable=self.company_name, width=40).grid(row=0, column=1, sticky=tk.W, padx=5)
        
        # Sender Name
        ttk.Label(company_frame, text="Sender Name:").grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Entry(company_frame, textvariable=self.sender_name, width=40).grid(row=1, column=1, sticky=tk.W, padx=5)
        
        # Sender Position
        ttk.Label(company_frame, text="Sender Position:").grid(row=2, column=0, sticky=tk.W, pady=5)
        ttk.Entry(company_frame, textvariable=self.sender_position, width=40).grid(row=2, column=1, sticky=tk.W, padx=5)
        
        # Files
        files_frame = ttk.LabelFrame(parent, text="Files", padding=10)
        files_frame.pack(fill=tk.X, pady=10)
        
        # CSV File
        ttk.Label(files_frame, text="Recipients CSV:").grid(row=0, column=0, sticky=tk.W, pady=5)
        ttk.Entry(files_frame, textvariable=self.csv_file_path, width=35).grid(row=0, column=1, sticky=tk.W, padx=5)
        ttk.Button(files_frame, text="Browse", command=self.browse_csv).grid(row=0, column=2, padx=5)
        
        # Attachment
        ttk.Label(files_frame, text="Attachment:").grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Entry(files_frame, textvariable=self.attachment_path, width=35).grid(row=1, column=1, sticky=tk.W, padx=5)
        ttk.Button(files_frame, text="Browse", command=self.browse_attachment).grid(row=1, column=2, padx=5)
            
    def content_tab(self, parent):
        # Email Subject
        subject_frame = ttk.Frame(parent)
        subject_frame.pack(fill=tk.X, pady=10)
        
        ttk.Label(subject_frame, text="Email Subject:").pack(side=tk.LEFT, padx=5)
        ttk.Entry(subject_frame, textvariable=self.email_subject, width=60).pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        # Email Body
        body_frame = ttk.LabelFrame(parent, text="Email Body (HTML)", padding=10)
        body_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        self.email_body = scrolledtext.ScrolledText(body_frame, wrap=tk.WORD, height=18)
        self.email_body.pack(fill=tk.BOTH, expand=True)
        
        # Default email template
        default_body = '''
<html>
<body style="font-family: Arial, sans-serif; line-height: 1.6;">
    <p>Dear Recipient,</p>
    
    <p>I hope this email finds you well. Please find attached an important document for your review.</p>
    
    <p>The attached PDF contains essential information regarding our recent project discussion. 
    We would appreciate your feedback at your earliest convenience.</p>
    
    <p>Should you have any questions or require further clarification, please do not hesitate to contact me directly.</p>
    
    <p>Thank you for your attention to this matter.</p>
    
    <p>Best regards,<br>
    {sender_name}<br>
    {sender_position}<br>
    {company_name}</p>
</body>
</html>
'''
        self.email_body.insert(tk.END, default_body)
        
        # Template tags helper
        tags_frame = ttk.Frame(parent)
        tags_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(tags_frame, text="Available Tags:").pack(side=tk.LEFT, padx=5)
        ttk.Label(tags_frame, text="{sender_name}, {sender_position}, {company_name}").pack(side=tk.LEFT, padx=5)
    
    def preview_tab(self, parent):
        # Preview Frame
        preview_frame = ttk.LabelFrame(parent, text="Email Preview", padding=10)
        preview_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # Recipients preview
        ttk.Label(preview_frame, text="Recipients:").pack(anchor=tk.W, pady=(0, 5))
        
        recipients_frame = ttk.Frame(preview_frame)
        recipients_frame.pack(fill=tk.X, pady=5)
        
        self.recipients_display = scrolledtext.ScrolledText(recipients_frame, wrap=tk.WORD, height=5, width=40)
        self.recipients_display.pack(fill=tk.X)
        
        # Attachment preview
        attachment_frame = ttk.Frame(preview_frame)
        attachment_frame.pack(fill=tk.X, pady=10)
        
        ttk.Label(attachment_frame, text="Attachment:").pack(side=tk.LEFT, padx=5)
        self.attachment_label = ttk.Label(attachment_frame, text="No attachment selected")
        self.attachment_label.pack(side=tk.LEFT, padx=5)
        
        # Email content preview
        ttk.Label(preview_frame, text="Email Content Preview:").pack(anchor=tk.W, pady=(10, 5))
        
        self.preview_content = scrolledtext.ScrolledText(preview_frame, wrap=tk.WORD, height=15)
        self.preview_content.pack(fill=tk.BOTH, expand=True)
        
        # Refresh button
        ttk.Button(preview_frame, text="Refresh Preview", command=self.refresh_preview).pack(pady=10)
    
    def browse_csv(self):
        filename = filedialog.askopenfilename(
            title="Select CSV File",
            filetypes=(("CSV files", "*.csv"), ("All files", "*.*"))
        )
        if filename:
            self.csv_file_path.set(filename)
            self.load_recipients_preview()
    
    def browse_attachment(self):
        filename = filedialog.askopenfilename(
            title="Select Attachment",
            filetypes=(("PDF files", "*.pdf"), ("All files", "*.*"))
        )
        if filename:
            self.attachment_path.set(filename)
            self.attachment_label.config(text=os.path.basename(filename))
    
    def load_recipients_preview(self):
        if not self.csv_file_path.get():
            return
            
        recipients = self.read_recipients_from_csv(self.csv_file_path.get())
        
        self.recipients_display.delete(1.0, tk.END)
        if recipients:
            display_text = "\n".join(recipients[:5])
            if len(recipients) > 5:
                display_text += f"\n... ({len(recipients) - 5} more)"
            self.recipients_display.insert(tk.END, display_text)
        else:
            self.recipients_display.insert(tk.END, "No recipients found in the CSV file.")
    
    def refresh_preview(self):
        # Update recipients preview
        self.load_recipients_preview()
        
        # Update content preview
        self.preview_content.delete(1.0, tk.END)
        
        # Get the email body with placeholders replaced
        body = self.email_body.get(1.0, tk.END)
        body = body.format(
            sender_name=self.sender_name.get(),
            sender_position=self.sender_position.get(),
            company_name=self.company_name.get()
        )
        
        # Show a simple version of the HTML
        self.preview_content.insert(tk.END, f"Subject: {self.email_subject.get()}\n\n")
        self.preview_content.insert(tk.END, "Body (HTML):\n")
        self.preview_content.insert(tk.END, body)
    
    def read_recipients_from_csv(self, csv_file):
        recipients = []
        try:
            with open(csv_file, mode='r') as file:
                csv_reader = csv.reader(file)
                header = next(csv_reader, None)  
                
                for row in csv_reader:
                    if row and len(row) > 0:  
                        if len(row) > 1:
                            recipients.append(row[1])
                        else:
                            recipients.append(row[0])
            
            self.total_emails = len(recipients)
            if not recipients:
                messagebox.showwarning("Warning", f"No email addresses found in {csv_file}")
                
            return recipients
        except FileNotFoundError:
            messagebox.showerror("Error", f"CSV file '{csv_file}' not found.")
            return []
        except Exception as e:
            messagebox.showerror("Error", f"Error reading CSV file: {str(e)}")
            return []
    
    def send_email(self, recipient_email, subject, body, sender_email, sender_password, smtp_server, smtp_port, attachment_path=None):
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = recipient_email
        msg['Subject'] = subject
        
        msg.attach(MIMEText(body, 'html'))  
        
        if attachment_path and os.path.exists(attachment_path):
            try:
                with open(attachment_path, "rb") as attachment:
                    part = MIMEApplication(attachment.read(), Name=os.path.basename(attachment_path))
                
                part['Content-Disposition'] = f'attachment; filename="{os.path.basename(attachment_path)}"'
                msg.attach(part)
            except Exception as e:
                self.log_message(f"Error attaching file: {str(e)}")
                return False
        elif attachment_path:
            self.log_message(f"Warning: Attachment file not found: {attachment_path}")
        
        try:
            server = smtplib.SMTP(smtp_server, smtp_port)
            server.starttls()  
            server.login(sender_email, sender_password)
            
            text = msg.as_string()
            server.sendmail(sender_email, recipient_email, text)
            
            server.quit()
            self.log_message(f"Email sent to {recipient_email}")
            return True
        except Exception as e:
            self.log_message(f"Failed to send email to {recipient_email}. Error: {str(e)}")
            return False
    
    def start_email_sending(self):
        if self.sending_in_progress:
            messagebox.showinfo("Info", "Email sending is already in progress")
            return
            
        # Validate fields
        if not self.validate_inputs():
            return
            
        # Ask for confirmation
        recipients = self.read_recipients_from_csv(self.csv_file_path.get())
        if not recipients:
            messagebox.showwarning("Warning", "No recipients found in CSV file")
            return
            
        confirm = messagebox.askyesno("Confirm", f"Send emails to {len(recipients)} recipients?")
        if not confirm:
            return
            
        # Start email sending thread
        self.sending_in_progress = True
        self.sent_emails = 0
        self.progress['maximum'] = len(recipients)
        self.progress['value'] = 0
        
        # Prepare the email body
        body = self.email_body.get(1.0, tk.END)
        body = body.format(
            sender_name=self.sender_name.get(),
            sender_position=self.sender_position.get(),
            company_name=self.company_name.get()
        )
        
        threading.Thread(
            target=self.send_emails_thread,
            args=(recipients, self.email_subject.get(), body)
        ).start()
    
    def send_emails_thread(self, recipients, subject, body):
        successful = 0
        
        for recipient in recipients:
            if self.send_email(
                recipient, 
                subject, 
                body, 
                self.sender_email.get(), 
                self.sender_password.get(), 
                self.smtp_server.get(), 
                self.smtp_port.get(), 
                self.attachment_path.get()
            ):
                successful += 1
                
            self.sent_emails += 1
            self.progress['value'] = self.sent_emails
            self.status_label.config(text=f"Sending: {self.sent_emails}/{len(recipients)}")
            self.root.update_idletasks()
        
        self.sending_in_progress = False
        messagebox.showinfo("Complete", f"Email sending complete. Successfully sent {successful} out of {len(recipients)} emails.")
        self.status_label.config(text="Ready")
    
    def validate_inputs(self):
        # Check required fields
        if not self.sender_email.get():
            messagebox.showerror("Error", "Sender email is required")
            return False
            
        if not self.sender_password.get():
            messagebox.showerror("Error", "App password is required")
            return False
            
        if not self.csv_file_path.get():
            messagebox.showerror("Error", "CSV file is required")
            return False
            
        # Validate file paths
        if not os.path.exists(self.csv_file_path.get()):
            messagebox.showerror("Error", "CSV file does not exist")
            return False
            
        if self.attachment_path.get() and not os.path.exists(self.attachment_path.get()):
            messagebox.showerror("Error", "Attachment file does not exist")
            return False
        
        return True
    
    def log_message(self, message):
        print(message)
        # You could also add a log display in the UI
        

if __name__ == "__main__":
    root = tk.Tk()
    app = EmailSenderApp(root)
    root.mainloop()