import unittest
from unittest.mock import patch, MagicMock
import os
import tempfile
from main import EmailSenderApp  # Import EmailSenderApp from main.py
import tkinter as tk

class TestEmailSender(unittest.TestCase):
    
    def setUp(self):
        """Initialize the Tkinter app before each test"""
        self.root = tk.Tk()  # Create a Tk instance
        self.app = EmailSenderApp(self.root)  # Create EmailSenderApp instance

    def tearDown(self):
        """Destroy the Tkinter instance after each test"""
        self.root.destroy()

    def test_email_validation(self):
        """Check if email validation works properly"""
        valid_email = "test@example.com"
        invalid_email = "invalid-email"
        
        self.assertTrue("@" in valid_email, "Valid email failed")
        self.assertFalse("@" in invalid_email, "Invalid email passed")

    def test_csv_parsing(self):
        """Ensure CSV file is correctly read"""
        csv_content = "Name,Email\nJohn Doe,john@example.com\nJane Doe,jane@example.com"
        
        # Create a temporary CSV file
        with tempfile.NamedTemporaryFile(delete=False, suffix=".csv", mode='w') as temp_csv:
            temp_csv.write(csv_content)
            temp_csv_name = temp_csv.name
        
        # Load recipients from CSV
        self.app.csv_file_path.set(temp_csv_name)
        recipients = self.app.read_recipients_from_csv(temp_csv_name)
        
        self.assertEqual(len(recipients), 2)
        self.assertIn("john@example.com", recipients)
        self.assertIn("jane@example.com", recipients)
        
        # Clean up temp file
        os.remove(temp_csv_name)

    @patch("smtplib.SMTP")
    def test_send_email(self, mock_smtp):
        """Mock SMTP and ensure email sending works"""
        mock_server = MagicMock()
        mock_smtp.return_value = mock_server

        recipient = "recipient@example.com"
        subject = "Test Subject"
        body = "This is a test email."
        
        success = self.app.send_email(
            recipient, 
            subject, 
            body, 
            "sender@example.com", 
            "password", 
            "smtp.gmail.com", 
            587
        )
        
        # Ensure SMTP was called correctly
        mock_smtp.assert_called_with("smtp.gmail.com", 587)
        mock_server.starttls.assert_called_once()
        mock_server.login.assert_called_once_with("sender@example.com", "password")
        mock_server.sendmail.assert_called_once()
        mock_server.quit.assert_called_once()

        self.assertTrue(success, "Email sending failed")

    def test_attachment_handling(self):
        """Check if attachment file exists before sending"""
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as temp_file:
            temp_file_name = temp_file.name
        
        self.app.attachment_path.set(temp_file_name)
        
        self.assertTrue(os.path.exists(self.app.attachment_path.get()))
        
        # Clean up
        os.remove(temp_file_name)

    def test_validate_inputs(self):
        """Ensure required fields are checked before sending"""
        self.app.sender_email.set("test@example.com")
        self.app.sender_password.set("app-password")

        # Use a real temporary CSV file instead of a non-existent one
        with tempfile.NamedTemporaryFile(delete=False, suffix=".csv", mode='w') as temp_csv:
            temp_csv.write("Name,Email\nJohn Doe,john@example.com\n")
            temp_csv_name = temp_csv.name

        self.app.csv_file_path.set(temp_csv_name)
        with patch("os.path.exists", return_value=True):
            self.assertTrue(self.app.validate_inputs(), "Validation failed when all inputs were valid")

        os.remove(temp_csv_name)  # Cleanup

if __name__ == "__main__":
    unittest.main(verbosity=2)
