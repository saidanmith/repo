import sys
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QListWidget, QLineEdit, QPushButton, 
                             QLabel, QTextEdit, QMessageBox, QFrame)
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from datetime import datetime

# Import the logic you already built
import email_logic 

class FetchEmailsThread(QThread):
    """Background thread to fetch emails so the UI doesn't freeze."""
    finished = pyqtSignal(list)
    error = pyqtSignal(str)

    def run(self):
        try:
            emails = email_logic.get_lewis_emails()
            self.finished.emit(emails)
        except Exception as e:
            self.error.emit(str(e))

class AsbestosBotApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Asbestos Survey Drafter")
        self.setMinimumSize(1000, 700)
        
        self.all_emails = []
        self.current_email = None
        self.client_label = "Parkingeye"

        self.init_ui()
        self.refresh_inbox()

    def init_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QHBoxLayout(central_widget)

        # --- LEFT COLUMN: Inbox ---
        left_column = QVBoxLayout()
        
        self.inbox_list = QListWidget()
        self.inbox_list.itemClicked.connect(self.on_email_selected)
        
        self.refresh_btn = QPushButton("Refresh Inbox")
        self.refresh_btn.clicked.connect(self.refresh_inbox)
        
        left_column.addWidget(QLabel("<b>Incoming Requests</b>"))
        left_column.addWidget(self.inbox_list)
        left_column.addWidget(self.refresh_btn)

        # --- RIGHT COLUMN: Editor ---
        right_column = QVBoxLayout()
        
        # Contact Details Form
        form_frame = QFrame()
        form_frame.setFrameShape(QFrame.Shape.StyledPanel)
        form_layout = QVBoxLayout(form_frame)

        self.name_input = QLineEdit()
        self.email_input = QLineEdit()
        self.date_input = QLineEdit()
        self.date_input.setPlaceholderText("dd/mm/yyyy")
        self.date_input.textChanged.connect(self.update_preview)

        form_layout.addWidget(QLabel("Site Contact Name:"))
        form_layout.addWidget(self.name_input)
        form_layout.addWidget(QLabel("Site Contact Email:"))
        form_layout.addWidget(self.email_input)
        form_layout.addWidget(QLabel("Proposed Visit Date:"))
        form_layout.addWidget(self.date_input)

        # Preview Area
        self.preview_box = QTextEdit()
        self.preview_box.setReadOnly(True)
        self.preview_box.setStyleSheet("background-color: #f9f9f9; color: #333;")

        # Actions
        btn_layout = QHBoxLayout()
        self.draft_btn = QPushButton("Create Outlook Draft")
        self.draft_btn.setFixedHeight(40)
        self.draft_btn.clicked.connect(self.open_draft)
        
        self.sent_btn = QPushButton("Mark as Sent")
        self.sent_btn.setFixedHeight(40)
        self.sent_btn.setStyleSheet("background-color: #2ecc71; color: white; font-weight: bold;")
        self.sent_btn.clicked.connect(self.mark_as_handled)

        btn_layout.addWidget(self.draft_btn)
        btn_layout.addWidget(self.sent_btn)

        right_column.addWidget(QLabel("<b>Email Composition</b>"))
        right_column.addWidget(form_frame)
        right_column.addWidget(QLabel("<b>Preview:</b>"))
        right_column.addWidget(self.preview_box)
        right_column.addLayout(btn_layout)

        # Add columns to main layout
        main_layout.addLayout(left_column, 1)
        main_layout.addLayout(right_column, 2)

    def refresh_inbox(self):
        self.inbox_list.clear()
        self.inbox_list.addItem("Fetching emails from Outlook...")
        self.thread = FetchEmailsThread()
        self.thread.finished.connect(self.on_emails_fetched)
        self.thread.error.connect(lambda err: QMessageBox.critical(self, "Error", err))
        self.thread.start()

    def on_emails_fetched(self, emails):
        self.inbox_list.clear()
        self.all_emails = emails
        if not emails:
            self.inbox_list.addItem("No new requests found.")
            return
            
        for e in emails:
            self.inbox_list.addItem(f"{e['received_time'].strftime('%d/%m %H:%M')} - {e['subject']}")

    def on_email_selected(self, item):
        idx = self.inbox_list.row(item)
        if idx < 0 or idx >= len(self.all_emails): return
        
        self.current_email = self.all_emails[idx]
        
        # Use your OCR logic
        print("[INFO] Running OCR on attachments...")
        pdfs = email_logic.extract_pdf_attachments(self.current_email)
        job = email_logic.detect_job_type(pdfs, self.current_email['subject'])
        self.client_label = "Parkingeye" if job == "parkingeye" else "G24"
        
        candidates = email_logic.extract_site_contact(self.current_email)
        
        if candidates:
            self.name_input.setText(candidates[0].get('name') or "")
            self.email_input.setText(candidates[0].get('email') or "")
        else:
            self.name_input.clear()
            self.email_input.clear()

        self.update_preview()

    def update_preview(self):
        if not self.current_email: return
        
        name = self.name_input.text() or "[Name]"
        date_str = self.date_input.text()
        
        try:
            day_name, formatted_date = email_logic.parse_visit_date(date_str)
        except:
            day_name, formatted_date = "[Day]", "[Date]"

        body = email_logic.build_email_body(name, self.client_label, day_name, formatted_date)
        self.preview_box.setText(body)

    def open_draft(self):
        if not self.current_email: return
        
        name = self.name_input.text()
        to_email = self.email_input.text()
        date_str = self.date_input.text()
        
        try:
            day_name, formatted_date = email_logic.parse_visit_date(date_str)
            plain_body = email_logic.build_email_body(name, self.client_label, day_name, formatted_date)
            html_intro = email_logic.build_email_body_html(name, self.client_label, day_name, formatted_date)
            subject = f"Asbestos Survey Booking - {formatted_date}"
            
            email_logic.open_forward_draft(self.current_email, to_email, subject, plain_body, html_intro)
        except Exception as e:
            QMessageBox.warning(self, "Input Error", str(e))

    def mark_as_handled(self):
        if not self.current_email: return
        
        # Log it using your logic
        sent_ids = email_logic.load_sent_log()
        sent_ids.add(self.current_email['message_id'])
        email_logic.save_sent_log(sent_ids)
        
        # Remove from UI list
        current_row = self.inbox_list.currentRow()
        self.inbox_list.takeItem(current_row)
        self.all_emails.pop(current_row)
        
        # Clear fields
        self.name_input.clear()
        self.email_input.clear()
        self.preview_box.clear()
        self.current_email = None
        
        QMessageBox.information(self, "Success", "Marked as handled.")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = AsbestosBotApp()
    window.show()
    sys.exit(app.exec())