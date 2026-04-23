import sys
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QListWidget, QLineEdit, QPushButton, 
                             QLabel, QTextEdit, QMessageBox, QFrame)
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from datetime import datetime
import pythoncom

# Import the logic you already built
import email_logic

class FetchEmailsThread(QThread):
    """Background thread to fetch emails so the UI doesn't freeze."""
    finished = pyqtSignal(list)
    error = pyqtSignal(str)

    def run(self):
        pythoncom.CoInitialize()
        try:
            emails = email_logic.get_lewis_emails()
            self.finished.emit(emails)
        except Exception as e:
            self.error.emit(str(e))
        finally:
            pythoncom.CoUninitialize()

class AnalyzeEmailThread(QThread):
    """Handles slow OCR and PDF parsing in the background."""
    finished = pyqtSignal(dict)
    error = pyqtSignal(str)

    def __init__(self, email_dict, parent=None):
        super().__init__(parent)
        self.email_dict = email_dict

    def run(self):
        pythoncom.CoInitialize()
        try:
            # Fetch a thread-local MailItem for processing
            item = email_logic.get_mail_item(self.email_dict['message_id'], self.email_dict.get('store_id'))
            if not item:
                raise Exception("Could not re-fetch email for background processing.")

            # 1. Extract PDFs & Detect Job (passing raw item)
            pdfs = email_logic.extract_pdf_attachments(item)
            job = email_logic.detect_job_type(pdfs, self.email_dict['subject'])
            client = "Parkingeye" if job == "parkingeye" else "G24"
            
            address = email_logic.extract_address_from_pdfs(pdfs)
            
            # 2. Extract Site Contacts via OCR (passing raw item)
            candidates = email_logic.extract_site_contact(item)
            
            self.finished.emit({
                "client": client,
                "candidates": candidates,
                "address": address
            })
        except Exception as e:
            self.error.emit(str(e))
        finally:
            pythoncom.CoUninitialize()

class AsbestosBotApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Asbestos Survey Drafter")
        self.setMinimumSize(1000, 700)
        
        self.all_emails = []
        self.current_email = None
        self.client_label = "Parkingeye"
        self.fetch_thread = None
        self.analysis_thread = None

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
        self.address_input = QLineEdit()
        self.address_input.textChanged.connect(self.update_preview)
        
        self.date_input = QLineEdit()
        self.date_input.setPlaceholderText("dd/mm/yyyy")
        self.date_input.textChanged.connect(self.update_preview)

        form_layout.addWidget(QLabel("Site Contact Name:"))
        form_layout.addWidget(self.name_input)
        form_layout.addWidget(QLabel("Site Contact Email:"))
        form_layout.addWidget(self.email_input)
        form_layout.addWidget(QLabel("Site Address:"))
        form_layout.addWidget(self.address_input)
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
        self.fetch_thread = FetchEmailsThread(self)
        self.fetch_thread.finished.connect(self.on_emails_fetched)
        self.fetch_thread.error.connect(lambda err: QMessageBox.critical(self, "Error", err))
        self.fetch_thread.start()

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
        self.preview_box.setText("<i>Analyzing email contents...</i>")
        
        # Start background analysis
        self.analysis_thread = AnalyzeEmailThread(self.current_email, self)
        self.analysis_thread.finished.connect(self.on_analysis_finished)
        self.analysis_thread.error.connect(lambda err: QMessageBox.warning(self, "Analysis Error", err))
        self.analysis_thread.start()

    def on_analysis_finished(self, results):
        self.client_label = results["client"]
        candidates = results["candidates"]
        self.address_input.setText(results.get("address") or "")

        if candidates:
            self.name_input.setText(candidates[0].get('name') or "")
            self.email_input.setText(candidates[0].get('email') or "")
            
        self.update_preview()

    def update_preview(self):
        if not self.current_email: return
        
        name = self.name_input.text() or "[Name]"
        address = self.address_input.text() or "[Site Address]"
        date_str = self.date_input.text()
        
        try:
            day_name, formatted_date = email_logic.parse_visit_date(date_str)
        except:
            day_name, formatted_date = "[Day]", "[Date]"

        body = email_logic.build_email_body(name, self.client_label, address, day_name, formatted_date)
        self.preview_box.setText(body)

    def open_draft(self):
        if not self.current_email: return
        
        name = self.name_input.text()
        to_email = self.email_input.text()
        address = self.address_input.text()
        date_str = self.date_input.text()
        
        try:
            day_name, formatted_date = email_logic.parse_visit_date(date_str)
            html_body = email_logic.build_email_body_html(name, self.client_label, address, day_name, formatted_date)
            subject = f"Asbestos Survey Booking - {formatted_date}"
            
            email_logic.open_new_draft(to_email, subject, html_body)
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