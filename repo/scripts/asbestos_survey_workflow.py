"""
Asbestos Survey Automation Workflow
- Reads emails from Lewis Dunkley
- Extracts PDFs and address
- Prints documents
- Duplicates Alpha Tracker job
- Updates job with new details
- Sends site contact email
"""
# git test
import os
import sys
from pathlib import Path
from datetime import datetime
import json
import re
import requests
import win32com.client
from PIL import Image
import pytesseract
import pdfplumber

# ============================================================================
# CONFIG & CREDENTIALS
# ============================================================================

CREDENTIALS_FILE = Path(r"c:\Users\Sherren\Desktop\feldman\scripts\credentials.txt")
TEMPLATE_JOBS = {
    "parkingeye": "G-25562",
    "g24": "G-25564"
}

def load_credentials():
    """Load API credentials from file."""
    creds = {}
    with open(CREDENTIALS_FILE, 'r') as f:
        for line in f:
            if '=' in line:
                key, value = line.strip().split('=', 1)
                creds[key.strip()] = value.strip().strip('"')
    return creds

CREDS = load_credentials()
API_URL = CREDS.get("API_URL", "https://manager.alphatracker.co.uk/api")
API_KEY = CREDS.get("API_KEY")
CLIENT_ID = CREDS.get("CLIENT_ID")

API_HEADERS = {
    "x-api-key": API_KEY,
    "client-id": CLIENT_ID,
    "Content-Type": "application/json"
}

# ============================================================================
# EMAIL READING
# ============================================================================

def get_outlook_emails_from_lewis(account_name=None):
    """Fetch emails from Lewis Dunkley using folder navigation."""
    print("\n[STEP 1] Fetching emails from Lewis Dunkley...")
    
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        
        # Get all available accounts
        accounts = namespace.Accounts
        print(f"  Available accounts: {[acc.DisplayName for acc in accounts]}")
        
        # Find the a.smith account
        a_smith_account = None
        for acc in accounts:
            if "a.smith" in acc.DisplayName.lower():
                a_smith_account = acc
                break
        
        if not a_smith_account:
            a_smith_account = accounts[1] if len(accounts) > 1 else accounts[0]
        
        print(f"  Using account: {a_smith_account.DisplayName}")
        
        # Get the root store folder for this account
        root_folder = namespace.Folders.Item(a_smith_account.DisplayName)
        print(f"  Root folder: {root_folder.Name}")
        
        # Find Inbox subfolder
        inbox = None
        for folder in root_folder.Folders:
            if folder.Name.lower() == "inbox":
                inbox = folder
                break
        
        if not inbox:
            print(f"  [ERROR] Could not find Inbox folder")
            return []
        
        print(f"  Inbox found: {inbox.Items.Count} emails")
        
        # Search for Lewis emails
        emails = []
        search_terms = ["lewis", "dunkley", "l.dunkley"]
        target_email = "l.dunkley@greenshieldenvironmental.co.uk"
        
        print(f"  Searching through all {inbox.Items.Count} emails...")
        item_count = 0
        
        for item in inbox.Items:
            item_count += 1
            if item_count % 100 == 0:
                print(f"    ... processed {item_count} emails...")
            
            try:
                sender_name_lower = str(item.SenderName).lower()
                sender_email = str(item.SenderEmailAddress).lower() if hasattr(item, 'SenderEmailAddress') else ""
                
                # Check if matches Lewis OR matches the specific email
                is_lewis = any(term in sender_name_lower for term in search_terms)
                is_target_email = target_email.lower() in sender_email
                
                if is_lewis or is_target_email:
                    emails.append({
                        "sender": item.SenderName,
                        "sender_email": sender_email,
                        "subject": item.Subject,
                        "body": item.Body,
                        "received_time": item.ReceivedTime,
                        "attachments": item.Attachments,
                        "message_id": item.EntryID
                    })
            except:
                pass
        
        print(f"  [OK] Found {len(emails)} email(s) from Lewis/Dunkley")
        return emails
    
    except Exception as e:
        print(f"  [ERROR] Error reading Outlook: {e}")
        import traceback
        traceback.print_exc()
        return []

def list_outlook_folders():
    """List all available folders in Outlook."""
    print("\n[DEBUG] Listing all Outlook folders...")
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        root_folder = namespace.Folders.Item(1)
        
        def print_folders(folder, indent=0):
            prefix = "  " * indent
            try:
                print(f"{prefix}- {folder.Name} ({folder.Items.Count} items)")
                if hasattr(folder, 'Folders'):
                    for subfolder in folder.Folders:
                        print_folders(subfolder, indent + 1)
            except:
                pass
        
        print_folders(root_folder)
    except Exception as e:
        print(f"  [ERROR] {e}")

def extract_attachments(email, output_dir=None):
    """Extract PDF attachments from email."""
    if output_dir is None:
        output_dir = Path.home() / "Downloads" / "survey_attachments"
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    
    print(f"  Extracting attachments from: {email['subject']}...")
    
    extracted_files = []
    for attachment in email['attachments']:
        if attachment.Filename.lower().endswith('.pdf'):
            filepath = output_dir / attachment.Filename
            attachment.SaveAsFile(str(filepath))
            extracted_files.append(filepath)
            print(f"    [OK] Saved: {attachment.Filename}")
    
    return extracted_files

# ============================================================================
# PDF EXTRACTION
# ============================================================================

def extract_address_from_pdf(pdf_path):
    """Extract address from survey PDF."""
    print(f"  Extracting address from: {pdf_path.name}...")
    
    try:
        with pdfplumber.open(pdf_path) as pdf:
            text = ""
            for page in pdf.pages:
                text += page.extract_text() or ""
        
        # Simple regex patterns for common UK address formats
        # Adjust as needed for your PDF structure
        lines = text.split('\n')
        address_candidates = [line.strip() for line in lines if line.strip() and len(line) > 10]
        
        # For now, return first plausible multi-line address
        if address_candidates:
            address = ", ".join(address_candidates[:3])
            print(f"    [OK] Extracted address: {address[:100]}...")
            return address
        return None
    
    except Exception as e:
        print(f"    [ERROR] Error extracting address: {e}")
        return None

def extract_po_number(email_body):
    """Extract PO number from email body (format: PO-xxxxxx)."""
    print("  Extracting PO number from email...")
    
    match = re.search(r'PO-(\d+)', email_body, re.IGNORECASE)
    if match:
        po_number = match.group(1)
        print(f"    [OK] Found PO: {po_number}")
        return po_number
    
    print("    [!] No PO number found")
    return None

def extract_site_contact_from_image(pdf_path):
    """Extract site contact details from 'SITE CONTACT DETAILS' image block."""
    print(f"  Extracting site contact from: {pdf_path.name}...")
    
    try:
        with pdfplumber.open(pdf_path) as pdf:
            # Try to find images in PDF
            for page_num, page in enumerate(pdf.pages):
                if page.images:
                    # Extract images from page
                    for img_info in page.images:
                        # Note: pdfplumber doesn't directly extract images
                        # You may need additional library like pdf2image
                        pass
        
        # Placeholder: would need pdf2image + pytesseract
        print("    ! Site contact extraction requires additional setup (pdf2image)")
        return None, None  # (email, name)
    
    except Exception as e:
        print(f"    [ERROR] Error extracting site contact: {e}")
        return None, None

def detect_job_type(pdf_paths, fallback_subject=None):
    """Detect if job is ParkingEye or G24 based on PDF content or subject."""
    print("  Detecting job type from PDFs...")
    
    for pdf in pdf_paths:
        try:
            with pdfplumber.open(pdf) as pdf_file:
                text = " ".join([page.extract_text() or "" for page in pdf_file.pages])
                
                if "parkingeye" in text.lower():
                    print(f"    [OK] Detected: ParkingEye")
                    return "parkingeye"
                elif "g24" in text.lower() or "G24" in text:
                    print(f"    [OK] Detected: G24")
                    return "g24"
        except:
            pass
    
    # Fallback to subject line detection
    if fallback_subject:
        if "parkingeye" in fallback_subject.lower():
            print(f"    [OK] Detected from subject: ParkingEye")
            return "parkingeye"
        elif "g24" in fallback_subject.lower():
            print(f"    [OK] Detected from subject: G24")
            return "g24"
    
    print("    ? Could not autodetect job type")
    return None

# ============================================================================
# PRINTING
# ============================================================================

def print_pdfs(pdf_paths):
    """Print PDF attachments."""
    print(f"\n[STEP 2] Printing {len(pdf_paths)} PDF(s)...")
    
    try:
        import subprocess
        for pdf in pdf_paths:
            print(f"  Sending to printer: {pdf.name}")
            try:
                subprocess.run(['print', str(pdf)], check=True, capture_output=True)
            except FileNotFoundError:
                # print command might not exist, try alternative
                print(f"    ! Print command not available, trying via Explorer...")
                os.startfile(str(pdf), 'print')
        print("[OK] PDFs sent to printer")
        return True
    except Exception as e:
        print(f"[ERROR] Error printing: {e}")
        return False

def print_sent_email_to_lewis(recipient_email):
    """Print the email that was sent TO Lewis (from Sent folder)."""
    print(f"\n[STEP 2b] Finding email sent to Lewis at {recipient_email}...")
    
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        sent_folder = namespace.GetDefaultFolder(5)  # 5 = Sent Items
        
        for item in sent_folder.Items:
            try:
                if recipient_email and recipient_email.lower() in str(item.To).lower():
                    print(f"  [OK] Found sent email: {item.Subject}")
                    # Print logic would go here (save as PDF then print)
                    print("  ! Email printing not yet implemented")
                    return True
            except:
                pass
        
        print(f"  [!] No sent email found to {recipient_email}")
        return False
    
    except Exception as e:
        print(f"  [ERROR] Error finding sent email: {e}")
        return False

# ============================================================================
# ALPHA TRACKER API
# ============================================================================

def get_project(project_number):
    """Fetch project details from Alpha Tracker."""
    print(f"\n  Fetching project {project_number} from Alpha Tracker...")
    
    try:
        resp = requests.get(
            f"{API_URL}/projects/{project_number}",
            headers=API_HEADERS,
            timeout=10
        )
        resp.raise_for_status()
        data = resp.json()
        print(f"    [OK] Project found: {data.get('projectType', 'unknown')}")
        return data
    except Exception as e:
        print(f"    [ERROR] Error fetching project: {e}")
        return None

def create_project_from_template(template_project_data, address, po_number, site_contact_email, site_contact_name):
    """
    Create a new project by duplicating template + inserting new data.
    NOTE: In beta mode, this will only show preview data.
    """
    print(f"\n[STEP 4] Preparing new project (PREVIEW MODE)...")
    
    # Extract key fields from template
    new_project = {
        "projectLetter": template_project_data.get("projectLetter", "A"),
        "clientId": template_project_data.get("clientId"),
        "clientOrderNumber": po_number,  # Insert PO number here (without "PO-" prefix)
        "clientProjectRef": template_project_data.get("clientProjectRef"),
        "siteId": template_project_data.get("siteId"),
        "reportRecipientName1": site_contact_name,
        "reportRecipientEmailAddress1": site_contact_email,
        "status": template_project_data.get("status", "New"),
        "projectTypeId": template_project_data.get("projectTypeId"),
    }
    
    print(f"    New project data prepared (not created yet):")
    print(f"      - Address: {address[:50]}...")
    print(f"      - PO Number: {po_number}")
    print(f"      - Site Contact: {site_contact_email}")
    
    return new_project

def update_project(project_number, address, po_number, site_contact_email, site_contact_name):
    """Update project with extracted data."""
    print(f"\n  Updating project {project_number} with new data...")
    
    payload = {
        "clientOrderNumber": po_number,
        "reportRecipientName1": site_contact_email,
        "reportRecipientEmailAddress1": site_contact_email,
    }
    
    try:
        resp = requests.patch(
            f"{API_URL}/projects/{project_number}",
            headers=API_HEADERS,
            json=payload,
            timeout=10
        )
        resp.raise_for_status()
        print(f"    [OK] Project updated")
        return True
    except Exception as e:
        print(f"    [ERROR] Error updating project: {e}")
        return False

def generate_desktop_study(project_number):
    """Generate desktop study report (placeholder - check if API supports this)."""
    print(f"\n[STEP 5] Checking if desktop study can be generated...")
    
    # The API docs show a /reports endpoint but may not have desktop study
    # This is a placeholder for now
    print(f"    ! Desktop study generation needs manual confirmation or web UI")
    return None

# ============================================================================
# EMAIL SENDING
# ============================================================================

def send_email_to_site_contact(recipient_email, recipient_name, job_number, address):
    """Send dynamic email to site contact with job details."""
    print(f"\n[STEP 6] Preparing email to site contact: {recipient_email}...")
    
    # TODO: Add dynamic template based on day of week
    # For now, placeholder template
    day_of_week = datetime.now().strftime("%A")
    
    subject = f"Asbestos Survey - Job {job_number}"
    
    body = f"""
Dear {recipient_name},

Your asbestos survey has been scheduled.

Job Number: {job_number}
Address: {address}
Day: {day_of_week}

[DYNAMIC CONTENT BASED ON {day_of_week}]

Please let us know if you have any questions.

Best regards,
[Your Company]
"""
    
    print(f"    [OK] Email prepared (not sent in beta mode)")
    print(f"    Subject: {subject}")
    print(f"    To: {recipient_email}")
    
    return True

# ============================================================================
# MAIN WORKFLOW
# ============================================================================

def main():
    print("="*70)
    print("ASBESTOS SURVEY AUTOMATION - BETA MODE")
    print("="*70)
    print(f"Started: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    
    # TEST MODE: Allow manual input if no emails found
    test_mode = input("\nTest mode? Enter 'test' to use mock data, or press Enter to read Outlook: ").lower()
    
    if test_mode == 'debug':
        list_outlook_folders()
        return
    elif test_mode == 'test':
        print("\n[TEST MODE] Using mock data for development...")
        emails = [{
            "sender": "Lewis Dunkley",
            "subject": "Asbestos Survey Request - ParkingEye",
            "body": "Please quote on the attached survey. PO-12345",
            "received_time": datetime.now(),
            "attachments": [],
            "message_id": "TEST"
        }]
        # Create mock PDF files for testing
        mock_pdf = Path.home() / "Downloads" / "survey_attachments" / "test_survey.pdf"
        mock_pdf.parent.mkdir(parents=True, exist_ok=True)
        if not mock_pdf.exists():
            mock_pdf.write_text("Mock PDF - test content\nParkingEye\n123 Main Street\nLondon\nE1 1AA")
        emails[0]["attachments"] = [mock_pdf]
    else:
        # STEP 1: Get email from Lewis - try a.smith account first
        emails = get_outlook_emails_from_lewis(account_name="a.smith")
        
        # Fallback: if no Lewis email found, show all emails and let user pick
        if not emails:
            print("\n  [!] Lewis email not found. Showing recent emails to select from...")
            try:
                outlook = win32com.client.Dispatch("Outlook.Application")
                namespace = outlook.GetNamespace("MAPI")
                inbox = namespace.GetDefaultFolder(6)
                
                all_items = list(inbox.Items)
                recent_items = all_items[:10]  # Show last 10
                
                print("\n  Recent emails:")
                for i, item in enumerate(recent_items):
                    try:
                        print(f"    [{i}] From: {item.SenderName} | {item.Subject[:60]}")
                    except:
                        pass
                
                # In non-interactive mode, just use first non-system email
                try:
                    email_index = int(input("  Select email index (or press Enter to skip): ") or "-1")
                    if 0 <= email_index < len(recent_items):
                        item = recent_items[email_index]
                        emails = [{
                            "sender": item.SenderName,
                            "subject": item.Subject,
                            "body": item.Body,
                            "received_time": item.ReceivedTime,
                            "attachments": item.Attachments,
                            "message_id": item.EntryID
                        }]
                except (ValueError, EOFError):
                    pass
            except:
                pass
    
    if not emails:
        print("[ERROR] No emails found from Lewis Dunkley. Exiting.")
        return
    
    # Try each Lewis email until we find one with PDF attachments
    selected_email = None
    pdf_files = None
    
    print(f"\n[Found {len(emails)} email(s) from Lewis - searching for one with PDF attachments...]")
    
    for idx, email in enumerate(emails):
        print(f"\n  [{idx+1}/{len(emails)}] Checking: {email['subject'][:60]}")
        
        # Check for attachments
        if isinstance(email['attachments'], list) and isinstance(email['attachments'][0], Path):
            # Test mode - already has file paths
            pdf_files = email['attachments']
            selected_email = email
            break
        else:
            # Normal mode - extract from Outlook
            try:
                pdf_files = extract_attachments(email)
                if pdf_files:
                    selected_email = email
                    break
            except:
                pass
    
    if not selected_email or not pdf_files:
        print("\n[ERROR] No Lewis email with PDF attachments found. Exiting.")
        return
    
    email = selected_email
    print(f"\n[OK] Selected email: {email['subject']}")
    
    # STEP 2: Detect job type
    job_type = detect_job_type(pdf_files, fallback_subject=email['subject'])
    if not job_type:
        # In test/non-interactive mode, default to parkingeye
        job_type = "parkingeye"
        print(f"  Using default: {job_type}")
    
    template_job = TEMPLATE_JOBS.get(job_type)
    print(f"\nUsing template job: {template_job}")
    
    # STEP 3: Extract data from PDFs and email
    address = extract_address_from_pdf(pdf_files[0])
    po_number = extract_po_number(email['body'])
    site_contact_email, site_contact_name = extract_site_contact_from_image(pdf_files[0])
    
    print("\n" + "="*70)
    print("EXTRACTED DATA SUMMARY")
    print("="*70)
    print(f"Job Type: {job_type.upper()}")
    print(f"Template Job: {template_job}")
    print(f"Address: {address or '[NOT EXTRACTED]'}")
    print(f"PO Number: {po_number or '[NOT FOUND]'}")
    print(f"Site Contact Email: {site_contact_email or '[NOT EXTRACTED]'}")
    print(f"Site Contact Name: {site_contact_name or '[NOT EXTRACTED]'}")
    print("="*70)
    
    # USER CHECKPOINT
    try:
        confirm = input("\n[?] Proceed with this data? (yes/no): ").lower()
    except EOFError:
        confirm = "yes"  # Default to yes in non-interactive mode
    
    if confirm != 'yes':
        print("[CANCELLED] Workflow cancelled by user.")
        return
    
    # STEP 4: Print documents
    print_pdfs(pdf_files)
    print_sent_email_to_lewis(site_contact_email or "N/A")
    
    # STEP 5: Get template project and prepare duplication
    template_project = get_project(template_job)
    if not template_project:
        print("[ERROR] Could not load template project. Exiting.")
        return
    
    new_project_data = create_project_from_template(
        template_project,
        address,
        po_number,
        site_contact_email,
        site_contact_name
    )
    
    print("\n" + "="*70)
    print("READY TO CREATE NEW PROJECT (BETA - NOT LIVE YET)")
    print("="*70)
    print(f"Template: {template_job}")
    print(f"New Project Data:\n{json.dumps(new_project_data, indent=2)}")
    print("="*70)
    
    # USER CHECKPOINT
    try:
        confirm = input("\n[?] Create new project in Alpha Tracker? (yes/no): ").lower()
    except EOFError:
        confirm = "no"  # Default to no in non-interactive mode (stay safe)
    
    if confirm == 'yes':
        print("\n[BETA] BETA MODE: Live project creation is disabled.")
        print("   In production, the new project would be created here.")
    
    
    # STEP 6: Send email to site contact
    send_email_to_site_contact(site_contact_email, site_contact_name, template_job, address)
    
    # STEP 7: Generate desktop study
    generate_desktop_study(template_job)
    
    print("\n" + "="*70)
    print("WORKFLOW COMPLETE")
    print("="*70)

if __name__ == "__main__":
    main()
