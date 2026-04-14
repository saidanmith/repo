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
import argparse
from pathlib import Path
from datetime import datetime
import json
import re
import time
import requests
import win32com.client
import pytesseract
import pdfplumber
import win32print
import win32api
from PIL import Image
try:
    from pdf2image import convert_from_path
except ImportError:
    convert_from_path = None

# ============================================================================
# CONFIG & CREDENTIALS
# ============================================================================

# Point pytesseract at the Tesseract install on Windows
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

BASE_DIR = Path(__file__).resolve().parent
CREDENTIALS_FILE = BASE_DIR / "credentials.txt"
TEMPLATE_JOBS = {
    "parkingeye": "G-25562",
    "g24": "G-25564"
}
ADDRESS_SKIP_TERMS = {
    "signage plan",
    "install doc",
    "staff only",
    "drawing",
    "revision",
    "parkingeye",
    "greenshield",
    "asbestos survey request",
}
UK_POSTCODE_RE = re.compile(
    r"\b[A-Z]{1,2}\d[A-Z\d]?\s*\d[A-Z]{2}\b",
    re.IGNORECASE,
)
IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".bmp", ".gif", ".tif", ".tiff"}

def load_credentials():
    """Load API credentials from file."""
    creds = {}
    if not CREDENTIALS_FILE.exists():
        return creds

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
                        "html_body": getattr(item, "HTMLBody", ""),
                        "received_time": item.ReceivedTime,
                        "attachments": item.Attachments,
                        "message_id": item.EntryID,
                        "outlook_item": item,
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
        if not Path(pdf_path).exists():
            print("    [!] PDF file not found; skipping address extraction")
            return None

        with pdfplumber.open(pdf_path) as pdf:
            text = ""
            for page in pdf.pages:
                text += page.extract_text() or ""
        
        address = extract_address_from_text(text)
        if address:
            print(f"    [OK] Extracted address: {address[:100]}...")
            return address

        print("    [!] No address block with town/postcode found")
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

def clean_address_line(line):
    """Normalize a candidate address line."""
    line = re.sub(r"\s+", " ", line.replace("|", " ")).strip(" ,.-")
    return line

def is_address_like_line(line):
    """Return True for likely address lines, False for document noise."""
    lowered = line.lower()
    if not line or len(line) < 3:
        return False
    if any(term in lowered for term in ADDRESS_SKIP_TERMS):
        return False
    if "@" in line:
        return False
    if re.search(r"\bpage\s+\d+\b", lowered):
        return False
    return True

def extract_address_from_text(text):
    """Extract an address block that includes town and postcode."""
    raw_lines = [clean_address_line(line) for line in text.splitlines()]
    lines = [line for line in raw_lines if is_address_like_line(line)]

    for index, line in enumerate(lines):
        if not UK_POSTCODE_RE.search(line):
            continue

        block = []
        start = max(0, index - 3)
        for candidate in lines[start:index + 1]:
            if candidate not in block:
                block.append(candidate)

        if len(block) < 2:
            continue

        postcode_match = UK_POSTCODE_RE.search(block[-1])
        if not postcode_match:
            continue

        town_line = block[-2]
        if any(char.isdigit() for char in town_line) and not re.search(r"\broad\b|\bstreet\b|\bavenue\b|\blane\b|\bdrive\b|\bclose\b|\bway\b", town_line, re.IGNORECASE):
            continue

        return ", ".join(block)

    return None

def extract_address_from_pdfs(pdf_paths):
    """Try all PDFs until an address with town and postcode is found."""
    for pdf_path in pdf_paths:
        address = extract_address_from_pdf(pdf_path)
        if address and UK_POSTCODE_RE.search(address):
            return address
    return None



def extract_site_contact_details_from_text(text):
    """Extract site contact details from OCR or plain text."""
    normalized_text = text.replace("\r", "\n")
    upper_text = normalized_text.upper()
    if "SITE CONTACT DETAILS" not in upper_text and "SITE CONTACT" not in upper_text:
        return None, None

    names = re.findall(r"Name[:\s]+([A-Z][a-zA-Z\s\-\']+)", normalized_text)
    emails = re.findall(
        r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}",
        normalized_text,
    )

    print(f"    Names found:  {names}")
    print(f"    Emails found: {emails}")

    if names and emails:
        return emails[0].strip(), names[0].strip()

    email = emails[0].strip() if emails else None
    name = names[0].strip() if names else None
    return email, name

def normalize_contact_name(name):
    """Clean OCR noise from a contact name."""
    if not name:
        return None
    name = re.sub(r"\s+", " ", name).strip(" ,.-")
    if "activate.ps1" in name.lower():
        return None
    return name

def normalize_contact_email(email):
    """Clean OCR noise from an email address."""
    if not email:
        return None
    email = email.strip().strip(" ,.;:")
    email = email.replace(" ", "")
    email = email.replace("..", ".")
    return email.lower()

def extract_contact_candidates_from_text(text):
    """Extract one or more contact candidates from text."""
    normalized_text = text.replace("\r", "\n")
    upper_text = normalized_text.upper()
    if "SITE CONTACT DETAILS" not in upper_text and "SITE CONTACT" not in upper_text:
        return []

    candidates = []
    lines = [line.strip() for line in normalized_text.splitlines() if line.strip()]

    current_name = None
    for line in lines:
        name_match = re.search(r"Name[:\s]+(.+)", line, re.IGNORECASE)
        if name_match:
            current_name = normalize_contact_name(name_match.group(1))
            continue

        email_match = re.search(r"([a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,})", line)
        if email_match:
            email = normalize_contact_email(email_match.group(1))
            candidates.append({"name": current_name, "email": email})
            current_name = None

    if candidates:
        return candidates

    email, name = extract_site_contact_details_from_text(normalized_text)
    if email or name:
        return [{"name": normalize_contact_name(name), "email": normalize_contact_email(email)}]
    return []

def dedupe_contact_candidates(candidates):
    """Remove duplicate contact candidates while preserving order."""
    deduped = []
    seen = set()
    for candidate in candidates:
        normalized = (
            normalize_contact_name(candidate.get("name")),
            normalize_contact_email(candidate.get("email")),
        )
        if normalized in seen:
            continue
        seen.add(normalized)
        deduped.append({"name": normalized[0], "email": normalized[1]})
    return deduped

def get_attachment_mime_type(attachment):
    """Read an Outlook attachment MIME type when available."""
    try:
        return attachment.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x370E001E")
    except Exception:
        return ""

def extract_inline_images(email, output_dir=None):
    """Save inline email images for OCR."""
    if output_dir is None:
        output_dir = Path.home() / "Downloads" / "survey_attachments" / "inline_images"
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    saved_files = []
    for index, attachment in enumerate(email["attachments"], start=1):
        filename = str(getattr(attachment, "Filename", "") or "").strip()
        suffix = Path(filename).suffix.lower()
        mime_type = str(get_attachment_mime_type(attachment) or "").lower()
        if suffix not in IMAGE_EXTENSIONS and not mime_type.startswith("image/"):
            continue

        if not suffix:
            suffix = ".png"
        safe_name = filename or f"inline_image_{index}{suffix}"
        target = output_dir / safe_name
        attachment.SaveAsFile(str(target))
        saved_files.append(target)
        print(f"    [OK] Saved inline image: {target.name}")

    return saved_files

def extract_site_contact_from_email(email):
    """Extract site contact details from the email body or inline images."""
    print("  Extracting site contact from email body...")

    candidates = []
    email_text = "\n".join(
        part for part in [email.get("body", ""), email.get("html_body", "")]
        if part
    )
    candidates.extend(extract_contact_candidates_from_text(email_text))

    image_paths = extract_inline_images(email)
    if not image_paths:
        print("    [!] No inline email images found")
    else:
        for image_path in image_paths:
            try:
                text = pytesseract.image_to_string(Image.open(image_path))
            except Exception as e:
                print(f"    [ERROR] OCR failed for {image_path.name}: {e}")
                continue

            image_candidates = extract_contact_candidates_from_text(text)
            if image_candidates:
                print(f"    [OK] Found {len(image_candidates)} contact candidate(s) in {image_path.name}")
                candidates.extend(image_candidates)

    candidates = dedupe_contact_candidates(candidates)
    if not candidates:
        print("    [!] No site contact details found in email body images")
        return None, None

    print("    Contact candidates:")
    for index, candidate in enumerate(candidates, start=1):
        print(f"      [{index}] {candidate.get('name') or '[NO NAME]'} | {candidate.get('email') or '[NO EMAIL]'}")

    if len(candidates) == 1:
        chosen = candidates[0]
        print(f"    [OK] Using single contact candidate: {chosen.get('name')} | {chosen.get('email')}")
        return chosen.get("email"), chosen.get("name")

    while True:
        choice = prompt_text(
            f"  Select site contact [1-{len(candidates)}]: ",
            default="1",
        )
        if choice.isdigit():
            index = int(choice)
            if 1 <= index <= len(candidates):
                chosen = candidates[index - 1]
                print(f"    [OK] Selected: {chosen.get('name')} | {chosen.get('email')}")
                return chosen.get("email"), chosen.get("name")
        print("[INFO] Enter a valid contact number.")

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

def get_default_printer():
    """Return the Windows default printer name, if available."""
    try:
        return win32print.GetDefaultPrinter()
    except Exception as e:
        print(f"  [!] Could not read default printer: {e}")
        return None

def show_printer_status():
    """Show printer configuration before attempting PDF prints."""
    print("\n[PRINT DIAGNOSTICS]")
    default_printer = get_default_printer()
    if not default_printer:
        print("  [ERROR] No default printer detected")
        return None

    print(f"  Default printer: {default_printer}")
    try:
        handle = win32print.OpenPrinter(default_printer)
        try:
            info = win32print.GetPrinter(handle, 2)
        finally:
            win32print.ClosePrinter(handle)

        status = info.get("Status", 0)
        attributes = info.get("Attributes", 0)
        print(f"  Driver: {info.get('pDriverName')}")
        print(f"  Port: {info.get('pPortName')}")
        print(f"  Status code: {status}")
        print(f"  Attributes: {attributes}")
    except Exception as e:
        print(f"  [!] Could not inspect printer details: {e}")

    return default_printer

def print_pdfs(pdf_paths):
    """Print PDF attachments."""
    print(f"\n[STEP 2] Printing {len(pdf_paths)} PDF(s)...")

    default_printer = show_printer_status()
    if not default_printer:
        return False

    all_sent = True
    for pdf in pdf_paths:
        print(f"  Sending to printer: {pdf.name}")
        try:
            win32api.ShellExecute(
                0,
                "printto",
                str(pdf),
                f'"{default_printer}"',
                str(pdf.parent),
                0,
            )
            time.sleep(2)
            print(f"    [OK] Print command sent to {default_printer}")
        except Exception as e:
            all_sent = False
            print(f"    [ERROR] Failed to print {pdf.name}: {e}")

    return all_sent

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
        "reportRecipientName1": site_contact_name,
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

def prepare_email_to_site_contact(recipient_email, recipient_name, job_number, address):
    """Prepare a draft email preview only; do not send automatically."""
    if not recipient_email:
        print("\n[STEP 6] Skipping email draft because no site contact email was found.")
        return False

    print(f"\n[STEP 6] Preparing email draft to site contact: {recipient_email}...")
    
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
    
    print(f"    [OK] Email drafted for manual review only")
    print(f"    Subject: {subject}")
    print(f"    To: {recipient_email}")
    print("    [MANUAL ACTION] Review and send this email yourself if appropriate")

    return True

def parse_args():
    """Parse command-line arguments."""
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--mode",
        choices=["test", "live", "debug"],
        help="Run in local test mode, live Outlook mode, or debug folder-listing mode.",
    )
    parser.add_argument(
        "--email-index",
        type=int,
        help="Select a recent inbox email by index when fallback email picking is needed.",
    )
    parser.add_argument(
        "--proceed",
        choices=["yes", "no"],
        help="Answer the extracted-data confirmation prompt.",
    )
    parser.add_argument(
        "--create-project",
        choices=["yes", "no"],
        help="Answer the Alpha Tracker create-project prompt.",
    )
    return parser.parse_args()

def resolve_mode(args=None):
    """Resolve runtime mode without depending on interactive stdin."""
    args = args or parse_args()
    if args.mode:
        return args.mode

    if not sys.stdin.isatty():
        print("\n[INFO] No interactive stdin detected. Defaulting to test mode.")
        return "test"

    try:
        return input(
            "\nMode? Enter 'test' for mock data, 'debug' for Outlook folders, or press Enter for live Outlook: "
        ).strip().lower()
    except (EOFError, KeyboardInterrupt):
        print("\n[INFO] Input unavailable. Defaulting to test mode.")
        return "test"

def prompt_text(message, default="", allowed=None):
    """Prompt safely, falling back to defaults when stdin is unavailable."""
    while True:
        try:
            value = input(message).strip().lower()
        except EOFError:
            print(f"{message}{default}")
            return default
        except KeyboardInterrupt:
            print("\n[INFO] Interrupted. Waiting for input...")
            continue

        if not value:
            return default
        if allowed and value not in allowed:
            print(f"[INFO] Invalid input '{value}'. Enter one of: {', '.join(sorted(allowed))}")
            continue
        return value

def prompt_required_text(message, current_value=None):
    """Keep prompting until a non-empty value is supplied."""
    prompt_suffix = f" [{current_value}]" if current_value else ""
    while True:
        try:
            value = input(f"{message}{prompt_suffix}: ").strip()
        except KeyboardInterrupt:
            print("\n[INFO] Interrupted. Waiting for input...")
            continue
        except EOFError:
            value = current_value or ""

        if value:
            return value
        if current_value:
            return current_value
        print("[INFO] This value is required.")

# ============================================================================
# MAIN WORKFLOW
# ============================================================================

def main():
    print("="*70)
    print("ASBESTOS SURVEY AUTOMATION - BETA MODE")
    print("="*70)
    print(f"Started: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    
    args = parse_args()
    mode = args.mode or resolve_mode(args)
    is_test_mode = mode == 'test'

    if mode == 'debug':
        list_outlook_folders()
        return
    elif is_test_mode:
        print("\n[TEST MODE] Using mock data for development...")
        emails = [{
            "sender": "Lewis Dunkley",
            "subject": "Asbestos Survey Request - ParkingEye",
            "body": "Please quote on the attached survey. PO-12345",
            "received_time": datetime.now(),
            "attachments": [],
            "message_id": "TEST"
        }]
        emails[0]["attachments"] = [BASE_DIR / "test_survey.pdf"]
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
                    selected_index = args.email_index
                    if selected_index is None:
                        selected_index = int(
                            prompt_text("  Select email index (or press Enter to skip): ", default="-1")
                        )
                    email_index = selected_index
                    if 0 <= email_index < len(recent_items):
                        item = recent_items[email_index]
                        emails = [{
                            "sender": item.SenderName,
                            "subject": item.Subject,
                            "body": item.Body,
                            "html_body": getattr(item, "HTMLBody", ""),
                            "received_time": item.ReceivedTime,
                            "attachments": item.Attachments,
                            "message_id": item.EntryID,
                            "outlook_item": item,
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
        if (
            isinstance(email['attachments'], list)
            and email['attachments']
            and isinstance(email['attachments'][0], Path)
        ):
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
    address = extract_address_from_pdfs(pdf_files)
    po_number = extract_po_number(email['body'])
    site_contact_email, site_contact_name = extract_site_contact_from_email(email)

    if email.get("message_id") == "TEST":
        address = address or "123 Main Street, London, E1 1AA"
        site_contact_email = site_contact_email or "site.contact@example.com"
        site_contact_name = site_contact_name or "Test Contact"

    if not address or not UK_POSTCODE_RE.search(address):
        print("\n[MANUAL CHECK] Address extraction needs help.")
        address = prompt_required_text("Enter full site address including town and postcode", address)

    if not site_contact_name:
        print("\n[MANUAL CHECK] Site contact name is required.")
        site_contact_name = prompt_required_text("Enter site contact name", site_contact_name)

    if not site_contact_email:
        print("\n[MANUAL CHECK] Site contact email is required.")
        site_contact_email = prompt_required_text("Enter site contact email", site_contact_email)
    
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
    confirm = args.proceed or prompt_text(
        "\n[?] Proceed with this data? (yes/no): ",
        default="yes",
        allowed={"yes", "no"},
    )
    
    if confirm != 'yes':
        print("[CANCELLED] Workflow cancelled by user.")
        return
    
    # STEP 4: Print documents
    if is_test_mode:
        print("\n[TEST MODE] Skipping printing and Sent Items lookup.")
    else:
        print_confirm = prompt_text(
            "\n[?] Print the PDF attachments now? (yes/no): ",
            default="no",
            allowed={"yes", "no"},
        )
        if print_confirm == "yes":
            print_pdfs(pdf_files)
        else:
            print("[INFO] Printing skipped by user.")
        print_sent_email_to_lewis(site_contact_email or "N/A")
    
    # STEP 5: Get template project and prepare duplication
    if is_test_mode:
        template_project = {
            "projectLetter": "G",
            "clientId": 1,
            "clientProjectRef": "TEST-REF",
            "siteId": 1,
            "status": "New",
            "projectTypeId": 1,
        }
    else:
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
    confirm = args.create_project or prompt_text(
        "\n[?] Create new project in Alpha Tracker? (yes/no): ",
        default="no",
        allowed={"yes", "no"},
    )
    
    if confirm == 'yes':
        print("\n[BETA] BETA MODE: Live project creation is disabled.")
        print("   In production, the new project would be created here.")
    
    
    # STEP 6: Prepare email draft only
    prepare_email_to_site_contact(site_contact_email, site_contact_name, template_job, address)
    
    # STEP 7: Generate desktop study
    generate_desktop_study(template_job)
    
    print("\n" + "="*70)
    print("WORKFLOW COMPLETE")
    print("="*70)

if __name__ == "__main__":
    main()
