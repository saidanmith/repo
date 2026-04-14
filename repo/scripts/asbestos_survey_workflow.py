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
EMAIL_SUBJECT_FILTER = "asbestos survey request"
BOT_DIR = Path(r"C:\Users\Sherren\Desktop\lewis\parkingeye bot")
TEMP_DIR = BOT_DIR / "temp"
LOG_DIR = BOT_DIR / "log"
TARGET_PRINTER_NAME = "Microsoft PCL6 Class Driver"
ADDRESS_SKIP_TERMS = {
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

def ensure_runtime_dirs():
    """Create runtime directories used for temp files and logs."""
    TEMP_DIR.mkdir(parents=True, exist_ok=True)
    LOG_DIR.mkdir(parents=True, exist_ok=True)

def append_run_note(run_notes, message):
    """Record a run note for the end-of-run summary."""
    print(message)
    run_notes.append(message)

def split_address_and_postcode(address):
    """Split full address into address body and postcode."""
    if not address:
        return None, None
    match = UK_POSTCODE_RE.search(address)
    if not match:
        return address, None
    postcode = match.group(0).upper().strip()
    address_without_postcode = address[:match.start()].strip(" ,")
    return address_without_postcode, postcode

def sanitize_address_for_tracker(address):
    """Remove punctuation for Alpha Tracker address fields."""
    if not address:
        return address
    cleaned = re.sub(r"[^\w\s]", " ", address)
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    return cleaned

def split_site_name_and_address(address_line):
    """Split the first line from the remaining address lines."""
    if not address_line:
        return None, None

    parts = [part.strip() for part in address_line.split(",") if part.strip()]
    if not parts:
        return None, None

    site_name = sanitize_address_for_tracker(parts[0])
    site_address = sanitize_address_for_tracker(" ".join(parts[1:])) if len(parts) > 1 else None
    return site_name, site_address

def write_run_summary(run_notes):
    """Persist a short run summary to disk."""
    ensure_runtime_dirs()
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_path = LOG_DIR / f"run_{timestamp}.log"
    log_path.write_text("\n".join(run_notes), encoding="utf-8")
    print(f"[LOG] Summary saved to {log_path}")

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
    """Fetch asbestos survey request emails from Lewis Dunkley."""
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
        
        # Search for Lewis emails with asbestos survey request subject
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
                
                subject = str(getattr(item, "Subject", "") or "")
                is_asbestos_request = EMAIL_SUBJECT_FILTER in subject.lower()

                if (is_lewis or is_target_email) and is_asbestos_request:
                    emails.append({
                        "sender": item.SenderName,
                        "sender_email": sender_email,
                        "subject": subject,
                        "body": item.Body,
                        "html_body": getattr(item, "HTMLBody", ""),
                        "received_time": item.ReceivedTime,
                        "attachments": item.Attachments,
                        "message_id": item.EntryID,
                        "outlook_item": item,
                    })
            except:
                pass
        
        emails.sort(key=lambda email: email["received_time"], reverse=True)
        print(f"  [OK] Found {len(emails)} asbestos survey request email(s) from Lewis/Dunkley")
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
        output_dir = TEMP_DIR / "pdfs"
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
            page_texts = [page.extract_text() or "" for page in pdf.pages]

        if "signage plan" in pdf_path.name.lower() and page_texts:
            address = extract_address_from_text(page_texts[0])
            if address:
                print(f"    [OK] Extracted address from signage plan first page: {address[:100]}...")
                return address

        text = "\n".join(page_texts)
        
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
    if lowered == "signage plan":
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

        block = [candidate for candidate in block if not re.fullmatch(r"car park\s*\d+", candidate, re.IGNORECASE)]
        if len(block) < 2:
            continue

        town_line = block[-2]
        if any(char.isdigit() for char in town_line) and not re.search(r"\broad\b|\bstreet\b|\bavenue\b|\blane\b|\bdrive\b|\bclose\b|\bway\b", town_line, re.IGNORECASE):
            continue

        return ", ".join(block)

    return None

def prioritize_address_pdfs(pdf_paths):
    """Prefer PDFs whose filenames contain 'signage plan'."""
    def sort_key(pdf_path):
        name = pdf_path.name.lower()
        return (0 if "signage plan" in name else 1, name)

    return sorted(pdf_paths, key=sort_key)

def extract_address_from_pdfs(pdf_paths):
    """Try all PDFs until an address with town and postcode is found."""
    for pdf_path in prioritize_address_pdfs(pdf_paths):
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
    name = re.sub(r"[^A-Za-z\s]", " ", name)
    name = re.sub(r"\s+", " ", name).strip()
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

def prompt_contact_candidates(candidates):
    """Let the user confirm which contact candidates to use."""
    if not candidates:
        return []

    print("    Contact candidates:")
    for index, candidate in enumerate(candidates, start=1):
        print(f"      [{index}] {candidate.get('name') or '[NO NAME]'} | {candidate.get('email') or '[NO EMAIL]'}")

    if len(candidates) == 1:
        confirm = prompt_text(
            "  Use this contact? (yes/no): ",
            default="yes",
            allowed={"yes", "no"},
        )
        return candidates if confirm == "yes" else []

    confirm = prompt_text(
        "  Use all detected contacts in the draft email? (yes/no): ",
        default="yes",
        allowed={"yes", "no"},
    )
    return candidates if confirm == "yes" else []

def get_attachment_mime_type(attachment):
    """Read an Outlook attachment MIME type when available."""
    try:
        return attachment.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x370E001E")
    except Exception:
        return ""

def extract_inline_images(email, output_dir=None):
    """Save inline email images for OCR."""
    if output_dir is None:
        output_dir = TEMP_DIR / "inline_images"
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
    """Extract site contact details from inline images in the email body."""
    print("  Extracting site contact from email body...")

    candidates = []
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
        return []

    selected = prompt_contact_candidates(candidates)
    if selected:
        return selected

    return []

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

def open_pdfs_for_manual_print(pdf_paths):
    """Open PDFs so the user can print them manually."""
    print("  Opening PDFs for manual printing...")
    for pdf in pdf_paths:
        try:
            os.startfile(str(pdf))
            print(f"    [OK] Opened: {pdf.name}")
        except Exception as e:
            print(f"    [ERROR] Could not open {pdf.name}: {e}")

def render_first_pages_for_printing(pdf_paths):
    """Render the first page of each PDF to a temporary PNG for test printing."""
    if convert_from_path is None:
        print("  [ERROR] pdf2image is not available; cannot render first pages for test printing.")
        return []

    output_dir = TEMP_DIR / "first_page_prints"
    output_dir.mkdir(parents=True, exist_ok=True)
    rendered_files = []

    for pdf in pdf_paths:
        try:
            pages = convert_from_path(str(pdf), dpi=200, first_page=1, last_page=1)
            if not pages:
                continue
            output_path = output_dir / f"{pdf.stem}_page1.png"
            pages[0].save(output_path, "PNG")
            rendered_files.append(output_path)
            print(f"  [OK] Rendered first page for test print: {output_path.name}")
        except Exception as e:
            print(f"  [ERROR] Could not render first page of {pdf.name}: {e}")

    return rendered_files

def print_pdfs(pdf_paths, first_page_only=False):
    """Print PDF attachments."""
    print(f"\n[STEP 2] Printing {len(pdf_paths)} PDF(s)...")

    files_to_print = pdf_paths
    if first_page_only:
        files_to_print = render_first_pages_for_printing(pdf_paths)
        if not files_to_print:
            open_pdfs_for_manual_print(pdf_paths)
            return False

    default_printer = show_printer_status()
    if not default_printer:
        open_pdfs_for_manual_print(files_to_print)
        return False

    if TARGET_PRINTER_NAME.lower() not in default_printer.lower():
        confirm_printer = prompt_text(
            f"  Default printer is '{default_printer}'. Continue with this printer? (yes/no): ",
            default="no",
            allowed={"yes", "no"},
        )
        if confirm_printer != "yes":
            open_pdfs_for_manual_print(files_to_print)
            return False
    else:
        confirm_printer = prompt_text(
            f"  Confirm printer '{default_printer}'? (yes/no): ",
            default="yes",
            allowed={"yes", "no"},
        )
        if confirm_printer != "yes":
            open_pdfs_for_manual_print(files_to_print)
            return False

    all_sent = True
    for pdf in files_to_print:
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

    printed = prompt_text(
        "  Did the pages actually print? (yes/no): ",
        default="no",
        allowed={"yes", "no"},
    )
    if printed != "yes":
        open_pdfs_for_manual_print(files_to_print)
        return False

    return all_sent

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

def build_site_preview(template_project_data, address_line, postcode, contacts):
    """Build a preview payload for Alpha Tracker site creation."""
    primary_contact = contacts[0] if contacts else {}
    site_name, site_address = split_site_name_and_address(address_line)
    site_payload = {
        "clientId": template_project_data.get("clientId"),
        "siteName": site_name,
        "siteAddress": site_address,
        "sitePostcode": postcode,
        "siteReference": None,
        "siteContactName": primary_contact.get("name"),
        "siteContactTelephone": None,
        "siteContactEmail": primary_contact.get("email"),
        "landlord": None,
    }
    return site_payload

def create_project_from_template(template_project_data, address_line, postcode, po_number, contacts):
    """
    Create a new project by duplicating template + inserting new data.
    NOTE: In beta mode, this will only show preview data.
    """
    print(f"\n[STEP 4] Preparing new project (PREVIEW MODE)...")

    primary_contact = contacts[0] if contacts else {}
    
    # Extract key fields from template
    new_project = {
        "projectLetter": template_project_data.get("projectLetter", "A"),
        "clientId": template_project_data.get("clientId"),
        "clientOrderNumber": po_number,  # Insert PO number here (without "PO-" prefix)
        "clientProjectRef": template_project_data.get("clientProjectRef"),
        "siteId": template_project_data.get("siteId"),
        "reportRecipientName1": primary_contact.get("name"),
        "reportRecipientEmailAddress1": primary_contact.get("email"),
        "status": template_project_data.get("status", "New"),
        "projectTypeId": template_project_data.get("projectTypeId"),
    }
    
    print(f"    New project data prepared (not created yet):")
    print(f"      - Address: {address_line[:50]}..." if address_line else "      - Address: [NONE]")
    print(f"      - Postcode: {postcode}")
    print(f"      - PO Number: {po_number}")
    print(f"      - Site Contact: {primary_contact.get('email')}")
    
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

def prepare_email_to_site_contact(contacts, job_number, address):
    """Create an Outlook draft only; do not send automatically."""
    valid_contacts = [contact for contact in contacts if contact.get("email")]
    if not valid_contacts:
        print("\n[STEP 6] Skipping email draft because no site contact email was found.")
        return False

    primary_contact = valid_contacts[0]
    cc_contacts = valid_contacts[1:]
    recipient_email = primary_contact["email"]
    recipient_name = primary_contact.get("name") or "Site Contact"

    print(f"\n[STEP 6] Preparing Outlook draft to site contact: {recipient_email}...")
    
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
    
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        draft = outlook.CreateItem(0)
        draft.Subject = subject
        draft.Body = body
        draft.To = recipient_email
        draft.CC = "; ".join(contact["email"] for contact in cc_contacts if contact.get("email"))
        draft.Save()
        draft.Display()
        print(f"    [OK] Outlook draft opened for manual review")
        print(f"    Subject: {subject}")
        print(f"    To: {draft.To}")
        print(f"    CC: {draft.CC}")
        print("    [MANUAL ACTION] Review and send this email yourself if appropriate")
        return True
    except Exception as e:
        print(f"    [ERROR] Could not create Outlook draft: {e}")
        return False

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
    """Resolve runtime mode and keep waiting for manual input."""
    args = args or parse_args()
    if args.mode:
        return args.mode

    while True:
        try:
            return input(
                "\nMode? Enter 'test' for mock data, 'debug' for Outlook folders, or press Enter for live Outlook: "
            ).strip().lower()
        except KeyboardInterrupt:
            print("\n[INFO] Interrupted. Waiting for input...")
            continue
        except EOFError:
            print("\n[INFO] Input unavailable. Waiting for manual input...")
            time.sleep(1)
            continue

def prompt_text(message, default="", allowed=None):
    """Prompt and keep waiting until input is available."""
    while True:
        try:
            value = input(message).strip().lower()
        except KeyboardInterrupt:
            print("\n[INFO] Interrupted. Waiting for input...")
            continue
        except EOFError:
            print("\n[INFO] Input unavailable. Waiting for manual input...")
            time.sleep(1)
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
            print("\n[INFO] Input unavailable. Waiting for manual input...")
            time.sleep(1)
            continue

        if value:
            return value
        if current_value:
            return current_value
        print("[INFO] This value is required.")

# ============================================================================
# MAIN WORKFLOW
# ============================================================================

def main():
    run_notes = []
    ensure_runtime_dirs()

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
            "html_body": "",
            "received_time": datetime.now(),
            "attachments": [],
            "message_id": "TEST",
        }]
        emails[0]["attachments"] = [BASE_DIR / "test_survey.pdf"]
    else:
        emails = get_outlook_emails_from_lewis(account_name="a.smith")
    
    if not emails:
        print("[ERROR] No emails found from Lewis Dunkley. Exiting.")
        write_run_summary(["No matching asbestos survey request emails found."])
        return
    
    # Try each Lewis email until we find one with PDF attachments
    selected_email = None
    pdf_files = None
    
    print(f"\n[Found {len(emails)} matching email(s) from Lewis - searching for the most recent one with PDF attachments...]")
    
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
        write_run_summary(["No matching email with PDF attachments found."])
        return
    
    email = selected_email
    print(f"\n[OK] Selected email: {email['subject']}")
    append_run_note(run_notes, f"Selected email: {email['subject']}")

    email_confirm = prompt_text(
        "\n[?] Continue with this email? (yes/no): ",
        default="yes",
        allowed={"yes", "no"},
    )
    if email_confirm != "yes":
        append_run_note(run_notes, "Stopped after email selection.")
        write_run_summary(run_notes)
        return
    
    # STEP 2: Detect job type
    job_type = detect_job_type(pdf_files, fallback_subject=email['subject'])
    if not job_type:
        job_type = prompt_text(
            "  Could not detect job type. Enter 'parkingeye' or 'g24': ",
            default="parkingeye",
            allowed={"parkingeye", "g24"},
        )
    
    template_job = TEMPLATE_JOBS.get(job_type)
    print(f"\nUsing template job: {template_job}")
    append_run_note(run_notes, f"Job type: {job_type}")
    
    # STEP 3: Extract data from PDFs and email
    address = extract_address_from_pdfs(pdf_files)
    po_number = extract_po_number(email['body'])
    contacts = extract_site_contact_from_email(email)

    if email.get("message_id") == "TEST":
        address = address or "123 Main Street, London, E1 1AA"
        contacts = contacts or [{"email": "site.contact@example.com", "name": "Test Contact"}]

    if not address or not UK_POSTCODE_RE.search(address):
        print("\n[MANUAL CHECK] Address extraction needs help.")
        address = prompt_required_text("Enter full site address including town and postcode", address)

    address_confirm = prompt_text(
        f"\n[?] Confirm extracted address '{address}'? (yes/no): ",
        default="yes",
        allowed={"yes", "no"},
    )
    if address_confirm != "yes":
        address = prompt_required_text("Enter corrected full site address including town and postcode", address)

    if not contacts:
        print("\n[MANUAL CHECK] Site contact details are required.")
        manual_name = prompt_required_text("Enter site contact name")
        manual_email = prompt_required_text("Enter site contact email")
        contacts = [{"name": manual_name, "email": manual_email}]
    else:
        for index, contact in enumerate(contacts, start=1):
            if not contact.get("name"):
                contact["name"] = prompt_required_text(f"Enter name for site contact {index}")
            if not contact.get("email"):
                contact["email"] = prompt_required_text(f"Enter email for site contact {index}")

    address_line, postcode = split_address_and_postcode(address)
    tracker_address = sanitize_address_for_tracker(address_line)
    tracker_postcode = sanitize_address_for_tracker(postcode) if postcode else None
    
    print("\n" + "="*70)
    print("EXTRACTED DATA SUMMARY")
    print("="*70)
    print(f"Job Type: {job_type.upper()}")
    print(f"Template Job: {template_job}")
    print(f"Address: {address or '[NOT EXTRACTED]'}")
    print(f"Tracker Address: {tracker_address or '[NOT EXTRACTED]'}")
    print(f"Tracker Postcode: {tracker_postcode or '[NOT EXTRACTED]'}")
    print(f"PO Number: {po_number or '[NOT FOUND]'}")
    for index, contact in enumerate(contacts, start=1):
        print(f"Site Contact {index}: {contact.get('name') or '[NOT EXTRACTED]'} | {contact.get('email') or '[NOT EXTRACTED]'}")
    print("="*70)
    append_run_note(run_notes, f"Address: {address}")
    append_run_note(run_notes, f"Contacts: {json.dumps(contacts)}")
    
    # USER CHECKPOINT
    confirm = args.proceed or prompt_text(
        "\n[?] Proceed with this data? (yes/no): ",
        default="yes",
        allowed={"yes", "no"},
    )
    
    if confirm != 'yes':
        print("[CANCELLED] Workflow cancelled by user.")
        append_run_note(run_notes, "Cancelled at extraction summary.")
        write_run_summary(run_notes)
        return
    
    # STEP 4: Print documents
    if is_test_mode:
        print_confirm = prompt_text(
            "\n[?] Test print the first page of each attached PDF? (yes/no): ",
            default="no",
            allowed={"yes", "no"},
        )
        if print_confirm == "yes":
            print_result = print_pdfs(pdf_files, first_page_only=True)
            append_run_note(run_notes, f"Test print result: {'success' if print_result else 'manual fallback or failed'}")
        else:
            append_run_note(run_notes, "Test printing skipped by user.")
    else:
        print_confirm = prompt_text(
            "\n[?] Print the PDF attachments now? (yes/no): ",
            default="no",
            allowed={"yes", "no"},
        )
        if print_confirm == "yes":
            print_result = print_pdfs(pdf_files)
            append_run_note(run_notes, f"Print result: {'success' if print_result else 'manual fallback or failed'}")
        else:
            print("[INFO] Printing skipped by user.")
            append_run_note(run_notes, "Printing skipped by user.")
    
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
        append_run_note(run_notes, f"Failed to load template project {template_job}.")
        write_run_summary(run_notes)
        return
    
    new_project_data = create_project_from_template(
        template_project,
        tracker_address,
        tracker_postcode,
        po_number,
        contacts,
    )
    site_preview = build_site_preview(template_project, tracker_address, tracker_postcode, contacts)
    
    print("\n" + "="*70)
    print("READY TO CREATE NEW PROJECT (BETA - NOT LIVE YET)")
    print("="*70)
    print(f"Template: {template_job}")
    print(f"New Project Data:\n{json.dumps(new_project_data, indent=2)}")
    print("\nSuggested site-create payload for POST /sites:")
    print(json.dumps(site_preview, indent=2))
    print("="*70)
    append_run_note(run_notes, f"Template project loaded: {template_job}")
    
    # USER CHECKPOINT
    confirm = args.create_project or prompt_text(
        "\n[?] Continue past Alpha Tracker preview? (yes/no): ",
        default="no",
        allowed={"yes", "no"},
    )
    
    if confirm != 'yes':
        append_run_note(run_notes, "Stopped at Alpha Tracker preview.")
        write_run_summary(run_notes)
        return
    
    print("\n[BETA] Live Alpha Tracker create/update remains disabled.")
    
    
    # STEP 6: Prepare email draft only
    draft_confirm = prompt_text(
        "\n[?] Create Outlook draft for site contacts now? (yes/no): ",
        default="yes",
        allowed={"yes", "no"},
    )
    if draft_confirm == "yes":
        draft_result = prepare_email_to_site_contact(contacts, template_job, address)
        append_run_note(run_notes, f"Draft email result: {'created' if draft_result else 'not created'}")
    else:
        append_run_note(run_notes, "Draft email skipped by user.")
    
    # STEP 7: Generate desktop study
    generate_desktop_study(template_job)
    append_run_note(run_notes, "Desktop study generation remains manual.")
    
    print("\n" + "="*70)
    print("WORKFLOW COMPLETE")
    print("="*70)
    write_run_summary(run_notes)

if __name__ == "__main__":
    main()
