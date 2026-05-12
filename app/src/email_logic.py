import json
import re
import win32com.client
import pytesseract
import pdfplumber
from pathlib import Path
from datetime import datetime
from PIL import Image, ImageOps

# ============================================================================
# CONFIG (Updated for Source Layout)
# ============================================================================

# Tesseract Path
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

# Path Logic
# __file__ is: .../app/src/email_logic.py
# .parent is: .../app/src/
# .parent.parent is: .../app/
SCRIPT_DIR   = Path(__file__).resolve().parent
PROJECT_ROOT = SCRIPT_DIR.parent

TEMP_DIR      = PROJECT_ROOT / "temp"
LOGS_DIR      = PROJECT_ROOT / "logs"
SENT_LOG_FILE = LOGS_DIR / "sent_log.json"
PROCESS_LOG   = PROJECT_ROOT / "process_log.txt"

# Ensure directories exist
TEMP_DIR.mkdir(parents=True, exist_ok=True)
LOGS_DIR.mkdir(parents=True, exist_ok=True)

EMAIL_SUBJECT_FILTER = "asbestos survey request"
IMAGE_EXTENSIONS     = {".png", ".jpg", ".jpeg"}

UK_POSTCODE_RE = re.compile(
    r"\b[A-Z]{1,2}\d[A-Z\d]?\s*\d[A-Z]{2}\b",
    re.IGNORECASE,
)

def log_operation(operation, target, outcome):
    """Structured logging per INSTRUCTIONS.md."""
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    log_entry = f"[{timestamp}, {operation}, {target}, {outcome}]\n"
    with open(PROCESS_LOG, "a", encoding="utf-8") as f:
        f.write(log_entry)

# ============================================================================
# SENT LOG
# ============================================================================

def load_sent_log():
    """Load the list of handled email IDs from the logs folder."""
    if SENT_LOG_FILE.exists():
        try:
            return set(json.loads(SENT_LOG_FILE.read_text(encoding="utf-8")))
        except Exception as e:
            print(f"[WARN] Could not read log file: {e}")
            return set()
    return set()


def save_sent_log(sent_ids):
    """Save the list of handled email IDs to the logs folder."""
    try:
        SENT_LOG_FILE.write_text(json.dumps(list(sent_ids), indent=2), encoding="utf-8")
    except Exception as e:
        print(f"[ERROR] Could not save log file: {e}")


# ============================================================================
# EMAIL FETCHING
# ============================================================================

def get_mail_item(message_id, store_id=None):
    """Fetches a fresh MailItem from Outlook (required for background threads)."""
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        if store_id:
            return namespace.GetItemFromID(message_id, store_id)
        return namespace.GetItemFromID(message_id)
    except Exception as e:
        print(f"[ERROR] Failed to fetch mail item: {e}")
        return None


def get_asbestos_request_emails():
    """Fetch all unhandled asbestos survey request emails."""
    print("\n[INFO] Connecting to Outlook and searching for survey requests...")
    sent_ids = load_sent_log()

    try:
        outlook   = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")

        account = None
        for acc in namespace.Accounts:
            if "a.smith" in acc.DisplayName.lower():
                account = acc
                break
        if not account:
            account = namespace.Accounts[0]

        root  = namespace.Folders.Item(account.DisplayName)
        inbox = None
        for folder in root.Folders:
            if folder.Name.lower() == "inbox":
                inbox = folder
                break

        if not inbox:
            print("[ERROR] Inbox not found.")
            return []

        emails       = []

        items = inbox.Items
        for i in range(1, items.Count + 1):
            item = items.Item(i)
            try:
                sender_email = str(getattr(item, "SenderEmailAddress", "")).lower()
                subject      = str(getattr(item, "Subject", "") or "")
                entry_id     = item.EntryID

                is_survey  = EMAIL_SUBJECT_FILTER in subject.lower()
                is_handled = entry_id in sent_ids

                if is_survey and not is_handled:
                    emails.append({
                        "sender":        item.SenderName,
                        "sender_email":  sender_email,
                        "subject":       subject,
                        "body":          item.Body,
                        "received_time": item.ReceivedTime,
                        "attachments":   item.Attachments,
                        "message_id":    entry_id,
                        "store_id":      item.Parent.StoreID,
                    })
            except Exception:
                pass

        emails.sort(key=lambda e: e["received_time"], reverse=True)
        return emails

    except Exception as e:
        print(f"[ERROR] Could not read Outlook: {e}")
        return []


# ============================================================================
# JOB TYPE DETECTION
# ============================================================================

def detect_job_type(pdf_paths, fallback_subject=None):
    """Parses PDF text or subject line to determine client."""
    for pdf_path in pdf_paths:
        pdf = Path(pdf_path)
        if not pdf.exists(): continue
        try:
            with pdfplumber.open(pdf) as f:
                text = " ".join(page.extract_text() or "" for page in f.pages)
            if "parkingeye" in text.lower():
                return "parkingeye"
            if "g24" in text.lower():
                return "g24"
        except Exception:
            pass
    if fallback_subject:
        s = fallback_subject.lower()
        if "parkingeye" in s:
            return "parkingeye"
        if "g24" in s:
            return "g24"
    return None


# ============================================================================
# OCR / CONTACT EXTRACTION
# ============================================================================

def get_attachment_mime_type(attachment):
    try:
        return attachment.PropertyAccessor.GetProperty(
            "http://schemas.microsoft.com/mapi/proptag/0x370E001E"
        )
    except Exception:
        return ""


def extract_pdf_attachments(email_item):
    """Saves PDF attachments from the raw Outlook MailItem."""
    pdf_paths = []
    pdf_dir = TEMP_DIR / "pdfs"
    pdf_dir.mkdir(parents=True, exist_ok=True)

    try:
        attachments = email_item.Attachments
        for i in range(1, attachments.Count + 1):
            att = attachments.Item(i)
            filename = str(getattr(att, "FileName", "") or "").strip()
            if filename.lower().endswith(".pdf"):
                dest = pdf_dir / filename
                att.SaveAsFile(str(dest))
                pdf_paths.append(dest)
    except Exception as e:
        print(f"  [ERROR] Failed to extract PDFs: {e}")
    return pdf_paths


def extract_address_from_pdfs(pdf_paths):
    """Attempts to find a site address and postcode within the provided PDFs."""
    for pdf_path in pdf_paths:
        try:
            with pdfplumber.open(pdf_path) as pdf:
                # Usually, the address is on the first page of signage plans or instructions
                text = pdf.pages[0].extract_text() or ""
                lines = [l.strip() for l in text.splitlines() if l.strip()]
                
                for i, line in enumerate(lines):
                    if UK_POSTCODE_RE.search(line):
                        # Capture the postcode line and the 1-2 lines preceding it
                        start = max(0, i - 2)
                        address_block = lines[start : i + 1]
                        return ", ".join(address_block)
        except Exception as e:
            print(f"  [DEBUG] Address extraction error on {pdf_path.name}: {e}")
    return ""


def extract_inline_images(email_item):
    """Saves image attachments from the raw Outlook MailItem."""
    output_dir = TEMP_DIR / "inline_images"
    if output_dir.exists():
        for old_file in output_dir.iterdir():
            try: old_file.unlink() 
            except Exception: pass

    output_dir.mkdir(parents=True, exist_ok=True)
    saved = []
    try:
        attachments = email_item.Attachments
        for i in range(1, attachments.Count + 1):
            att = attachments.Item(i)
            filename = str(getattr(att, "FileName", "") or "").strip()
            suffix = Path(filename).suffix.lower()
            mime_type = str(get_attachment_mime_type(att) or "").lower()
            
            if suffix not in IMAGE_EXTENSIONS and not mime_type.startswith("image/"):
                continue
            
            safe_name = filename or f"inline_image_{i}{suffix or '.png'}"
            dest = output_dir / safe_name
            att.SaveAsFile(str(dest))
            saved.append(dest)
    except Exception as e:
        print(f"  [ERROR] Failed to extract images: {e}")
    return saved


def normalize_contact_name(name):
    if not name:
        return None
    # Strip smart/curly quotes and other OCR noise before removing non-alpha chars
    name = name.replace("\u2018", "").replace("\u2019", "").replace("'", "").replace("`", "")
    name = re.sub(r"[^A-Za-z\s]", " ", name)
    name = re.sub(r"\s+", " ", name).strip()
    if "activate.ps1" in name.lower():
        return None
    return name or None


def normalize_contact_email(email):
    if not email:
        return None
    # Remove whitespace and common OCR artefacts like pipes or slashes
    email = re.sub(r"[\s|\\/]", "", email)
    email = email.strip(" ,.;:").replace("..", ".")
    email = re.sub(r"^[^a-zA-Z0-9]+", "", email)
    return email.lower() or None


def extract_contact_candidates_from_text(text):
    normalized = text.replace("\r", "\n")

    candidates = []
    lines = [l.strip() for l in normalized.splitlines() if l.strip()]
    current_name = None

    for line in lines:
        # Look for name patterns: Name:, Contact:, etc.
        name_match = re.search(r"(?:Name|Contact|Site Contact)\b[:\s]+(.+)", line, re.IGNORECASE)
        if name_match:
            current_name = normalize_contact_name(name_match.group(1))
            continue
            
        # Look for email patterns: Email:, Contact Email:, etc.
        email_label_match = re.search(r"(?:Email|Contact Email|Email Address)\b[:\s]+(.+)", line, re.IGNORECASE)
        if email_label_match:
            email_match = re.search(r"([a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,})", email_label_match.group(1))
            if email_match:
                candidates.append({"name": current_name, "email": normalize_contact_email(email_match.group(1))})
                current_name = None
            continue

        # Catch bare emails anywhere
        email_match = re.search(r"([a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,})", line)
        if email_match:
            candidates.append({"name": current_name, "email": normalize_contact_email(email_match.group(1))})
            current_name = None

    # Global fallback: find all emails and try to associate with nearby names
    if not candidates:
        all_emails = re.findall(r"([a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,})", normalized)
        # Look for names near emails (within a few lines)
        for match in re.finditer(r"([a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,})", normalized):
            email = normalize_contact_email(match.group(1))
            # Look backwards for a name
            start = match.start()
            before_text = normalized[:start]
            name_match = re.search(r"([A-Za-z\s]{3,20})(?:\s|$)", before_text[-100:], re.IGNORECASE)
            name = normalize_contact_name(name_match.group(1)) if name_match else None
            candidates.append({"name": name, "email": email})

    return dedupe_candidates(candidates)


def dedupe_candidates(candidates):
    seen, deduped = set(), []
    for c in candidates:
        key = (normalize_contact_name(c.get("name")), normalize_contact_email(c.get("email")))
        if key not in seen:
            seen.add(key)
            deduped.append({"name": key[0], "email": key[1]})
    return deduped


def ocr_image(img_path):
    """OCR an image with enhanced preprocessing for better accuracy."""
    img = Image.open(img_path).convert("RGB")
    
    # Convert to grayscale
    gray = img.convert("L")
    
    # Apply thresholding to get binary image
    thresh = gray.point(lambda x: 0 if x < 128 else 255, 'L')
    
    # Try OCR on original, grayscale, and thresholded
    texts = []
    for img_variant in [img, gray, thresh]:
        try:
            text = pytesseract.image_to_string(img_variant, config='--psm 6')
            texts.append(text)
        except Exception as e:
            print(f"[DEBUG] OCR failed on variant: {e}")
            texts.append("")
    
    # Choose the best text (longest with letters)
    best_text = max(texts, key=lambda t: len(re.findall(r'[a-zA-Z]', t)))
    
    # If still no useful text, try inverting the thresholded image
    if not re.search(r"[a-zA-Z]{3,}", best_text):
        inverted_thresh = ImageOps.invert(thresh)
        try:
            inv_text = pytesseract.image_to_string(inverted_thresh, config='--psm 6')
            if len(re.findall(r'[a-zA-Z]', inv_text)) > len(re.findall(r'[a-zA-Z]', best_text)):
                best_text = inv_text
        except Exception as e:
            print(f"[DEBUG] Inverted OCR failed: {e}")
    
    print(f"[DEBUG] OCR result for {img_path.name}: '{best_text[:100]}...'")
    return best_text


def extract_site_contact(email_item):
    """OCR inline images and parse email body for contact candidates."""
    candidates = []
    
    # First, try OCR on images
    image_paths = extract_inline_images(email_item)
    print(f"[DEBUG] Found {len(image_paths)} images to OCR")
    for img_path in image_paths:
        try:
            img = Image.open(img_path)
            if img.width < 200 or img.height < 50:
                print(f"[DEBUG] Skipping small image: {img.width}x{img.height}")
                continue
                
            text = ocr_image(img_path)
            img_candidates = extract_contact_candidates_from_text(text)
            print(f"[DEBUG] Extracted {len(img_candidates)} candidates from {img_path.name}")
            candidates.extend(img_candidates)
        except Exception as e:
            print(f"[DEBUG] Error OCRing {img_path.name}: {e}")
    
    # If no candidates from images, try parsing the email body
    if not candidates:
        try:
            body_text = str(getattr(email_item, 'Body', '') or '')
            print(f"[DEBUG] Parsing email body for contacts")
            body_candidates = extract_contact_candidates_from_text(body_text)
            print(f"[DEBUG] Extracted {len(body_candidates)} candidates from email body")
            candidates.extend(body_candidates)
        except Exception as e:
            print(f"[DEBUG] Error parsing email body: {e}")
    
    final_candidates = dedupe_candidates(candidates)
    print(f"[DEBUG] Final deduped candidates: {final_candidates}")
    return final_candidates


# ============================================================================
# DATE PARSING
# ============================================================================

def parse_visit_date(date_str):
    try:
        dt = datetime.strptime(date_str.strip(), "%d/%m/%Y")
        return dt.strftime("%A"), dt.strftime("%d/%m/%Y")
    except ValueError:
        raise ValueError(f"Invalid date '{date_str}'. Use dd/mm/yyyy.")


# ============================================================================
# EMAIL BODY
# ============================================================================

def build_email_body(contact_name, client_label, site_address, day_name, formatted_date):
    hour     = datetime.now().hour
    greeting = "Good morning" if hour < 12 else "Good afternoon" if hour < 18 else "Good evening"
    first_name = contact_name.split()[0] if contact_name else "[Name]"
    
    return (
        f"{greeting} {first_name},\n\n"
        f"My name is Aidan and I work for Greenshield Environmental. I have been provided your contact details by {client_label}, "
        f"in regards to booking in a small targeted asbestos survey for the below-named premises, prior to their installation works.\n\n"
        f"{site_address}\n\n"
        f"The survey is predominately external so will not cause any disruption to any on-site members of staff or guests, "
        f"and should only take around 30-40 minutes, would it be possible to send a surveyor on {day_name} {formatted_date} "
        f"to undertake the survey please?\n\n"
        f"Any issues please do not hesitate to reply to this email.\n"
        f"Kind regards,\nAidan Smith."
    )


def build_email_body_html(contact_name, client_label, site_address, day_name, formatted_date):
    hour     = datetime.now().hour
    greeting = "Good morning" if hour < 12 else "Good afternoon" if hour < 18 else "Good evening"
    first_name = contact_name.split()[0] if contact_name else "[Name]"
    addr_html = site_address.replace(", ", "<br>")

    return (
        f"<p>{greeting} {first_name},</p>"
        f"<p>My name is Aidan and I work for Greenshield Environmental. I have been provided your contact details by {client_label}, "
        f"in regards to booking in a small targeted asbestos survey for the below-named premises, prior to their installation works.</p>"
        f"<p><b>{addr_html}</b></p>"
        f"<p>The survey is predominately external so will not cause any disruption to any on-site members of staff or guests, "
        f"and should only take around 30-40 minutes, would it be possible to send a surveyor on {day_name} {formatted_date} "
        f"to undertake the survey please?</p>"
        f"<p>Any issues please do not hesitate to reply to this email.<br>"
        f"Kind regards,<br>Aidan Smith.</p>"
    )


# ============================================================================
# OUTLOOK DRAFTING
# ============================================================================

def open_new_draft(to_email, subject, html_body):
    """Creates a brand new email draft in Outlook."""
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        draft = outlook.CreateItem(0) # 0 = olMailItem
        
        draft.To      = to_email
        draft.Subject = subject
        draft.HTMLBody = html_body

        draft.Save()
        draft.Display()
        print("[OK] New draft opened in Outlook.")

    except Exception as e:
        print(f"[ERROR] Could not create draft: {e}")


# ============================================================================
# PROMPTS
# ============================================================================

def prompt(message, allowed=None, default=None):
    while True:
        raw = input(message).strip()
        val = raw.lower() if raw else (default or "")
        if allowed and val not in allowed:
            print(f"  Please enter one of: {', '.join(sorted(allowed))}")
            continue
        return val


def prompt_required(message, current=None):
    suffix = f" [{current}]" if current else ""
    while True:
        val = input(f"{message}{suffix}: ").strip()
        if val:
            return val
        if current:
            return current
        print("  This field is required.")


# ============================================================================
# MAIN
# ============================================================================

def main():
    print("=" * 60)
    print("  QUICK SITE CONTACT EMAIL")
    print("=" * 60)

    # --- Step 1: Fetch emails ---
    emails = get_asbestos_request_emails()
    if not emails:
        print("\n[INFO] No unhandled asbestos survey request emails found.")
        return

    print(f"\nFound {len(emails)} unhandled email(s):\n")
    for i, e in enumerate(emails, start=1):
        received = e["received_time"].strftime("%d/%m/%Y %H:%M")
        print(f"  [{i}] {received} — {e['subject']}")

    # --- Step 2: Pick email ---
    while True:
        raw = input("\nEnter number to select email: ").strip()
        if raw.isdigit() and 1 <= int(raw) <= len(emails):
            selected = emails[int(raw) - 1]
            break
        print(f"  Please enter a number between 1 and {len(emails)}.")

    print(f"\n[OK] Selected: {selected['subject']}")

    # --- Step 3: Detect job type ---
    print("\n[INFO] Extracting PDF attachments...")
    pdf_files = extract_pdf_attachments(selected)
    job_type  = detect_job_type(pdf_files, fallback_subject=selected["subject"])
    if not job_type:
        job_type = prompt(
            "Could not detect job type. Enter 'parkingeye' or 'g24': ",
            allowed={"parkingeye", "g24"},
        )
    client_label = "Parkingeye" if job_type == "parkingeye" else "G24"
    print(f"[OK] Job type: {client_label}")

    # --- Step 4: Extract contact via OCR ---
    print("\n[INFO] Extracting site contact details via OCR...")
    candidates = extract_site_contact(selected)

    # Always initialise contact so it is never unbound below
    contact = {"name": None, "email": None}

    if candidates:
        contact = candidates[0]
        print(f"\n  Name:  {contact.get('name') or '[not found]'}")
        print(f"  Email: {contact.get('email') or '[not found]'}")
        ok = prompt("\nAre these details correct? (y/n): ", allowed={"y", "n"})
        if ok == "n":
            contact = {
                "name":  prompt_required("  Enter contact name"),
                "email": prompt_required("  Enter contact email"),
            }
    else:
        print("  [!] No contact details found via OCR.")
        contact = {
            "name":  prompt_required("  Enter contact name"),
            "email": prompt_required("  Enter contact email"),
        }

    # Fill any missing fields
    if not contact.get("name"):
        contact["name"] = prompt_required("  Enter contact name")
    if not contact.get("email"):
        contact["email"] = prompt_required("  Enter contact email")

    # --- Step 5: Visit date ---
    while True:
        date_str = input("\nVisit date (dd/mm/yyyy): ").strip()
        try:
            day_name, formatted_date = parse_visit_date(date_str)
            break
        except ValueError as e:
            print(f"  [!] {e}")

    # --- Step 7: Mark as sent ---
    sent = prompt("\nHave you sent the email? (y/n): ", allowed={"y", "n"})
    if sent == "y":
        log = load_sent_log()
        log.add(selected["message_id"])
        save_sent_log(log)
        log_operation("MARK_HANDLED", selected["message_id"], "SUCCESS")
        print("[OK] Marked as handled. Won't appear in future runs.")
    else:
        print("[INFO] Not marked as sent. Will appear again next time.")


if __name__ == "__main__":
    main()