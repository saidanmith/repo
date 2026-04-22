import json
import re
import win32com.client
import pytesseract
import pdfplumber
from pathlib import Path
from datetime import datetime
from PIL import Image

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

# Ensure directories exist
TEMP_DIR.mkdir(parents=True, exist_ok=True)
LOGS_DIR.mkdir(parents=True, exist_ok=True)

EMAIL_SUBJECT_FILTER = "asbestos survey request"
IMAGE_EXTENSIONS     = {".png", ".jpg", ".jpeg"}

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

def get_lewis_emails():
    """Fetch unhandled asbestos survey request emails from Lewis Dunkley."""
    print("\n[INFO] Connecting to Outlook...")
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
        search_terms = ["lewis", "dunkley", "l.dunkley"]
        target_email = "l.dunkley@greenshieldenvironmental.co.uk"

        for item in inbox.Items:
            try:
                sender_name  = str(item.SenderName).lower()
                sender_email = str(getattr(item, "SenderEmailAddress", "")).lower()
                subject      = str(getattr(item, "Subject", "") or "")
                entry_id     = item.EntryID

                is_lewis   = any(t in sender_name for t in search_terms) or target_email in sender_email
                is_survey  = EMAIL_SUBJECT_FILTER in subject.lower()
                is_handled = entry_id in sent_ids

                if is_lewis and is_survey and not is_handled:
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

def extract_pdf_attachments(email):
    output_dir = TEMP_DIR / "pdfs"
    output_dir.mkdir(parents=True, exist_ok=True)
    files = []
    for att in email["attachments"]:
        if att.Filename.lower().endswith(".pdf"):
            dest = output_dir / att.Filename
            att.SaveAsFile(str(dest))
            files.append(dest)
    return files


def detect_job_type(pdf_paths, fallback_subject=None):
    for pdf in pdf_paths:
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


def extract_inline_images(email):
    output_dir = TEMP_DIR / "inline_images"
    # Clear stale images from previous runs
    if output_dir.exists():
        for old_file in output_dir.iterdir():
            try:
                old_file.unlink()
            except Exception:
                pass
    output_dir.mkdir(parents=True, exist_ok=True)
    saved = []
    for i, att in enumerate(email["attachments"], start=1):
        filename  = str(getattr(att, "Filename", "") or "").strip()
        suffix    = Path(filename).suffix.lower()
        mime_type = str(get_attachment_mime_type(att) or "").lower()
        if suffix not in IMAGE_EXTENSIONS and not mime_type.startswith("image/"):
            continue
        if not suffix:
            suffix = ".png"
        safe_name = filename or f"inline_image_{i}{suffix}"
        dest = output_dir / safe_name
        att.SaveAsFile(str(dest))
        saved.append(dest)
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
    # Remove all whitespace and pipe characters (OCR table border artefacts)
    email = re.sub(r"[\s|]", "", email)
    email = email.strip(" ,.;:").replace("..", ".")
    # Strip any leading punctuation OCR artefacts
    email = re.sub(r"^[^a-zA-Z0-9]+", "", email)
    return email.lower() or None


def extract_contact_candidates_from_text(text):
    normalized = text.replace("\r", "\n")

    candidates   = []
    lines        = [l.strip() for l in normalized.splitlines() if l.strip()]
    current_name = None

    for line in lines:
        name_match = re.search(r"^Name\b[:\s]+(.+)", line, re.IGNORECASE)
        if name_match:
            current_name = normalize_contact_name(name_match.group(1))
            continue
        # Match "Email Address:", "Email Address", "Email:", "Email" followed by content
        email_label_match = re.search(r"^Email(?:\s*Address)?\b[:\s]+(.+)", line, re.IGNORECASE)
        if email_label_match:
            raw = email_label_match.group(1)
            email_in_raw = re.search(r"([a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,})", re.sub(r"\s+", "", raw))
            if email_in_raw:
                candidates.append({
                    "name":  current_name,
                    "email": normalize_contact_email(email_in_raw.group(1)),
                })
                current_name = None
            continue
        # Catch bare email addresses on their own line (fallback for any format)
        email_match = re.search(r"([a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,})", line)
        if email_match:
            candidates.append({
                "name":  current_name,
                "email": normalize_contact_email(email_match.group(1)),
            })
            current_name = None

    if candidates:
        return candidates

    # Fallback: grab whatever name/email exist anywhere in the text
    names  = re.findall(r"Name[:\s]+([A-Z][a-zA-Z\s\-\']+)", normalized)
    emails = re.findall(r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}", normalized)
    if names or emails:
        return [{
            "name":  normalize_contact_name(names[0]) if names else None,
            "email": normalize_contact_email(emails[0]) if emails else None,
        }]
    return []


def dedupe_candidates(candidates):
    seen, deduped = set(), []
    for c in candidates:
        key = (normalize_contact_name(c.get("name")), normalize_contact_email(c.get("email")))
        if key not in seen:
            seen.add(key)
            deduped.append({"name": key[0], "email": key[1]})
    return deduped


def ocr_image(img_path):
    """OCR an image, retrying with inverted colours if the image is dark."""
    img = Image.open(img_path).convert("RGB")
    text = pytesseract.image_to_string(img)
    # If no useful text found and image appears dark, try inverting
    if not re.search(r"[a-zA-Z]{3,}", text):
        from PIL import ImageOps
        inverted = ImageOps.invert(img)
        text = pytesseract.image_to_string(inverted)
    return text


def extract_site_contact(email):
    """OCR inline images and return deduplicated contact candidates."""
    candidates  = []
    image_paths = extract_inline_images(email)
    for img_path in image_paths:
        try:
            img = Image.open(img_path)
            # Skip tiny images (logos, icons)
            if img.width < 200 or img.height < 50:
                continue
            text = ocr_image(img_path)
            candidates.extend(extract_contact_candidates_from_text(text))
        except Exception:
            pass
    return dedupe_candidates(candidates)


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

def build_email_body(contact_name, client_label, day_name, formatted_date):
    hour     = datetime.now().hour
    greeting = "Good morning" if hour < 12 else "Good afternoon" if hour < 18 else "Good evening"
    return (
        f"{greeting} {contact_name},\n\n"
        f"I have been provided your contact details by {client_label}, in regards to booking in a small targeted "
        f"asbestos survey for the above-named premises, prior to their installation works.\n\n"
        f"The survey is predominately external so will not cause any disruption to any on-site members of staff "
        f"or guests, and should only take around 30-40 minutes, would it be possible to send a surveyor on "
        f"{day_name} {formatted_date} to undertake the survey please?\n\n"
        f"Any issues please do not hesitate to reply to this email.\n"
        f"Kind regards,\n"
        f"Aidan Smith."
    )


def build_email_body_html(contact_name, client_label, day_name, formatted_date):
    hour     = datetime.now().hour
    greeting = "Good morning" if hour < 12 else "Good afternoon" if hour < 18 else "Good evening"
    return (
        f"<p>{greeting} {contact_name},</p>"
        f"<p>I have been provided your contact details by {client_label}, in regards to booking in a small targeted "
        f"asbestos survey for the above-named premises, prior to their installation works.</p>"
        f"<p>The survey is predominately external so will not cause any disruption to any on-site members of staff "
        f"or guests, and should only take around 30-40 minutes, would it be possible to send a surveyor on "
        f"{day_name} {formatted_date} to undertake the survey please?</p>"
        f"<p>Any issues please do not hesitate to reply to this email.<br>"
        f"Kind regards,<br>Aidan Smith.</p>"
    )


# ============================================================================
# OUTLOOK FORWARD DRAFT
# ============================================================================

def open_forward_draft(email, to_email, subject, plain_body, html_intro):
    """Forward the original Lewis email with our intro prepended."""
    try:
        outlook   = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")

        message_id  = email["message_id"]
        store_id    = email["store_id"]
        source_item = (
            namespace.GetItemFromID(message_id, store_id)
            if store_id
            else namespace.GetItemFromID(message_id)
        )

        draft         = source_item.Forward()
        draft.To      = to_email
        draft.Subject = subject

        if draft.HTMLBody:
            draft.HTMLBody = html_intro + draft.HTMLBody
        else:
            draft.Body = plain_body + "\n\n" + "-" * 40 + "\n" + draft.Body

        draft.Save()
        draft.Display()
        print("[OK] Forward draft opened in Outlook. Review and send manually.")

    except Exception as e:
        print(f"[ERROR] Could not create forward draft: {e}")


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
    emails = get_lewis_emails()
    if not emails:
        print("\n[INFO] No unhandled asbestos survey request emails found from Lewis.")
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

    # --- Step 6: Preview & open forward draft ---
    plain_body = build_email_body(contact["name"], client_label, day_name, formatted_date)
    html_intro = build_email_body_html(contact["name"], client_label, day_name, formatted_date)
    subject    = f"Asbestos Survey Booking - {formatted_date}"

    print("\n" + "=" * 60)
    print(f"To:      {contact['email']}")
    print(f"Subject: {subject}")
    print("-" * 60)
    print(plain_body)
    print("=" * 60)

    if prompt("\nOpen forward draft in Outlook? (y/n): ", allowed={"y", "n"}) != "y":
        print("[INFO] Cancelled.")
        return

    open_forward_draft(selected, contact["email"], subject, plain_body, html_intro)

    # --- Step 7: Mark as sent ---
    sent = prompt("\nHave you sent the email? (y/n): ", allowed={"y", "n"})
    if sent == "y":
        log = load_sent_log()
        log.add(selected["message_id"])
        save_sent_log(log)
        print("[OK] Marked as handled. Won't appear in future runs.")
    else:
        print("[INFO] Not marked as sent. Will appear again next time.")


if __name__ == "__main__":
    main()