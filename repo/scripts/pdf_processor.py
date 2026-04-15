import re
from pathlib import Path
import pytesseract
import pdfplumber
from PIL import Image
try:
    from pdf2image import convert_from_path
except ImportError:
    convert_from_path = None

# ============================================================================
# CONFIG & CONSTANTS
# ============================================================================

# Point pytesseract at the Tesseract install on Windows
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

BOT_DIR = Path(r"C:\Users\Sherren\Desktop\lewis\parkingeye bot")
TEMP_DIR = BOT_DIR / "temp"

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

IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg"}

# ============================================================================
# PDF & ADDRESS EXTRACTION
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
    match = re.search(r'PO-(\d+)', email_body, re.IGNORECASE)
    if match:
        return match.group(1)
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

# ============================================================================
# CONTACT & OCR EXTRACTION
# ============================================================================

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

def detect_job_type(pdf_paths, fallback_subject=None):
    """Detect if job is ParkingEye or G24 based on PDF content or subject."""
    for pdf in pdf_paths:
        try:
            with pdfplumber.open(pdf) as pdf_file:
                text = " ".join([page.extract_text() or "" for page in pdf_file.pages])
                
                if "parkingeye" in text.lower():
                    return "parkingeye"
                elif "g24" in text.lower() or "G24" in text:
                    return "g24"
        except:
            pass
    
    if fallback_subject:
        if "parkingeye" in fallback_subject.lower():
            return "parkingeye"
        elif "g24" in fallback_subject.lower():
            return "g24"
    return None