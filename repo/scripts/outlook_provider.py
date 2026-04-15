import win32com.client
from pathlib import Path
import re
import logging

# Initialize logger for this module
logger = logging.getLogger(__name__)

# ============================================================================
# CONFIG & CONSTANTS
# ============================================================================

EMAIL_SUBJECT_FILTER = "asbestos survey request"
BOT_DIR = Path(r"C:\Users\Sherren\Desktop\lewis\parkingeye bot")
TEMP_DIR = BOT_DIR / "temp"
IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg"}

# ============================================================================
# EMAIL READING & ATTACHMENT HANDLING
# ============================================================================

def get_outlook_emails_from_lewis(account_name=None):
    """Fetch asbestos survey request emails from Lewis Dunkley."""
    logger.info("\n[STEP 1] Fetching emails from Lewis Dunkley...")
    
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        
        target_account = None
        # Prioritize account_name if provided
        if account_name:
            logger.info(f"  Searching for Outlook account: '{account_name}'...")
            for acc in namespace.Accounts:
                if account_name.lower() in acc.DisplayName.lower() or \
                   (hasattr(acc, 'SmtpAddress') and account_name.lower() in acc.SmtpAddress.lower()):
                    target_account = acc
                    logger.info(f"  [OK] Found account: {acc.DisplayName}")
                    break
        
        # Fallback to "a.smith" if account_name not provided or not found
        if not target_account:
            logger.info("  Searching for 'a.smith' account...")
            for acc in namespace.Accounts:
                if "a.smith" in acc.DisplayName.lower():
                    target_account = acc
                    logger.info(f"  [OK] Found 'a.smith' account: {acc.DisplayName}")
                    break
        
        # Fallback to the first available account if neither specific account is found
        if not target_account and namespace.Accounts.Count > 0:
            target_account = namespace.Accounts.Item(1)
            logger.warning(f"  Specific account not found, using default account: {target_account.DisplayName}")
        elif not target_account:
            logger.error("  No Outlook accounts found.")
            return []
        
        root_folder = namespace.Folders.Item(target_account.DisplayName)
        
        # Find Inbox subfolder
        inbox = None
        for folder in root_folder.Folders:
            if folder.Name.lower() == "inbox":
                inbox = folder
                break
        
        if not inbox:
            logger.error(f"  Could not find Inbox folder")
            return []
        
        emails = []
        search_terms = ["lewis", "dunkley", "l.dunkley"]
        target_email = "l.dunkley@greenshieldenvironmental.co.uk"
        
        for item in inbox.Items:
            try:
                sender_name_lower = str(item.SenderName).lower()
                sender_email = str(item.SenderEmailAddress).lower() if hasattr(item, 'SenderEmailAddress') else ""
                
                # Check if matches Lewis OR matches the specific email
                is_lewis = any(term in sender_name_lower for term in search_terms)
                is_target_email = sender_email == target_email.lower()
                
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
                        "store_id": item.Parent.StoreID,
                        "outlook_item": item, # Store the actual Outlook item for later operations
                    })
            except Exception as e:
                logger.warning(f"  Could not process email item (Subject: {getattr(item, 'Subject', 'N/A')}). Skipping. Error: {e}")
        
        emails.sort(key=lambda email: email["received_time"], reverse=True)
        logger.info(f"  [OK] Found {len(emails)} asbestos survey request email(s) from Lewis/Dunkley")
        return emails
    
    except Exception as e:
        logger.error(f"  Error reading Outlook: {e}")
        return []

def extract_attachments(email, output_dir=None):
    """Extract PDF attachments from email."""
    if output_dir is None:
        output_dir = TEMP_DIR / "pdfs"
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    
    extracted_files = []
    for attachment in email['attachments']:
        if attachment.Filename.lower().endswith('.pdf'):
            filepath = output_dir / attachment.Filename
            attachment.SaveAsFile(str(filepath))
            extracted_files.append(filepath)
    logger.info(f"  Saved {len(extracted_files)} PDF attachment(s)")
    return extracted_files

def get_attachment_mime_type(attachment):
    """Read an Outlook attachment MIME type when available."""
    try:
        return attachment.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x370E001E")
    except Exception:
        return ""

def list_outlook_folders():
    """List all available folders in Outlook (for debugging)."""
    # This function is primarily for debugging and might be moved or adapted
    # depending on how the main orchestrator handles debug modes.
    # For now, it's kept here as it directly interacts with Outlook MAPI.
    logger.info("\n[DEBUG] Listing all Outlook folders...")
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        root_folder = namespace.Folders.Item(1)
        
        def print_folders(folder, indent=0):
            prefix = "  " * indent
            try:
                logger.info(f"{prefix}- {folder.Name} ({folder.Items.Count} items)")
                if hasattr(folder, 'Folders'):
                    for subfolder in folder.Folders:
                        print_folders(subfolder, indent + 1)
            except Exception as e:
                logger.warning(f"{prefix}  Error accessing folder {folder.Name}: {e}")
        
        print_folders(root_folder)
    except Exception as e:
        logger.error(f"  Error listing Outlook folders: {e}")

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
        try:
            attachment.SaveAsFile(str(target))
            saved_files.append(target)
        except Exception as e:
            logger.warning(f"  Could not save inline image {safe_name}: {e}")

    return saved_files