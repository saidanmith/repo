import pandas as pd
import requests
from concurrent.futures import ThreadPoolExecutor
from tqdm import tqdm
from rapidfuzz import fuzz
import time
import logging
from datetime import datetime

# -------- SETTINGS --------
EXCEL_FILE = r"C:\Users\Sherren\Desktop\feldman\2Copy of Property List and Asbestos Reinspections 2026.xlsx"
CREDENTIALS_FILE = r"C:\Users\Sherren\Desktop\feldman\credentials.txt"
MAX_WORKERS = 5   # number of parallel API calls
SAVE_INTERVAL = 50  # save every N matches
FUZZY_THRESHOLD = 80  # fuzzy match percentage
API_URL = "https://manager.alphatracker.co.uk/api"
MAX_RETRIES = 3
RETRY_DELAY = 5  # seconds
BETWEEN_CALL_DELAY = 0.2  # seconds between API calls (rate limiting)
PROJECT_PREFIX = "G-"  # only projects starting with this prefix are matched

# -------- COLUMN INDICES --------
# Project number column
COL_PROJECT_NUMBER = 13
# Report date column
COL_REPORT_DATE = 14

# -------- LOGGING SETUP --------
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(r'C:\Users\Sherren\Desktop\feldman\search_log.txt'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# -------- LOAD CREDENTIALS --------
client_id = ""
api_key = ""

with open(CREDENTIALS_FILE, "r") as f:
    for line in f:
        if line.startswith("CLIENT_ID="):
            client_id = line.strip().split("=", 1)[1]
        elif line.startswith("API_KEY="):
            api_key = line.strip().split("=", 1)[1]

if not client_id or not api_key:
    logger.error("CLIENT_ID or API_KEY not found in credentials file")
    raise ValueError("CLIENT_ID or API_KEY not found in credentials file")

logger.info("Credentials loaded successfully")

headers = {
    "accept": "application/json",
    "X-API-KEY": api_key,
    "X-CLIENT-ID": client_id
}

# -------- LOAD EXCEL --------
df = pd.read_excel(EXCEL_FILE, header=None)
logger.info(f"Loaded Excel file with {len(df)} rows")

# -------- SEARCH FUNCTION --------
def search_project(row_index):
    address = str(df.iloc[row_index, 0]).strip()

    if not address or pd.isna(address):
        return None

    # skip rows already filled
    if pd.notna(df.iloc[row_index, COL_PROJECT_NUMBER]) and df.iloc[row_index, COL_PROJECT_NUMBER] != "":
        return None

    for attempt in range(MAX_RETRIES):
        try:
            # Rate limiting - small delay between calls
            time.sleep(BETWEEN_CALL_DELAY)
            
            response = requests.get(
                API_URL,
                headers=headers,
                params={"search": address},
                timeout=10
            )

            if response.status_code == 200:
                projects = response.json()

                for p in projects:
                    project_number = str(p.get("projectNumber", ""))
                    if not project_number.startswith(PROJECT_PREFIX):
                        continue

                    api_address = str(p.get("siteAddress", "")).lower()

                    # fuzzy match addresses
                    if fuzz.partial_ratio(address.lower(), api_address) >= FUZZY_THRESHOLD:
                        logger.debug(f"Row {row_index}: Matched '{address}' -> {project_number}")
                        return (
                            row_index,
                            project_number,
                            p.get("reportProduced")
                        )

                # No match found in results
                logger.debug(f"Row {row_index}: No matching {PROJECT_PREFIX} projects found")
                return None

            elif response.status_code == 429:
                logger.warning(f"Row {row_index}: Rate limited (429). Retrying in {RETRY_DELAY}s")
                time.sleep(RETRY_DELAY)
                continue
            else:
                logger.warning(f"Row {row_index}: API returned {response.status_code} on attempt {attempt+1}")
                time.sleep(RETRY_DELAY)
                continue

        except requests.exceptions.Timeout:
            logger.warning(f"Row {row_index}: Request timeout (attempt {attempt+1}/{MAX_RETRIES})")
            time.sleep(RETRY_DELAY)
        except requests.exceptions.ConnectionError:
            logger.warning(f"Row {row_index}: Connection error (attempt {attempt+1}/{MAX_RETRIES})")
            time.sleep(RETRY_DELAY)
        except ValueError as e:
            logger.error(f"Row {row_index}: JSON decode error: {e}")
            return None
        except Exception as e:
            logger.error(f"Row {row_index}: Unexpected error (attempt {attempt+1}/{MAX_RETRIES}): {type(e).__name__}: {e}")
            time.sleep(RETRY_DELAY)

    logger.error(f"Row {row_index}: Failed after {MAX_RETRIES} retries")
    return None

# -------- RUN SEARCHES --------
results = []
no_match_rows = []

logger.info(f"Starting search with {MAX_WORKERS} workers...")
start_time = datetime.now()

with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
    futures = [executor.submit(search_project, i) for i in range(len(df))]

    for f in tqdm(futures, total=len(futures), desc="Searching projects"):
        try:
            result = f.result(timeout=30)
            if result:
                row_index, project_number, report_date = result
                df.iloc[row_index, COL_PROJECT_NUMBER] = project_number
                df.iloc[row_index, COL_REPORT_DATE] = report_date
                results.append(result)
            else:
                # Track rows that didn't match
                row_index = futures.index(f)
                no_match_rows.append(row_index)

            # incremental save
            if len(results) % SAVE_INTERVAL == 0:
                df.to_excel(EXCEL_FILE, index=False, header=False)
                logger.info(f"Saved progress: {len(results)} matches, {len(no_match_rows)} no-match rows")
        except Exception as e:
            logger.error(f"Error processing future result: {e}")

# -------- FINAL SAVE & REPORT --------
df.to_excel(EXCEL_FILE, index=False, header=False)
elapsed = datetime.now() - start_time

logger.info(f"\n{'='*60}")
logger.info(f"Finished in {elapsed}")
logger.info(f"Total rows: {len(df)}")
logger.info(f"Projects matched: {len(results)}")
logger.info(f"No matches found: {len(no_match_rows)}")
logger.info(f"Success rate: {len(results)/len(df)*100:.1f}%")
logger.info(f"{'='*60}")
logger.info(f"Results saved to {EXCEL_FILE}")
logger.info(f"Log saved to C:\\Users\\Sherren\\Desktop\\feldman\\search_log.txt")