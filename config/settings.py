from pathlib import Path


BASE_DIR = Path(__file__).resolve().parent.parent

DATA_DIR = BASE_DIR / "data"
OUTPUT_DIR = BASE_DIR / "output"
LOG_DIR = BASE_DIR / "logs"

# Start small while testing.
TARGET_RAW_RECORDS = 100

MIN_FINAL_RECORDS = 13_000


OUTPUT_EXCEL = OUTPUT_DIR / "trader_leads.xlsx"
OUTPUT_CSV = OUTPUT_DIR / "trader_leads.csv"

REQUEST_TIMEOUT = 30

USER_AGENT = (
    "TraderLeadGenerator/1.0 "
    "Authorized data collection project"
)