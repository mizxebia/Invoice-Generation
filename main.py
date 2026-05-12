import json
import logging
import traceback
import requests
import pandas as pd
from msal import ConfidentialClientApplication


# ==============================
# 📦 LOAD CONFIG
# ==============================

print("🔧 Loading config...")
logging.info("Loading config file...")

with open("config.json", "r") as f:
    config = json.load(f)

auth = config["auth"]
dv = config["dataverse"]
log_cfg = config["logging"]

print("✅ Config loaded")


# ==============================
# 🧾 LOGGING SETUP
# ==============================

logging.basicConfig(
    filename=log_cfg["log_file"],
    level=getattr(logging, log_cfg["log_level"]),
    format="%(asctime)s | %(levelname)s | %(message)s"
)

logger = logging.getLogger(__name__)
logger.info("Logging initialized")
print("✅ logger loaded")


# ==============================
# 🔐 AUTHENTICATION
# ==============================

def get_access_token():
    try:
        logger.info("🔐 Starting authentication...")

        authority_url = f"https://login.microsoftonline.com/{auth['tenant_id']}"
        logger.info(f"Authority URL: {authority_url}")
        logger.info(f"Client ID: {auth['client_id']}")

        app = ConfidentialClientApplication(
            auth["client_id"],
            authority=authority_url,
            client_credential=auth["client_secret"]
        )

        logger.info("Requesting token from MSAL...")

        token_response = app.acquire_token_for_client(
            scopes=[auth["scope"]]
        )

        logger.debug(f"Token response: {token_response}")

        if "access_token" not in token_response:
            logger.error("❌ Token response missing access_token")
            raise Exception(token_response)

        logger.info("✅ Authentication successful")
        return token_response["access_token"]

    except Exception as e:
        logger.error("❌ Authentication failed")
        logger.error(traceback.format_exc())
        print("❌ AUTH ERROR:", e)
        return None


# ==============================
# 📥 FETCH DATA FUNCTION
# ==============================

def fetch_table(table_name, token, ticket_column, ticket_value):
    try:
        logger.info(f"📥 Fetching table: {table_name}")
        logger.info(f"Filter column: {ticket_column}")
        logger.info(f"Filter value: {ticket_value}")

        headers = {
            "Authorization": f"Bearer {token}",
            "Accept": "application/json"
        }

        base_url = auth["dataverse_url"]

        url = (
            f"{base_url}/api/data/v9.2/{table_name}"
            f"?$filter={ticket_column} eq '{ticket_value}'"
        )

        logger.info(f"Initial API URL: {url}")

        all_records = []
        page_count = 1

        while url:
            logger.info(f"➡️ Requesting page {page_count}")
            logger.debug(f"Request URL: {url}")

            response = requests.get(url, headers=headers)

            logger.info(f"Response status: {response.status_code}")

            if response.status_code != 200:
                logger.error(f"❌ API ERROR: {response.text}")
                print("❌ API ERROR:", response.text)
                raise Exception(response.text)

            data = response.json()

            records = data.get("value", [])
            logger.info(f"Records fetched this page: {len(records)}")

            all_records.extend(records)
            logger.info(f"Total records so far: {len(all_records)}")

            # Pagination
            url = data.get("@odata.nextLink")

            if url:
                logger.info("➡️ More pages detected (pagination)")
            else:
                logger.info("✅ No more pages")

            page_count += 1

        logger.info(f"✅ Finished fetching {table_name}. Total records: {len(all_records)}")

        return all_records

    except Exception as e:
        logger.error(f"❌ Error fetching table: {table_name}")
        logger.error(traceback.format_exc())
        print(f"❌ FETCH ERROR ({table_name}):", e)
        return []


# ==============================
# 🚀 MAIN FUNCTION
# ==============================

def main():
    try:
        logger.info("🚀 ===== PROCESS START =====")

        # Step 1: Auth
        logger.info("Step 1: Authentication")
        token = get_access_token()

        if not token:
            logger.error("❌ Token is None, stopping execution")
            return None, None

        logger.info("Step 2: Reading config values")

        ticket_column = dv["columns"]["ticket_id"]
        ticket_value = dv["filter"]["ticket_id"]

        logger.info(f"Ticket column: {ticket_column}")
        logger.info(f"Ticket value: {ticket_value}")

        # Step 3: Fetch tables
        logger.info("Step 3: Fetching closing ticket details")

        closing_ticket_data = fetch_table(
            dv["tables"]["closing_ticket_details"],
            token,
            ticket_column,
            ticket_value
        )

        logger.info("Step 4: Fetching invoice details")

        invoice_data = fetch_table(
            dv["tables"]["invoice_details"],
            token,
            ticket_column,
            ticket_value
        )

        # Step 5: Convert to DataFrames
        logger.info("Step 5: Converting to DataFrames")

        df_closing_ticket_details = pd.DataFrame(closing_ticket_data)
        df_invoice_details = pd.DataFrame(invoice_data)

        logger.info(f"Closing Ticket DF shape: {df_closing_ticket_details.shape}")
        logger.info(f"Invoice DF shape: {df_invoice_details.shape}")

        print("\n=== CLOSING TICKET DETAILS ===")
        print(df_closing_ticket_details.head())

        print("\n=== INVOICE DETAILS ===")
        print(df_invoice_details.head())

        logger.info("✅ ===== PROCESS COMPLETE =====")

        return df_closing_ticket_details, df_invoice_details

    except Exception as e:
        logger.critical("❌ Fatal error in main()")
        logger.critical(traceback.format_exc())
        print("❌ MAIN ERROR:", e)
        return None, None


# ==============================
# ▶️ RUN SCRIPT
# ==============================

if __name__ == "__main__":
    logger.info("Script execution started")

    df_closing_ticket_details, df_invoice_details = main()

    if df_closing_ticket_details is None:
        logger.error("Script failed")
        print("\n❌ Script failed. Check app.log for details.")
    else:
        logger.info("Script completed successfully")
        print("\n✅ Data fetched successfully!")