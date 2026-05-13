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

with open("config.json", "r") as f:
    config = json.load(f)

auth = config["auth"]
dv = config["dataverse"]
storage = config["storage"]
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

logger.info("✅ Logger initialized")

print("✅ Logger loaded")


# ==============================
# 🔐 AUTHENTICATION
# ==============================

def get_access_token(scope):

    try:

        logger.info(
            f"🔐 Generating token for scope: {scope}"
        )

        authority_url = (
            f"https://login.microsoftonline.com/"
            f"{auth['tenant_id']}"
        )

        app = ConfidentialClientApplication(
            auth["client_id"],
            authority=authority_url,
            client_credential=auth["client_secret"]
        )

        token_response = app.acquire_token_for_client(
            scopes=[scope]
        )

        if "access_token" not in token_response:

            logger.error(token_response)

            raise Exception(token_response)

        logger.info(
            "✅ Token generated successfully"
        )

        return token_response["access_token"]

    except Exception as e:

        logger.error(
            "❌ Authentication Failed"
        )

        logger.error(traceback.format_exc())

        print("\n❌ AUTH ERROR:")
        print(e)

        return None


# ==============================
# 📥 FETCH DATAVERSE TABLE
# ==============================

def fetch_table(
    table_name,
    token,
    ticket_column,
    ticket_value
):

    try:

        logger.info(
            f"📥 Fetching table: {table_name}"
        )

        headers = {
            "Authorization": f"Bearer {token}",
            "Accept": "application/json"
        }

        base_url = auth["dataverse_url"]

        url = (
            f"{base_url}/api/data/v9.2/{table_name}"
            f"?$filter={ticket_column} eq '{ticket_value}'"
        )

        logger.info(f"Dataverse URL: {url}")

        print("\n🌐 DATAVERSE URL:")
        print(url)

        all_records = []

        while url:

            response = requests.get(
                url,
                headers=headers
            )

            print(
                "\n📡 Dataverse Response:",
                response.status_code
            )

            if response.status_code != 200:

                logger.error(response.text)

                print(
                    "\n❌ DATAVERSE ERROR:"
                )

                print(response.text)

                raise Exception(response.text)

            data = response.json()

            records = data.get("value", [])

            all_records.extend(records)

            url = data.get("@odata.nextLink")

        logger.info(
            f"✅ Total Records: "
            f"{len(all_records)}"
        )

        return all_records

    except Exception as e:

        logger.error(
            f"❌ FETCH ERROR ({table_name})"
        )

        logger.error(traceback.format_exc())

        print(
            f"\n❌ FETCH ERROR ({table_name}):"
        )

        print(e)

        raise


# ==============================
# 📁 FETCH ONEDRIVE FILE
# ==============================

def get_onedrive_file(
    token,
    user_email,
    file_path
):

    try:

        logger.info(
            "📁 Fetching OneDrive File"
        )

        headers = {
            "Authorization": f"Bearer {token}"
        }

        # CORRECT GRAPH FORMAT
        url = (
            f"https://graph.microsoft.com/v1.0/"
            f"users/{user_email}"
            f"/drive/root:/{file_path}:/"
        )

        logger.info(f"Graph URL: {url}")

        print("\n🌐 GRAPH URL:")
        print(url)

        response = requests.get(
            url,
            headers=headers
        )

        print(
            "\n📡 Graph Response:",
            response.status_code
        )

        if response.status_code != 200:

            logger.error(response.text)

            print(
                "\n❌ OneDrive ERROR:"
            )

            print(response.text)

            raise Exception(response.text)

        data = response.json()

        logger.info(
            "✅ OneDrive File Found"
        )

        print(
            "\n✅ TEMPLATE FILE FOUND"
        )

        print("📄 File Name :", data["name"])
        print("🆔 File ID   :", data["id"])

        return data

    except Exception as e:

        logger.error(
            "❌ OneDrive Fetch Failed"
        )

        logger.error(traceback.format_exc())

        print(
            "\n❌ ONEDRIVE FETCH ERROR:"
        )

        print(e)

        raise


# ==============================
# 🚀 MAIN
# ==============================

def main():

    try:

        logger.info(
            "🚀 ===== PROCESS START ====="
        )

        # ==============================
        # STEP 1: GENERATE TOKENS
        # ==============================

        print(
            "\n🔐 Generating Dataverse Token..."
        )

        dv_token = get_access_token(
            auth["dataverse_scope"]
        )

        if not dv_token:

            raise Exception(
                "Dataverse token failed"
            )

        print(
            "✅ Dataverse Token Generated"
        )

        print(
            "\n🔐 Generating Graph Token..."
        )

        graph_token = get_access_token(
            auth["graph_scope"]
        )

        if not graph_token:

            raise Exception(
                "Graph token failed"
            )

        print(
            "✅ Graph Token Generated"
        )

        # ==============================
        # STEP 2: READ CONFIG VALUES
        # ==============================

        ticket_column = (
            dv["columns"]["ticket_id"]
        )

        ticket_value = (
            dv["filter"]["ticket_id"]
        )

        invoice_template_path = (
            storage["paths"]
            ["invoice_template_path"]
        )

        # IMPORTANT FIX
        user_email = (
            storage["user_email_dev"]
        )

        print(
            "\n📄 Ticket Column:",
            ticket_column
        )

        print(
            "📄 Ticket Value :",
            ticket_value
        )

        print(
            "\n📁 Template Path:",
            invoice_template_path
        )

        print(
            "\n👤 OneDrive User:",
            user_email
        )

        # ==============================
        # STEP 3: FETCH TEMPLATE FILE
        # ==============================

        template_file = get_onedrive_file(
            graph_token,
            user_email,
            invoice_template_path
        )

        # ==============================
        # STEP 4: FETCH DATAVERSE TABLES
        # ==============================

        print(
            "\n📥 Fetching Closing Ticket Details..."
        )

        closing_ticket_data = fetch_table(
            dv["tables"]["closing_ticket_details"],
            dv_token,
            ticket_column,
            ticket_value
        )

        print(
            "✅ Closing Ticket Details Retrieved"
        )

        print(
            "\n📥 Fetching Invoice Details..."
        )

        invoice_data = fetch_table(
            dv["tables"]["invoice_details"],
            dv_token,
            ticket_column,
            ticket_value
        )

        print(
            "✅ Invoice Details Retrieved"
        )

        # ==============================
        # STEP 5: CREATE DATAFRAMES
        # ==============================

        df_closing_ticket_details = (
            pd.DataFrame(
                closing_ticket_data
            )
        )

        df_invoice_details = (
            pd.DataFrame(
                invoice_data
            )
        )

        print(
            "\n=== CLOSING TICKET DETAILS ==="
        )

        print(
            df_closing_ticket_details.head()
        )

        print(
            "\n=== INVOICE DETAILS ==="
        )

        print(
            df_invoice_details.head()
        )

        logger.info(
            "✅ ===== PROCESS COMPLETE ====="
        )

        print(
            "\n✅ ALL OPERATIONS COMPLETED"
        )

        return (
            df_closing_ticket_details,
            df_invoice_details
        )

    except Exception as e:

        logger.critical(
            "❌ MAIN PROCESS FAILED"
        )

        logger.critical(
            traceback.format_exc()
        )

        print(
            "\n❌ MAIN ERROR:"
        )

        print(e)

        return None, None


# ==============================
# ▶️ RUN SCRIPT
# ==============================

if __name__ == "__main__":

    main()