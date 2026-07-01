import os
import json
import base64
import traceback
from datetime import datetime
from zoneinfo import ZoneInfo

import pandas as pd
import requests
from arcgis.gis import GIS


ARCGIS_URL = "https://www.arcgis.com"
SURVEY_NAME_OR_ID = "Individual Daily Activity Report 214 - NYC Cool"

OUTPUT_FILE = "Individual_Daily_Activity_Report_214_NYC_Cool.xlsx"
EASTERN = ZoneInfo("America/New_York")


def status(message):
    print(f"[{datetime.now(EASTERN):%Y-%m-%d %H:%M:%S %Z}] {message}")


def should_run_now():
    """
    GitHub cron runs in UTC and does not handle Eastern DST automatically.
    Workflow runs at both possible UTC hours, and this prevents duplicate emails.
    """
    now_et = datetime.now(EASTERN)
    return now_et.hour == 22


def convert_arcgis_dates_to_eastern(df):
    """
    Converts ArcGIS UTC date/time fields to Eastern Time for Excel output.
    """

    date_fields = [
        "CreationDate",
        "created_date",
        "createddate",
        "EditDate",
        "editdate",
        "last_edited_date",
    ]

    date_field_lookup = {field.lower() for field in date_fields}

    for col in df.columns:
        if col.lower() in date_field_lookup:
            df[col] = pd.to_datetime(df[col], errors="coerce", utc=True)
            df[col] = df[col].dt.tz_convert(EASTERN)
            df[col] = df[col].dt.strftime("%Y-%m-%d %I:%M %p %Z")

    return df


def connect_arcgis():
    gis = GIS(
        ARCGIS_URL,
        username=os.environ["ARCGIS_USERNAME"],
        password=os.environ["ARCGIS_PASSWORD"]
    )

    if not gis.users.me:
        raise Exception("ArcGIS login failed.")

    status(f"Connected to ArcGIS as {gis.users.me.username}")
    return gis


def find_survey_item(gis):
    item = gis.content.get(SURVEY_NAME_OR_ID)

    if item:
        status(f"Found survey by item ID: {item.title}")
        return item

    items = gis.content.search(
        query=f'title:"{SURVEY_NAME_OR_ID}"',
        item_type="Form",
        max_items=10
    )

    if not items:
        raise Exception(f"Survey123 form not found: {SURVEY_NAME_OR_ID}")

    for i, item in enumerate(items, start=1):
        status(f"{i}. {item.title} | {item.id} | {item.owner}")

    selected = items[0]
    status(f"Using survey: {selected.title}")
    return selected


def get_feature_layer(survey_item):
    related = survey_item.related_items(
        rel_type="Survey2Service",
        direction="forward"
    )

    if not related:
        related = survey_item.related_items(
            rel_type="Service2Survey",
            direction="reverse"
        )

    if not related:
        raise Exception("Could not find related feature service.")

    service_item = related[0]

    if not service_item.layers:
        raise Exception("Related feature service has no layers.")

    layer = service_item.layers[0]
    status(f"Using layer: {layer.properties.name}")
    return layer


def pull_all_records(layer):
    status("Pulling all survey records...")

    features = layer.query(
        where="1=1",
        out_fields="*",
        return_geometry=True,
        return_all_records=True
    )

    df = features.sdf

    if df.empty:
        status("No records found.")
        return df

    df = convert_arcgis_dates_to_eastern(df)

    # Convert geometry to text if present so Excel export is cleaner.
    if "SHAPE" in df.columns:
        df["SHAPE_JSON"] = df["SHAPE"].apply(
            lambda x: json.dumps(x) if x is not None else None
        )
        df = df.drop(columns=["SHAPE"])

    status(f"Pulled {len(df)} record(s).")
    return df


def save_excel(df):
    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Raw Records", index=False)

    status(f"Saved Excel file: {OUTPUT_FILE}")
    return OUTPUT_FILE


def get_graph_token():
    tenant_id = os.environ["MS_TENANT_ID"]
    client_id = os.environ["MS_CLIENT_ID"]
    client_secret = os.environ["MS_CLIENT_SECRET"]

    token_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"

    payload = {
        "client_id": client_id,
        "scope": "https://graph.microsoft.com/.default",
        "client_secret": client_secret,
        "grant_type": "client_credentials",
    }

    response = requests.post(token_url, data=payload, timeout=30)
    response.raise_for_status()

    return response.json()["access_token"]


def send_email_with_attachment(file_path, record_count):
    sender = os.environ["EMAIL_SENDER"]
    recipients = [
        email.strip()
        for email in os.environ["EMAIL_RECIPIENTS"].split(",")
        if email.strip()
    ]

    if not recipients:
        raise Exception("No EMAIL_RECIPIENTS provided.")

    token = get_graph_token()

    with open(file_path, "rb") as f:
        attachment_content = base64.b64encode(f.read()).decode("utf-8")

    today = datetime.now(EASTERN).strftime("%Y-%m-%d")

    message = {
        "message": {
            "subject": f"NYC Cool Individual Daily Activity Report 214 - {today}",
            "body": {
                "contentType": "Text",
                "content": (
                    "Attached is the Individual Daily Activity Report 214 - NYC Cool export.\n\n"
                    f"Record count: {record_count}\n"
                    f"Generated: {datetime.now(EASTERN):%Y-%m-%d %I:%M %p %Z}"
                ),
            },
            "toRecipients": [
                {"emailAddress": {"address": email}}
                for email in recipients
            ],
            "attachments": [
                {
                    "@odata.type": "#microsoft.graph.fileAttachment",
                    "name": os.path.basename(file_path),
                    "contentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    "contentBytes": attachment_content,
                }
            ],
        },
        "saveToSentItems": "true",
    }

    url = f"https://graph.microsoft.com/v1.0/users/{sender}/sendMail"

    response = requests.post(
        url,
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
        },
        json=message,
        timeout=60,
    )

    response.raise_for_status()
    status(f"Email sent to: {', '.join(recipients)}")


def main():
    status("Script started.")

    try:
        # Prevent duplicate sends because GitHub runs both possible UTC hours.
        if os.environ.get("ENFORCE_10PM_ET", "true").lower() == "true":
            if not should_run_now():
                status("Not 10 PM Eastern. Exiting without sending.")
                return

        gis = connect_arcgis()
        survey_item = find_survey_item(gis)
        layer = get_feature_layer(survey_item)

        df = pull_all_records(layer)

        if df.empty:
            status("No records to email.")
            return

        file_path = save_excel(df)
        send_email_with_attachment(file_path, len(df))

        status("Script completed successfully.")

    except Exception:
        status("Script failed.")
        traceback.print_exc()
        raise


if __name__ == "__main__":
    main()
