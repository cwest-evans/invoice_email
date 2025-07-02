import pandas as pd
from collections import defaultdict, OrderedDict
import pyodbc
import os
import smtplib
from jinja2 import Environment, FileSystemLoader
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime
from dotenv import load_dotenv
import msal
import requests


load_dotenv()

# SQL Server connection settings
conn_str = (
    f"DRIVER={{{os.getenv('SQL_DRIVER')}}};"
    f"SERVER={os.getenv('SQL_SERVER')};"
    f"DATABASE={os.getenv('SQL_DATABASE')};"
    f"UID={os.getenv('SQL_UID')};"
    f"PWD={os.getenv('SQL_PWD')};"
)

conn = pyodbc.connect(conn_str)


# ----- Authenticate with Microsoft Graph API -----
def send_email_via_graph(to_address, subject, html_body):
    client_id = os.getenv("GRAPH_CLIENT_ID")
    tenant_id = os.getenv("GRAPH_TENANT_ID")
    client_secret = os.getenv("GRAPH_CLIENT_SECRET")
    sender = os.getenv("GRAPH_SENDER")

    authority = f"https://login.microsoftonline.com/{tenant_id}"
    scopes = ["https://graph.microsoft.com/.default"]

    app = msal.ConfidentialClientApplication(
        client_id, authority=authority, client_credential=client_secret
    )

    token = app.acquire_token_for_client(scopes=scopes)
    if "access_token" not in token:
        raise Exception(f"Token acquisition failed: {token.get('error_description')}")

    email_payload = {
        "message": {
            "subject": subject,
            "body": {"contentType": "HTML", "content": html_body},
            "toRecipients": [{"emailAddress": {"address": to_address}}],
        },
        "saveToSentItems": "true",
    }

    headers = {
        "Authorization": f"Bearer {token['access_token']}",
        "Content-Type": "application/json",
    }

    response = requests.post(
        "https://graph.microsoft.com/v1.0/users/" + sender + "/sendMail",
        headers=headers,
        json=email_payload,
    )

    if response.status_code != 202:
        raise Exception(
            f"Graph sendMail failed: {response.status_code} - {response.text}"
        )


# ---- Load SQL queries ----
def load_sql(filename):
    with open(f"sql/{filename}", "r") as file:
        return file.read()


invoice_query = load_sql("unapproved_invoices.sql")
# Load and format data
invoice_df = pd.read_sql(invoice_query, conn)
invoice_df["Invoice Line Total"] = invoice_df["Invoice Line Total"].apply(
    lambda x: f"${x:,.2f}"
)
invoice_df["Date Assigned"] = pd.to_datetime(invoice_df["Date Assigned"]).dt.strftime(
    "%m/%d/%Y"
)
invoice_df["Invoice Date"] = pd.to_datetime(invoice_df["Invoice Date"]).dt.strftime(
    "%m/%d/%Y"
)

# Extract the top 20 rows (or whatever preview size you want)
invoice_preview_df = invoice_df.head(20)

# Build preview HTML table
invoice_preview_table_html = "<table border='0' cellpadding='0' cellspacing='0' style='border-collapse:collapse; width: 100%; font-size: 14px;'>"
invoice_preview_table_html += """
<thead>
<tr>
    <th align="left" style="background-color: #e0e0e0; padding: 10px 12px;">Reviewer</th>
    <th align="left" style="background-color: #e0e0e0; padding: 10px 12px;">Job</th>
    <th align="left" style="background-color: #e0e0e0; padding: 10px 12px;">Date Assigned</th>
    <th align="left" style="background-color: #e0e0e0; padding: 10px 12px;">Invoice Date</th>
    <th align="left" style="background-color: #e0e0e0; padding: 10px 12px;">Vendor Name</th>
    <th align="right" style="background-color: #e0e0e0; padding: 10px 12px;">Invoice Line Total</th>
</tr>
</thead>
<tbody>
"""

for _, row in invoice_preview_df.iterrows():
    invoice_preview_table_html += f"""
    <tr>
        <td style="padding: 10px 12px;">{row['Reviewer']}</td>
        <td style="padding: 10px 12px;">{row['Job']}</td>
        <td style="padding: 10px 12px;">{row['Date Assigned']}</td>
        <td style="padding: 10px 12px;">{row['Invoice Date']}</td>
        <td style="padding: 10px 12px;">{row['Vendor Name']}</td>
        <td align="right" style="padding: 10px 12px;">{row['Invoice Line Total']}</td>
    </tr>
    """

invoice_preview_table_html += "</tbody></table>"

# ---- Set up Jinja2 template environment ----
env = Environment(loader=FileSystemLoader("templates"))
template = env.get_template("template.html")

# ---- Prepare context ----
tables = [
    ("Top 20 Unapproved Invoices (Older than 7 Days)", invoice_preview_table_html)
]
title = "Unapproved Invoices"
date_today = datetime.today().strftime("%m/%d/%Y")


# ---- Render and save preview ----
csv_link = "test"

html_output = template.render(
    title=title,
    tables=tables,
    csv_link=csv_link,
    date_today=date_today,
    total_count=len(invoice_df),
)
filename = "preview_unapproved_invoices.html"
with open(f"test_outputs/{filename}", "w", encoding="utf-8") as f:
    f.write(html_output)
