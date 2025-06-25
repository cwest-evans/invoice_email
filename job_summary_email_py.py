import pandas as pd
from collections import defaultdict, OrderedDict
import pyodbc
import os
import smtplib
from jinja2 import Environment, FileSystemLoader
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# SQL Server connection settings
conn_str = (
    "DRIVER={ODBC Driver 17 for SQL Server};"
    "SERVER=10.150.29.4;"
    "DATABASE=Viewpoint;"
    "UID=cwest.sql;"
    "PWD=Bluepony12"
)
conn = pyodbc.connect(conn_str)

# FOR TESTING: Set the project ID directly
project_id = "24-2902"

# Step 1: Load the data
project_row = (
    pd.read_sql(
        "SELECT * FROM EGC_ProjectionCoversheet WHERE Project = ?",
        conn,
        params=[project_id],
    )
    .iloc[0]
    .to_dict()
)
field_metadata = pd.read_sql("SELECT * FROM EGC_ProjectionCoversheet_SortOrder", conn)

# Step 2: Group fields into sections
section_order = ["Project Description"]  # force this section first
# Extend with all other sections sorted alphabetically after that
other_sections = sorted(set(field_metadata["Section"]) - {"Project Description"})
final_section_order = section_order + other_sections

from collections import OrderedDict

# Step 1: Group fields
grouped_fields = defaultdict(list)
for _, row in field_metadata.sort_values(by=["Section", "SortOrder"]).iterrows():
    field = row["FieldName"]
    label = row["DisplayName"]
    section = row["Section"]
    value = project_row.get(field)

    if (
        isinstance(value, (float, int))
        and "Percent" not in field
        and "Rate" not in field
    ):
        value = f"${value:,.0f}"
    elif "Percent" in field or "Rate" in field or "Margin" in field:
        try:
            value = f"{float(value):.1%}"
        except:
            value = value
    elif isinstance(value, bool):
        value = "Y" if value else "N"
    elif value is None:
        value = "—"

    grouped_fields[section].append((label, value))

# ✅ Step 2: Reorder grouped_fields after building it
grouped_fields = OrderedDict(
    (section, grouped_fields[section]) for section in final_section_order
)

# Step 1: Set up Jinja2 environment
env = Environment(loader=FileSystemLoader("templates"))
template = env.get_template("template.html")  # This loads /templates/template.html

# Step 2: (your grouped_fields logic)
tables = []
for section, rows in grouped_fields.items():
    html = "<table border='0' cellpadding='0' cellspacing='0' style='border-collapse:collapse;'>"
    for label, val in rows:
        html += f"""
        <tr>
            <th align="left" style="background-color: #e0e0e0; padding: 10px 12px;">{label}</th>
            <td style="background-color: #ffffff; padding: 10px 12px;">{val}</td>
        </tr>
        """

    html += "</table>"
    tables.append((section, html))

# Step 3: Render HTML
html_output = template.render(
    title=project_id,  # pushes variance to the html title 
    tables=tables
)

# Step 4: Optional – Save output locally for testing
with open("test_project_health_output.html", "w", encoding="utf-8") as f:
    f.write(html_output)
