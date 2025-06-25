import pandas as pd
from collections import defaultdict, OrderedDict
import pyodbc
import os
import smtplib
from jinja2 import Environment, FileSystemLoader
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import datetime
from dotenv import load_dotenv

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

# ---- Load SQL queries ----
def load_sql(filename):
    with open(f"sql/{filename}", "r") as file:
        return file.read()


project_query = load_sql("project_details.sql")
field_query = load_sql("field_metadata.sql")
roster_query = load_sql("job_team_roster.sql")

# ---- Run queries ----
field_metadata = pd.read_sql(field_query, conn)
roster_df = pd.read_sql(roster_query, conn)

email_to_name = (
    roster_df.drop_duplicates(
        subset="EMail"
    )  # in case same person is on multiple projects
    .set_index("EMail")["Name"]
    .to_dict()
)


# Create a mapping: email → list of project IDs
email_to_projects = roster_df.groupby("EMail")["Project"].apply(list).to_dict()

# Set up Jinja2 template environment
env = Environment(loader=FileSystemLoader("templates"))
template = env.get_template("template.html")

for _, row in roster_df.iterrows():
    project_id = row["Project"]
    email = row["EMail"]
    name = row["Name"]

    # Load project data
    project_row = (
        pd.read_sql(project_query, conn, params=[project_id]).iloc[0].to_dict()
    )

    # Build title
    title = f"{project_id} – {name}"

    # Group and format fields
    grouped_fields = defaultdict(list)
    for _, meta in field_metadata.sort_values(by=["Section", "SortOrder"]).iterrows():
        field = meta["FieldName"]
        label = meta["DisplayName"]
        section = meta["Section"]
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
        elif isinstance(value, (datetime.date, datetime.datetime)):
             value = value.strftime("%B %#d, %Y")  # e.g., "June 4, 2025"
        elif isinstance(value, str) and "00:00:00" in value and "-" in value:
            try:
                parsed = pd.to_datetime(value)
                value = parsed.strftime("%B %-d, %Y")
            except:
                pass
        elif value is None:
            value = "—"

        grouped_fields[section].append((label, value))

    # Reorder grouped fields
    section_order = ["Project Description"]
    other_sections = sorted(set(field_metadata["Section"]) - {"Project Description"})
    final_section_order = section_order + other_sections
    grouped_fields = OrderedDict(
        (section, grouped_fields[section])
        for section in final_section_order
        if section in grouped_fields
    )

    # Build HTML table
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

    # Render and save
    html_output = template.render(title=title, tables=tables)

    filename = (
        f"preview_{project_id}_{email.replace('@', '_at_').replace('.', '_')}.html"
    )
    with open(filename, "w", encoding="utf-8") as f:
        f.write(html_output)

    print(f"Saved preview for {name} ({email}) — Project {project_id}")
