"""
Daily digest email — sends new Specsavers jobs to recruiters.
Uses Microsoft Graph with app-only auth (no user login needed).
"""

import json
import os
import requests
import msal
from collections import Counter
from datetime import date
from dotenv import load_dotenv

load_dotenv()

ROB_EMAIL = os.environ.get("ROB_EMAIL", "robert.mould@talentshed.co.uk")
RECIPIENTS_ENV = os.environ.get("RECIPIENTS", "")

REGION_SUFFIXES = sorted([
    "greater-london", "g-london", "east-midlands", "west-midlands",
    "east-england", "east-central", "west-central", "north-east",
    "south-east", "south-west", "north-west", "northern-ireland",
    "north", "south", "east", "west", "scotland", "wales", "ireland",
    "guernsey", "isle-of-man",
], key=len, reverse=True)

ROLE_GROUPS = [
    ("Optometrist",                         ["optometrist"]),
    ("Dispensing Optician",                 ["dispensing optician"]),
    ("Optical Assistant",                   ["optical assistant"]),
    ("Audiologist / Hearing Aid Dispenser", ["audiologist", "hearing aid dispenser"]),
    ("Contact Lens Optician",               ["contact lens optician"]),
    ("Store Manager",                       ["store manager"]),
    ("Assistant Manager",                   ["assistant manager"]),
    ("Lab Technician",                      ["lab technician"]),
    ("Team Leader / Supervisor",            ["team leader", "supervisor"]),
]


def ordinal(n):
    if 11 <= n % 100 <= 13:
        suffix = "th"
    else:
        suffix = {1: "st", 2: "nd", 3: "rd"}.get(n % 10, "th")
    return f"{n}{suffix}"


def format_location(slug):
    slug_lower = slug.lower()
    parts = slug_lower.split("-")
    mid = len(parts) // 2
    if len(parts) % 2 == 0 and parts[:mid] == parts[mid:]:
        return " ".join(parts[:mid]).title()
    for suffix in REGION_SUFFIXES:
        if slug_lower.endswith("-" + suffix):
            return slug_lower[:-len(suffix)-1].replace("-", " ").title()
    return slug_lower.replace("-", " ").title()


def classify_role(title):
    title_lower = title.lower()
    for group_name, keywords in ROLE_GROUPS:
        if any(kw in title_lower for kw in keywords):
            return group_name
    return "Other"


def build_html(formatted_date, all_stores, role_counts):
    total_jobs = sum(len(j) for j in all_stores.values())

    role_bullets = ""
    for role, count in sorted(role_counts.items(), key=lambda x: -x[1]):
        if role != "Other":
            role_bullets += f"&nbsp;&nbsp;&bull; {count} x {role}<br>\n"
    if role_counts.get("Other"):
        role_bullets += f"&nbsp;&nbsp;&bull; {role_counts['Other']} x Other<br>\n"

    html = f"""<html><body style="font-family:Arial,sans-serif;font-size:14px;
color:#333;max-width:800px;margin:0 auto;">

<p>Hi All,</p>

<p>Please see below for a list of all new jobs posted by Specsavers today.
Please check Tracker and contact the store to see if you can help &mdash;
where we&rsquo;re not already working on them.</p>

<p><strong>Today&rsquo;s roles ({total_jobs} jobs):</strong><br>
{role_bullets}</p>

<hr style="border:none;border-top:2px solid #C85A1A;margin:20px 0;">
<h3 style="color:#1B2A4A;margin-bottom:4px;">&#128205; New Jobs by Store</h3>
<p style="color:#666;font-size:12px;margin-top:0;">
Check Tracker before contacting &mdash; confirm no live job or recent
placement first. Home visits / domiciliary roles included &mdash; check
the right territory contact in Tracker.
</p>
"""

    for loc, jobs in sorted(all_stores.items()):
        store_name = format_location(loc)
        html += f"""
<div style="margin-bottom:12px;padding:10px 14px;
border-left:4px solid #C85A1A;background:#fafafa;
border-radius:0 4px 4px 0;">
  <strong style="color:#1B2A4A;">{store_name} Specsavers</strong><br>
  <div style="margin-top:4px;">
"""
        for job in jobs:
            tag = ""
            if job.get("email_type") == "trainee":
                tag = ('&nbsp;<span style="background:#e8f4e8;color:#2a7a2a;'
                       'padding:2px 6px;border-radius:3px;font-size:11px;'
                       'font-weight:bold;">TRAINEE</span>')
            html += (f'&nbsp;&nbsp;&bull; <a href="{job["url"]}" '
                     f'style="color:#C85A1A;text-decoration:none;">'
                     f'{job["title"]}</a>{tag}<br>\n')
        html += "  </div>\n</div>\n"

    html += """
<hr style="border:none;border-top:1px solid #ddd;margin:24px 0 12px;">
<p>Thanks</p>
<p>Rob</p>
<p style="color:#aaa;font-size:11px;margin-top:20px;">
Generated automatically &mdash; runs daily at 7:30am weekdays.
</p>
</body></html>"""

    return html


def get_app_token():
    app = msal.ConfidentialClientApplication(
        client_id=os.environ.get("MS_CLIENT_ID"),
        client_credential=os.environ.get("MS_CLIENT_SECRET"),
        authority=f"https://login.microsoftonline.com/{os.environ.get('MS_TENANT_ID')}",
    )
    result = app.acquire_token_for_client(
        scopes=["https://graph.microsoft.com/.default"]
    )
    if "access_token" not in result:
        raise RuntimeError(f"Auth failed: {result.get('error_description')}")
    return result["access_token"]


# --- Load data ---
with open("new_jobs_today.json") as f:
    data = json.load(f)

all_stores_raw = data["stores"]
today_date = date.today()
formatted_date = f"{ordinal(today_date.day)} {today_date.strftime('%B')}"

# --- Build store list and role counts ---
all_stores = {}
role_counts = Counter()

for loc, store in all_stores_raw.items():
    if store["jobs"]:
        all_stores[loc] = store["jobs"]
        for job in store["jobs"]:
            role_counts[classify_role(job["title"])] += 1

total_jobs = sum(len(j) for j in all_stores.values())

print(f"Date: {formatted_date}")
print(f"Total jobs: {total_jobs} across {len(all_stores)} stores")

# --- Build email ---
html_body = build_html(formatted_date, all_stores, role_counts)
subject = (f"{formatted_date} BD: New Specsavers Jobs to chase "
           f"({total_jobs} jobs)")

# --- Authenticate ---
print("\nAuthenticating with Microsoft...")
token = get_app_token()
print("Token obtained.")

# --- Build recipients ---
to_list = [
    {"emailAddress": {"address": e.strip()}}
    for e in RECIPIENTS_ENV.split(",") if e.strip()
]

# --- Send ---
send_url = f"https://graph.microsoft.com/v1.0/users/{ROB_EMAIL}/sendMail"

payload = {
    "message": {
        "subject": subject,
        "importance": "high",
        "body": {"contentType": "HTML", "content": html_body},
        "toRecipients": to_list,
        
    },
    "saveToSentItems": True,
}

print("Sending email...")
r = requests.post(
    send_url,
    headers={
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    },
    json=payload,
    timeout=30,
)

if r.status_code == 202:
    print(f"\nSent successfully.")
    print(f"Subject: {subject}")
    print(f"To: {RECIPIENTS_ENV}")
    print(f"CC: {ROB_EMAIL}")
else:
    print(f"Error: {r.status_code} — {r.text[:300]}")