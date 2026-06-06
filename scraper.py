import requests
from bs4 import BeautifulSoup
import json
import os
from collections import defaultdict
from datetime import date

headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
}

CATEGORIES = {
    "Optometry": "193",
    "Audiology": "194",
    "Home Visits": "195",
    "Store Support": "196"
}

# Jobs containing these keywords get excluded entirely - no email sent
EXCLUDED_KEYWORDS = [
    "community champion",
    "receptionist",
    "call centre",
    "scheduling assistant",
    "meet & greet",
    "front of house",
    "customer service",
    "manufacturing assistant",
    "spectacle technician",
    "pre screener",
    "audiology admin",
    "personal assistant",
    "office manager",
]

# Jobs containing these keywords get a trainee-angle email
TRAINEE_KEYWORDS = [
    "trainee optical assistant",
    "optical assistant apprentice",
    "trainee lab",
]

def get_jobs_from_page(category_id, page_number):
    url = f"https://join.specsavers.com/uk/jobs?options={category_id}%2C{category_id}&page={page_number}"
    response = requests.get(url, headers=headers)
    soup = BeautifulSoup(response.text, "html.parser")
    job_links = soup.find_all("a", href=True)
    jobs = []
    seen_urls = []
    for link in job_links:
        href = link["href"]
        if "/uk/job/" in href and href not in seen_urls:
            job_id = href.split("-jid-")[-1]
            title = link.get_text(strip=True)
            if title and title != "Find out more":
                jobs.append({
                    "id": job_id,
                    "title": title,
                    "url": "https://join.specsavers.com" + href
                })
                seen_urls.append(href)
    return jobs

def get_location(url):
    slug = url.split('/uk/job/')[-1]
    if '-in-' in slug and '-jid-' in slug:
        after_in = slug.split('-in-', 1)[1]
        location = after_in.split('-jid-')[0]
        return location
    return "unknown"

def tag_job(job):
    title_lower = job['title'].lower()
    # Check excluded first - takes priority
    for keyword in EXCLUDED_KEYWORDS:
        if keyword in title_lower:
            return "excluded"
    # Then check trainee
    for keyword in TRAINEE_KEYWORDS:
        if keyword in title_lower:
            return "trainee"
    # Everything else is standard
    return "standard"

def deduplicate_jobs(jobs):
    # Same title + same location: keep highest ID (newest posting)
    best = {}
    for job in jobs:
        loc = get_location(job['url'])
        key = f"{job['title'].lower().strip()}|{loc}"
        if key not in best or int(job['id']) > int(best[key]['id']):
            best[key] = job
    return list(best.values())

# --- STEP 1: SCRAPE EVERYTHING ---
print("Scraping all categories...")
all_jobs = []
seen_ids_this_run = set()

for category_name, category_id in CATEGORIES.items():
    print(f"\n  Category: {category_name}")
    page = 1
    while True:
        print(f"    Fetching page {page}...")
        jobs = get_jobs_from_page(category_id, page)
        if not jobs:
            print(f"    Empty page - moving to next category.")
            break
        new_this_page = [j for j in jobs if j["id"] not in seen_ids_this_run]
        if not new_this_page:
            print(f"    Page {page} all duplicates - end of category.")
            break
        for job in new_this_page:
            seen_ids_this_run.add(job["id"])
            all_jobs.append(job)
        page += 1

print(f"\nTotal jobs scraped today: {len(all_jobs)}")

# Save full snapshot
with open("all_jobs_full.json", "w") as f:
    json.dump(all_jobs, f, indent=2)

# --- STEP 2: FIND NEW JOBS ---
storage_file = "seen_jobs.json"
if os.path.exists(storage_file):
    with open(storage_file, "r") as f:
        seen_job_ids = json.load(f)
    print(f"Previously seen IDs: {len(seen_job_ids)}")
else:
    seen_job_ids = []
    print("No previous data - this is the first run.")

raw_new_jobs = [job for job in all_jobs if job["id"] not in seen_job_ids]
print(f"Raw new jobs found: {len(raw_new_jobs)}")

# --- STEP 3: SAVE ALL IDs TO MEMORY ---
all_ids = list(set(seen_job_ids + [job["id"] for job in all_jobs]))
with open(storage_file, "w") as f:
    json.dump(all_ids, f)
print(f"Memory updated: {len(all_ids)} total IDs stored.")

# --- STEP 4: TAG AND FILTER ---
tagged_jobs = []
excluded_count = 0
for job in raw_new_jobs:
    tag = tag_job(job)
    if tag == "excluded":
        excluded_count += 1
    else:
        job['email_type'] = tag
        tagged_jobs.append(job)

print(f"After filtering: {len(tagged_jobs)} jobs kept, {excluded_count} excluded.")
print(f"  Standard: {sum(1 for j in tagged_jobs if j['email_type'] == 'standard')}")
print(f"  Trainee:  {sum(1 for j in tagged_jobs if j['email_type'] == 'trainee')}")

# --- STEP 5: DEDUPLICATE ---
deduped_jobs = deduplicate_jobs(tagged_jobs)
print(f"After deduplication: {len(deduped_jobs)} jobs.")

# --- STEP 6: GROUP BY LOCATION ---
by_location = defaultdict(list)
for job in deduped_jobs:
    loc = get_location(job['url'])
    by_location[loc].append(job)

# --- STEP 7: SAVE TO new_jobs_today.json ---
today = date.today().isoformat()
output = {
    "date": today,
    "total_new_jobs": len(deduped_jobs),
    "total_stores": len(by_location),
    "stores": {}
}

for loc, loc_jobs in sorted(by_location.items()):
    output["stores"][loc] = {
        "job_count": len(loc_jobs),
        "jobs": loc_jobs
    }

with open("new_jobs_today.json", "w") as f:
    json.dump(output, f, indent=2)

# --- PRINT SUMMARY ---
if deduped_jobs:
    print(f"\nNew jobs for {today} ({len(by_location)} stores to contact):")
    for loc, loc_jobs in sorted(by_location.items()):
        print(f"\n  {loc} ({len(loc_jobs)} role{'s' if len(loc_jobs) > 1 else ''}):")
        for job in loc_jobs:
            print(f"    [{job['email_type'].upper()}] {job['title']}")
else:
    print(f"\nNo new jobs today ({today}).")

print("\nDone. new_jobs_today.json is ready for the email step.")