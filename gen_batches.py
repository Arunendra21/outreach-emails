import csv
import os
import re

# Splits emails.csv into batches of 100 for easy mail-merge sending.
# Each batch CSV has: #, Company, Email, Person, Title, Region, Notes
# Picks the BEST email per company: HR/CEO direct email if present, else general contact.

BATCH_SIZE = 100

def region_of(notes, company_id):
    text = notes.lower()
    if any(k in text for k in ("bangalore", "bengaluru", "bellandur", "koramangala", "hsr layout", "indiranagar", "whitefield")):
        return "Bangalore"
    if any(k in text for k in ("gandhinagar", "gift city", "ahmedabad", "kudasan", "infocity", "sargasan", "motera", "rajkot")):
        return "Gandhinagar/Ahmedabad"
    if any(k in text for k in ("dubai", "uae", "sharjah", "abu dhabi")):
        return "Dubai/UAE"
    try:
        n = int(str(company_id).rstrip("b"))
        if n <= 143:
            return "Dubai/UAE"
        if n <= 221:
            return "Gandhinagar/GIFT"
        return "Bangalore"
    except ValueError:
        return "Other"

seen = set()
rows = []
with open("emails.csv", encoding="utf-8") as f:
    for r in csv.DictReader(f):
        direct = r["HR Person Direct Email"].strip()
        general = r["General Contact Email"].strip()
        email = direct if "@" in direct and "(" not in direct and "*" not in direct else general
        email = email.strip()
        if "@" not in email:
            continue
        key = email.lower()
        if key in seen:
            continue
        seen.add(key)
        rows.append({
            "#": r["#"],
            "Company": r["Company"],
            "Email": email,
            "Person": r["HR Person Name"] or "Hiring Team",
            "Title": r["HR Person Title"] or "—",
            "Region": region_of(r["Notes"], r["#"]),
            "Notes": r["Notes"][:120],
        })

for old in os.listdir("."):
    if re.match(r"batch_\d+\.csv$", old):
        os.remove(old)

total = 0
for i in range(0, len(rows), BATCH_SIZE):
    batch = rows[i:i + BATCH_SIZE]
    name = f"batch_{i // BATCH_SIZE + 1:02d}.csv"
    with open(name, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=["#", "Company", "Email", "Person", "Title", "Region", "Notes"])
        w.writeheader()
        w.writerows(batch)
    total += len(batch)
    print(f"{name}: {len(batch)} emails")

print(f"TOTAL: {total} unique emails across {(len(rows) + BATCH_SIZE - 1) // BATCH_SIZE} batches")
