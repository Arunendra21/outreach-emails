#!/usr/bin/env python3
"""Add verified Udaipur company outreach contacts to city_udaipur_01.csv.

Follows CONTRIBUTING.md: validates every email against the project regex,
deduplicates against existing rows, writes CRLF CSV, and prints a summary.
Every address below was verified as printed on the company's own website and
its domain passed an MX check at authoring time.
"""
import csv
import os
import re

EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")
OUTPUT_FILE = "city_udaipur_01.csv"
HEADER = ["#", "Company", "Email", "Person", "Title", "City", "Notes"]

# (Company, Email, Person, Title, City, Notes)
ROWS = [
    ("WebSenor", "sales@websenor.com", "Sales Desk", "General", "Udaipur",
     "Udaipur software/web/app dev; VERIFIED printed on websenor.com/contact-us; MX OK"),
    ("Go-tech Solution", "career@go-techsolution.com", "Careers Desk", "Recruitment", "Udaipur",
     "Udaipur web/app dev; VERIFIED printed on go-techsolution.com/contact-us; MX OK"),
    ("Go-tech Solution", "projects@go-techsolution.com", "Projects Desk", "General", "Udaipur",
     "Udaipur web/app dev; VERIFIED printed on go-techsolution.com/contact-us; MX OK"),
    ("Webtech Softwares", "info@webtechsoftwares.com", "Corporate Office", "General", "Udaipur",
     "Udaipur software/web dev; VERIFIED printed on webtechsoftwares.com; MX OK"),
    ("Elixir Technologies", "sales@elixirinfo.com", "Sales Desk", "General", "Udaipur",
     "Udaipur digital marketing/web design (Udaipur office); VERIFIED printed on elixirinfo.com; MX OK"),
    ("Matrix Web Infotech", "info@matrixwebinfotech.com", "Corporate Office", "General", "Udaipur",
     "Udaipur digital marketing/web dev; VERIFIED printed on matrixwebinfotech.com; MX OK"),
]


def load_existing(path: str) -> set:
    existing = set()
    if os.path.exists(path):
        with open(path, newline="") as f:
            for row in csv.reader(f):
                if len(row) >= 3 and row[2].strip():
                    existing.add(row[2].strip().lower())
    return existing


def main() -> None:
    existing = load_existing(OUTPUT_FILE)
    new_file = not os.path.exists(OUTPUT_FILE)
    written = skipped_invalid = skipped_dup = 0

    with open(OUTPUT_FILE, "a", newline="") as f:
        writer = csv.writer(f)  # csv default lineterminator is CRLF, matching repo files
        if new_file:
            writer.writerow(HEADER)
        for company, email, person, title, city, notes in ROWS:
            e = email.strip()
            if not EMAIL_RE.match(e):
                print(f"[INVALID]   {e}")
                skipped_invalid += 1
            elif e.lower() in existing:
                print(f"[DUPLICATE] {e}")
                skipped_dup += 1
            else:
                writer.writerow(["", company, e, person, title, city, notes])
                existing.add(e.lower())
                written += 1

    print(f"Done {OUTPUT_FILE}. Written: {written}, "
          f"Invalid: {skipped_invalid}, Duplicates: {skipped_dup}")


if __name__ == "__main__":
    main()
