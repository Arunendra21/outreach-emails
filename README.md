# Internship Outreach - Email Database

Private repository containing all scraped company emails and HR contacts for paid internship outreach.

## Files

- `emails.csv` — 143 companies with general contact emails, categories, priority levels
- `hr_contacts.csv` — 38 named HR/TA contacts with direct emails and LinkedIn profiles
- `make_sheet.py` — Python script to generate formatted Excel spreadsheet

## Stats

- **143** companies (Dubai/UAE/Remote tech)
- **38** named HR contacts with direct emails
- **~85** High-priority targets
- **~15** verified direct HR/careers/recruitment inboxes

## Categories

Fintech, PropTech, HealthTech, EdTech, AI/ML, SaaS, Web Dev, App Dev, FoodTech, Logistics, Payments, Cybersecurity, Digital Marketing

## Usage

```bash
pip install openpyxl
python make_sheet.py
```

Generates `Dubai_Internship_Outreach.xlsx` with 4 sheets:
1. All Companies
2. Priority Dubai Companies (High only)
3. Named HR Contacts
4. Email Template
