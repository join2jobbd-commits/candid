# -*- coding: utf-8 -*-
"""
Convert master_database.xlsx -> data.json
Maps all Excel columns to the JSON keys expected by index.html.
"""
import pandas as pd
import json
import os

EXCEL_FILE = 'master_database.xlsx'
OUTPUT_FILE = 'data.json'

print(f"🔄 Reading {EXCEL_FILE}...")

# Try reading with header on row 0; if 'Candidate ID' is missing, try row 1 (skips "DATABASE" title)
df = pd.read_excel(EXCEL_FILE)
if 'Candidate ID' not in df.columns:
    print("⚠️  Title row detected. Re-reading with header=1...")
    df = pd.read_excel(EXCEL_FILE, header=1)

# Exact column mapping: Excel header -> JSON key
COLUMN_MAP = {
    'Candidate ID':                            'id',
    'Name':                                    'name',
    'Gender':                                  'gender',
    'Last Update':                             'last_update',
    'Special Degree Like (CA/FCA/IBA/Phd)':    'special_qual',
    'Special Degree Institute Name':           'special_degree_inst',
    'Last degree':                             'last_degree',
    'Subject':                                 'subject',
    'Last Degree Year':                        'last_degree_year',
    'Last degree Institute':                   'last_degree_inst',
    'Army':                                    'army',
    'University Local or Foregin':             'degree_local_foreign',
    'Phone':                                   'phone',
    'Email':                                   'email',
    'Country':                                 'country',
    'Location Preferred':                      'location_pref',
    'Industry Experience':                     'industry_exp',
    'Prefered Industry':                       'prefered_industry',
    'Product Area':                            'product_area',
    'Product Area 2':                          'product_area_2',
    'Department':                              'department',
    'Current Company':                         'current_org',
    'Current Designations':                    'current_position',
    'Previous Company , Designation & Length': 'previous',
    'Salary Expectation':                      'salary_expect',
    'Current Salary':                          'current_salary',
    'Date of Birth':                           'dob',
    'Total Experience (Years)':                'total_exp',
    'Join date (Last Company)':                'join_date',
    'Academic Qualification':                  'academic_qual',
    'Training/Certification':                  'certification',
    'Core Competencies':                       'core_competencies',
    'TECHNICAL SKILLS':                        'technical_skills',
    'Language':                                'language',
    'PERSONAL DETAILS':                        'personal_details',
    'Key Achievements & Responsibilities:':    'achievements',
    'Management Level':                        'mgmt_level',
}

# Rename only columns that actually exist
rename_dict = {k: v for k, v in COLUMN_MAP.items() if k in df.columns}
df = df.rename(columns=rename_dict)

print(f"✅ Mapped {len(rename_dict)} columns.")
missing = [k for k in COLUMN_MAP if k not in df.columns]
if missing:
    print(f"⚠️  These Excel columns were not found (skipped): {missing}")

# Clean helper
def clean(val):
    if pd.isna(val):
        return ""
    s = str(val).strip()
    return "" if s.lower() == 'nan' else s

# Build records
records = []
for _, row in df.iterrows():
    rec = {
        "id":                  clean(row.get('id', '')),
        "name":                clean(row.get('name', '')),
        "gender":              clean(row.get('gender', '')).lower(),
        "special_qual":        clean(row.get('special_qual', '')),
        "special_degree_inst": clean(row.get('special_degree_inst', '')),
        "last_degree":         clean(row.get('last_degree', '')),
        "subject":             clean(row.get('subject', '')),
        "last_degree_year":    clean(row.get('last_degree_year', '')),
        "last_degree_inst":    clean(row.get('last_degree_inst', '')),
        "army":                clean(row.get('army', '')).lower(),
        "degree_local_foreign":clean(row.get('degree_local_foreign', '')).lower(),
        "phone":               clean(row.get('phone', '')),
        "email":               clean(row.get('email', '')),
        "country":             clean(row.get('country', '')),
        "location_pref":       clean(row.get('location_pref', '')),
        "industry_exp":        clean(row.get('industry_exp', '')),
        "prefered_industry":   clean(row.get('prefered_industry', '')),
        "product_area":        clean(row.get('product_area', '')),
        "product_area_2":      clean(row.get('product_area_2', '')),
        "department":          clean(row.get('department', '')),
        "current_org":         clean(row.get('current_org', '')),
        "current_position":    clean(row.get('current_position', '')),
        "previous":            clean(row.get('previous', '')),
        "salary_expect":       clean(row.get('salary_expect', '')),
        "current_salary":      clean(row.get('current_salary', '')),
        "dob":                 clean(row.get('dob', '')),
        "total_exp":           clean(row.get('total_exp', '')),
        "join_date":           clean(row.get('join_date', '')),
        "academic_qual":       clean(row.get('academic_qual', '')),
        "certification":       clean(row.get('certification', '')),
        "core_competencies":   clean(row.get('core_competencies', '')),
        "technical_skills":    clean(row.get('technical_skills', '')),
        "language":            clean(row.get('language', '')),
        "personal_details":    clean(row.get('personal_details', '')),
        "achievements":        clean(row.get('achievements', '')),
        "mgmt_level":          clean(row.get('mgmt_level', '')),
    }
    if rec["id"]:
        records.append(rec)

# Save
with open(OUTPUT_FILE, 'w', encoding='utf-8') as f:
    json.dump({"records": records}, f, ensure_ascii=False, indent=2)

print(f"\n🎉 SUCCESS! Created {OUTPUT_FILE} with {len(records)} candidates.")
print(f"📁 File size: {os.path.getsize(OUTPUT_FILE)/1024/1024:.2f} MB")