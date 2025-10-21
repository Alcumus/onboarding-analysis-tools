#!/usr/bin/env python3
import pandas as pd
from fuzzywuzzy import fuzz
import sys

def clean_company_name(company_name):
    """Clean company name for fuzzy matching"""
    if not company_name:
        return ""
    
    # Convert to lowercase and handle special characters
    clean_name = str(company_name).lower()
    
    # Remove common business suffixes and words
    replacements = [
        ('incorporated', ''), ('incorporation', ''), ('inc.', ''), ('inc', ''),
        ('limited', ''), ('ltd.', ''), ('ltd', ''), ('ltée', ''), ('ltee', ''),
        ('corporation', ''), ('corp.', ''), ('corp', ''), ('company', ''), 
        ('co.', ''), ('co', ''), ('société', ''), ('societe', ''),
        ('enterprises', ''), ('entreprises', ''), ('enterprise', ''), ('entreprise', ''),
        ('services', ''), ('service', ''), ('solutions', ''), ('solution', ''),
        ('systems', ''), ('system', ''), ('technologies', ''), ('technology', ''),
        ('tech', ''), ('group', ''), ('groupe', ''), ('international', ''),
        ('canada', ''), ('québec', ''), ('quebec', ''), ('ontario', ''),
        ('&', 'and'), ('o/a', ''), ('dba', ''), ('d/b/a', '')
    ]
    
    for old, new in replacements:
        clean_name = clean_name.replace(old, new)
    
    # Remove numbers, special characters, and extra spaces
    import re
    clean_name = re.sub(r'[0-9\-\(\)\.\/\\\,\;\:\!\@\#\$\%\^\&\*\+\=\[\]\{\}\|\<\>\?\_]', ' ', clean_name)
    clean_name = re.sub(r'\s+', ' ', clean_name).strip()
    
    return clean_name

# Read CBX data
print("Reading CBX data...")
cbx_df = pd.read_csv('data/Oct3.csv', encoding='utf-8-sig')
print(f"CBX data loaded: {len(cbx_df)} records")

# Read HC data
print("Reading QSLMAUDE1 data...")
hc_df = pd.read_excel('data/QSLMAUDE1.xlsx')
print(f"HC data loaded: {len(hc_df)} records")

# Find FMS USITECH
fms_row = hc_df[hc_df['contractor_name'].str.contains('FMS USITECH', na=False, case=False)]
if fms_row.empty:
    print("FMS USITECH not found!")
    sys.exit(1)

fms_company = fms_row.iloc[0]['contractor_name']
print(f"\nSearching for matches to: {fms_company}")

clean_fms = clean_company_name(fms_company)
print(f"Cleaned FMS company: '{clean_fms}'")

# Test fuzzy matching against CBX data
print("\nTesting fuzzy matching...")
matches = []

for idx, cbx_row in cbx_df.iterrows():
    cbx_company_en = str(cbx_row.get('company_name_en', ''))
    cbx_company_fr = str(cbx_row.get('company_name_fr', ''))
    
    for cbx_company in [cbx_company_en, cbx_company_fr]:
        if cbx_company and cbx_company != 'nan':
            clean_cbx = clean_company_name(cbx_company)
            
            if clean_cbx.strip():  # Only test non-empty names
                ratio = fuzz.ratio(clean_fms, clean_cbx)
                token_sort = fuzz.token_sort_ratio(clean_fms, clean_cbx)
                token_set = fuzz.token_set_ratio(clean_fms, clean_cbx)
                partial = fuzz.partial_ratio(clean_fms, clean_cbx)
                max_score = max(ratio, token_sort, token_set, partial)
                
                if max_score >= 50:  # Lower threshold to see more matches
                    matches.append({
                        'cbx_company': cbx_company,
                        'clean_cbx': clean_cbx,
                        'ratio': ratio,
                        'token_sort': token_sort,
                        'token_set': token_set,
                        'partial': partial,
                        'max_score': max_score
                    })

# Sort by max score
matches.sort(key=lambda x: x['max_score'], reverse=True)

print(f"\nFound {len(matches)} potential matches (score >= 50):")
for match in matches[:10]:  # Top 10
    print(f"Score {match['max_score']:3d}: {match['cbx_company']}")
    print(f"           Clean: '{match['clean_cbx']}'")
    print(f"           Ratios: R={match['ratio']} TS={match['token_sort']} TSe={match['token_set']} P={match['partial']}")
    print()

if matches:
    best_match = matches[0]
    print(f"Best match: {best_match['cbx_company']} (Score: {best_match['max_score']})")
    
    if best_match['max_score'] >= 70:
        print("This would trigger a MATCH")
    elif best_match['max_score'] >= 50:
        print("This might trigger AMBIGUOUS matching")
    else:
        print("This would be NO MATCH")
else:
    print("No matches found - should trigger company-only matching logic")