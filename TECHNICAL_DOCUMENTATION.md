# Onboarding Analysis Tools - Technical Documentation

## Overview
This document provides a comprehensive technical breakdown of the main processing script (`main.py`) that matches hiring client contractors with existing CBX (ContractorBase) contractors. The script performs fuzzy matching, business logic evaluation, and categorizes contractors into different action categories.

## Script Flow and Decision Tree

### 1. Initialization and Setup

#### Input Processing
- **CBX Data Source**: CSV file containing existing contractors (`args.cbx_list`)
- **HC Data Source**: Excel file containing hiring client contractors (`args.hc_list`)
- **Output**: Multi-sheet Excel workbook with analysis results

#### Configuration Parameters
```python
--min_company_match_ratio (default: 70)    # Company name fuzzy matching threshold
--min_address_match_ratio (default: 80)    # Address fuzzy matching threshold
--additional_generic_domain                 # Custom generic email domains
--additional_generic_name_word             # Custom generic company name words
```

#### Data Validation Checks
```
IF CBX headers != expected CBX headers AND NOT ignore_warnings:
    EXIT with error
    
IF HC headers != expected HC headers AND NOT ignore_warnings:
    EXIT with error

IF file size > (10,000 rows OR 250 columns) AND NOT ignore_warnings:
    EXIT with error
```

### 2. Main Processing Loop - For Each HC Contractor

The script processes each hiring client contractor through a comprehensive decision tree:

#### 2.1 Priority Check: HC_DO_NOT_MATCH Flag
```
IF HC_DO_NOT_MATCH flag is TRUE:
    IF all mandatory fields are present:
        → ACTION: "onboarding" (no CBX matching needed)
        → SKIP fuzzy matching
    ELSE:
        → Attempt company-only fuzzy matching
        IF company match score >= 70:
            → Use business logic with company match
        ELSE:
            → ACTION: "missing_info"
```

#### 2.2 Priority Check: Force CBX ID
```
IF HC_FORCE_CBX_ID is specified:
    Search CBX database for exact ID match
    IF forced CBX contractor found:
        → Use forced match (score = 100)
        → Set analysis = "FORCED MATCH"
        → SKIP fuzzy matching
    ELSE:
        → Log warning about missing forced ID
        → Continue with normal fuzzy matching
```

#### 2.3 Fuzzy Matching Algorithm (if no force override)

##### 2.3.1 Company Name Preprocessing
```python
# Clean and normalize company names
clean_hc_company = clean_company_name(hc_company)
# Removes: accents, punctuation, generic words, legal suffixes
# Preserves: word order (critical for fuzzy matching)
```

##### 2.3.2 Candidate Filtering Loop
For each CBX contractor:
```python
# Calculate fuzzy match scores
ratio_company_en = max(
    fuzz.token_sort_ratio(cbx_company_en, clean_hc_company),
    fuzz.partial_ratio(cbx_company_en, clean_hc_company), 
    fuzz.token_set_ratio(cbx_company_en, clean_hc_company)
)

ratio_company_fr = max(
    fuzz.token_sort_ratio(cbx_company_fr, clean_hc_company),
    fuzz.partial_ratio(cbx_company_fr, clean_hc_company),
    fuzz.token_set_ratio(cbx_company_fr, clean_hc_company)
)

ratio_company = max(ratio_company_en, ratio_company_fr)

# Address matching
cbx_address_normalized = normalize_address(cbx_address)
hc_address_normalized = normalize_address(hc_address)
ratio_address = max(
    fuzz.token_sort_ratio(cbx_address, hc_address_normalized),
    fuzz.partial_ratio(cbx_address, hc_address_normalized),
    fuzz.token_set_ratio(cbx_address, hc_address_normalized)
)
```

##### 2.3.3 Candidate Qualification Criteria
```
IF (ratio_company >= min_company_match_ratio OR name_substring_match):
    ADD to candidates list
    
    # Calculate additional factors:
    contact_match = (email_match AND first_name_match AND last_name_match)
    zip_match_strict = (normalized_postal_codes_match)
    suite_match_strict = (suite/bureau_keywords_match)
    location_bonus = calculate_location_bonus() # Includes country matching penalties
    
    # Business scoring algorithm
    business_score = (
        module_count * 3 +                      # CBX modules (qualification breadth)
        min(hc_count * 0.5, 5) +               # Hiring client relationships (capped)
        (30 if contact_match else 0) +         # Contact verification bonus
        (25 if zip_match_strict else 0) +     # Postal code accuracy bonus  
        (15 if suite_match_strict else 0) +   # Address precision bonus
        location_bonus * 2                     # Geographic proximity bonus
    )
```

#### 2.4 Candidate Selection Algorithm

##### Priority-Based Selection Logic
```
IF candidates exist:
    eligible_candidates = [c for c in candidates if c.name_score >= 70]
    
    IF eligible_candidates exist:
        # Priority 1: Perfect company matches (>=95)
        perfect_matches = [c for c in eligible_candidates if c.name_score >= 95]
        IF perfect_matches:
            → Sort by: (relationship_exists, relationship_depth, contact_match, zip_match, address_score, business_score)
            → SELECT top candidate
        
        # Priority 2: High company matches (>=90) 
        ELSE IF high_matches exist (>=90):
            → Sort by: (relationship_exists, relationship_depth, contact_match, business_score, name_score, address_score)
            → SELECT top candidate
            
        # Priority 3: Good matches with postal code (>=70 + zip match)
        ELSE IF good_matches_with_zip exist:
            → Sort by: (business_score, name_score, address_score)
            → SELECT top candidate
            
        # Priority 4: Best overall combination
        ELSE:
            # Enhanced scoring prioritizing relationships
            FOR each candidate:
                relationship_bonus = hc_count * 20  # 20 points per client relationship
                combined_score = (
                    relationship_bonus +                # Relationship depth (20 pts/client)
                    name_score * 0.3 +                 # Company match (30%)
                    business_score * 0.3 +             # Business factors (30%)
                    address_score * 0.2                # Address match (20%)
                )
            → Sort by: (relationship_exists, combined_score)
            → SELECT top candidate
```

#### 2.5 Match Quality Assessment
```
IF best_match exists:
    IF force_cbx_id was used:
        is_ambiguous = HC_AMBIGUOUS flag value
        contact_match = calculate_contact_match()
        
    ELSE (fuzzy matching):
        IF company_score >= 70 AND address_score < 70:
            is_ambiguous = TRUE  # Good company name but poor address
        ELSE IF company_score < 50:
            best_match = NULL   # Too poor to be considered a match
        
        contact_match = calculate_contact_match()
```

### 3. Business Logic Processing

#### 3.1 Action Determination Function
```python
def action(hc_data, cbx_data, create, subscription_update, expiration_date, is_qualified, ignore):
```

##### Primary Business Rules (in order of evaluation):

###### Rule 1: Missing Information Check
```
IF create == TRUE AND NOT core_mandatory_provided(hc_data):
    → ACTION: "missing_info"
    # Mandatory fields: company, first_name, last_name, email, phone, street, city, country, zip
    # Note: state optional for non-CA/US countries
```

###### Rule 2: Association Fee Processing
```
IF HC_IS_ASSOCIATION_FEE == TRUE AND cbx_data exists:
    IF registration_status IN ('Active', 'Non Member') AND NOT has_relationship:
        IF expiration_date exists:
            IF expiration_date > (today + 60 days):
                → ACTION: "association_fee"
            ELSE:
                → ACTION: "add_questionnaire"
        ELSE:
            → ACTION: "association_fee"
```

###### Rule 3: New Contractor Creation (create=TRUE)
```
IF create == TRUE:
    IF HC_IS_TAKE_OVER == TRUE:
        → ACTION: "activation_link"
    ELSE:
        # Priority fix: Check relationships even for ambiguous contractors
        IF cbx_data exists AND has_relationship:
            IF registration_status == 'Active':
                IF is_qualified:
                    → ACTION: "already_qualified"
                ELSE:
                    → ACTION: "follow_up_qualification"
            ELSE IF registration_status == 'Suspended':
                → ACTION: "restore_suspended"  
            ELSE:
                → ACTION: "add_questionnaire"
        ELSE IF HC_AMBIGUOUS == TRUE:
            → ACTION: "ambiguous_onboarding"
        ELSE IF core_mandatory_provided(hc_data):
            → ACTION: "onboarding"
        ELSE:
            → ACTION: "missing_info"  # Fallback
```

###### Rule 4: Existing Contractor Processing (create=FALSE)
```
IF create == FALSE:
    registration_status = cbx_data.registration_status
    
    IF HC_IS_TAKE_OVER == TRUE:
        IF registration_status == 'Suspended':
            → ACTION: "restore_suspended"
        ELSE IF registration_status == 'Active':
            → ACTION: "add_questionnaire"
        ELSE IF registration_status == 'Non Member':
            → ACTION: "activation_link"
        ELSE:
            → ERROR: Invalid registration status
            
    ELSE (normal processing):
        IF registration_status == 'Active':
            IF has_relationship:
                IF is_qualified:
                    → ACTION: "already_qualified"
                ELSE:
                    → ACTION: "follow_up_qualification"
            ELSE:
                IF subscription_update:
                    → ACTION: "subscription_upgrade"
                ELSE IF HC_IS_ASSOCIATION_FEE AND NOT has_relationship:
                    # Apply association fee logic (same as Rule 2)
                ELSE:
                    → ACTION: "add_questionnaire"
                    
        ELSE IF registration_status == 'Suspended':
            → ACTION: "restore_suspended"
            
        ELSE IF registration_status IN ('Non Member', '', NULL):
            → ACTION: "re_onboarding"
            
        ELSE:
            → ERROR: Invalid registration status
```

### 4. Output Generation

#### 4.1 Analysis String Building
```
IF match found AND NOT ambiguous:
    analysis_string = ">>> SELECTED BEST MATCH: [CBX details]
                      >>> ALL CANDIDATES CONSIDERED:
                      [Detailed candidate analysis with scores]"

ELSE IF match found AND ambiguous:
    analysis_string = "[Candidate analysis with ambiguous match notation]"

ELSE IF company-only match:
    analysis_string = "Company-only matching attempt:
                      [Match details and limitations]"
                      
ELSE:
    analysis_string = "[No match found details]"
```

#### 4.2 Output Row Construction
```
output_row = hc_row[0:HC_HEADER_LENGTH]  # Original HC data

# Add analysis columns based on match quality:
IF good_match (not ambiguous):
    # Populate all CBX analysis columns
    FOR each analysis_header:
        output_row.append(match_data.get(header, ''))
        
ELSE IF ambiguous_match:
    IF has_relationship:
        # Full CBX data population
    ELSE:
        # Limited data (analysis, scores, action only)
        
ELSE (no match):
    # Empty analysis columns except action and index
    FOR each analysis_header:
        IF header IN ('action', 'index', 'analysis'):
            output_row.append(appropriate_value)
        ELSE:
            output_row.append('')

# Add metadata columns if present
IF metadata_indexes exist:
    output_row.extend(metadata_columns)
```

#### 4.3 Multi-Sheet Output Generation

The script creates 15 specialized worksheets:

1. **"all"** - Complete dataset with all contractors
2. **"onboarding"** - New contractors needing full onboarding
3. **"association_fee"** - Contractors requiring association fee processing
4. **"re_onboarding"** - Existing contractors needing re-onboarding
5. **"subscription_upgrade"** - Active contractors needing subscription upgrades
6. **"ambiguous_onboarding"** - Contractors with unclear matching requiring manual review
7. **"restore_suspended"** - Suspended contractors needing restoration
8. **"activation_link"** - Takeover contractors needing activation links
9. **"already_qualified"** - Contractors already qualified with hiring client
10. **"add_questionnaire"** - Active contractors needing questionnaire addition
11. **"missing_info"** - Contractors with incomplete information
12. **"follow_up_qualification"** - Contractors needing qualification follow-up
13. **"Data to import"** - Formatted data for system import (add_questionnaire subset)
14. **"Existing Contractors"** - All non-onboarding contractors for reference
15. **"Data for HS"** - HubSpot-formatted data export

### 5. Data Quality and Normalization

#### 5.1 Address Normalization
```python
# Unicode normalization (removes accents)
normalized = unicodedata.normalize('NFKD', address)
normalized = ''.join([c for c in normalized if not unicodedata.combining(c)])

# French/English standardization
translations = {
    'chemin': 'road', 'rue': 'street', 'route': 'road',
    'boulevard': 'boulevard', 'blvd': 'boulevard',
    'ouest': 'west', 'o': 'west', 'est': 'east', 'e': 'east',
    # ... additional translations
}

# Suite/office information preservation
suite_keywords = ['app', 'apartment', 'suite', 'bureau', 'office', 'room', 'unit']
```

#### 5.2 Company Name Cleaning  
```python
# Remove generic legal suffixes
generic_suffixes = ['inc', 'ltd', 'ltée', 'co', 'corp', 'llc', 'enr', 
                   'industriel', 'services', 'solutions', 'systems']

# Preserve word order (critical for fuzzy matching)
# Remove duplicates while maintaining sequence
```

#### 5.3 Postal Code Normalization
```python
# Remove all non-alphanumeric characters
# Convert to uppercase
# Unicode normalization for international codes
```

### 6. Error Handling and Validation

#### 6.1 Data Consistency Checks
- Currency/country validation (CAD for CA, USD for others)
- Phone number parsing and extension extraction
- Date format validation and conversion
- Header column count validation

#### 6.2 Fuzzy Matching Quality Control
- Minimum score thresholds prevent poor matches
- Business scoring prevents high-volume contractors from dominating selection
- Location penalties for international mismatches
- Contact verification bonuses for identity confirmation

#### 6.3 Action Logic Safeguards
- Missing information checks prevent incomplete processing
- Relationship prioritization ensures existing connections are maintained
- Ambiguous flagging for manual review of uncertain matches

## Performance Considerations

### Optimization Strategies
1. **Pre-filtering**: Length and character-based filtering before expensive fuzzy matching
2. **Candidate limiting**: Only detailed analysis for scores >= 70%
3. **Business scoring**: Weighted factors to prioritize high-value matches
4. **Memory management**: Streaming Excel processing for large datasets

### Scalability
- Supports up to 10,000 HC contractors and 250 columns
- Parallel processing capability via `parallel_process.sh`
- Configurable thresholds for different matching requirements

## Common Decision Points

### When to Use Force CBX ID
- Manual override for known relationships
- Correction of algorithm mistakes
- Special business circumstances

### When Contractors Are Marked Ambiguous
- Company match >= 70% but address match < 70%
- Multiple high-scoring candidates with similar scores
- Conflicting business signals (relationship vs. contact mismatch)

### When Missing Info vs. Onboarding
- **missing_info**: Incomplete mandatory fields prevent processing
- **onboarding**: Complete information but no existing CBX match

This technical documentation provides the complete decision tree and processing logic for understanding how the onboarding analysis tool categorizes contractors and makes business decisions.