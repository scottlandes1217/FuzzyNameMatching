import pandas as pd
from fuzzywuzzy import fuzz

# Load the data
sheet1_df = pd.read_excel("/Users/scott/Desktop/Book3.xlsx", sheet_name="Sheet1")
sheet2_df = pd.read_excel("/Users/scott/Desktop/Book3.xlsx", sheet_name="Sheet2")

# List of terms to ignore in the matching process
ignore_terms = {"property", "management", "managers", "group", "llc", "properties", "Real Estate", "Realty"}

# Custom preprocessing function to remove specified terms
def preprocess_name(name):
    # Split the names into words, convert to lowercase, and remove ignored terms
    return set(word.lower() for word in name.split() if word.lower() not in ignore_terms)

# Apply preprocessing to the names in both sheets for tokenization
sheet1_df['Tokens'] = sheet1_df['SF Name'].apply(preprocess_name)
sheet2_df['Tokens'] = sheet2_df['Admin Name'].apply(preprocess_name)

# Custom scoring function for matching based on the preprocessed tokens
def custom_score(sf_tokens, admin_tokens):
    # Find the intersection of tokens and calculate score based on common tokens
    common_tokens = sf_tokens.intersection(admin_tokens)
    score = len(common_tokens) / max(len(sf_tokens), len(admin_tokens))
    return score

# Iterate over Sheet1 names and find the best match from Sheet2
matches = []
matched_admin_names = set()  # To ensure 1:1 matching

for index, sf_row in sheet1_df.iterrows():
    sf_name = sf_row['SF Name']
    sf_tokens = sf_row['Tokens']
    best_match = None
    best_score = 0
    best_id = None
    
    for admin_index, admin_row in sheet2_df.iterrows():
        admin_name = admin_row['Admin Name']
        admin_tokens = admin_row['Tokens']
        admin_id = admin_row['Id']
        
        # Skip if already matched
        if admin_name in matched_admin_names:
            continue
        
        # Calculate custom score based on token sets
        score = custom_score(sf_tokens, admin_tokens)
        
        if score > best_score:
            best_score = score
            best_match = admin_name
            best_id = admin_id
    
    # Only accept matches above a threshold to avoid incorrect matches
    if best_score > 0.5:  # Threshold can be adjusted based on empirical results
        matches.append({'SF Name': sf_name, 'Matched Name': best_match, 'Id': best_id})
        matched_admin_names.add(best_match)  # Add to the set of matched names
    else:
        matches.append({'SF Name': sf_name, 'Matched Name': None, 'Id': None})

# Convert the list of matches to a DataFrame
matches_df = pd.DataFrame(matches)

# Save the matched results to an Excel file
matches_df.to_excel("/Users/scott/Desktop/MatchedFinalResults.xlsx", index=False)
