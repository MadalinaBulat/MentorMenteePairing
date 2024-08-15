import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
from fuzzywuzzy import fuzz

# Define the scope for Google Sheets and Google Drive API
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]

# Load the credentials.json file
creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)

# Authorize the client
client = gspread.authorize(creds)

# Open the Google Sheets by name
mentor_sheet = client.open("Mentor Sheet").sheet1
mentee_sheet = client.open("Mentee Sheet").sheet1

# Fetch all data from the sheets
mentors_data = mentor_sheet.get_all_records()
mentees_data = mentee_sheet.get_all_records()

# Convert the data to pandas DataFrames
mentors_df = pd.DataFrame(mentors_data)
mentees_df = pd.DataFrame(mentees_data)

# Strip any extra spaces from column names
mentors_df.columns = mentors_df.columns.str.strip()
mentees_df.columns = mentees_df.columns.str.strip()

# Debug: Print column names to identify any issues
# print("Mentor Columns:", mentors_df.columns)
# print("Mentee Columns:", mentees_df.columns)

# Function to calculate similarity between mentor and mentee
def calculate_similarity(mentor, mentee):
    # Calculate similarity scores for each criterion
    major_score = fuzz.token_sort_ratio(mentor['major'], mentee['major'])
    cs_topics_score = fuzz.token_sort_ratio(mentor['CStopics'], mentee['CStopics'])
    hobbies_score = fuzz.token_sort_ratio(mentor['Hobbies'], mentee['Hobbies'])
    activities_score = fuzz.token_sort_ratio(mentor.get('Activities', ''), mentee.get('Activities', ''))
    
    # Sum all the scores to get a total similarity score
    total_score = major_score + cs_topics_score + hobbies_score + activities_score
    
    return total_score

# Pairing logic: find the best match for each mentee
pairings = []

for idx, mentee in mentees_df.iterrows():
    best_match = None
    best_score = 0
    
    for _, mentor in mentors_df.iterrows():
        score = calculate_similarity(mentor, mentee)
        if score > best_score:
            best_score = score
            best_match = mentor['Full Name']  # Use the correct name column for mentors
    
    pairings.append({'Mentee': mentee['Full Name'], 'Mentor': best_match, 'Score': best_score})

# Convert the pairings list to a DataFrame for easier visualization
pairings_df = pd.DataFrame(pairings)

# Output the pairings
print(pairings_df)

# Optionally, save the pairings to a Google Sheet or CSV file
pairings_df.to_csv('mentor_mentee_pairings.csv', index=False)

# Save the pairings to a new Google Sheet
pairing_sheet = client.create("Mentor-Mentee Pairings")
pairing_sheet.sheet1.update([pairings_df.columns.values.tolist()] + pairings_df.values.tolist())
