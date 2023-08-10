# Importing required libraries
import pygsheets

# User must specify
# 1. Name of the Open Question Bank, 
# e.g. "Statistics and Probability"
# 2. Name of the University and Course, separated by a "-" 
# e.g. "MMUST - STA 142 Introduction to Probability"

course_tracker_name = "MMUST - STA 141 Introduction to Statistics"

# --------------- RETRIEVE DATA FROM COURSE TRACKER

# Open question bank
try:
    client = pygsheets.authorize(service_account_file="C:/Users/James/OneDrive/Documents/008_VisualStudio/IDEMS/Quesiton-Bank-manager/service_account.json")
except:
    print("Error: Service account .json file not found.")
    exit()
try:
    gss_course_tracker = client.open(course_tracker_name + " - Question Tracker")
except:
    print("Error: Could not find this Question Bank.")
    exit()
try:
    gws_course_tracker = gss_course_tracker.worksheet("title", "Questions")
except:
    print("The Question Bank does not have a Questions sheet.")
    exit()

# Download cell data from the question bank
course_tracker_cells = gws_course_tracker.get_all_values()

# Separate column headers and question data
course_tracker_headers = course_tracker_cells[0]
course_tracker_data = course_tracker_cells[1:]

# --------------- FUNCTIONS FOR CONVERTING TO MATTERMOST CARDS

