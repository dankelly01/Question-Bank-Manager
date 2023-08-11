# WOKRING ON
# Code exports questions selected in question bank to course tracker
# Currently reformats the tracker each time (could by updated)
# Add transporting of lists sheet

# Importing required libraries
import pygsheets

# User must specify
# 1. Name of the Open Question Bank, 
# e.g. "Statistics and Probability"
# 2. Name of the University and Course, separated by a "-" 
# e.g. "MMUST - STA 142 Introduction to Probability"

question_bank_name = "Statistics and Probability"
course_tracker_name = "MMUST - STA 141 Introduction to Statistics"


# --------------- RETRIEVE DATA FROM OPEN QUESTION BANK

# Open question bank
try:
    client = pygsheets.authorize(service_account_file="C:/Users/James/OneDrive/Documents/008_VisualStudio/IDEMS/Quesiton-Bank-manager/service_account.json")
except:
    print("Error: Service account .json file not found.")
    exit()
try:
    gss_question_bank = client.open(question_bank_name + " - Question Bank")
except:
    print("Error: Could not find this Question Bank.")
    exit()
try:
    gws_question_bank = gss_question_bank.worksheet("title", "Questions")
except:
    print("The Question Bank does not have a Questions sheet.")
    exit()
try:
    gws_question_bank_lists = gss_question_bank.worksheet("title", "Questions")
except:
    print("The Question Bank does not have a Lists sheet.")
    exit()

# Download cell data from the question bank
question_bank_cells = gws_question_bank.get_all_values()

# Separate column headers and question data
question_bank_headers = question_bank_cells[0]
question_bank_data = question_bank_cells[1:]

# Find the column of the relevant course
[course_institution, course_name] = course_tracker_name.split("-")
course_header = course_institution + "\n" + course_name[1:]
try:
    course_index = question_bank_headers.index(course_header)
except:
    print("The Course Name you have given could not be found in the Question Bank.")
    exit()

# Extract the questions which should be transferred to the Course Tracker
course_questions = []
for question_row in question_bank_data:
    if question_row[course_index] != "":
        course_questions.append(question_row)

# Sort the questions (alphabitically/numerically)
course_questions.sort(key = lambda e : e[course_index])


# --------------- RETRIEVE EXISTING DATA FROM COURSE TRACKER

# Open course tracker
try:
    gss_course_tracker = client.open(course_tracker_name + " - Question Tracker")
except:
    print("Error: Could not find this Course Tracker.")
    exit()
try:
    gws_course_tracker = gss_course_tracker.worksheet("title", "Questions")
except:
    print("The Course Tracker does not have a Questions sheet.")
    exit()
try:
    gws_course_tracker_lists = gss_course_tracker.worksheet("title", "Lists")
except:
    print("The Course Tracker does not have a Lists sheet.")
    exit()

# Download cell data from the course tracker
course_tracker_cells = gws_course_tracker.get_all_values()

# Separate column headers and question data
course_tracker_headers = course_tracker_cells[0]
course_tracker_data = course_tracker_cells[1:]


# --------------- SETING UP DATA FOR TRANSFER

# Order the question data retrieved from the quesiton bank 
# to match the column order in the course tracker
ordered_course_questions = []

# Create a dictionary taking header of course tracker to index 
# of the matching header in the question tracker
header_index_dict = {
    0: course_index
}
for header in course_tracker_headers[1:]:
    try:
        header_index_dict[course_tracker_headers.index(header)] = question_bank_headers.index(header)
    except:
        print("Data for the column \"" + header + "\" could not be found in the Question Bank.")

# Parse through the questions for transfer and reorder the entries to fit the columns
# in the course tracker
for question_row in course_questions:
    row = []
    for i in range(len(course_tracker_headers)):
        if i in header_index_dict:
            row.append(question_row[header_index_dict[i]])
        else:
            row.append("")
    ordered_course_questions.append(row)

# Order rows to match the existing structure of the course tracker
'''
NEEDS WORK
'''


# --------------- UPLOAD NEW DATA TO COURSE TRACKER

# Push the updated question data to the course tracker
no_questions = len(ordered_course_questions)
gws_course_tracker.update_values(
    "A2:K" + str(no_questions + 1), 
    ordered_course_questions)

# Delete any extra rows in the course tracker
course_tracker_rows = gws_course_tracker.rows
if course_tracker_rows > no_questions + 1:
    gws_course_tracker.delete_rows(
        no_questions + 2, 
        gws_course_tracker.rows - (no_questions + 1)
    )
