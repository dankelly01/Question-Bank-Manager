# Importing required libraries
import pygsheets
import csv
import json

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
print(ordered_course_questions[0])


# --------------- UPLOAD NEW DATA TO COURSE TRACKER

# Push the updated data to the database
no_questions = len(ordered_course_questions)
gws_course_tracker.update_values(
    "A2:J" + str(no_questions + 1), 
    ordered_course_questions)



'''
question_bank_question_codes = gws_question_bank.get_col(
    1,
    include_tailing_empty=False,
    )[1:]
question_bank_data = gws_question_bank.get_values(
    start = "A2", 
    end = "R" + str(len(question_bank_question_codes) + 1),
    value_render = "FORMULA"
    )

# Open database
gss_database = client.open(database_name)
gws_database = gss_database.worksheet("title", "Question Database")

# Download database headers, question codes and data
database_heads = gws_database.get_row(
    1,
    include_tailing_empty=False,
    )
database_question_codes = gws_database.get_col(
    1,
    include_tailing_empty=False,
    )[1:]
if len(database_question_codes) >= 1:
    database_data = gws_database.get_values(
        start = "A2", 
        end = "L" + str(len(database_question_codes) + 1),
        value_render = "FORMULA"
        )
else:
    database_data = []

# Set up key-value mapping from the header of the column to position in the row
question_bank_headsToIndex = {}
for i in question_bank_heads:
    question_bank_headsToIndex[i] = question_bank_heads.index(i)
database_headsToIndex = {}
for i in database_heads:
    database_headsToIndex[i] = database_heads.index(i)

# Loop through the data in the question bank being uploaded
for question in question_bank_data:

    # Proceed if the question is not marked as backed up
    if question[question_bank_headsToIndex["Backed up"]] == False:

        # If the question already has a entry in the database, update the data
        if question[question_bank_headsToIndex["Question Code"]] in database_question_codes:
            row_to_edit = database_question_codes.index(question[question_bank_headsToIndex["Question Code"]])
            
            for header in database_heads:
                if header == "IDEMS Question Bank":
                    database_data[row_to_edit][database_headsToIndex[header]] = question_bank_name
                elif header == "Status":
                    if question[question_bank_headsToIndex[header]] == "Approved":
                        database_data[row_to_edit][database_headsToIndex[header]] = "Completed"
                    else:
                        database_data[row_to_edit][database_headsToIndex[header]] = "Under Review"
                else:
                    database_data[row_to_edit][database_headsToIndex[header]] = question[question_bank_headsToIndex[header]]
        
        # If the question does not have an entry, add one with the relevant data
        else:
            row_to_edit = len(database_question_codes)
            database_question_codes.append(question[question_bank_headsToIndex["Question Code"]])
            database_data.append(["" for x in database_heads])

            for header in database_heads:
                if header == "IDEMS Question Bank":
                    database_data[row_to_edit][database_headsToIndex[header]] = "Mock-up Question Tracker 1"
                elif header == "Status":
                    if question[question_bank_headsToIndex[header]] == "Approved":
                        database_data[row_to_edit][database_headsToIndex[header]] = "Completed"
                    else:
                        database_data[row_to_edit][database_headsToIndex[header]] = "Under Review"
                else:
                    database_data[row_to_edit][database_headsToIndex[header]] = question[question_bank_headsToIndex[header]]

# Push the updated data to the database
database_no_rows = len(database_data)
gws_database.update_values("A2:L" + str(database_no_rows + 1), database_data)

# Mark the questions in the question bank as updated
question_bank_no_rows = len(question_bank_data)
gws_question_bank.update_values("R2:R" + str(database_no_rows + 1), [[True] for x in question_bank_data])
'''