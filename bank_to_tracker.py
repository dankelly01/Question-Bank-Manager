# Importing required libraries
import pygsheets

# User must specify
# 1. Name of the Open Question Bank, e.g. "Statistics and Probability"
# 2. Name of the University and Course separated by a "-", e.g. "MMUST - STA 142 Introduction to Probability"

question_bank_name = "Calculus"
course_tracker_name = "BDU - Basic Maths SS"


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
    gws_question_bank_lists = gss_question_bank.worksheet("title", "Lists")
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


# --------------- RETRIEVE DATA FROM LISTS SHEETS

# Retrieve data from the question bank lists
question_bank_list_cells = gws_question_bank_lists.get_all_values(
    include_tailing_empty_rows = False
)
question_bank_list_headers = question_bank_list_cells[0]
question_bank_list_cells = question_bank_list_cells[1:]

# Determine the structure titles for the course being exported
try:
    qb_course_structure_index = question_bank_list_headers.index(course_header)
    qb_course_structure = [row[qb_course_structure_index] for row in question_bank_list_cells if row[qb_course_structure_index] != ""]
except:
    print("The Course Name you have given could not be found in the Lists sheet of the Question Bank.")
    exit()

# Determine the STACK leads for the quesiton bank
try:
    qb_stack_leads_index = question_bank_list_headers.index("STACK Leads")
    qb_stack_leads = [row[qb_stack_leads_index] for row in question_bank_list_cells if row[qb_stack_leads_index] != ""]
except:
    print("The Course Name you have given could not be found in the Lists sheet of the Question Bank.")
    exit()

# Retrieve existing data from course tracker lists
course_tracker_list_cells = gws_course_tracker_lists.get_all_values(
    include_tailing_empty_rows = False
)
course_tracker_list_headers = course_tracker_list_cells[0]
course_tracker_list_cells = course_tracker_list_cells[1:]

# Find the existing structure titles for the course
try:
    ct_course_structure_index = course_tracker_list_headers.index("Course Structure Labels")
    ct_course_structure = [row[ct_course_structure_index] for row in course_tracker_list_cells if row[ct_course_structure_index] != ""]
except:
    print("The Lists sheet for the course tracker does not have a Course Structure Labels column.")
    exit()

# Find the existing STACK leads for the course
try:
    ct_stack_leads_index = course_tracker_list_headers.index("STACK Leads")
    ct_stack_leads = [row[ct_stack_leads_index] for row in course_tracker_list_cells if row[ct_stack_leads_index] != ""]
except:
    print("The Lists sheet for the course tracker does not have a STACK Leads column.")
    exit()

# Add missing values to course structure and STACK leads in the course tracker
for x in qb_course_structure:
    if not x in ct_course_structure:
        ct_course_structure.append(x)
gws_course_tracker_lists.update_col(
    ct_course_structure_index+1,
    ct_course_structure,
    1
)

for x in qb_stack_leads:
    if not x in ct_stack_leads:
        ct_stack_leads.append(x)
gws_course_tracker_lists.update_col(
    ct_stack_leads_index+1,
    ct_stack_leads,
    1
)


# --------------- SETING UP QUESTION DATA FOR TRANSFER

# Create a dictionary taking the index of each header in the course tracker to index of the matching header in the question bank
header_index_dict = {
    0: course_index
}
for header in course_tracker_headers[1:]:
    try:
        header_index_dict[course_tracker_headers.index(header)] = question_bank_headers.index(header)
        header_index_list = list(header_index_dict.keys())
    except:
        print("Data for the column \"" + header + "\" could not be found in the Question Bank.")

# Determine the index of the Question Codes in each of the course tracker and question bank
ct_question_code_index = course_tracker_headers.index("Question Code")
qb_question_code_index = question_bank_headers.index("Question Code")
# Isolate a list of the existing question codes in the course tracker
course_tracker_question_codes = [row[ct_question_code_index] for row in course_tracker_data]

# Parse through questions for transfer, determine if the question exists already in the course tracker by comparing the question code.
for question_row in course_questions:
    question_code = question_row[qb_question_code_index]
    
    # If the question exists already update the relevant values.
    if question_code in course_tracker_question_codes:
        
        # Find the relevant row in the course tracker
        course_tracker_row_index = course_tracker_question_codes.index(question_code)
        course_tracker_row = course_tracker_data[course_tracker_row_index]
        
        # Update values in the question tracker row
        for x in range(len(course_tracker_headers)):
            if x in header_index_list:
                course_tracker_row[x] = question_row[header_index_dict[x]]
        course_tracker_data[course_tracker_row_index] = course_tracker_row

    # If the question does not exist in the course tracker, add it to the end of the course tracker.
    else:
        course_tracker_row = []
        for x in range(len(course_tracker_headers)):
            if x in header_index_list:
                course_tracker_row.append(question_row[header_index_dict[x]])
            elif course_tracker_headers[x] == "Question Bank":
                course_tracker_row.append(question_bank_name)
            else:
                course_tracker_row.append("")
        course_tracker_data.append(course_tracker_row)


# --------------- UPLOAD NEW DATA TO COURSE TRACKER

# Push the updated question data to the course tracker
no_questions = len(course_tracker_data)
no_columns = len(course_tracker_headers)
gws_course_tracker.update_values(
    "A2:" + chr(no_columns + 64) + str(no_questions + 1), 
    course_tracker_data)

# Delete any extra rows in the course tracker
course_tracker_rows = gws_course_tracker.rows
if course_tracker_rows > no_questions + 1:
    gws_course_tracker.delete_rows(
        no_questions + 2, 
        gws_course_tracker.rows - (no_questions + 1)
    )
