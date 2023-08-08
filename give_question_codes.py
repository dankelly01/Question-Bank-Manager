# FULLY FUNCTIONAL

# Importing required libraries
import pygsheets
import json

# User must specify
# 1. Name of the Open Question Bank, 
# e.g. "Statistics and Probability"
# 2. Name of the University and Course, separated by a "-" 
# e.g. "MMUST - STA 142 Introduction to Probability"

question_bank_name = "Statistics and Probability"


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

# Download cell data from the question bank
question_bank_cells = gws_question_bank.get_all_values()

# Separate column headers and question data
question_bank_headers = question_bank_cells[0]
question_bank_data = question_bank_cells[1:]


# --------------- RETRIEVE DATA FROM THE USED QUESTION CODES .JSON FILE

try:
    with open("used_question_codes.json") as codes_json:
        codes_dict = json.load(codes_json)
    used_question_codes = list(codes_dict.keys())
except:
    print("The .json file with the used question codes could not be located.")

used_code_numbers = [int(x) for x in used_question_codes]
max_code_number = max(used_code_numbers)


# --------------- GIVE QUESTION CODES TO THOSE WITHOUT

try:
    code_index = question_bank_headers.index("Question Code")
except:
    question_bank_headers.insert(0, "Question Code")
    for row in question_bank_data:
        row.insert(0, "")
    code_index = 0
    gws_question_bank.insert_cols(0)
    print("A Question Code column has been added to the Question Bank.")

for x in range(len(question_bank_data)):
    row = question_bank_data[x]
    # If the question has no code, give the next available code and update the .json variables
    if row[code_index] == "":

        # Prepare the code to be inserted
        max_code_number += 1
        insert_code = str(max_code_number)
        while len(insert_code) < 6:
            insert_code = "0" + insert_code

        # Add the code to the used codes list and .json dictionary
        used_question_codes.append(insert_code)
        codes_dict[insert_code] = ["000"]

        # Add the code to the row
        insert_code = "#" + insert_code + "-000"
        row[code_index] = insert_code
    
    # If the quesiton code is in the .json file, we need do nothing
    elif row[code_index][1:7] in used_question_codes:
        pass

    # Otherwise we inform the user that there was a malformed code in the tracker and replace it
    else:

        # Prepare the code to be inserted
        max_code_number += 1
        insert_code = str(max_code_number)
        while len(insert_code) < 6:
            insert_code = "0" + insert_code
        used_question_codes.append(insert_code)

        print("The question code " + row[code_index] + " is malformed.\n" +
              "It has been replaced by the code: " + insert_code)
        
        # Add the code to the used codes list and .json dictionary
        used_question_codes.append(insert_code)
        codes_dict[insert_code] = ["000"]

        # Add the code to the row
        insert_code = "#" + insert_code + "-000"
        row[code_index] = insert_code

    # Update the question bank data
    question_bank_data[x] = row


# --------------- UPADATE .JSON FILE

json_codes_object = json.dumps(codes_dict, indent=4)
with open("used_question_codes.json", "w") as outfile:
    outfile.write(json_codes_object)


# --------------- UPADATE QUESTION BANK

number_rows = len(question_bank_data)
number_cols = len(question_bank_headers)
gws_question_bank.update_values(
    "A1:" + chr(number_cols) + str(number_rows + 1), 
    [question_bank_headers] + question_bank_data
)
