# Adds/Updates question data from Mattermost .csv to question bank
# Assigns question codes to new questions and returns required updates to Mattermost

# Importing required libraries
import pygsheets
import csv
import json

# User must specify 
# 1. The .csv file with downloaded board data 
# 2. Question Bank name

csv_data_file = "board_downloads/advanced_mathematics.csv"
question_bank_name = "Copy of **Blank**"


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

# Set up key-value mapping from the header of the column in the question bank to position in the row
question_bank_headsToIndex = {}
for i in question_bank_headers:
    question_bank_headsToIndex[i] = question_bank_headers.index(i)

# Extract a list of the question codes already in the question bank
qb_question_codes_index = question_bank_headers.index("Question Code")
qb_question_codes = [x[qb_question_codes_index] for x in question_bank_data]

# Importing the translator file to convert board user codes to names
with open("indexing_codes_files/author_codes.json") as author_codes_file:
    author_codes = json.load(author_codes_file)
codes_to_author = {}
for author in author_codes.keys():
    codes_to_author[author_codes[author]] = author


# --------------- PARSE .CSV FILE DATA AND IMPORT TO QUESTION BANK

with open(csv_data_file, mode='r') as csv_file:

    # Opening the .csv file as a iterable object of dictionaries
    csv_reader = csv.DictReader(csv_file)
    csv_row_counter = 0

    # We parse through the rows of the .csv file and update the question bank data accordingly
    for csv_row_dict in csv_reader:

        # On the first iteration, create a list of the headers in the .csv file
        if csv_row_counter == 0:
            csv_heads = csv_row_dict.keys()

        # Proceed if the row relates to an active question card (not "Useful Info" or "To Do List")
        if csv_row_dict["Status"] not in ["Useful Info", "To Do List"]:
        
            # If the card has no question code, assign one and add an empty row to the end of the question bank
            if csv_row_dict["Question Code"] == "":

                # Assign question code
                with open("used_question_codes.json") as used_question_codes:
                    used_codes_list = json.load(used_question_codes)
                    last_code_num = int(used_codes_list[-1][0:6])
                    new_code_num = last_code_num + 1
                    new_code = str(new_code_num) + "-000"
                    while len(new_code) < 10:
                        new_code = "0" + new_code
                    csv_row_dict["Question Code"] = new_code
                    
                # Add used question code to .json file
                with open("used_question_codes.json", "w") as used_question_codes:
                    used_codes_list.append(new_code)
                    used_question_codes.write(json.dumps(used_codes_list, indent=4))

                # Add new empty row in question bank data to edit
                edit_row = len(question_bank_data)
                question_bank_data.append(["" for x in question_bank_headers])

                # Add the new question code to the list of codes in the question bank
                qb_question_codes.append(new_code)

                # Output the Mattermost link to the question card which requires new question code
                if csv_row_dict["Card Link"] == "":
                    print("The code " + new_code + " has been assigned to one of the questions." +
                          "\n The link to the question card was not found.")
                else:
                    print("The code " + new_code + " has been assigned to one of the questions." +
                          "\n The question card is: " + csv_row_dict["Card Link"])

            # If the card has a question code, determine if the card has an existing entry in the question bank
            else:
                
                # If the question code has a row in the question bank, then set the edit row accordingly
                if csv_row_dict["Question Code"].replace("___hash_sign___", "#") in qb_question_codes:
                    edit_row = qb_question_codes.index(csv_row_dict["Question Code"].replace("___hash_sign___", "#"))

                else:
                # If the Question code is not in the data, add a new row to edit
                    edit_row = len(question_bank_data)
                    question_bank_data.append(["" for x in question_bank_headers])

            # Finally, edit the row
            for header in csv_heads:

                # Determine the column in the question bank to be editted
                if header == "Name":
                    edit_col = question_bank_headsToIndex["Question Title"]
                elif header == "Status":
                    edit_col = question_bank_headsToIndex["Internal Review Status"]
                else:
                    edit_col = question_bank_headsToIndex[header]

                # Edit the relevant column
                if header == "Question Code":
                    question_bank_data[edit_row][edit_col] = csv_row_dict[header].replace("___hash_sign___", "#")
                elif header == "STACK Lead" or header == "Peer Reviewer" or header == "Second Reviewer":
                    question_bank_data[edit_row][edit_col] = codes_to_author[csv_row_dict[header]]
                else:
                    question_bank_data[edit_row][edit_col] = csv_row_dict[header]

        csv_row_counter += 1

# Update cell data in google sheet
question_bank_no_rows = len(question_bank_data)
gws_question_bank.update_values("A2:R" + str(question_bank_no_rows + 1), question_bank_data)
