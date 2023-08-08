# REQUIRES UPDATES

# Importing required libraries
import pygsheets
import csv
import json

# Adds/Updates question data from Matter most .csv to question bank
# Assigns question codes to new questions and returns required updates to Mattermost

# User specifies .csv file with board data and Google spreadsheet question tracker name

csv_data_file = "C:/Users/James/OneDrive/Documents/008_VisualStudio/IDEMS/Quesiton Bank manager/Mock-up_Authoring_Board_2.csv"
gss_name = "Mock-up Question Bank 2"

# <---------- SET UP ------------->

# Open the question bank
client = pygsheets.authorize(service_account_file="C:/Users/James/OneDrive/Documents/008_VisualStudio/IDEMS/Quesiton Bank manager/service_account.json")
gss_question_bank = client.open(gss_name)
gws_question_bank = gss_question_bank.worksheet("title", "Question Tracker")

# Download question headers, question codes and data
question_bank_heads = gws_question_bank.get_row(
    1,
    include_tailing_empty=False,
    )
question_bank_question_codes = gws_question_bank.get_col(
    1,
    include_tailing_empty=False,
    )[1:]
question_bank_data = gws_question_bank.get_all_values(
    include_tailing_empty=False,
    include_tailing_empty_rows=False,
    value_render = "FORMULA"
    )[1:]

# Set up key-value mapping from the header of the column in the question bank to position in the row
question_bank_headsToIndex = {}
for i in question_bank_heads:
    question_bank_headsToIndex[i] = question_bank_heads.index(i)

# Importing the translator file to convert board user codes to names
with open("user_translator.json") as user_translator_file:
    translator = json.load(user_translator_file)
translated_words = translator.keys()

# <----------- SCRIPT ------------->

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
                question_bank_data.append(["" for x in question_bank_heads])
                needs_edit = True

                # Add the new question code to the list of codes in the question bank
                question_bank_question_codes.append(new_code)

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
                if csv_row_dict["Question Code"] in question_bank_question_codes:
                    edit_row = question_bank_question_codes.index(csv_row_dict["Question Code"])

                    # Decide if the row requires editing
                    if [csv_row_dict[header] for header in csv_heads] == [question_bank_data[edit_row][question_bank_headsToIndex[header]] for header in csv_heads]:
                        needs_edit = False
                    else:
                        needs_edit = True

                else:
                # If the Question code is not in the data, add a new row to edit
                    edit_row = len(question_bank_data)
                    question_bank_data.append(["" for x in question_bank_heads])
                    needs_edit = True

            # Finally, proceed if the row requires editting
            if needs_edit:

                for header in csv_heads:
                    edit_col = question_bank_headsToIndex[header]
                    if csv_row_dict[header] in translated_words:
                        question_bank_data[edit_row][edit_col] = translator[csv_row_dict[header]]
                    
                    elif csv_row_dict[header][0 : min(8, len(csv_row_dict[header]))] == "https://":
                        question_bank_data[edit_row][edit_col] = "=HYPERLINK(\"" + csv_row_dict[header] + "\",\"" + header + "\")"
                    
                    else:
                        question_bank_data[edit_row][edit_col] = csv_row_dict[header]
                
                # Set backed up to false
                question_bank_data[edit_row][17] = "False"

        csv_row_counter += 1

# Update cell data in google sheet
question_bank_no_rows = len(question_bank_data)
gws_question_bank.update_values("A2:R" + str(question_bank_no_rows + 1), question_bank_data)

# Format "Edit Tags" column with drop down lists
gws_question_bank.apply_format(
    "F2:F" + str(question_bank_no_rows + 1),
    {"backgroundColor":{
        "red":255/255,
        "green":242/255,
        "blue":204/255,
        "alpha":0
        }
    }
)

gws_question_bank.set_data_validation(
    start = "F2",
    end = "F" + str(question_bank_no_rows + 1),
    condition_type = "ONE_OF_RANGE",
    condition_values = ["=Tags!A:A"],
    strict = True,
    showCustomUi = True,
    inputMessage = "This input is not a recognised tag."
)

# Format "Backed up" column with tick boxes
gws_question_bank.set_data_validation(
    start = "R2",
    end = "R"+str(question_bank_no_rows + 1),
    condition_type = "BOOLEAN",
    condition_values = []
)
