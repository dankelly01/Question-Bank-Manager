# Importing required libraries
import pygsheets
import csv
import json

# Uploads non-backed up questions to the database from the question banks

# User specifies .csv file with board data and Google spreadsheet question tracker name

question_bank_name = "Mock-up Question Bank 2"
database_name = "Mock-up Database"

# <---------- SET UP ------------->

# Open question bank
client = pygsheets.authorize(service_account_file="service_account.json")
gss_question_bank = client.open(question_bank_name)
gws_question_bank = gss_question_bank.worksheet("title", "Question Tracker")

# Download question headers and data
question_bank_heads = gws_question_bank.get_row(
    1,
    include_tailing_empty=False,
    )
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
