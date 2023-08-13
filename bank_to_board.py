# Importing required libraries
import pygsheets
import manageBoards

# User must specify
# 1. Name of the Open Question Bank, e.g. "Statistics and Probability"
# 2. Name of the .boardarchive file from which the board is created

question_bank_name = "Statistics and Probability"
boardArchiveFile =  "board_archives/statistics_and_probability.boardarchive"


# --------------- RETRIEVE DATA FROM COURSE TRACKER

# Open question bank
try:
    client = pygsheets.authorize(service_account_file="C:/Users/James/OneDrive/Documents/008_VisualStudio/IDEMS/Quesiton-Bank-manager/service_account.json")
except:
    print("Error: Service account .json file not found.")
    exit()
try:
    gss_course_tracker = client.open(question_bank_name + " - Question Bank")
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

# --------------- GENERATE MATTERMOST BOARD

board = manageBoards.board(question_bank_name + " - Open Question Bank")
authorCodes = manageBoards.get_author_codes("indexing_codes_files/author_codes.json")

header_index = {}
for header in course_tracker_headers:
    header_index[header] = course_tracker_headers.index(header)

for row in course_tracker_data:
    manageBoards.card(
        row[header_index["Question Title"]],
        "\U0001F4BB",
        row[header_index["Question Code"]],
        row[header_index["Description of Question"]],
        row[header_index["Internal Review Status"]],
        row[header_index["STACK Lead"]],
        row[header_index["Peer Review"]],
        row[header_index["Second Review"]],
        authorCodes,
        row[header_index["Link to Concept"]],
        row[header_index["Link to STACK"]],
        row[header_index["Link to Card"]]
    ).add_to_board(board)

manageBoards.write_to_file(board.create_board(), boardArchiveFile)
