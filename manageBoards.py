import time
import math
import random
import string
import json

def get_random_string(length):
    letters = string.ascii_lowercase + string.digits
    result_str = ''.join(random.choice(letters) for i in range(length))
    return(result_str)

def get_formatted_id():
    return(get_random_string(8) + "-" 
           + get_random_string(4) + "-" 
           + get_random_string(4) + "-" 
           + get_random_string(4) + "-" 
           + get_random_string(12))

def get_author_codes(file_location):
    with open(file_location, "r") as file:
        return(json.load(file))

def write_to_file(data, file_location):
    with open(file_location, "w", encoding="utf-8") as file:
        for line in data:
            file.write(str(line)
                    .replace("False", "false")
                    .replace("True", "true")
                    .replace(": ", ":")
                    .replace(", ", ",")
                    .replace("{\'", "{\"")
                    .replace("[\'", "[\"")
                    .replace(":\'", ":\"")
                    .replace(",\'", ",\"")
                    .replace("\'}", "\"}")
                    .replace("\']", "\"]")
                    .replace("\':", "\":")
                    .replace("\',", "\","))
            file.write("\n")


class board:

    def __init__(self, title):
        # Board creation data
        self.title = title
        self.date = math.floor(1000*time.time())
        self.id = get_random_string(27)
        self.viewId = get_random_string(27)
        # Card properties data
        self.statusId = get_formatted_id()
        self.statuses = {
            "Useful Info":get_formatted_id(),
            "To Do List":get_formatted_id(),
            "Backlog":get_formatted_id(),
            "In Progress":get_formatted_id(),
            "Ready for Peer Review":get_formatted_id(),
            "Peer Changes Requested":get_formatted_id(),
            "Ready for Second Review":get_formatted_id(),
            "Changes Requested":get_formatted_id(),
            "Under External Review":get_formatted_id(),
            "Rejected":get_formatted_id(),
            "Approved":get_formatted_id()
        }
        self.questionCodeId = get_formatted_id()
        self.descriptionId = get_formatted_id()
        self.stackLeadId = get_formatted_id()
        self.peerReviewId = get_formatted_id()
        self.secondReviewId = get_formatted_id()
        self.conceptLinkId = get_formatted_id()
        self.stackLinkId = get_formatted_id()
        self.cardLinkId = get_formatted_id()
        self.cards = []

    # Function to create board as a py dictionary to be written to a .boardarchive file
    def create_board(self):
        return([
        # Creation info
        {
            "version":1,
            "date":self.date
        },
        # Board info
        {
            "type":"board",
            "data":{
                "id":self.id,
                "teamId":"",
                "channelId":"",
                "createdBy":"",
                "modifiedBy":"",
                "type":"P",
                "minimumRole":"",
                "title":self.title,
                "description":"Board Description",
                "icon":"",
                "showDescription":True,
                "isTemplate":False,
                "templateVersion":0,
                "properties":{},
                "cardProperties":[{
                    # Status property
                    "id":self.statusId,
                    "name":"Status",
                    "type":"select",
                    "options":[{
                        "id":self.statuses["Useful Info"],
                        "value":"Useful Info",
                        "color":"Default"
                    },
                    {
                        "id":self.statuses["To Do List"],
                        "value":"To Do List",
                        "color":"propColorYellow"
                    },
                    {
                        "id":self.statuses["Backlog"],
                        "value":"Backlog",
                        "color":"propColorYellow"
                    },
                    {
                        "id":self.statuses["In Progress"],
                        "value":"In Progress",
                        "color":"propColorPink"
                    },
                    {
                        "id":self.statuses["Ready for Peer Review"],
                        "value":"Ready for Peer Review",
                        "color":"propColorBlue"
                    },
                    {
                        "id":self.statuses["Peer Changes Requested"],
                        "value":"Peer Changes Requested",
                        "color":"propColorOrange"
                    },
                    {
                        "id":self.statuses["Ready for Second Review"],
                        "value":"Ready for Second Review",
                        "color":"propColorBlue"
                    },
                    {
                        "id":self.statuses["Changes Requested"],
                        "value":"Changes Requested",
                        "color":"propColorOrange"
                    },
                    {
                        "id":self.statuses["Under External Review"],
                        "value":"Under External Review",
                        "color":"propColorPurple"
                    },
                    {
                        "id":self.statuses["Rejected"],
                        "value":"Rejected",
                        "color":"propColorRed"
                    },
                    {
                        "id":self.statuses["Approved"],
                        "value":"Accepted",
                        "color":"propColorGreen"
                    }]
                },
                {
                    # Question Code Property
                    "id":self.questionCodeId,
                    "name":"Question Code",
                    "type":"text"
                },
                {
                    # Question Description Property
                    "id":self.descriptionId,
                    "name":"Question Description",
                    "type":"text"
                },
                {
                    # Lead Author Property
                    "id":self.leadAuthorId,
                    "name":"STACK Lead",
                    "type":"person"
                },
                {
                    # Peer Reviewer Property
                    "id":self.peerReviewId,
                    "name":"Peer Reviewer",
                    "type":"person"
                },
                {
                    # Second Reviewer Property
                    "id":self.secondReviewId,
                    "name":"Second Reviewer",
                    "type":"person"
                },
                {
                    # Concept Link Property
                    "id":self.conceptLinkId,
                    "name":"Concept Link",
                    "type":"url"
                },
                {
                    # STACK Link Property
                    "id":self.stackLinkId,
                    "name":"STACK Link",
                    "type":"url"
                },
                {
                    # Card Link Property
                    "id":self.cardLinkId,
                    "name":"Card Link",
                    "type":"url"
                }]
            },
            "createAt":self.date,
            "updateAt":self.date,
            "deleteAt":0
        },
        # "Progress Tracker" view info
        {
            "type":"block",
            "data":
            {
                "id":self.viewId,
                "schema":1,
                "boardId":self.id,
                "parentId":self.id,
                "createdBy":"",
                "modifiedBy":"",
                "type":"view",
                "fields":
                {
                    "viewType":"board",
                    "sortOptions":[],
                    "visiblePropertyIds":[],
                    "visibleOptionIds":[],
                    "hiddenOptionIds":[],
                    "collapsedOptionIds":[],
                    "filter":
                    {
                        "operation":"and",
                        "filters":[]
                    },
                    "cardOrder":[],
                    "columnWidths":{},
                    "columnCalculations":{},
                    "kanbanCalculations":{},
                    "defaultTemplateId":""
                },
                "title":"Progress Tracker",
                "createAt":self.date,
                "updateAt":self.date,
                "deleteAt":0,
                "limited":False
            }
        }] + self.cards)


class card:

    def __init__(self, questionTitle, icon, questionCode, description, status, stackLead, peerReviewer, secondReviewer, authorCodes, conceptLink, stackLink, cardLink):
        self.questionTitle = questionTitle
        self.date = math.floor(1000*time.time())
        self.id = get_random_string(27)
        self.icon = icon
        self.questionCode = questionCode
        self.description = description
        self.status = status
        self.stackLead = stackLead
        self.peerReviewer = peerReviewer
        self.secondReviewer = secondReviewer
        self.authorCodes = authorCodes
        self.conceptLink = conceptLink
        self.stackLink = stackLink
        self.cardLink = cardLink

    def add_to_board(self, board):
        board.cards.append({
            "type":"block",
            "data":
            {
                "id":self.id,
                "schema":1,
                "boardId":board.id,
                "parentId":board.id,
                "createdBy":"",
                "modifiedBy":"",
                "type":"card",
                "fields":
                {
                    "icon": self.icon,
                    "properties":
                    {   
                        board.statusId:board.statuses[self.status],
                        board.questionCodeId:self.questionCode,
                        board.descriptionId:self.description,
                        board.stackLead:self.authorCodes[self.stackLead],
                        board.peerReviewId:self.authorCodes[self.peerReviewer],
                        board.secondReviewId:self.authorCodes[self.secondReviewer],
                        board.conceptLinkId:self.conceptLink,
                        board.stackLinkId:self.stackLink,
                        board.cardLinkId:self.cardLink,
                    },
                    "contentOrder":[],
                    "isTemplate":False
                },
                "title":self.questionTitle,
                "createAt":self.date,
                "updateAt":self.date,
                "deleteAt":0,
                "limited":False
            }
        })
