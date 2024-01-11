""" Checklist Collections

Idea
Checklist: list or dynamically construct a list
Inside:
checkbox: str
choice: dict (item and choices). The last number of the choices list controls widget width
textbox: list: Tuple. The number controls the widget's width and height

Checked for type and take necessary action. If needs be, a list can be constructed
from different checklists.
"""

YES_NO = "Yes, No, 50"
NO_YES = "No, Yes, 50"
NIL_YES_NO = " , Yes, No, 50"
NA_YES_NO = "NA, Yes, No, 50"
NIL_YES_NO_NA = " , Yes, No, NA, 50"
DELIVERY_TERMS = "EXW, FOB, CIF, CPT, FCA, DAP, DDU, DDP, 50"
CREDIT_TERMS = "30 Days, 45 Days, Advanced T/T, COD, 7 Days, 10 Days, 14 Days, LC at Sight, 70"
CLASS_SOCIETY = "NA, DNV, ABS, LR, BV, Others, 50"
VALIDITY = "30 Days, 45 Days, 60 Days, 90 Days, 120 Days, 7 Days, 15 Days, 70"
TEXTBOX_HEIGHT = 17

leave_application_checklist = [
    "Have you marked the leave in the team calendar?",
    "For AM or PM leave, have you marked the exact time in the calendar?",
    {
        "Is the leave longer than 10 days duration including weekends and holidays?": NO_YES.split(','),
        """If the above is 'Yes', it is required to put the note in the email signature 
        two weeks before the due leave. Have you put the reminder for yourself for this?""": NA_YES_NO.split(','),
        "You are responsible for filling out this checklist. Have you answered all the checklist items carefully?": NIL_YES_NO.split(','),
    },
]

ikigai_checklists = [
    "Is it something you are good at?",
    "Is it something you love to do?",
    "Is it something the world needs?",
    "Does it generate money?",
]


textbox_checklist_example = [
    ("Name let us write something long here so that I know how it will behave.", 120),
    (
        "Firstname let us write something long here so that I know how it will behavaviour.",
        100,
    ),
]


sales_checklist = [   #type:ignore
    ("Project name", 200, TEXTBOX_HEIGHT),
    ("Job code", 200, TEXTBOX_HEIGHT),
    ("Customer name", 200, TEXTBOX_HEIGHT),
    {
        "Customer type": ["Existing", "New", 70],
        "If 'Existing Customer', do we have any past issues with the customer we need to be aware of?": ['No', 'Yes', 'Unknown', 70],
    },
    ("If above is 'Yes', state the reason.", 300, TEXTBOX_HEIGHT*2),
    ("Yard name the vessel will be built in", 200, TEXTBOX_HEIGHT),
    ("End user or owner name", 200, TEXTBOX_HEIGHT),
    ("Infrastructure Type", 200, TEXTBOX_HEIGHT),
    ("The operating country/region of the vessel", 200, TEXTBOX_HEIGHT),
    {
        "Classification society": CLASS_SOCIETY.split(','),
    },
    {
        "Jason entity to be quoted under": ["Jason Electronics", "Jason Energy", 120],
        "Currency to be quoted in": ['SGD', 'USD', 70],
        "Type of quotation": ['Budgetary', 'Firmed', 70],
        "Have we received all the required information.": [' ', 'Yes', 'No', 'Not Sure', 70]
    },
    ("Preferred margin to be quoted in percentage", 30, TEXTBOX_HEIGHT),
    ("Preferred milestone payment terms", 300, TEXTBOX_HEIGHT),
    {
        "Preferred credit terms": CREDIT_TERMS.split(','),
        "Preferred delivery terms": DELIVERY_TERMS.split(','),
    },
    ("Delivery location based above delivery terms (to or from)", 150, TEXTBOX_HEIGHT),
    {
        "Quotation validity": VALIDITY.split(','),
    },
    ("Warranty duration and/or warranty end date", 200, TEXTBOX_HEIGHT),
    ("Commissioning location", 200, TEXTBOX_HEIGHT),
    ("Estimated project delivery date", 200, TEXTBOX_HEIGHT),
    {
        "Any special requirement?": ['No', 'Yes', 'Unknown', 70],
    },
    ("If above is 'Yes', state the requirement.", 300, TEXTBOX_HEIGHT*5),
    ("Any known competitor?", 200, TEXTBOX_HEIGHT),
    ("Any known concern?", 200, TEXTBOX_HEIGHT),
    {
        "Have you answered all the checklist items carefully?": NIL_YES_NO.split(','),
    },

]


test = []
test.append(leave_application_checklist)
test.append(textbox_checklist_example)  # type:ignore
