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


leave_application_checklist = [
    "Have you marked the leave in the team calendar?",
    "For AM or PM leave, have you marked the exact time in the calendar?",
    {
        "Is the leave longer than 10 days duration including weekends and holidays?": ["No", "Yes", 30],
        """If the above is 'Yes', it is required to put the note in the email signature 
        two weeks before the due leave. Have you put the reminder for yourself for this?""": ["NA", "Yes", "No", 30],
        "You are responsible for filling out this checklist. Have you answered all the checklist items carefully?": [" ", "Yes", "No", 30],
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
    ("Project name", 200, 17),
    ("Customer name", 200, 17),
    {
        "Customer type": ["Existing", "New", 70],
        "If 'Existing Customer', do we have any past issues with the customer we need to be aware of?": ['No', 'Yes', 'Unknown', 70],
    },
    ("If above is 'Yes', state the reason.", 300, 17*2),
    ("Yard name the vessel will be built in", 200, 17),
    ("End user or owner name", 200, 17),
    ("Infrastructure Type", 200, 17),
    ("The operating region of the vessel", 200, 17),
    {
        "Classification society": [' ', 'DNV', 'ABS', 'BV', 'LR', 'Others', 70]
    },
    {
        "Jason entity to be quoted under": ["Jason Electronics", "Jason Energy", 120],
        "Currency to be quoted in": ['SGD', 'USD', 70],
        "Type of quotation": ['Budgetary', 'Firmed', 70],
        "Have we received all the required information.": [' ', 'Yes', 'No', 'Not Sure', 70]
    },
    ("Preferred margin to be quoted in percentage", 30, 17),
    ("Preferred payment terms", 250, 17),
    ("Preferred delivey terms", 200, 17),
    ("Warranty duration or/and warranty end date", 200, 17),
    ("Commissioning location", 200, 17),
    ("Estimated project delivery date", 200, 17),
    {
        "Any special requirement?": ['No', 'Yes', 'Unknown', 70],
    },
    ("If above is 'Yes', state the requirement.", 200, 17),
    ("Any known competitor?", 200, 17),
    ("Any known concern?", 200, 17),

]


test = []
test.append(leave_application_checklist)
test.append(textbox_checklist_example)  # type:ignore
