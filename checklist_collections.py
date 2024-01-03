""" Checklist Collections

Idea
Checklist: list or dynamically construct a list
Inside:
checkbox: list: first item str
choice: dict (item and choices)
textbox: list: first item tuple (item and width)

Checked for type and take necessary action. If needs be, a list can be constructed
from different checklists.
"""

YES_NO = [' ', 'Yes', 'No']
YES_NO_NA = [' ', 'Yes', 'No', 'NA']

general_checklist2 = {"Have you done it? How does it really work?": [('', ' '), ('Yes'), ('No')],
                     "Have you not done it?": [('', ' '), ('Yes'), ('No')]}

leave_application_checklist =[
    "Have you marked the leave in the team calendar?",
    "Have you put reminder to put in the email signature for long leave?",
]

ikigai_checklists = [
    "Is it something you are good at? Is it something you are good at? Is it something you are good at? Is it something you are good at?",
    "Is it something you are good at? Is it something you are good at? Is it something you are good at? Is it something you are good at?",
    "Is it something you are good at? Is it something you are good at? Is it something you are good at? Is it something you are good at?",
    "Is it something you are good at? Is it something you are good at? Is it something you are good at? Is it something you are good at?",
    "Is it something you are good at? Is it something you are good at? Is it something you are good at? Is it something you are good at?",
    "Is it something you are good at? Is it something you are good at? Is it something you are good at? Is it something you are good at?",
    "Is it something you are good at? Is it something you are good at? Is it something you are good at? Is it something you are good at?",
    "Is it something you are good at? Is it something you are good at? Is it something you are good at? Is it something you are good at?",
    "Is it something you are good at? Is it something you are good at? Is it something you are good at? Is it something you are good at?",
    "Is it something you are good at? Is it something you are good at? Is it something you are good at? Is it something you are good at?",
    "Is it something you are good at? Is it something you are good at? Is it something you are good at? Is it something you are good at?",
    "Is it something you are good at? Is it something you are good at? Is it something you are good at? Is it something you are good at?",
    "Is it something you are good at? Is it something you are good at? Is it something you are good at? Is it something you are good at?",
    "Is it something you are good at? Is it something you are good at? Is it something you are good at? Is it something you are good at?",
    "Is it something you are good at? Is it something you are good at? Is it something you are good at? Is it something you are good at?"
    "Is it something you are good at? Is it something you are good at? Is it something you are good at? Is it something you are good at?",
    "Is it something you are good at? Is it something you are good at? Is it something you are good at? Is it something you are good at?",
    "Is it something you are good at? Is it something you are good at? Is it something you are good at? Is it something you are good at?",
    "Is it something you are good at? Is it something you are good at? Is it something you are good at? Is it something you are good at?",
    "Is it something you are good at? Is it something you are good at? Is it something you are good at? Is it something you are good at?",
    "Is it something you are good at? Is it something you are good at? Is it something you are good at? Is it something you are good at?",
    "Is it something you are good at? Is it something you are good at? Is it something you are good at? Is it something you are good at?",
    "Is it something you are good at? Is it something you are good at? Is it something you are good at? Is it something you are good at?",
    "Is it something you love to do?",
    "Is it something the world needs?",
    "Does it generate money?",
    "Is it something you are good at?",
    "Is it something you love to do?",
    "Is it something the world needs?",
    "Is it something you are good at? Is it something you are good at? Is it something you are good at? Is it something you are good at?",
    "Does it generate money?",
    "Is it something you are good at?",
    "Is it something you love to do?",
    "Is it something the world needs?",
    "Does it generate money?",
    "Is it something you are good at?",
    "Is it something you love to do?",
    "Is it something the world needs?",
    "Does it generate money?",
    "Is it something you are good at?",
    "Is it something you love to do?",
    "Is it something the world needs?",
    "Does it generate money?",
    "Is it something you are good at?",
    "Is it something you love to do?",
    "Is it something the world needs?",
    "Does it generate money?",
    "Is it something you are good at?",
    "Is it something you love to do?",
    "Is it something the world needs?",
    "Does it generate money?",
    "Is it something you are good at?",
    "Is it something you love to do?",
    "Is it something the world needs?",
    "Does it generate money?",
    "Is it something you are good at?",
    "Is it something you love to do?",
    "Is it something the world needs?",
    "Does it generate money?",
    "Is it something you are good at?",
    "Is it something you love to do?",
    "Is it something the world needs?",
    "Does it generate money?",
    "Is it something you are good at?",
    "Is it something you love to do?",
    "Is it something the world needs?",
    "Does it generate money?",
    "Is it something you are good at?",
    "Is it something you love to do?",
    "Is it something the world needs?",
    "Does it generate money?",
    "Is it something you are good at?",
    "Is it something you love to do?",
    "Is it something the world needs?",
    "Does it generate money?",
    "Is it something you are good at?",
    "Is it something you love to do?",
    "Is it something the world needs?",
    "Does it generate money?",
    "Is it something you are good at?",
    "Is it something you love to do?",
    "Is it something the world needs?",
    "Does it generate money?",
    "Is it something you are good at?",
    "Is it something you love to do?",
    "Is it something the world needs?",
    "Does it generate money?",
    "Is it something you are good at?",
    "Is it something you love to do?",
    "Is it something the world needs?",
    "Does it generate money?",
    "Is it something you are good at?",
    "Is it something you love to do?",
    "Is it something the world needs?",
    "Does it generate money?",
]


general_checklist = {"Have you done it? How does it really work?": [' ', 'Yes', 'No'],
                     "Have you not done it?": [' ', 'Yes', 'No', 'NA']}

textbox_checklist = [('Name let us write something long here so that I know how it will behave.', 120), 
                     ('Firstname let us write something long here so that I know how it will behavaviour.', 100)]