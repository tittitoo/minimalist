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

NIL_YES_NO = [' ', 'Yes', 'No']
YES_NO = ['Yes', 'No']
NO_YES = ['No', 'Yes']
NIL_YES_NO_NA = [' ', 'Yes', 'No', 'NA']
NA_YES_NO = ['NA', 'Yes', 'No']

general_checklist2 = {"Have you done it? How does it really work?": [('', ' '), ('Yes'), ('No')],
                     "Have you not done it?": [('', ' '), ('Yes'), ('No')]}

leave_application_checklist =[
    "Have you marked the leave in the team calendar?",
    "For AM or PM leave, have you marked the exact time in the calendar?",
    {"Is the leave longer than 10 days duration including weekends and holidays?": NO_YES,
     """If the above is 'Yes', it is required to put the note in the email signature
       two weeks before the due leave. Have you put the reminder for yourself for this?""": NA_YES_NO,
    "Have you answered all the checklist items carefully?": NO_YES}
]

ikigai_checklists = [
    "Is it something you are good at?",
    "Is it something you love to do?",
    "Is it something the world needs?",
    "Does it generate money?",
]


general_checklist = {"Have you done it? How does it really work?": [' ', 'Yes', 'No'],
                     "Have you not done it?": [' ', 'Yes', 'No', 'NA']}

textbox_checklist = [('Name let us write something long here so that I know how it will behave.', 120), 
                     ('Firstname let us write something long here so that I know how it will behavaviour.', 100)]