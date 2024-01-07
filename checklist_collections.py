""" Checklist Collections

Idea
Checklist: list or dynamically construct a list
Inside:
checkbox: str
choice: dict (item and choices). The last number of the choices list controls widget width
textbox: list: Tuple. The number controls the widget's width.

Checked for type and take necessary action. If needs be, a list can be constructed
from different checklists.
"""

# The last item controls the widget's width and wrap_width.
NIL_YES_NO = [" ", "Yes", "No", 30]
YES_NO = ["Yes", "No", 35]
NO_YES = ["No", "Yes", 30]
NIL_YES_NO_NA = [" ", "Yes", "No", "NA", 30]
NA_YES_NO = ["NA", "Yes", "No", 30]

general_checklist2 = {
    "Have you done it? How does it really work?": [("", " "), ("Yes"), ("No"), 30],
    "Have you not done it?": [("", " "), ("Yes"), ("No"), 30],
}

leave_application_checklist = [
    "Have you marked the leave in the team calendar?",
    "For AM or PM leave, have you marked the exact time in the calendar?",
    {
        "Is the leave longer than 10 days duration including weekends and holidays?": NO_YES,
        """If the above is 'Yes', it is required to put the note in the email signature 
        two weeks before the due leave. Have you put the reminder for yourself for this?""": NA_YES_NO,
        "You are responsible for filling out this checklist. Have you answered all the checklist items carefully?": NIL_YES_NO,
    },
]

ikigai_checklists = [
    "Is it something you are good at?",
    "Is it something you love to do?",
    "Is it something the world needs?",
    "Does it generate money?",
]


general_checklist = {
    "Have you done it? How does it really work?": [" ", "Yes", "No", 30],
    "Have you not done it?": [" ", "Yes", "No", "NA", 30],
}

textbox_checklist = [
    ("Name let us write something long here so that I know how it will behave.", 120),
    (
        "Firstname let us write something long here so that I know how it will behavaviour.",
        100,
    ),
]

test = []
test.append(leave_application_checklist)
test.append(textbox_checklist)  # type:ignore
