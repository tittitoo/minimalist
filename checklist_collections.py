""" Checklist Collections

Idea
Checklist: list or dynamically construct a list
Inside:
checkbox: str
choice: dict (item and choices). The last number of the choices list controls widget width
textbox: list: Tuple. The number controls the widget's width and height

Checked for type and take necessary action. If needs be, a list can be constructed
from different checklists.

It is advisable to have checklist only once in particular checklist. The pop() method
could get in the way otherwise.
© Thiha Aung (infowizard@gmail.com)
"""

YES_NO = "Yes, No, 70"
NO_YES = "No, Yes, 70"
NIL_YES_NO = " , Yes, No, 70"
NA_YES_NO = "NA, Yes, No, 70"
NIL_YES_NO_NA = " , Yes, No, NA, 70"
DELIVERY_TERMS = "EXW, FOB, CIF, CPT, FCA, DAP, DDU, DDP, 70"
CREDIT_TERMS = (
    "30 Days, 45 Days, Advanced T/T, COD, 7 Days, 10 Days, 14 Days, LC at Sight, 70"
)
CLASS_SOCIETY = "NA, DNV, ABS, LR, BV, Others, 70"
VALIDITY = "30 Days, 45 Days, 60 Days, 90 Days, 120 Days, 7 Days, 15 Days, 70"
TEXTBOX_HEIGHT = 17
PIC = " , Lin Zar, Oliver, Sahib, Thiha, 70"
SALES = " , Shaun, Derick, Thiha, 70"

leave_application_checklist = [
    "Have you marked the leave in the team calendar?",
    "For AM or PM leave, have you marked the exact time in the calendar?",
    {
        "Is the leave longer than 10 days duration including weekends and holidays?": NO_YES.split(
            ","
        ),
        """If the above is 'Yes', it is required to put the note in the email signature 
        two weeks before the due leave. Have you put the reminder for yourself for this?""": NA_YES_NO.split(
            ","
        ),
        "You are responsible for filling out this checklist. Have you answered all the checklist items carefully?": NIL_YES_NO.split(
            ","
        ),
        "Prepared by:": PIC.split(","),
    },
]

ikigai_checklist = [
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


sales_checklist = [  # type:ignore
    ("Job code", 200, TEXTBOX_HEIGHT),
    ("Project name", 200, TEXTBOX_HEIGHT),
    ("Customer name", 200, TEXTBOX_HEIGHT),
    {
        "Customer type": ["Existing", "New", 70],
        "If 'Existing Customer', do we have any past concerns/issues with the customer we need to be aware of?": [
            "No",
            "Yes",
            "Unknown",
            70,
        ],
    },
    ("If above is 'Yes', please state here.", 300, TEXTBOX_HEIGHT * 2),
    ("Yard name the vessel will be built in", 200, TEXTBOX_HEIGHT),
    ("End user or owner name", 200, TEXTBOX_HEIGHT),
    ("Infrastructure Type", 200, TEXTBOX_HEIGHT),
    ("The operating country/region of the vessel", 200, TEXTBOX_HEIGHT),
    {
        "Classification society": CLASS_SOCIETY.split(","),
        "Is there any NDA in place?": NO_YES.split(","),
    },
    {
        "Jason entity to be quoted under": ["Jason Electronics", "Jason Energy", 120],
        "Currency to be quoted in": ["SGD", "USD", 70],
        "Type of quotation": ["Budgetary", "Firmed", 70],
        # "Have we received all the required information.": [' ', 'Yes', 'No', 'Not Sure', 70],
    },
    ("Preferred margin to be quoted in percentage", 70, TEXTBOX_HEIGHT),
    ("Project budget if known", 70, TEXTBOX_HEIGHT),
    ("Preferred milestone payment terms", 300, TEXTBOX_HEIGHT * 3),
    {
        "Preferred credit terms": CREDIT_TERMS.split(","),
        "Preferred delivery terms": DELIVERY_TERMS.split(","),
    },
    ("Delivery location based above delivery terms (to or from)", 150, TEXTBOX_HEIGHT),
    {
        "Quotation validity": VALIDITY.split(","),
    },
    ("Warranty duration and/or warranty end date", 300, TEXTBOX_HEIGHT * 2),
    ("Commissioning location", 200, TEXTBOX_HEIGHT),
    ("Estimated project delivery date", 200, TEXTBOX_HEIGHT),
    ("Any special requirement?", 300, TEXTBOX_HEIGHT * 2),
    ("Any known competitor?", 300, TEXTBOX_HEIGHT * 2),
    ("Any known risk that you want to highlight?", 300, TEXTBOX_HEIGHT * 2),
    ("Any remark you want to add?", 300, TEXTBOX_HEIGHT * 2),
    {
        "Prepared by": SALES.split(","),
    },
]

general = [  # type:ignore
    "Here",
]


engineering_services = [  # type:ignore
    ("Job code", 200, TEXTBOX_HEIGHT),
    ("Project name", 200, TEXTBOX_HEIGHT),
    ("Customer name", 200, TEXTBOX_HEIGHT),
]


paga = [
    "There",
]

vhf_am = [
    "Everywhere",
    "Nowhere",
    "Anywhere",
    "Another item",
    {
        "Classification society": CLASS_SOCIETY.split(","),
        "Is there any NDA in place?": NO_YES.split(","),
    },
]

vhf_fm = [
    "Everywhere",
    "Nowhere",
    "Anywhere",
    "Another item",
    {
        "Classification society": CLASS_SOCIETY.split(","),
        "Is there any NDA in place?": NO_YES.split(","),
    },
]
confirmation = [
    {
        "This is an important document for quality control. Have you checked all the items carefully?": NIL_YES_NO.split(
            ","
        ),
        "Have you affixed your signature to this affect and printed(virtually)/kept frozen copy of this document for downstreams/audit purpose?": NIL_YES_NO.split(
            ","
        ),
    },
]


# Handover checklist
rfqs = [  # type:ignore
    """
Produce the contract version of the costing sheet:
(1) fix unit prices (FUP), 
(2) update the latest cost (where applicable),
(3) append "CONTRACT" to the filename.
(4) seek manager's approval for the contract file once prepared.
""",
    "Organize and clean up '00-ITB' folder. The folders inside are to be named by date and the date format shall be 'yyyy-mm-dd', e.g. '2024-01-29'.",
    "Save the latest CQ in '01-Commercial' folder.",
    "Organize and clean up '02-Technical' folder. All the technical clarifications are to be organized and included along with project schedule.",
    "Organize and clean up '03-Supplier' folder. The latest emails from the supplier must be outside and historical reference emails must be in '00-Arc' inside this folder.",
    "Organize and clean up '04-Datasheet' folder.",
    "Save any relevant drawings (block diagrams, DMD, Rack GA, etc.) inside the '05-Drawing' folder.",
    "Keep the PO in '06-PO' folder.",
    "Work out the engineering cost estimater excel and save in '08-Toolkit' folder.",
]

costing = [  # type:ignore
    "Create the folder with the same project name in '@costing' folder.",
    "Put in the latest commercial proposal PDF.",
    "Put in the contract version of the costing sheet.",
    "Put in the latest CQ or commercial clarification.",
]

handover = [  # type:ignore
    "Create a folder with the same project name in '@handover' folder.",
    "Crate a folder called '00-MAIN' for main order and '01-VO', '02-VO' for subsequent orders inside the above created folder. For VO items, also include description, e.g. '01-VO SET-TOP BOX'.",
    "Copy '00-ITB' folder in.",
    "Create a new folder '01-PO' and keep the PO inside.",
    "Copy '02-Technical' folder in.",
    "Copy '03-Supplier' folder in.",
    "Copy '04-Datasheet' folder in.",
    "Generate internal costing sheet from the contract version. Make sure to either do 'Summary' or 'Summary Discount' first. The value to give to project management is 'COST', 'MATERIAL'.",
    "Create a folder called '05-Cost.'",
    "Put in generated internal costing sheet this folder.",
    "Put in enginnering cost estimator PDF in this folder.",
    "Put in the latest CQ or commercial clarification if applicable in this folder.",
    "If drawing exists, create a folder called '06-Drawing' and copy the content from '05-Drawing' folder from '@rfqs'.",
]

in_closing = [  # type:ignore
    "Once all the preparation is done, let the manager review the folder content.",
    "After approval, send the link for '@handover' folder to project management side.",
    "Send the link for '@costing' folder to sales support side.",
    "Keep the printed copy of the handover checklist in '06-PO' folder.",
]
