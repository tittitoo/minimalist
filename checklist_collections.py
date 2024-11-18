""" Checklist Collections

Idea
Checklist: list or dynamically construct a list
Inside:
checkbox: str
choice: dict (item and choices). The last number of the choices list controls widget width
textbox: tuple The number controls the widget's width and height and the last indiciates default value.

Checked for type and take necessary action. If needs be, a list can be constructed
from different checklists.

It is advisable to have checklist only once in particular checklist. The pop() method
could get in the way otherwise.
© Thiha Aung (infowizard@gmail.com)
"""

# The last number is meant for choice box size
YES_NO = "Yes, No, 70"
NO_YES = "No, Yes, 70"
NIL_YES_NO = " , Yes, No, 70"
NA_YES_NO = "NA, Yes, No, 70"
NA_YES_NO_PROGRESS = "NA, Yes, No, In Progress, 70"
NIL_YES_NO_NA = " , Yes, No, NA, 70"
DELIVERY_TERMS = "EXW, FOB, CIF, CPT, FCA, DAP, DDU, DDP, 70"
CREDIT_TERMS = (
    "30 Days, 45 Days, Advanced T/T, COD, 7 Days, 10 Days, 14 Days, LC at Sight, 70"
)
CLASS_SOCIETY = "DNV, ABS, LR, BV, NA, Others, 70"
VALIDITY = "30 Days, 45 Days, 60 Days, 90 Days, 120 Days, 7 Days, 15 Days, 70"
TEXTBOX_HEIGHT = 17
PIC = " , Lin Zar, Oliver, Sahib, Thiha, 70"
SALES = " , Derick, Don, Thiha, Shuan, 70"

LATEST_MINOR_REVISION = "M1"

# Proposal checklist
# Register system here will be availble in drop_down list in excel 'Technical_Notes'
# Need to update the LATEST_MINOR_REVISION number after register update.
available_system_checklist_register = [
    "paga",
]

# This is for generating dynamic checklist
# Need to update the LATEST_MINOR_REVISION number after register update.
available_checklist_register = [
    "ikigai_checklist",
    "cisco_rfq_checklist",
    "rfq_checklist",
    "ndb_checklist",
]

leave_application_checklist = [
    "Have you marked the leave in the team calendar?",
    {
        "For AM or PM leave, have you marked the exact time in the calendar?": NA_YES_NO.split(
            ","
        ),
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


sales_checklist = [  # type:ignore
    ("Job code (As registered in system)", 70, TEXTBOX_HEIGHT, "J12"),
    ("Project name (As registered in system)", 250, TEXTBOX_HEIGHT, ""),
    ("Infrastructure Type (Acronym registered in system)", 70, TEXTBOX_HEIGHT, ""),
    ("Customer name (Acronym registered in system)", 70, TEXTBOX_HEIGHT, ""),
    {
        "Customer type": [
            "Existing",
            "New",
            70,
        ],
        "If existing customer, do we have anything to be aware of, such as difficulity collecting payment, argument on VO, etc.?": [
            "No",
            "Yes",
            "NA",
            "Unknown",
            70,
        ],
    },
    ("If above is 'Yes', eleborate here.", 300, TEXTBOX_HEIGHT * 2, "NA"),
    {
        "If new customer, have we done our due deligence (research and analysis of a company or organization done in preparation for a business transaction)?": NA_YES_NO_PROGRESS.split(
            ","
        ),
    },
    ("Remark if any on the new customer", 300, TEXTBOX_HEIGHT * 2, "NA"),
    (
        "Yard name the vessel will be built in (for new-bulit) or the location the project will be carried out",
        200,
        TEXTBOX_HEIGHT,
        "",
    ),
    ("End user or owner name", 200, TEXTBOX_HEIGHT, ""),
    ("The operating country/region of the vessel", 200, TEXTBOX_HEIGHT, ""),
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
    ("Preferred margin to be quoted in percentage", 70, TEXTBOX_HEIGHT, r"25%"),
    ("Project budget if known", 70, TEXTBOX_HEIGHT, "Not known"),
    (
        "Preferred milestone payment terms",
        300,
        TEXTBOX_HEIGHT * 3,
        r"Default: 30% upon order PO confirmation, 60% before delivery and 10% after commissioning",
    ),
    {
        "Preferred credit terms": CREDIT_TERMS.split(","),
        "Preferred delivery terms": DELIVERY_TERMS.split(","),
    },
    (
        "Delivery location based on above delivery terms (from or to)",
        150,
        TEXTBOX_HEIGHT,
        "Jason Premises (Singapore)",
    ),
    {
        "Quotation validity": VALIDITY.split(","),
    },
    (
        "Warranty duration and/or warranty end date",
        300,
        TEXTBOX_HEIGHT * 2,
        "Default: Twelve (12) months after commissioning or eighteen (18) months after delivery, whichever is earlier.",
    ),
    ("Commissioning location(s) (City or Country)", 200, TEXTBOX_HEIGHT, "Singapore"),
    ("Estimated project delivery date/quarter", 200, TEXTBOX_HEIGHT, ""),
    ("Any special requirement?", 300, TEXTBOX_HEIGHT * 2, "NIL"),
    ("Any known competitor?", 300, TEXTBOX_HEIGHT * 2, "NIL"),
    (
        "Any known risk that you want to highlight, such as project being fast track?",
        300,
        TEXTBOX_HEIGHT * 2,
        "NIL",
    ),
    ("Any remark you want to add?", 300, TEXTBOX_HEIGHT * 2, "NIL"),
    """
Have you gone through all the above points carefully, including the default answers
and answered all of them to the best of your ability? This is an important document that will
be kept as part of the ITB and as a frozen set of information at the point of submission of/sending out the form.
Once you are confident you can attach your signiture to this document, choose your name below and submit/send
the form. As the project progresses and more information is gathered, you may volunteer/be asked to fill up
the form as many times as necessary throughout the project tendering lifecycle.
""",
    {
        "Prepared and confirmed by": SALES.split(","),
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
(5) after approval, do "Summary' or 'Summary Discount' on the file and remove 'discount simulation' in Summary sheet.
""",
    "Organize and clean up '00-ITB' folder. The folders inside are to be named by date and the date format shall be 'yyyy-mm-dd', e.g. '2024-01-29'.",
    {
        "Save the latest CQ or Commercial Clarification in '01-Commercial' folder.": [
            "NA",
            "Done",
            40,
        ],
    },
    "Organize and clean up '02-Technical' folder. All the technical clarifications are to be organized and included along with project schedule.",
    "Organize and clean up '03-Supplier' folder. The latest emails from the supplier \
must be outside and historical reference emails must be in '00-Arc' inside this folder.",
    "Organize and clean up '04-Datasheet' folder.",
    {
        "Save any relevant drawings (block diagrams, DMD, Rack GA, etc.) inside the '05-Drawing' folder.": [
            "NA",
            "Done",
            40,
        ],
    },
    "Keep the PO in '06-PO' folder.",
    {
        "Work out the engineering cost estimater excel and save in '08-Toolkit' folder.": [
            "NA",
            "Done",
            40,
        ]
    },
]

costing = [  # type:ignore
    "Create the folder with the same project name in '@costing' folder.",
    """Create a folder called '00-MAIN' for main order and '01-VO', '02-VO' for
subsequent orders inside the above created folder. For VO items, also include description,
e.g. '01-VO SET-TOP BOX'.
""",
    "Put in the latest commercial proposal PDF.",
    "Put in the contract version of the costing sheet.",
    {"Put in the latest CQ or commercial clarification.": ["NA", "Done", 40]},
]

handover = [  # type:ignore
    "Create a folder with the same project name in '@handover' folder.",
    """Create a folder called '00-MAIN' for main order and '01-VO', '02-VO' for
subsequent orders inside the above created folder. For VO items, also include description,
e.g. '01-VO SET-TOP BOX'.
""",
    "Copy '00-ITB' folder in.",
    "Create a new folder '01-PO' and keep the PO inside.",
    "Copy '02-Technical' folder in.",
    "Copy '03-Supplier' folder in.",
    "Copy '04-Datasheet' folder in.",
    "Create a folder called '05-Cost.'",
    """Generate internal costing sheet from contract file and put in this folder.
Make sure 'COST' value is indicated in summary instead of 'MATERIAL' value.
""",
    {
        "Put in enginnering cost estimator PDF in this folder.": [
            "NA",
            "Done",
            40,
        ]
    },
    {
        "Put in the latest CQ or commercial clarification in this folder.": [
            "NA",
            "Done",
            40,
        ],
        "If drawing exists, create a folder called '06-Drawing' and copy the content from '05-Drawing' folder from '@rfqs'.": [
            "NA",
            "Done",
            40,
        ],
    },
]

in_closing = [  # type:ignore
    "Once all the preparation is done, let the manager review the folder content.",
    "After approval, send the link for '@handover' folder to project management side. Attention to project management head and cc 'Email Group - Project Energy <project@jason.com.sg>'.",
    "Send the link for '@costing' folder to sales support side.",
    """Keep the original and printed copy of the handover checklist in '06-PO' folder in '@rfqs'.
Append the filename of the printed copy of the file with 'Printed', e.g. 'J12473 Handover
Checklist 2024-02-02 Printed.pdf'. Original file is meant to keep track of your progress 
while working, and 'Printed' copy is a frozen information point, which serves audit purpose.
""",
]

rfq_checklist = [  # type:ignore
    "Is there any AVL?",
]

general = [  # type:ignore
    "Have you updated the date and revision number?",
    "Have you done final common sense check of all the numbers? Do they make sense?",
    {
        "Have you checked the cabinet size requirement and quantity? Would all the equipment fit inside the provided cabinet(s)?": NIL_YES_NO_NA.split(
            ","
        ),
        "Have you checked the sever specification requirement, e.g memory and SSD sizes?": NIL_YES_NO_NA.split(
            ","
        ),
        "Have you clarified on the cable gland requirement? Have you put the note to this effect?": NIL_YES_NO_NA.split(
            ","
        ),
        "Is there any AVL? Have you followed as per requirement, such as JB?": NIL_YES_NO_NA.split(
            ","
        ),
        "Have you checked all the IP ratings requirment?": NIL_YES_NO_NA.split(","),
        "Have you checked the material requirement, e.g. SS316?": NIL_YES_NO_NA.split(
            ","
        ),
    },
    "Have you adjusted all the T&C's as per the sales checklist?",
    "Is the class submission required?",
    ("What is the warranty period catered?", 300, TEXTBOX_HEIGHT, ""),
]

engineering_services = [  # type:ignore
    ("Job code", 200, TEXTBOX_HEIGHT, ""),
    ("Project name", 200, TEXTBOX_HEIGHT, ""),
    ("Customer name", 200, TEXTBOX_HEIGHT, ""),
]

confirmation = [
    {
        "This is an important document for quality control. Have you checked all the items carefully?": NIL_YES_NO.split(
            ","
        ),
        "Have you affixed your signature to this affect and printed(virtually)/kept\
        flatten copy of this document for downstreams/audit purpose?": NIL_YES_NO.split(
            ","
        ),
    },
]

paga = [
    {"Loudspeaker required to be fused?": NO_YES.split(",")},
    {"JB required for Loudspeaker": NO_YES.split(",")},
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

lan = [{"Redundancy required?": YES_NO.split(",")}]


# Sales Briefing
sales_briefing = [
    "Something",
]


# New Sales Onborading
sales_onboarding = [
    "Grant edit access to '@commercial-review' folder",
    "Grant read only accesss to '@tools' folder",
    "Grant editor access to Airtable",
    "Setup python & excel",
]

# Setting up python and excel
python_excel_setup = [
    "Have Anaconda distribution installed.",
    """
Search for 'where python' and if it points to 'Windowsapp', take the location of
'Windowsapp' out from the environmental 'Path' variable. We don't need Windows interfering.
""",
    "Put new 'Path' variable pointing to 'anaconda3... python'.",
    "Put new 'Path' variable pointing to 'anaconda3... Script' folder.",
    "Make sure @tools folder is set to 'Always Keep On This Device'.",
    "Close excel if open and run 'xlwings addin install.'",
    "Install reportlab by running 'pip install reportlab.'",
    "In excel xlwings add-in, set the intrepreter path to anaconda python path.",
    "In excel xlwings add-in, set the PYTHONPATH to @tools folder.",
    "Setup the necessary toolbar.",
    "Ask IT to allow the scripts if necessary.",
    "Take note or inform that the excel file may need to be local to the machine to run the tools.",
]

cisco_rfq_checklist = [
    ("End User Full Company Name:", 250, TEXTBOX_HEIGHT, ""),
    ("End User Personnel Name:", 250, TEXTBOX_HEIGHT, ""),
    ("End User Email Address:", 250, TEXTBOX_HEIGHT, ""),
    ("End User Contact Number:", 250, TEXTBOX_HEIGHT, ""),
    ("End User Address:", 300, TEXTBOX_HEIGHT * 2, ""),
    ("Timeline to Award:", 250, TEXTBOX_HEIGHT, ""),
]

ndb_checklist = [
    {"Transmitter Configuration": ["Dual", "Single", 70]},
    {"Transmitter Mounting Type": ["Wall/Bulkhead", "Inside Separate Cabinet", 150]},
    {"Surge Protection For Transmitter Required?": NO_YES.split(",")},
    {"Transmitter Power": ["125W", "250W", 70]},
    {
        "Operating Frequency": [
            "190kHz - 530kHz",
            "536kHz - 1250kHz",
            "1600kHz - 1800kHz",
            120,
        ]
    },
    {"Antenna Type": ["Whip", "Helideck", 70]},
    {"Remote Control At Site (With Serial Connection) Required?": NO_YES.split(",")},
    {
        "Status Monitoring Beacon Receiver (Off-Air Monitoring) Required?": NO_YES.split(
            ","
        )
    },
    {"Remote Connection (Ethernet) Required?": NO_YES.split(",")},
]
