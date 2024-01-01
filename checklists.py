""" 
Creating checklists. This may later be turned into a class.

"""
import os
import tempfile
from pathlib import Path
import subprocess
import requests
from datetime import datetime

from reportlab.pdfgen import canvas
from reportlab.lib.units import inch
from reportlab.lib.pagesizes import A4
from reportlab.lib.colors import lightcyan, black, white, lightyellow

import utilities

ants_logo = './resources/ants.png'
# The checklist will be written to the download folder of the user.
# Meant to be saved by the user at the desired location.
downloads_folder = os.path.join(os.path.expanduser('~'), 'Downloads')
filename = 'Ikigai Checklist '+ datetime.now().date().strftime("%Y-%m-%d") + '.pdf'
file_path = Path(downloads_folder, filename)

leave_checklists =[
    "Have you marked the leave in the team calendar?",
    "Have you put reminder to put in the email signature for long leave?",
    "Another",
    "Still another"
]

ikigai_checklists = [
    "Is it something you are good at?",
    "Is it something you love to do?",
    "Is it something the world needs?",
    "Does it generate money?"
    "Is it something you are good at?",
    "Is it something you love to do?",
    "Is it something the world needs?",
    "Does it generate money?"    "Is it something you are good at?",
    "Is it something you love to do?",
    "Is it something the world needs?",
    "Does it generate money?"    "Is it something you are good at?",
    "Is it something you love to do?",
    "Is it something the world needs?",
    "Does it generate money?"    "Is it something you are good at?",
    "Is it something you love to do?",
    "Is it something the world needs?",
    "Does it generate money?"    "Is it something you are good at?",
    "Is it something you love to do?",
    "Is it something the world needs?",
    "Does it generate money?"    "Is it something you are good at?",
    "Is it something you love to do?",
    "Is it something the world needs?",
    "Does it generate money?"    "Is it something you are good at?",
    "Is it something you love to do?",
    "Is it something the world needs?",
    "Does it generate money?"    "Is it something you are good at?",
    "Is it something you love to do?",
    "Is it something the world needs?",
    "Does it generate money?"
]

c = canvas.Canvas(str(file_path), pagesize=A4)
utilities.page_color(c, color=lightyellow)
utilities.put_logo(c)
c.setFont('Helvetica-Bold', 15)
c.drawCentredString(c._pagesize[0]/2, 750, filename[:-15].upper())
c.setFont('Helvetica-Oblique', 12)
c.drawRightString(A4[0]-50, 730, datetime.now().date().strftime("%Y-%m-%d"))

last_position = utilities.draw_checkbox(c, ikigai_checklists, x=50, y=700)
print(last_position)
# x, y = 50, 700
# form = c.acroForm
# for i, checklist in enumerate(ikigai_checklists):
#     c.setFont('Helvetica', 12)
#     c.drawString(x, y, str(i+1) + '. ' + checklist)
#     form.checkbox(
#         name=str(i+1),
#         tooltip=f"Item {i+1}",
#         x= x+475,
#         y=y,
#         buttonStyle="check",
#         size=13,
#         borderColor=black,
#         borderStyle='solid',
#         # borderStyle='dashed',
#         fillColor=white,
#         # textColor=black,
#         # forceBorder=False,
#         )
#     y -= 20
#     if y <= 80:
#         c.showPage()
#         utilities.page_color(c)
#         utilities.put_logo(c, ants_logo)
#         y = 750

# c.showPage()
c.save()

# Open in system pdf application
try:
    if os.name == 'posix':
        subprocess.call(['open', str(file_path)])
    elif os.name == 'nt':
        subprocess.call(['start', str(file_path)], shell=True)
except Exception as e:
    print(f'Unsupported os {e}.')


