""" 
Creating checklists. This may later be turned into a class.

"""
import os
from pathlib import Path
import subprocess
from datetime import datetime

from reportlab.pdfgen import canvas
from reportlab.lib.units import inch
from reportlab.lib.pagesizes import A4
from reportlab.lib.colors import lightcyan, black, white, lightyellow

import utilities
import checklist_collections as cc


# The checklist will be written to the download folder of the user.
# Meant to be saved by the user at the desired location.
downloads_folder = os.path.join(os.path.expanduser('~'), 'Downloads')
filename = 'Leave Checklist '+ datetime.now().date().strftime("%Y-%m-%d") + '.pdf'
file_path = Path(downloads_folder, filename)


c = canvas.Canvas(str(file_path), pagesize=A4)
# utilities.page_color(c)
utilities.put_logo(c)
c.setFont('Helvetica-Bold', 15)
c.drawCentredString(c._pagesize[0]/2, 750, filename[:-15].upper())
c.setFont('Helvetica-Oblique', 11)
c.drawRightString(A4[0]-50, 730, datetime.now().date().strftime("%Y-%m-%d"))
c.setFont('Helvetica', 11)

# last_position = utilities.draw_checkbox(c, leave_checklists, x=70, y=700)
# # last_position = utilities.yes_no_choices(c, general_checklist, x=70, initial=0, y=700)
# last_position = utilities.draw_choice(c, general_checklist2, x=70, initial=(last_position[0]), y=last_position[1])
# last_position = utilities.draw_checkbox(c, leave_checklists, x=70, initial=(last_position[0]), y=last_position[1])
# pdf_scratch.create_simple_choices()

# last_position = utilities.draw_textfield(c, cc.textbox_checklist, x=70, y=700)
# last_position = utilities.draw_checkbox(c, cc.leave_application_checklist, x=70, initial=(last_position[0]), y=last_position[1])
# last_position = utilities.draw_choice(c, cc.general_checklist2, x=70, initial=(last_position[0]), y=last_position[1])

# utilities.produce_checklist(c, leave_checklist)
leave_checklist = ['Something here']
leave_checklist.append(cc.leave_application_checklist)   #type:ignore
# leave_checklist.append(cc.general_checklist)   #type:ignore
# leave_checklist.append(cc.textbox_checklist)   #type:ignore
# leave_checklist.append('Something more here And something else here I can make this sentence long so that I can see')
# leave_checklist.append('Something more here And something else here I can make this sentence long so that I can see')
# leave_checklist.append('Something more here And something else here')
# leave_checklist.append(cc.ikigai_checklists)   #type:ignore

utilities.produce_checklist(c, leave_checklist)
c.showPage
c.save()

# Open in system pdf application
try:
    if os.name == 'posix':
        subprocess.call(['open', str(file_path)])
    elif os.name == 'nt':
        subprocess.call(['start', str(file_path)], shell=True)
except Exception as e:
    print(f'Unsupported os {e}.')


