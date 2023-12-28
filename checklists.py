""" 
Creating checklists. This may later be turned into a class.

"""

from reportlab.pdfgen import canvas
from reportlab.lib.units import inch
from reportlab.lib.pagesizes import A4
from reportlab.lib.colors import lightcyan, black, white, lightyellow
# from reportlab.lib.colors import blue, green, yellow
import os
import tempfile
from pathlib import Path
import subprocess
import requests
from datetime import datetime

# Download necessary files to local machine in 'Documents' folder
def download_file(local_path, filename, url):
    """
    local_path: directory
    filename: filename with extendsion
    url: url to download
    """
    local_file_path = Path(local_path, filename)
    if not os.path.exists(local_file_path):
        response = requests.get(url)
        if response.status_code == 200:
            with open(local_file_path, 'wb') as fd:
                for chunk in response.iter_content(chunk_size=8192):
                    fd.write(chunk)
            print(f"Downloaded {local_file_path}")
        else:
            print("Failed to download file.")

try:
    bid = os.path.join(os.path.expanduser('~/Documents'), 'Bid')
    if not os.path.exists(bid):
        os.makedirs(bid)
    # Download Jason Logo
    download_file(bid, 
                  'Jason_Transparent_Logo_SS.png', 
                  'https://filedn.com/liTeg81ShEXugARC7cg981h/Bid/Jason_Transparent_Logo_SS.png')
except Exception as e:
    pass

def put_logo(c: canvas.Canvas):
    c.saveState()
    logo = os.path.join(os.path.expanduser('~/Documents/Bid'), 'Jason_Transparent_Logo_SS.png')
    c.drawImage(logo, 6.5*inch, 780, width=1.25*inch, height=(1.25*inch)*0.224, mask='auto' )
    c.restoreState

def page_color(c: canvas.Canvas, color=lightyellow):
    c.saveState()
    c.setFillColor(color, alpha=1)
    c.rect(0, 0, c._pagesize[0], c._pagesize[1], stroke=0, fill=1)
    c.restoreState()

# The checklist will be written to the download folder of the user.
# Meant to be saved by the user at the desired location.
downloads_folder = os.path.join(os.path.expanduser('~'), 'Downloads')

# temp_directory = tempfile.TemporaryDirectory()
filename = 'Leave Application Checklist.pdf'
file_path = Path(downloads_folder, filename)

leave_checklists =[
    "Have you mark the leave in the team calendar?",
    "Have you put reminder to put in the email signature for long leave?"
]



c = canvas.Canvas(str(file_path), pagesize=A4)
page_color(c)
put_logo(c)
c.setFont('Helvetica-Bold', 15)
c.drawCentredString(300, 750, filename[:-4].upper())
c.setFont('Helvetica-Oblique', 12)
c.drawRightString(A4[0]-50, 730, datetime.now().date().strftime("%Y-%m-%d"))
x, y = 70, 700
form = c.acroForm
for i, checklist in enumerate(leave_checklists):
    c.setFont('Helvetica', 12)
    c.drawString(x, y, str(i+1) + '. ' + checklist)
    form.checkbox(
        name=str(i+1),
        tooltip=f"Item {i+1}",
        x= x+450,
        y=y,
        buttonStyle="check",
        size=13,
        borderColor=None,
        # fillColor=white,
        # textColor=black,
        # forceBorder=False,
        )
    y = y - 25
    if y <= 80:
        c.showPage()
        page_color(c)
        put_logo(c)
        y = 750

c.showPage()
c.save()

# Open in system pdf application
try:
    if os.name == 'posix':
        subprocess.call(['open', str(file_path)])
    elif os.name == 'nt':
        subprocess.call(['start', str(file_path)])
except Exception as e:
    print(f'Unsupported os {e}.')


