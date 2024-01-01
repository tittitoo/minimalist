# Useful utilities

import os
from pathlib import Path
import subprocess
import requests
from datetime import datetime

from reportlab.pdfgen import canvas
from reportlab.lib.units import inch
from reportlab.lib.pagesizes import A4
from reportlab.lib.colors import lightcyan, black, white, lightyellow

def download_file(path, filename, url):
    """
    path: directory
    filename: filename with extension
    url: url to download
    """
    local_file_path = Path(path, filename)
    if not os.path.exists(local_file_path):
        response = requests.get(url)
        if response.status_code == 200:
            with open(local_file_path, 'wb') as fd:
                for chunk in response.iter_content(chunk_size=8192):
                    fd.write(chunk)
            print(f"Downloaded {local_file_path}")
        else:
            print("Failed to download file.")

# Download necessary files to local machine in 'Documents' folder
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

# Logo as Form. This is for A4 paper currently.
def put_logo(c: canvas.Canvas, logo = ('./resources/Jason_Transparent_Logo_SS.png')):
    c.saveState()
    c.beginForm('logo_Form')
    c.drawImage(logo, 6.5*inch, 780, width=1.25*inch, height=(1.25*inch)*0.224, mask='auto')
    c.endForm()
    c.restoreState()

    c.doForm('logo_Form')

def page_color(c: canvas.Canvas, color=lightcyan):
    c.saveState()
    c.setFillColor(color, alpha=1)
    c.rect(0, 0, c._pagesize[0], c._pagesize[1], stroke=0, fill=1)
    c.restoreState()

def draw_checkbox(c: canvas.Canvas, checklists: list, x: int, y: int, step=20) -> int:
# def draw_checkbox(c, checklists, x, y, step=20):
    """
    Draw checkboxes on the canvas form a list.
    """
    c.saveState()
    form = c.acroForm
    for i, checklist in enumerate(checklists):
        c.setFont('Helvetica', 12)
        print(i, checklist, x, y)
        c.drawString(x, y, str(i+1) + '. ' + checklist)
        form.checkbox(
            name=str(i+1),
            tooltip=f"{i+1}",
            x=x+475,
            y=y,
            buttonStyle="check",
            size=13,
            borderColor=black,
            borderStyle="solid",
            fillColor=white,
            # textColor=black,
            # forceBorder=False,
        )
        y -= step
        if y <= 80:
            c.showPage()
            page_color(c)
            put_logo(c)
            y = 750
    c.restoreState()
    c.showPage()
    return y
