# Useful utilities

import os
from pathlib import Path
import subprocess
# import requests
from datetime import datetime
from textwrap import wrap

from reportlab.pdfgen import canvas
from reportlab.lib.units import inch
from reportlab.lib.pagesizes import A4
from reportlab.lib.colors import lightcyan, black, white, lightyellow

# def download_file(path, filename, url):
#     """
#     path: directory
#     filename: filename with extension
#     url: url to download
#     """
#     local_file_path = Path(path, filename)
#     if not os.path.exists(local_file_path):
#         response = requests.get(url)
#         if response.status_code == 200:
#             with open(local_file_path, 'wb') as fd:
#                 for chunk in response.iter_content(chunk_size=8192):
#                     fd.write(chunk)
#             print(f"Downloaded {local_file_path}")
#         else:
#             print("Failed to download file.")

# # Download necessary files to local machine in 'Documents' folder
# try:
#     bid = os.path.join(os.path.expanduser('~/Documents'), 'Bid')
#     if not os.path.exists(bid):
#         os.makedirs(bid)
#     # Download Jason Logo
#     download_file(bid, 
#                   'Jason_Transparent_Logo_SS.png', 
#                   'https://filedn.com/liTeg81ShEXugARC7cg981h/Bid/Jason_Transparent_Logo_SS.png')
# except Exception as e:
#     pass

# Logo as Form. This is for A4 paper currently.
# Flag is required because form needs to be defined once and the function does not return value
FORM_FLAG = True
def put_logo(c: canvas.Canvas, logo = ('./resources/Jason_Transparent_Logo_SS.png')):
    c.saveState()
    global FORM_FLAG
    if FORM_FLAG:
        c.beginForm('logo_Form')
        c.drawImage(logo, 6.5*inch, 780, width=1.25*inch, height=(1.25*inch)*0.224, mask='auto')
        c.endForm()
        FORM_FLAG = False
    c.restoreState()

    c.doForm('logo_Form')

def page_color(c: canvas.Canvas, color=lightcyan):
    c.saveState()
    c.setFillColor(color, alpha=1)
    c.rect(0, 0, c._pagesize[0], c._pagesize[1], stroke=0, fill=1)
    c.restoreState()

def draw_checkbox(c: canvas.Canvas, checklists: list, x: int, y: int, step=20, initial=0, color=None) -> tuple:
    """
    Draw checkboxes on the canvas form a list.
    """
    form = c.acroForm
    offset = 3
    for i, checklist in enumerate(checklists):
        i += initial
        c.setFont('Helvetica', 12)
        if i < 9:
            spacer = c.stringWidth('0')
            c.drawString(x+spacer, y, str(i+1) + '. ')
            skip = c.stringWidth(str(i+10) + '. ')
        else:
            c.drawString(x, y, str(i+1) + '. ')
            skip = c.stringWidth(str(i+1) + '. ')
        for n, line in enumerate(wrap(checklist, 80)):
            c.drawString(x+skip, y, line)
            if n == 0:
                form.checkbox(
                    name=str(i+1),
                    tooltip=f"{i+1}",
                    x=x+465,
                    y=y-offset,
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
                if color:
                    page_color(c, color)
                put_logo(c)
                y = 750
        y -= offset
        if y <= 80:
            c.showPage()
            if color:
                page_color(c, color)
            put_logo(c)
            y = 750
    # c.showPage()
    return (i+1, y)

def yes_no_choices(c: canvas.Canvas, checklists: dict, x=0, y=0, step=20, initial=0, color=None) -> tuple:
    form = c.acroForm
    i = initial
    offset = 3
    c.setFont('Helvetica', 12)
    for k, v in checklists.items():
        print(k)
        # print(v)
        if i < 9:
            spacer = c.stringWidth('0')
            c.drawString(x+spacer, y, str(i+1) + '. ')
            skip = c.stringWidth(str(i+10) + '. ')
        else:
            c.drawString(x, y, str(i+1) + '. ')
            skip = c.stringWidth(str(i+1) + '. ')
        for n, line in enumerate(wrap(k, 80)):
            c.drawString(x+skip, y, line)
            if n == 0:
                form.choice(# name='', 
                            # tooltip='',
                            value=v[0][1], # Take the second value of the first tuple
                            options=v,
                            x=x+465-(40/2),   # 40 is width
                            y=y-offset, 
                            width=40, 
                            height=18,
                            # borderColor=black, 
                            fillColor=white,
                            fontSize=11, 
                            # textColor=black, 
                            # forceBorder=True,
                            )
            y -= step
            if y <= 80:
                c.showPage()
                if color:
                    page_color(c, color)
                put_logo(c)
                y = 750
        y -= offset
        if y <= 80:
            c.showPage()
            if color:
                page_color(c, color)
            put_logo(c)
            y = 750
        i += 1
    return (i, y)
