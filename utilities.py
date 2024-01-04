# Useful utilities

import os
from pathlib import Path
# import subprocess
# import requests
# from datetime import datetime
from textwrap import wrap

from reportlab.pdfgen import canvas
from reportlab.lib.units import inch
from reportlab.lib.pagesizes import A4
from reportlab.lib.colors import lightcyan, black, white, lightyellow, blue

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

# Global Variables

RIGHT_MARGIN = 50
PAPERWIDTH = A4[0]

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

def page_color(c: canvas.Canvas, color=lightyellow):
    c.saveState()
    c.setFillColor(color, alpha=1)
    c.rect(0, 0, c._pagesize[0], c._pagesize[1], stroke=0, fill=1)
    c.restoreState()

def draw_checkbox(c: canvas.Canvas, checklists: str, x: int, y: int, step=20, initial=0, color=None) -> tuple[int, int]:
    """
    Draw checkboxes on the canvas form a list.
    """
    form = c.acroForm
    offset = 3
    if isinstance(checklists, str):
        i = initial
        # c.setFont('Helvetica', 12)
        if i < 9:
            spacer = c.stringWidth('0')
            c.drawString(x+spacer, y, str(i+1) + '. ')
            skip = c.stringWidth(str(i+10) + '. ')
        else:
            c.drawString(x, y, str(i+1) + '. ')
            skip = c.stringWidth(str(i+1) + '. ')
        form.checkbox(
            name=str(i+1),
            tooltip=f"{i+1}",
            x=PAPERWIDTH - RIGHT_MARGIN - 13, # 13 is the size
            y=y-offset,
            buttonStyle="check",
            size=13,
            borderColor=black,
            borderStyle="solid",
            fillColor=white,
            # textColor=black,
            # forceBorder=False,
        )
        for line in wrap(checklists, 80):
            c.drawString(x+skip, y, line)
            y -= step
        i += 1
        # y -= step
        if y <= 80:
            number_page(c)
            c.showPage()
            if color:
                page_color(c, color)
            put_logo(c)
            y = 750
        return(i, y)
    return (i, y)

def draw_choice(c: canvas.Canvas, checklists: dict, x=0, y=0, step=20, width=40, initial=0, color=None) -> tuple[int, int]:
    form = c.acroForm
    i = initial
    offset = 3
    # c.setFont('Helvetica', 12)
    for k, options in checklists.items():
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
                            value=options, 
                            options=options,
                            width=width, 
                            height=17,
                            x=PAPERWIDTH - RIGHT_MARGIN - width,
                            y=y-offset, 
                            # borderColor=black, 
                            fillColor=white,
                            fontSize=11, 
                            # textColor=black, 
                            # forceBorder=True,
                            )
            y -= step
            if y <= 80:
                number_page(c)
                c.showPage()
                if color:
                    page_color(c, color)
                put_logo(c)
                y = 750
        y -= offset
        if y <= 80:
            number_page(c)
            c.showPage()
            if color:
                page_color(c, color)
            put_logo(c)
            y = 750
        i += 1
    return (i, y)

def draw_textfield(c: canvas.Canvas, checklist: tuple, x=0, y=0, step=20, initial=0, color=None) -> tuple[int, int]:
    """ Checklists here is a list of tuples of 'str' and 'width: int'"""
    form = c.acroForm
    i = initial
    offset = 3
    # c.setFont('Helvetica', 12)
    name, width = checklist   # Unpack tuple
    if i < 9:
        spacer = c.stringWidth('0')
        c.drawString(x+spacer, y, str(i+1) + '. ')
        skip = c.stringWidth(str(i+10) + '. ')
    else:
        c.drawString(x, y, str(i+1) + '. ')
        skip = c.stringWidth(str(i+1) + '. ')
    wrap_width = int((PAPERWIDTH - width - RIGHT_MARGIN) / c.stringWidth('0'))
    if wrap_width > 80:
        wrap_width = 80
    # print(wrap_width)
    for n, line in enumerate(wrap(name, wrap_width)):
        c.drawString(x+skip, y, line)
        if n == 0:
            form.textfield(
                # name="fname",
                # tooltip="First Name",
                x=PAPERWIDTH - RIGHT_MARGIN - width,
                y=y-offset,
                borderStyle="solid",
                borderColor=black,
                fillColor=white,
                width=width,
                height=17,
                textColor=blue,
                fontSize=11,
                forceBorder=True,
            )
        y -= step
        if y <= 80:
            number_page(c)
            c.showPage()
            if color:
                page_color(c, color)
            put_logo(c)
            y = 750
        i += 1
    return (i, y)

def number_page(c: canvas.Canvas):
    c.saveState()
    c.setFont('Helvetica-Oblique', 11)
    page_number = 'Page %s' % c.getPageNumber()
    # c.drawRightString(PAPERWIDTH - RIGHT_MARGIN, 60, page_number)
    c.drawCentredString(PAPERWIDTH/2, 60, page_number)
    c.restoreState()

LAST_POSITION = (int, int)
def produce_checklist(c: canvas.Canvas, checklists: list, x=70, y=700, step=20, initial=0, width=40, color=None):
    global LAST_POSITION
    LAST_POSITION = (initial, y)
    for checklist in checklists:
        if isinstance(checklist, str):
            LAST_POSITION = draw_checkbox(c, checklist, x, initial=LAST_POSITION[0], y=LAST_POSITION[1])  #type:ignore 
        if isinstance(checklist, dict):
            LAST_POSITION = draw_choice(c, checklist, x, initial=LAST_POSITION[0], y=LAST_POSITION[1], width=width)   #type:ignore
        if isinstance(checklist, tuple):
            LAST_POSITION = draw_textfield(c, checklist, x, initial=LAST_POSITION[0], y=LAST_POSITION[1])   #type:ignore
        if isinstance(checklist, list):
            produce_checklist(c, checklist, x, initial=LAST_POSITION[0], y=LAST_POSITION[1])