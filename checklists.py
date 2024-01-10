""" 
Creating checklists. This may later be turned into a class.

"""
import os
from pathlib import Path
from datetime import datetime
from textwrap import wrap
import subprocess
import requests
import xlwings as xw  # type:ignore

from reportlab.pdfgen import canvas
from reportlab.lib.units import inch
from reportlab.lib.pagesizes import A4
from reportlab.lib.colors import black, lightyellow, blue, lightcyan

import checklist_collections as cc
import hide


# Global Variables
RIGHT_MARGIN = 50
PAPERWIDTH = A4[0]
LOGO = os.path.join(
    os.path.expanduser("~/Documents"), "Bid/Jason_Transparent_Logo_SS.png"
)
MAX_TEXTBOX_WIDTH = 250 # If greater than this number, textbox will flow to next line item
# MACRO_NB = xw.Book('PERSONAL.XLSB')


def show_checklist(checklist: list, title="Checklist", color=None):
    """Take checklist and generates pdf in user download folder"""
    downloads_folder = os.path.join(os.path.expanduser("~"), "Downloads")
    filename = f'{title.title()} {datetime.now().date().strftime("%Y-%m-%d")}.pdf'
    file_path = Path(downloads_folder, filename)

    # Create canvas and initialize
    c = canvas.Canvas(str(file_path), pagesize=A4)
    if color:
        page_color(c, color)
    put_logo(c)
    c.setFont("Helvetica-Bold", 15)
    c.drawCentredString(c._pagesize[0] / 2, 750, title.upper())
    c.setFont("Helvetica-Oblique", 11)
    c.drawRightString(A4[0] - 50, 730, datetime.now().date().strftime("%Y-%m-%d"))
    c.setFont("Helvetica", 11)

    produce_checklist(c, checklist, color=color)
    number_page(c)

    c.showPage()
    c.save()
    open_file(file_path)


def open_file(file_path):
    """Open PDF or other files in system default application"""
    try:
        if os.name == "posix":
            subprocess.call(["open", str(file_path)])
        elif os.name == "nt":
            # subprocess.call(["start", str(file_path)], shell=True)
            # subprocess.Popen(["explorer", str(file_path)],
                            #  creationflags=subprocess.DETACHED_PROCESS)
            os.startfile(file_path)
    except Exception as e:
        print(f"Unsupported os {e}.")


# Logo as Form. This is for A4 paper currently.
# Flag is required because form needs to be defined once and the function does not return value
FORM_FLAG = True


def put_logo(c: canvas.Canvas, logo=LOGO):
    c.saveState()
    global FORM_FLAG
    if FORM_FLAG:
        width = 1.25 * inch
        c.beginForm("logo_Form")
        c.drawImage(
            logo,
            PAPERWIDTH - RIGHT_MARGIN - width,
            788,
            width=width,
            height=(1.25 * inch) * 0.224,
            mask="auto",
        )
        c.endForm()
        FORM_FLAG = False
    c.restoreState()

    c.doForm("logo_Form")


def page_color(c: canvas.Canvas, color=lightyellow):
    c.saveState()
    c.setFillColor(color, alpha=1)
    c.rect(0, 0, c._pagesize[0], c._pagesize[1], stroke=0, fill=1)
    c.restoreState()


def draw_checkbox(
    c: canvas.Canvas, checklists: str, x: int, y: int, step=20, initial=0, color=None
) -> tuple:
    """
    Draw checkboxes on the canvas form a list.
    """
    form = c.acroForm
    offset = 3
    if isinstance(checklists, str):
        i = initial
        # c.setFont('Helvetica', 12)
        if i < 9:
            spacer = c.stringWidth("0")
            c.drawString(x + spacer, y, str(i + 1) + ". ")
            skip = c.stringWidth(str(i + 10) + ". ")
        else:
            c.drawString(x, y, str(i + 1) + ". ")
            skip = c.stringWidth(str(i + 1) + ". ")
        form.checkbox(
            name=str(i + 1),
            tooltip=f"{i+1}",
            x=PAPERWIDTH - RIGHT_MARGIN - 13,  # Numerical value is the size
            y=y - offset,
            buttonStyle="check",
            size=13,
            borderColor=black,
            borderWidth=0.5,
            borderStyle="solid",
            fillColor=color,
            # textColor=black,
            # forceBorder=False,
        )
        for line in wrap(checklists, 80):
            c.drawString(x + skip, y, line)
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
        return (i, y)
    return (i, y)


def draw_choice(
    c: canvas.Canvas,
    checklists: dict,
    x=0,
    y=0,
    step=20,
    # width=30,
    initial=0,
    color=None,
) -> tuple:
    form = c.acroForm
    i = initial
    offset = 3
    # c.setFont('Helvetica', 12)
    for k, options in checklists.items():
        if i < 9:
            spacer = c.stringWidth("0")
            c.drawString(x + spacer, y, str(i + 1) + ". ")
            skip = c.stringWidth(str(i + 10) + ". ")
        else:
            c.drawString(x, y, str(i + 1) + ". ")
            skip = c.stringWidth(str(i + 1) + ". ")

        # Get width from last item of the options
        width = float(options.pop())
        wrap_width = int((PAPERWIDTH - width - RIGHT_MARGIN) / c.stringWidth("0"))
        if wrap_width > 80:
            wrap_width = 80
        for n, line in enumerate(wrap(k, wrap_width)):
            c.drawString(x + skip, y, line)
            if n == 0:
                form.choice(  # name='',
                    # tooltip='',
                    value=options,
                    options=options,
                    width=width,
                    height=17,
                    x=PAPERWIDTH - RIGHT_MARGIN - width,
                    y=y - offset,
                    # borderColor=black,
                    borderWidth=0.5,
                    fillColor=color,
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


def draw_textfield(
    c: canvas.Canvas, checklist: tuple, x=0, y=0, step=20, initial=0, color=None
) -> tuple:
    """Checklists here is a list of tuples of 'str' and 'width: int'"""
    form = c.acroForm
    i = initial
    offset = 3
    name, width, height = checklist
    if i < 9:
        spacer = c.stringWidth("0")
        c.drawString(x + spacer, y, str(i + 1) + ". ")
        skip = c.stringWidth(str(i + 10) + ". ")
    else:
        c.drawString(x, y, str(i + 1) + ". ")
        skip = c.stringWidth(str(i + 1) + ". ")
    if width <= MAX_TEXTBOX_WIDTH:
        wrap_width = int((PAPERWIDTH - width - RIGHT_MARGIN) / c.stringWidth("0"))
        if wrap_width > 80:
            wrap_width = 80
    else:
        wrap_width = 80
    # print(wrap_width)
    for n, line in enumerate(wrap(name, wrap_width)):
        c.drawString(x + skip, y, line)
        if n == 0 and width <= MAX_TEXTBOX_WIDTH:
            form.textfield(
                # name="fname",
                # tooltip="First Name",
                x=PAPERWIDTH - RIGHT_MARGIN - width,
                y=y - offset,
                borderStyle="solid",
                borderColor=black,
                borderWidth=0.5,
                fillColor=color,
                width=width,
                height=height,
                # textColor=black,
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
    if width > MAX_TEXTBOX_WIDTH:
        width = PAPERWIDTH - x - RIGHT_MARGIN
        # If the textbox does not fit in the current page, start at next page
        if y - height <= 80:
            number_page(c)
            c.showPage()
            if color:
                page_color(c, color)
            put_logo(c)
            y = 750
        # if y != 750:
        #     y -= step
        form.textfield(
            # name="fname",
            # tooltip="First Name",
            x=PAPERWIDTH - RIGHT_MARGIN - width + skip,
            y=y - offset,
            borderStyle="solid",
            borderColor=black,
            borderWidth=0.5,
            fillColor=color,
            width=width - skip,
            height=height,
            # textColor=black,
            fontSize=11,
            forceBorder=True,
            fieldFlags='multiline'
        )
        y -= step
    return (i, y)


def number_page(c: canvas.Canvas):
    c.saveState()
    c.setFont("Helvetica-Oblique", 11)
    page_number = "Page %s" % c.getPageNumber()
    c.drawCentredString(PAPERWIDTH/2, 60, page_number)
    c.restoreState()


LAST_POSITION = (int, int)


def produce_checklist(
    c: canvas.Canvas,
    checklists: list,
    x=70,
    y=700,
    step=20,
    initial=0,
    # width=30,
    color=None,
):
    global LAST_POSITION
    LAST_POSITION = (initial, y)
    for checklist in checklists:
        if isinstance(checklist, str):
            LAST_POSITION = draw_checkbox(
                c,
                checklist,
                x,
                initial=LAST_POSITION[0],
                y=LAST_POSITION[1],   #type:ignore
                color=color,
            )
        if isinstance(checklist, dict):
            LAST_POSITION = draw_choice(
                c,
                checklist,
                x,
                initial=LAST_POSITION[0],
                y=LAST_POSITION[1],
                # width=width,
                color=color,
            )  # type:ignore
        if isinstance(checklist, tuple):
            LAST_POSITION = draw_textfield(
                c,
                checklist,
                x,
                initial=LAST_POSITION[0],
                y=LAST_POSITION[1],
                color=color,
            )  # type:ignore
        if isinstance(checklist, list):
            produce_checklist(
                c,
                checklist,
                x,
                initial=LAST_POSITION[0],
                y=LAST_POSITION[1],
                color=color,
            )


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
            with open(local_file_path, "wb") as fd:
                for chunk in response.iter_content(chunk_size=8192):
                    fd.write(chunk)
            print(f"Downloaded {local_file_path}")
        else:
            print("Download is not necessary.")


# Download necessary files to local machine in 'Documents' folder
def download_logo():
    try:
        bid = os.path.join(os.path.expanduser("~/Documents"), "Bid")
        if not os.path.exists(bid):
            os.makedirs(bid)
        # Download Jason Logo
        download_file(
            bid,
            "Jason_Transparent_Logo_SS.png",
            "https://filedn.com/liTeg81ShEXugARC7cg981h/Bid/Jason_Transparent_Logo_SS.png",
        )
    except Exception as e:
        print(f"{e} has occured.")


# Can be done as tempfile
def download_template():
    try:
        bid = os.path.join(os.path.expanduser("~/Documents"), "Bid")
        filename = "Template.xlsx"
        file_path = Path(bid, filename)
        # Delete the file if exists
        if os.path.exists(file_path):
            os.remove(file_path)
        download_file(
            bid, filename, "https://filedn.com/liTeg81ShEXugARC7cg981h/Template.xlsx"
        )
        wb = xw.Book.caller()
        wb.app.books.open(file_path.absolute(), password=hide.legacy)
    except Exception as e:
        print(f"Failed to download template -> {e}")


# Can be done as tempfile
def download_planner():
    try:
        bid = os.path.join(os.path.expanduser("~/Documents"), "Bid")
        filename = "Planner.xlsx"
        file_path = Path(bid, filename)
        # Delete the file if exists
        if os.path.exists(file_path):
            os.remove(file_path)
        download_file(
            bid,
            filename,
            "https://filedn.com/liTeg81ShEXugARC7cg981h/Project_Planner_R0.xlsx",
        )
        wb = xw.Book.caller()
        wb.app.books.open(file_path.absolute())
    except Exception as e:
        print(f"Failed to download template -> {e}")


def leave_application_checklist():
    show_checklist(
        cc.leave_application_checklist,
        title="Leave Application Checklist",
        color=lightyellow,
    )


def sales_checklist():
    show_checklist(
        cc.sales_checklist,
        title="Sales Checklist",
        color=lightcyan,
    )

sales_checklist()

# TODO
def proposal_checklist():
    pass