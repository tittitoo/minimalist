""" 
Creating checklists. This may later be turned into a class.
© Thiha Aung (infowizard@gmail.com)
"""

import os
from pathlib import Path
from datetime import datetime
from textwrap import wrap
import subprocess
import requests
import xlwings as xw  # type:ignore
import numpy as np
import pandas as pd

from reportlab.pdfgen import canvas
from reportlab.lib.units import inch
from reportlab.lib.pagesizes import A4
from reportlab.lib.colors import black, lightyellow, lightcyan, honeydew, lavender, blue

import checklist_collections as cc
import hide


# Global Variables
LEFT_MARGIN = 70
RIGHT_MARGIN = 50
PAPERWIDTH = A4[0]
LAST_POSITION = (int, int)
WORD_WRAP = 80
TITLE_LINE = 750
FIRST_NORMAL_LINE = 700
LAST_NORMAL_LINE = 80

MAX_TEXTBOX_WIDTH = (
    250  # If greater than this number, textbox will flow to next line item
)

LOGO = os.path.join(
    os.path.dirname(os.path.realpath(__file__)),
    "resources/Jason_Transparent_Logo_SS.png",
)


def generate_single_checklist(
    checklist: list, title="Checklist", font="Helvetica", font_size=9, color=None
):
    """Take checklist and generates pdf in user download folder."""
    downloads_folder = os.path.join(os.path.expanduser("~"), "Downloads")
    filename = f'{title.title()} {datetime.now().date().strftime("%Y-%m-%d")}.pdf'
    file_path = Path(downloads_folder, filename)

    # Create canvas and initialize
    c = canvas.Canvas(str(file_path), pagesize=A4)
    if color:
        page_color(c, color)
    put_logo(c)
    c.setFont("Helvetica-Bold", 15)
    c.drawCentredString(c._pagesize[0] / 2, TITLE_LINE, title.upper())
    c.setFont("Helvetica-Oblique", font_size)
    c.drawRightString(A4[0] - 50, 730, datetime.now().date().strftime("%Y-%m-%d"))
    c.setFont(font, font_size)

    produce_checklist(c, checklist, font=font, font_size=font_size, color=color)
    number_page(c)

    c.showPage()
    c.save()
    open_file(file_path)


def generate_combined_checklist(
    checklists: list, title="Checklist", font="Helvetica", font_size=9, color=None
):
    """
    Take combined checklists and generates pdf in user download folder.
    Checklist names are printed as titles.
    """
    downloads_folder = os.path.join(os.path.expanduser("~"), "Downloads")
    filename = f'{title.title()} {datetime.now().date().strftime("%Y-%m-%d")}.pdf'
    file_path = Path(downloads_folder, filename)

    # Create canvas and initialize
    c = canvas.Canvas(str(file_path), pagesize=A4)
    if color:
        page_color(c, color)
    put_logo(c)
    c.setFont("Helvetica-Bold", 15)
    c.drawCentredString(c._pagesize[0] / 2, TITLE_LINE, title.upper())
    c.setFont("Helvetica-Oblique", font_size)
    c.drawRightString(A4[0] - 50, 730, datetime.now().date().strftime("%Y-%m-%d"))
    c.setFont(font, font_size)

    global LAST_POSITION
    LAST_POSITION = (0, FIRST_NORMAL_LINE)  # type:ignore
    for item in checklists:
        item = item.lower().replace("-", "_")
        initial = 0
        try:
            checklist = getattr(cc, item)
            LAST_POSITION = draw_title(  # type:ignore
                c,
                item.upper().replace("_", " "),
                initial=initial,
                # initial=LAST_POSITION[0]  # If numbers not to reset
                y=LAST_POSITION[1],
            )
            produce_checklist(
                c,
                checklist,
                initial=LAST_POSITION[0],
                y=LAST_POSITION[1],
                font=font,
                font_size=font_size,
                color=color,
            )
        except Exception as e:
            print(f"Not found {e}")
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
            # creationflags=subprocess.DETACHED_PROCESS)
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
            780,
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


def draw_title(
    c: canvas.Canvas,
    text: str,
    x=LEFT_MARGIN,
    y=700,
    step=20,
    initial=0,
    font="Helvetica-Bold",
    font_size=11,
    color=None,
) -> tuple:
    global LAST_POSITION
    LAST_POSITION = (initial, y)
    c.saveState()
    c.setFont("Helvetica-Bold", font_size)
    wrap_width = WORD_WRAP
    for line in wrap(text, wrap_width):
        c.drawString(x, y, line)
        y -= step
        if y <= LAST_NORMAL_LINE:
            number_page(c, font_size)
            c.showPage()
            if color:
                page_color(c, color)
            put_logo(c)
            c.setFont(font, font_size)
            y = TITLE_LINE
    c.restoreState()
    return (initial, y)


def draw_checkbox(
    c: canvas.Canvas,
    checklists: str,
    x: int,
    y: int,
    step=20,
    initial=0,
    font="Helvetica",
    font_size=11,
    color=None,
) -> tuple:
    """
    Draw checkboxes on the canvas form a list.
    """
    form = c.acroForm
    offset = 3
    word_wrap = WORD_WRAP
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
        for line in wrap(checklists, word_wrap):
            c.drawString(x + skip, y, line)
            y -= step
            if y <= LAST_NORMAL_LINE:
                number_page(c, font_size)
                c.showPage()
                if color:
                    page_color(c, color)
                put_logo(c)
                c.setFont(font, font_size)
                y = TITLE_LINE
        i += 1
        # y -= step
        if y <= LAST_NORMAL_LINE:
            number_page(c)
            c.showPage()
            if color:
                page_color(c, color)
            put_logo(c)
            c.setFont(font, font_size)
            y = TITLE_LINE
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
    font="Helvetica",
    font_size=11,
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
        if wrap_width > WORD_WRAP:
            wrap_width = WORD_WRAP
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
                    fontSize=font_size,
                    # textColor=black,
                    # forceBorder=True,
                )
            y -= step
            if y <= LAST_NORMAL_LINE:
                number_page(c, font_size)
                c.showPage()
                if color:
                    page_color(c, color)
                put_logo(c)
                c.setFont(font, font_size)
                y = TITLE_LINE
        y -= offset
        if y <= LAST_NORMAL_LINE:
            number_page(c, font_size)
            c.showPage()
            if color:
                page_color(c, color)
            put_logo(c)
            c.setFont(font, font_size)
            y = TITLE_LINE
        i += 1
    return (i, y)


def draw_textfield(
    c: canvas.Canvas,
    checklist: tuple,
    x=0,
    y=0,
    step=20,
    initial=0,
    font="Helvetica",
    font_size=11,
    color=None,
) -> tuple:
    """Checklists here is a list of tuples of 'str' and 'width: int'"""
    form = c.acroForm
    i = initial
    offset = 3
    name, width, height, value = checklist
    if i < 9:
        spacer = c.stringWidth("0")
        c.drawString(x + spacer, y, str(i + 1) + ". ")
        skip = c.stringWidth(str(i + 10) + ". ")
    else:
        c.drawString(x, y, str(i + 1) + ". ")
        skip = c.stringWidth(str(i + 1) + ". ")
    if width <= MAX_TEXTBOX_WIDTH:
        wrap_width = int((PAPERWIDTH - width - RIGHT_MARGIN) / c.stringWidth("0"))
        if wrap_width > WORD_WRAP:
            wrap_width = WORD_WRAP
    else:
        wrap_width = WORD_WRAP
    # print(wrap_width)
    for n, line in enumerate(wrap(name, wrap_width)):
        c.drawString(x + skip, y, line)
        if n == 0 and width <= MAX_TEXTBOX_WIDTH:
            form.textfield(
                # name="fname",
                # tooltip="First Name",
                value=value,
                x=PAPERWIDTH - RIGHT_MARGIN - width,
                y=y - offset,
                borderStyle="solid",
                borderColor=black,
                borderWidth=0.5,
                fillColor=color,
                width=width,
                height=height,
                # textColor=black,
                fontName=font,
                fontSize=font_size,
                forceBorder=True,
            )
        y -= step
        if y <= LAST_NORMAL_LINE:
            number_page(c, font_size)
            c.showPage()
            if color:
                page_color(c, color)
            put_logo(c)
            c.setFont(font, font_size)
            y = TITLE_LINE
        i += 1
    if width > MAX_TEXTBOX_WIDTH:
        width = PAPERWIDTH - x - RIGHT_MARGIN
        # If the textbox does not fit in the current page, start at next page
        if y - height <= LAST_NORMAL_LINE:
            number_page(c, font_size)
            c.showPage()
            if color:
                page_color(c, color)
            put_logo(c)
            c.setFont(font, font_size)
            y = TITLE_LINE
        if height > 17:
            offset = offset + (height - 17)
        form.textfield(
            # name="fname",
            # tooltip="First Name",
            value=value,
            x=PAPERWIDTH - RIGHT_MARGIN - width + skip,
            y=y - offset,
            borderStyle="solid",
            borderColor=black,
            borderWidth=0.5,
            fillColor=color,
            width=width - skip,
            height=height,
            # textColor=black,
            fontName=font,
            fontSize=font_size,
            forceBorder=True,
            fieldFlags="multiline",
            maxlen=500,
        )
        if height > 17:
            y -= step + (height - 17)
        else:
            y -= step
    return (i, y)


def number_page(c: canvas.Canvas, font_size=9):
    c.saveState()
    c.setFont("Helvetica-Oblique", font_size)
    page_number = "Page %s" % c.getPageNumber()
    c.drawCentredString(PAPERWIDTH / 2, 60, page_number)
    c.restoreState()


def produce_checklist(
    c: canvas.Canvas,
    checklists: list,
    x=LEFT_MARGIN,
    y=700,
    step=20,
    initial=0,
    # width=30,
    font="Helevtica",
    font_size=10,
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
                y=LAST_POSITION[1],  # type:ignore
                font=font,
                font_size=font_size,
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
                font=font,
                font_size=font_size,
                color=color,
            )  # type:ignore
        if isinstance(checklist, tuple):
            LAST_POSITION = draw_textfield(
                c,
                checklist,
                x,
                initial=LAST_POSITION[0],
                y=LAST_POSITION[1],
                font=font,
                font_size=font_size,
                color=color,
            )  # type:ignore
        if isinstance(checklist, list):
            produce_checklist(
                c,
                checklist,
                x,
                y=LAST_POSITION[1],
                step=step,
                initial=LAST_POSITION[0],
                font=font,
                font_size=font_size,
                color=color,
            )


def leave_application_checklist():
    generate_single_checklist(
        cc.leave_application_checklist,
        title="Leave Application Checklist",
        font_size=11,
        color=lightyellow,
    )


def generate_sales_checklist():
    generate_single_checklist(
        cc.sales_checklist,
        title="Sales Checklist",
        font_size=10,
        color=lightcyan,
    )


def generate_sales_onboarding_checklist():
    generate_single_checklist(
        cc.sales_onboarding,
        title="Sales Onboarding Checklist",
        font_size=10,
        color="",
    )


def generate_proposal_checklist(
    wb,
    proposal_type="firmed",
    title="Firmed Proposal Checklist",
    font="Helvetica",
    font_size=10,
    color=lavender,
):
    # Get system names from the proposal
    ws = wb.sheets["Technical_Notes"]
    last_row = ws.range("F1048576").end("up").row
    job_code = wb.sheets["Config"].range("B29").value
    job_title = ws.range("A1").value
    pic = wb.sheets["Config"].range("B27").value
    data = ws.range(f"F4:F{last_row}").options(pd.DataFrame, index=False).value
    data.columns = ["Systems"]
    data = data.dropna()
    checklist_titles = data.Systems.to_list()
    checklist_titles = ["general"] + checklist_titles
    checklist_titles.append("engineering-services")
    checklist_titles.append("Confirmation")

    # Create file
    downloads_folder = os.path.join(os.path.expanduser("~"), "Downloads")
    filename = (
        f'{job_code} {title.title()} {datetime.now().date().strftime("%Y-%m-%d")}.pdf'
    )
    file_path = Path(downloads_folder, filename)

    # Create canvas and initialize
    c = canvas.Canvas(str(file_path), pagesize=A4)
    if color:
        page_color(c, color)
    put_logo(c)
    c.setFont("Helvetica-Bold", 15)
    c.drawCentredString(c._pagesize[0] / 2, TITLE_LINE, title.upper())
    c.setFont("Helvetica", font_size - 1)
    c.drawString(LEFT_MARGIN, 700, job_title.upper())
    c.setFont("Helvetica-Bold", font_size)
    c.setFillColor(blue)
    c.drawString(LEFT_MARGIN, 680, f"Prepared by: {pic.title()}")
    c.setFillColor(black)
    c.setFont("Helvetica-Oblique", font_size)
    c.drawRightString(A4[0] - 50, 730, datetime.now().date().strftime("%Y-%m-%d"))
    c.setFont(font, font_size)

    global LAST_POSITION
    LAST_POSITION = (0, 660)
    if proposal_type == "firmed":
        for item in checklist_titles:
            item = item.lower().replace("-", "_")
            initial = 0
            try:
                checklist = getattr(cc, item)
                # print(checklist)
                if item == "confirmation":
                    item = item + f" BY {pic.upper()}"
                    LAST_POSITION = draw_title(
                        c,
                        item.upper().replace("_", " "),
                        initial=initial,  # initial=LAST_POSITION[0] if numbers not to reset
                        y=LAST_POSITION[1],
                    )
                else:
                    LAST_POSITION = draw_title(
                        c,
                        item.upper().replace("_", " "),
                        initial=initial,  # initial=LAST_POSITION[0] if numbers not to reset
                        y=LAST_POSITION[1],
                    )
                produce_checklist(
                    c,
                    checklist,
                    initial=LAST_POSITION[0],
                    y=LAST_POSITION[1],
                    font=font,
                    font_size=font_size,
                    color=color,
                )
                initial = 0
            except Exception as e:
                print(f"Not found {e}")
        pass
    else:
        checklist_titles = ["GENERAL", "ENGINEERING-SERVICES"]
        for item in checklist_titles:
            initial = 0
            try:
                checklist = getattr(cc, item.lower().replace("-", "_"))
                # print(checklist)
                LAST_POSITION = draw_title(c, item, initial=initial, y=LAST_POSITION[1])
                produce_checklist(
                    c,
                    checklist,
                    initial=LAST_POSITION[0],
                    y=LAST_POSITION[1],
                    font=font,
                    font_size=font_size,
                    color=color,
                )
                initial = 0
            except Exception as e:
                print(f"Not found {e}")

    number_page(c)
    c.showPage()
    c.save()
    open_file(file_path)


def generate_handover_checklist(
    wb,
    title="Handover Checklist",
    font="Helvetica",
    font_size=10,
    color=honeydew,
):
    # Get system names from the proposal
    job_title = wb.sheets["Summary"].range("A1").value
    job_code = wb.sheets["Config"].range("B29").value
    pic = wb.sheets["Config"].range("B27").value

    # Create file
    downloads_folder = os.path.join(os.path.expanduser("~"), "Downloads")
    filename = (
        f'{job_code} {title.title()} {datetime.now().date().strftime("%Y-%m-%d")}.pdf'
    )
    file_path = Path(downloads_folder, filename)
    # Create canvas and initialize
    c = canvas.Canvas(str(file_path), pagesize=A4)
    if color:
        page_color(c, color)
    put_logo(c)
    c.setFont("Helvetica-Bold", 15)
    c.drawCentredString(c._pagesize[0] / 2, TITLE_LINE, title.upper())
    c.setFont("Helvetica", font_size - 1)
    c.drawString(LEFT_MARGIN, 700, job_title.upper())
    c.setFont("Helvetica-Bold", font_size)
    c.setFillColor(blue)
    c.drawString(LEFT_MARGIN, 680, f"Prepared by: {pic.title()}")
    c.setFillColor(black)
    c.setFont("Helvetica-Oblique", font_size)
    c.drawRightString(A4[0] - 50, 730, datetime.now().date().strftime("%Y-%m-%d"))
    c.setFont(font, font_size)

    global LAST_POSITION
    LAST_POSITION = (0, 660)
    checklist_titles = [
        "@rfqs",
        "@handover",
        "@costing",
        "in_closing",
    ]
    for item in checklist_titles:
        initial = 0
        try:
            if "@" in item:
                checklist = getattr(cc, item.lower().replace("@", ""))
                LAST_POSITION = draw_title(
                    c,
                    (item + " folder"),
                    initial=initial,
                    y=LAST_POSITION[1],
                )
            else:
                checklist = getattr(cc, item.lower())
                # print(checklist)
                LAST_POSITION = draw_title(
                    c,
                    item.capitalize().replace("_", " "),
                    initial=initial,
                    y=LAST_POSITION[1],
                )
            produce_checklist(
                c,
                checklist,
                initial=LAST_POSITION[0],
                y=LAST_POSITION[1],
                font=font,
                font_size=font_size,
                color=color,
            )
            initial = 0
        except Exception as e:
            print(f"Not found {e}")

    number_page(c)
    c.showPage()
    c.save()
    open_file(file_path)


if __name__ == "__main__":
    generate_single_checklist(
        cc.sales_checklist,
        title="Sales Checklist",
        font_size=10,
        color="",
    )

    # checklists = [
    #     "vhf-fm",
    #     "engineering_services",
    #     "none",
    #     "sales_onboarding",
    #     "ikigai_checklist",
    #     "sales_checklist",
    #     "handover",
    #     "costing",
    # ]
    # generate_combined_checklist(
    #     checklists=checklists,
    #     title="Test",
    #     font_size=10,
    #     color=lavender,
    # )