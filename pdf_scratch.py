# simple_checkboxes.py

from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfform
from reportlab.lib.colors import magenta, pink, blue, green


def create_simple_checkboxes():
    c = canvas.Canvas("/Users/infowizard/Downloads/simple_checkboxes.pdf")

    c.setFont("Courier", 20)
    c.drawCentredString(300, 700, "Pets")
    form = c.acroForm
    c.setFont("Courier", 14)

    c.drawString(10, 650, "Dog:")
    form.checkbox(
        name="cb1",
        tooltip="Field cb1",
        x=110,
        y=645,
        buttonStyle="check",
        borderColor=magenta,
        fillColor=pink,
        textColor=blue,
        forceBorder=True,
    )

    c.drawString(10, 600, "Cat:")
    form.checkbox(
        name="cb2",
        tooltip="Field cb2",
        x=110,
        y=595,
        buttonStyle="cross",
        borderWidth=2,
        forceBorder=True,
    )

    c.drawString(10, 550, "Pony:")
    form.checkbox(
        name="cb3",
        tooltip="Field cb3",
        x=110,
        y=545,
        buttonStyle="star",
        borderWidth=1,
        forceBorder=True,
    )

    c.drawString(10, 500, "Python:")
    form.checkbox(
        name="cb4",
        tooltip="Field cb4",
        x=110,
        y=495,
        buttonStyle="circle",
        borderWidth=3,
        forceBorder=True,
    )

    c.drawString(10, 450, "Hamster:")
    form.checkbox(
        name="cb5",
        tooltip="Field cb5",
        x=110,
        y=445,
        buttonStyle="diamond",
        borderWidth=None,
        checked=True,
        forceBorder=True,
    )

    c.save()

# simple_form.py

from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfform
from reportlab.lib.colors import magenta, pink, blue, green


def create_simple_form():
    c = canvas.Canvas("/Users/infowizard/Downloads/simple_form.pdf")

    c.setFont("Courier", 18)
    c.drawCentredString(300, 750, "Employment Form")
    c.setFont("Courier", 14)
    form = c.acroForm

    c.drawString(10, 650, "First Name:")
    form.textfield(
        name="fname",
        tooltip="First Name",
        x=110,
        y=635,
        borderStyle="inset",
        borderColor=magenta,
        fillColor=pink,
        width=300,
        textColor=blue,
        forceBorder=True,
    )

    c.drawString(10, 600, "Last Name:")
    form.textfield(
        name="lname",
        tooltip="Last Name",
        x=110,
        y=585,
        borderStyle="inset",
        borderColor=green,
        fillColor=magenta,
        width=300,
        textColor=blue,
        forceBorder=True,
    )

    c.drawString(10, 550, "Address:")
    form.textfield(
        name="address",
        tooltip="Address",
        x=110,
        y=535,
        borderStyle="inset",
        width=400,
        forceBorder=True,
    )

    c.drawString(10, 500, "City:")
    form.textfield(
        name="city", tooltip="City", x=110, y=485, borderStyle="inset", forceBorder=True
    )

    c.drawString(250, 500, "State:")
    form.textfield(
        name="state",
        tooltip="State",
        x=350,
        y=485,
        borderStyle="inset",
        forceBorder=True,
    )

    c.drawString(10, 450, "Zip Code:")
    form.textfield(
        name="zip_code",
        tooltip="Zip Code",
        x=110,
        y=435,
        borderStyle="inset",
        forceBorder=True,
    )

    c.save()

# simple_choices.py

from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfform
from reportlab.lib.colors import magenta, pink, blue, green, red

def create_simple_choices():
    c = canvas.Canvas('/Users/infowizard/Downloads/simple_choices.pdf')
    
    c.setFont("Courier", 20)
    c.drawCentredString(300, 700, 'Choices')
    c.setFont("Courier", 14)
    form = c.acroForm
    
    c.drawString(10, 650, 'Choose a letter:')
    options = [('A','Av'),'B',('C','Cv'),('D','Dv'),'E',('F',),('G','Gv')]
    form.choice(name='choice1', tooltip='Field choice1',
                value='A',
                x=165, y=645, width=72, height=20,
                borderColor=magenta, fillColor=pink, 
                textColor=blue, forceBorder=True, options=options)
  
    c.drawString(10, 600, 'Choose an animal:')
    options = [('Cat', 'cat'), ('Dog', 'dog'), ('Pig', 'pig')]
    form.choice(name='choice2', tooltip='Field choice2',
                value='Cat',
                options=options, 
                x=165, y=595, width=72, height=20,
                borderStyle='solid', borderWidth=1,
                forceBorder=True)
    
    c.save()
    
if __name__ == "__main__":
    # create_simple_checkboxes()
    # create_simple_choices()
    create_simple_form()
    # pass
