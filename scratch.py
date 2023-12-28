from reportlab.platypus import SimpleDocTemplate, Flowable, Paragraph
from reportlab.lib.styles import getSampleStyleSheet


style = getSampleStyleSheet()["BodyText"]


class InteractiveCheckBox(Flowable):
    def __init__(self, name, tooltip="", checked=False, size=12, button_style="check"):
        Flowable.__init__(self)
        self.name = name
        self.tooltip = tooltip
        self.size = size
        self.checked = checked
        self.buttonStyle = "check"

    def draw(self):
        self.canv.saveState()
        form = self.canv.acroForm
        form.checkbox(checked=self.checked,
                      buttonStyle=self.buttonStyle,
                      name=self.name,
                      tooltip=self.tooltip,
                      relative=True,
                      size=self.size)
        self.canv.restoreState()


doc = SimpleDocTemplate("/Users/infowizard/Downloads/hello.pdf")
checkbox1 = InteractiveCheckBox("first_name", "First name")
Story = [Paragraph("First Name can this be longer? Lorem ipsum dolor sit amet. Est earum dolorem a minus culpa qui porro amet hic consequatur debitis. Id sequi inventore ut quam aliquid qui totam aperiam aut facilis corporis et repudiandae voluptatibus et vero eveniet eum voluptates Quis. Ut aspernatur delectus rem odit cumque non quod vitae. Sed quae doloremque et magni itaque aut atque quos a omnis blanditiis qui mollitia tempore et asperiores voluptas. Sed sint internos qui nobis tenetur rem nulla labore eos amet quae eos tempora fuga eos delectus recusandae. Id deleniti illum ad delectus expedita vel mollitia voluptatibus rem dicta voluptatem quo repellat quis. Et assumenda alias sed velit assumenda ex neque nihil et veniam rerum qui quia enim eos dignissimos autem non reiciendis fugiat.", style=style), checkbox1]
doc.build(Story)