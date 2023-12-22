from reportlab.pdfgen import canvas
from reportlab.lib.colors import blue, green, yellow
import os
import tempfile
from pathlib import Path
import subprocess

temp_directory = tempfile.TemporaryDirectory()
filename = 'checklist.pdf'
c = canvas.Canvas(str(Path(temp_directory.name, filename)))
c.drawString(10, 650, "Dog:")
c.showPage()
c.save()
# os.system(str(Path(temp_directory.name, filename)))
# subprocess.Popen(Path(temp_directory.name, filename))
# subprocess.call(Path(temp_directory.name, filename))