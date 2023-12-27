""" 
Creating checklists. This may later be turned into a class.

"""

from reportlab.pdfgen import canvas
# from reportlab.lib.colors import blue, green, yellow
import os
import tempfile
from pathlib import Path
import subprocess
import requests

# The checklist will be written to the download folder of the user.
# Meant to be saved by the user at the desired location.
downloads_folder = os.path.join(os.path.expanduser('~'), 'Downloads')

# temp_directory = tempfile.TemporaryDirectory()
filename = 'checklist.pdf'
file_path = Path(downloads_folder, filename)

# Download necessary files to local machine in 'Documents' folder
try:
    bid = os.path.join(os.path.expanduser('~/Documents'), 'Bid')
    if not os.path.exists(bid):
        os.makedirs(bid)
except Exception as e:
    pass



c = canvas.Canvas(str(file_path))
c.drawString(100, 650, "Dog:")
c.drawString(100, 550, "Mee Mee:")
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

print(file_path)

