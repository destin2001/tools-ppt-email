#!/usr/bin/env python
import os
import win32com.client as win32
from pptx_tools import utils  # needs python-pptx-interface installed

# virtual environment activate
venv_path = os.path.join(
    r"C:\Destin_Nguyen\develop\dev2001\tools-ppt-email\.venv\Scripts\activate_this.py"
)
activate_this = os.path.abspath(venv_path)
exec(open(activate_this).read(), {"__file__": activate_this})

# you should use full paths, to make sure PowerPoint can handle the paths
fixed_width = 1024  # default of pptx
fixed_height = 768  # default of pptx
png_folder = r"C:\Destin_Nguyen\develop\dev2001\tools-ppt-email\temp"
presentation_path = r"C:\Destin_Nguyen\develop\dev2001\tools-ppt-email\input\demo.pptx"

# Save slides to png
utils.save_pptx_as_png(png_folder, presentation_path, overwrite_folder=True)

# Initialize outlook application
outlook = win32.Dispatch("Outlook.Application")

# Create a new mail item
mail = outlook.CreateItem(0)

# Sorted list of files
file_list = sorted(
    os.listdir(png_folder), key=lambda x: int("".join(filter(str.isdigit, x)))
)

# Append images into body email Outlook
for file_name in file_list:
    image_path = os.path.join(png_folder, file_name)

    attachment = mail.Attachments.Add(image_path)
    attachment.PropertyAccessor.SetProperty(
        "http://schemas.microsoft.com/mapi/proptag/0x3712001E", file_name
    )

    image_html = (
        f"<img src='cid:{file_name}' width='{fixed_width}' height='{fixed_height}'>"
    )
    mail.HTMLBody += "<br>" + image_html

mail.Display()
