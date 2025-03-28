#solidworks_dxf_gib_R.py

# -*- coding: utf-8 -*-
import win32com.client
import os
import sys
import time
import builtins

print("Running the script...")

# Restore str
str = builtins.str
print(f"Type of str after restoration: {type(str)}")

# Connect to an existing SolidWorks instance
try:
    swApp = win32com.client.GetActiveObject("SldWorks.Application")
    print("Connected to an existing SolidWorks instance")
    swApp.Visible = True
    time.sleep(2)
except Exception as e:
    print(f"Error connecting to SolidWorks (GetActiveObject): {e}")

try:
    swApp = win32com.client.Dispatch("SldWorks.Application")
    print("Created a new SolidWorks instance")
    swApp.Visible = True
    time.sleep(2)
except Exception as e:
    print(f"Error creating a new SolidWorks instance: {e}")
    sys.exit(1)

# Path to the file
part_path = r"C:\path\to\your\sldprt_file.SLDPRT"
folder_path = r"C:\path\to\your\folder"

# Check if the folder and files exist
if os.path.exists(folder_path):
    print(f"Folder found: {folder_path}")
    print("Folder contents:")
    for file in os.listdir(folder_path):
        print(f" - {file}")
else:
    print(f"Folder not found: {folder_path}")
    sys.exit(1)

# Check if the file exists
if not os.path.exists(part_path):
    print(f"File not found: {part_path}")
    sys.exit(1)
else:
    print(f"File found: {part_path}")

# Close all documents
swApp.CloseAllDocuments(True)
time.sleep(2)

# Open the part
try:
    print(f"Type of part_path: {type(part_path)}")
    swModel = swApp.OpenDoc(part_path, 1) # 1 = swDocPART
    if swModel is None:
        print("Failed to open the file (swModel is None)")
    else:
        title = swModel.GetTitle # Call separately
        print(f"File successfully opened: {title}")
        
        # Export the flat pattern
        dxf_path = r"C:\path\to\your\folder\{}.dxf".format(title)
        success = swModel.ExportFlatPatternView(dxf_path, 1)
        if success:
            print(f"DXF saved: {dxf_path}")
        else:
            print("DXF export failed")
        
        # Close the file
        swApp.CloseDoc(title)
except Exception as e:
    print(f"Error working with the file: {e}")
    sys.exit(1)
