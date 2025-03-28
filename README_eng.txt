## Introduction

This repository contains two Python scripts designed to work with 3D models and export data to the DXF format. The first script, `step_to_dxf.py`, loads STEP files and exports the contours of the largest face to DXF. The second script, `solidworks_dxf_gib_R.py`, interacts with SolidWorks to export the flat pattern of a part to DXF.

## step_to_dxf.py

### Purpose

This script is designed to load STEP files and export the contours of the largest face to DXF. It uses the Open CASCADE library for working with 3D models and ezdxf for creating DXF files.

### Functionality

- Loading STEP Files: The script loads a STEP file and extracts all faces from it.
- Identifying the Largest Face: It finds the face with the maximum area.
- Exporting Contours: It extracts the contours of the largest face and exports them to DXF, projecting onto the XY, YZ, or XZ plane based on the longest contour length.
- Calculating Areas and Lengths: It calculates the area of the largest face and the total length of the contours.

### Usage

1. Install the necessary libraries: Open CASCADE and ezdxf.
2. Run the script, passing the path to the STEP file as an argument.
3. The script will automatically create a DXF file in the same directory as the original STEP file.

## solidworks_dxf_gib_R.py

### Purpose

This script is designed to export the flat pattern of a part from SolidWorks to DXF. It uses the win32com library to interact with SolidWorks.

### Functionality

- Interacting with SolidWorks: It connects to an existing SolidWorks instance or creates a new one.
- Opening a File: It opens the specified SLDPRT file.
- Exporting Flat Pattern: It exports the flat pattern of the part to DXF.

### Usage

1. Install SolidWorks and the win32com library.
2. Ensure SolidWorks is running or configured to start automatically.
3. Run the script, specifying the path to the SLDPRT file.
4. The script exports the flat pattern to a DXF file in the same directory.

---

These scripts can be useful for automating tasks related to exporting data from 3D models to DXF for further processing or use in other applications.

