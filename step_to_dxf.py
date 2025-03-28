# -*- coding: utf-8 -*-

from OCC.Core.STEPControl import STEPControl_Reader
from OCC.Core.TopExp import TopExp_Explorer
from OCC.Core.TopAbs import TopAbs_FACE, TopAbs_WIRE, TopAbs_EDGE
from OCC.Core.GProp import GProp_GProps
from OCC.Core.BRepGProp import brepgprop
from OCC.Core.BRep import BRep_Tool
from OCC.Core.TopoDS import topods
from OCC.Core.gp import gp_Pnt, gp_Vec, gp_Trsf, gp_Ax1, gp_Dir
from OCC.Core.GCPnts import GCPnts_UniformAbscissa, GCPnts_AbscissaPoint
from OCC.Core.GeomAdaptor import GeomAdaptor_Curve
from OCC.Core.BRepBuilderAPI import BRepBuilderAPI_MakeFace, BRepBuilderAPI_MakePolygon, BRepBuilderAPI_Transform

import ezdxf
import os

# Scale factor constant (if 1, measurements are in mm)
SCALE_FACTOR = 0.001

# Function to load a STEP file
def load_step_file(file_path):
    reader = STEPControl_Reader()
    reader.ReadFile(file_path)
    reader.TransferRoots()
    return reader.OneShape()

# Function to extract faces from the model
def extract_faces(shape):
    explorer = TopExp_Explorer(shape, TopAbs_FACE)
    faces = []
    while explorer.More():
        faces.append(topods.Face(explorer.Current()))
        explorer.Next()
    return faces

# Function to calculate the area of a face
def calculate_face_area(face):
    props = GProp_GProps()
    brepgprop.SurfaceProperties(face, props)
    return props.Mass() * SCALE_FACTOR**2

# Function to find the largest face by area
def find_largest_face(faces):
    return max(faces, key=calculate_face_area)

# Function to compute the normal vector of a face
def compute_face_normal(face):
    props = GProp_GProps()
    brepgprop.SurfaceProperties(face, props)
    center = props.CentreOfMass()
    normal_vec = gp_Vec(0, 0, 1)
    try:
        u, v = 0.5, 0.5
        brepgprop.Normal(face, u, v, normal_vec)
        normal_vec = normal_vec.Normalized()
    except:
        pass
    return gp_Pnt(center.X() * SCALE_FACTOR, center.Y() * SCALE_FACTOR, center.Z() * SCALE_FACTOR), normal_vec

# Function to compute transformation to the XY plane
def compute_transformation_to_xy(face):
    face_center, face_normal = compute_face_normal(face)
    target_normal = gp_Vec(0, 0, 1)
    if face_normal.IsParallel(target_normal, 1e-6):
        if face_normal.Dot(target_normal) > 0:
            return gp_Trsf()
        else:
            rotation = gp_Trsf()
            rotation.SetRotation(gp_Ax1(gp_Pnt(0, 0, 0), gp_Dir(1, 0, 0)), 3.141592653589793)
            return rotation
    axis = face_normal.Crossed(target_normal)
    axis_dir = gp_Dir(axis)
    angle = face_normal.Angle(target_normal)
    rotation = gp_Trsf()
    rotation.SetRotation(gp_Ax1(gp_Pnt(0, 0, 0), axis_dir), angle)
    return rotation

# Function to discretize curves
def discretize_curve(curve, u1, u2, min_points=5, max_points=100):
    adaptor = GeomAdaptor_Curve(curve, u1, u2)
    # Calculate curve length using GCPnts_AbscissaPoint
    try:
        length = GCPnts_AbscissaPoint.Length(adaptor)
    except:
        # If length calculation fails, return start and end points
        return [(curve.Value(u1).X() * SCALE_FACTOR, curve.Value(u1).Y() * SCALE_FACTOR, curve.Value(u1).Z() * SCALE_FACTOR),
                (curve.Value(u2).X() * SCALE_FACTOR, curve.Value(u2).Y() * SCALE_FACTOR, curve.Value(u2).Z() * SCALE_FACTOR)]
    num_points = max(min(int(length * 10), max_points), min_points)
    try:
        discretizer = GCPnts_UniformAbscissa(adaptor, num_points)
    except:
        return [(curve.Value(u1).X() * SCALE_FACTOR, curve.Value(u1).Y() * SCALE_FACTOR, curve.Value(u1).Z() * SCALE_FACTOR),
                (curve.Value(u2).X() * SCALE_FACTOR, curve.Value(u2).Y() * SCALE_FACTOR, curve.Value(u2).Z() * SCALE_FACTOR)]
    return [(curve.Value(discretizer.Parameter(i+1)).X() * SCALE_FACTOR,
             curve.Value(discretizer.Parameter(i+1)).Y() * SCALE_FACTOR,
             curve.Value(discretizer.Parameter(i+1)).Z() * SCALE_FACTOR) for i in range(discretizer.NbPoints())]

# Function to extract contours without transformation
def extract_contours_from_face(face):
    contours = []
    wire_explorer = TopExp_Explorer(face, TopAbs_WIRE)
    while wire_explorer.More():
        wire = topods.Wire(wire_explorer.Current())
        contour = []
        edge_explorer = TopExp_Explorer(wire, TopAbs_EDGE)
        while edge_explorer.More():
            edge = topods.Edge(edge_explorer.Current())
            curve, u1, u2 = BRep_Tool.Curve(edge)
            # Discretize the curve
            points = discretize_curve(curve, u1, u2)
            contour.append(points)
            edge_explorer.Next()
        contours.append(contour)
        wire_explorer.Next()
    return contours

# Function to project points onto the XY plane
def project_points_to_xy(points):
    return [(x, y) for x, y, z in points]

# Function to project points onto the YZ plane
def project_points_to_yz(points):
    return [(y, z) for x, y, z in points]

# Function to project points onto the XZ plane
def project_points_to_xz(points):
    return [(x, z) for x, y, z in points]

# Function to calculate the total length of lines
def calculate_total_length(contours, projection_func):
    total_length = 0
    for contour in contours:
        for points in contour:
            if len(points) > 1:
                dxf_points = projection_func(points)
                for i in range(1, len(dxf_points)):
                    dx, dy = dxf_points[i][0] - dxf_points[i-1][0], dxf_points[i][1] - dxf_points[i-1][1]
                    total_length += (dx**2 + dy**2)**0.5
    return total_length

# Function to save contours to DXF (2D version)
def create_dxf_from_contours_2d(contours, dxf_file_path, projection_func):
    doc = ezdxf.new('R2010')
    msp = doc.modelspace()
    for contour in contours:
        for points in contour:
            if len(points) > 1:
                # Convert list of tuples to ezdxf format
                dxf_points = projection_func(points)
                msp.add_polyline2d(dxf_points)
    doc.saveas(dxf_file_path)

# Function to process a STEP file
def process_step_file(step_file_path):
    try:
        # Load the STEP file
        shape = load_step_file(step_file_path)
        
        # Extract faces
        faces = extract_faces(shape)
        
        # Determine the largest face
        largest_face = find_largest_face(faces)
        
        print(f"Area of the largest face: {calculate_face_area(largest_face):.2f} cm²")
        
        # Extract contours without transformation
        contours = extract_contours_from_face(largest_face)
        
        # Debug information
        for i, contour in enumerate(contours):
            print(f"Contour {i+1}: {len(contour)} polylines")
            for j, points in enumerate(contour):
                print(f"  Polyline {j+1}: {len(points)} points")
                print(f"  First point: {points[0]}")
                print(f"  Last point: {points[-1]}")
        
        # Calculate line lengths for each projection type
        length_xy = calculate_total_length(contours, project_points_to_xy)
        length_yz = calculate_total_length(contours, project_points_to_yz)
        length_xz = calculate_total_length(contours, project_points_to_xz)
        
        # Choose the projection with the longest length
        max_length = max(length_xy, length_yz, length_xz)
        
        if max_length == length_xy:
            dxf_file = os.path.join(os.path.dirname(step_file_path), f"{os.path.splitext(os.path.basename(step_file_path))[0]}.dxf")
            create_dxf_from_contours_2d(contours, dxf_file, project_points_to_xy)
            print(f"DXF file created: {dxf_file}")
        elif max_length == length_yz:
            dxf_file = os.path.join(os.path.dirname(step_file_path), f"{os.path.splitext(os.path.basename(step_file_path))[0]}.dxf")
            create_dxf_from_contours_2d(contours, dxf_file, project_points_to_yz)
            print(f"DXF file created: {dxf_file}")
        else:
            dxf_file = os.path.join(os.path.dirname(step_file_path), f"{os.path.splitext(os.path.basename(step_file_path))[0]}.dxf")
            create_dxf_from_contours_2d(contours, dxf_file, project_points_to_xz)
            print(f"DXF file created: {dxf_file}")
    
    except Exception as e:
        print(f"Error: {e}")

# Main code
if __name__ == "__main__":
    step_file = r"C:\path\to\your\step_file.step"
    process_step_file(step_file)
