from PyQt5.QtWidgets import (
    QApplication, QWidget, QLabel, QLineEdit, QPushButton,
    QVBoxLayout, QHBoxLayout, QFileDialog, QMessageBox
)
import sys
import threading
import asyncio
import os 
import pythoncom
import win32com.client
from win32com.client import VARIANT
import time
from pyautocad import Autocad, APoint
import array


def lire_fichier(path):
    """Read and validate file"""
    ext = os.path.splitext(path)[1].lower()
    try:
        if ext == '.dwg':
            print("✅ Import success")
            return True
        else:
            print("❌ Extension invalide.")
            return False
    except Exception as e:
        print(f"❌ Erreur lors de la lecture du fichier: {e}")
        return False


def inserer_component(source_path, insertion_point=(0, 0, 0), destination_doc=None):
    """
    Universal component insertion method - works for all cases
    """
    pythoncom.CoInitialize()
    
    try:
        # Get destination document
        if destination_doc is None:
            try:
                app = win32com.client.GetActiveObject("AutoCAD.Application")
                destination_doc = app.ActiveDocument
            except:
                app = win32com.client.Dispatch("AutoCAD.Application")
                destination_doc = app.ActiveDocument
        else:
            app = destination_doc.Application
        
        destination_doc.Activate()
        print(f"📄 Destination document: {destination_doc.Name}")
        
        try:

            x, y, z = insertion_point
            print("🔄 copie du component...")
            source_doc = app.Documents.Open(source_path)
            source_doc.Activate()
            time.sleep(1)
            
            # Select all and copy
            source_doc.SendCommand('_AI_SELALL\n')
            time.sleep(1)
            source_doc.SendCommand('_COPYCLIP\n')
            time.sleep(1)
            
            # Switch back and paste
            destination_doc.Activate()
            time.sleep(1)
            destination_doc.SendCommand(f'_PASTECLIP\n{x},{y}\n')
            time.sleep(2)
            
            # Close source
            source_doc.Close(False)
            destination_doc.SendCommand('_ZOOM\nE\n')
            print("✅ Copy-paste method successful")

            return True
            
        except Exception as e:
            print(f"❌ copie du component failed: {e}")
            return False
    
    except Exception as e:
        print(f"💥 Component insertion failed: {e}")
        return False
    
    finally:
        try:
            pythoncom.CoUninitialize()
        except:
            pass


def calculer_bbox(file_path):
    """Calculate bounding box of all entities"""
    # pythoncom.CoInitialize()
    
    try:
        app = win32com.client.GetActiveObject("AutoCAD.Application")
        doc = app.Documents.Open(file_path)
        doc.Activate()
        time.sleep(1)


        acad = Autocad(create_if_not_exists = False)
        main_doc = acad.doc
        model = main_doc.ModelSpace

        min_x = float('inf')
        min_y = float('inf')
        max_x = float('-inf')
        max_y = float('-inf')

        entity_count = 0
        for entity in model:
            obj_type = entity.ObjectName

            if obj_type == "AcDbLine":
                x1, y1 = entity.StartPoint[:2]
                x2, y2 = entity.EndPoint[:2]

                min_x = min(min_x, x1, x2)
                min_y = min(min_y, y1, y2)
                max_x = max(max_x, x1, x2)
                max_y = max(max_y, y1, y2)

            elif obj_type == "AcDbPolyline":
                try:
                    coords = entity.Coordinates
                    num_vertices = len(coords) // 2  # Divide by 2 for 2D points
                    print(f"Nombre de sommets : {num_vertices}")
                    
                    for i in range(num_vertices):
                        x = coords[i * 2]
                        y = coords[i * 2 + 1]
                        
                        min_x = min(min_x, x)
                        min_y = min(min_y, y)
                        max_x = max(max_x, x)
                        max_y = max(max_y, y)
                        
                except Exception as e:
                        print(f"Échec méthode de fallback: {e}")

            elif obj_type == "AcDbCircle":
                center = entity.Center
                radius = entity.Radius
                x, y = center[:2]
                min_x = min(min_x, x - radius)
                min_y = min(min_y, y - radius)
                max_x = max(max_x, x + radius)
                max_y = max(max_y, y + radius)

            elif obj_type == "AcDbBlockReference":
                ins_point = entity.InsertionPoint
                x, y = ins_point[:2]
                # ici on ne gère pas la rotation ni les entités internes
                min_x = min(min_x, x)
                min_y = min(min_y, y)
                max_x = max(max_x, x)
                max_y = max(max_y, y)

        if min_x == float('inf'):
            print("Aucune entité traitée.")
        else:
            print(f"BBOX globale: Min ({min_x:.2f}, {min_y:.2f}) / Max ({max_x:.2f}, {max_y:.2f})")

            # try:

            import array

            points = array.array('d', [
                min_x, min_y, 0.0,
                max_x, min_y, 0.0,
                max_x, max_y, 0.0,
                min_x, max_y, 0.0,
                min_x, min_y, 0.0
            ])

            doc.Close(False)
            return points
                
    except Exception as e:
        print(f"Erreur lors du calcul de la BBOX: {e}")
        return 

def get_dimensions(bbox_data):
    width = bbox_data[3] - bbox_data[0]
    height = bbox_data[7] - bbox_data[1]
    return width, height


def inserer_plan(source_path, destination_doc=None):
    """Insert main plan at origin"""
    print(f"📋 Inserting plan: {os.path.basename(source_path)}")
    inserer_component(source_path, (0, 0, 0), destination_doc)
    print("📌 Calcul du bounding box...")


def inserer_legende(source_path, bbox_data, destination_doc=None):
    """Insert legend based on bounding box"""
    if not bbox_data:
        print("❌ No bbox data for legend insertion")
        return False
    pythoncom.CoInitialize()
    try:
        if destination_doc is None:
            app = win32com.client.GetActiveObject("AutoCAD.Application")
            destination_doc = app.ActiveDocument
            print(f'📄 Active Document: {app.ActiveDocument}')
            model = destination_doc.ModelSpace
        time.sleep(3)
        destination_doc.Activate()
        

        leg_x = bbox_data[0]
        leg_y = bbox_data[7] + (bbox_data[7] + bbox_data[1]) 
        insert_pt = (leg_x, leg_y, 0.0)
        
        print (f'Selection des objets avant l insertion de la legende...')
        before = [obj for obj in model]

        print(f"🏷️ Inserting legend at: ({leg_x:.2f}, {leg_y:.2f})")
        inserer_component(source_path, insert_pt, destination_doc)
        time.sleep(2)

        after = [obj for obj in model]
        legend_objects = [obj for obj in after if obj not in before]

        coords_before = []
        for obj in legend_objects:
            bounds = obj.GetBoundingBox()
            center_x = (bounds[0][0] + bounds[1][0]) / 2
            center_y = (bounds[0][1] + bounds[1][1]) / 2
            coords_before.append({"center_x" : center_x, "center_y": center_y})
        
        print (legend_objects)
        print(f'coords before scaling: {coords_before}')

        time.sleep(2)
        
        #scale legende
        bbox_data_legende = calculer_bbox(source_path)
        time.sleep(1)
            # base_pt = APoint(center_x, center_y)

        plan_width, plan_height = get_dimensions(bbox_data)
        legend_width, legend_height = get_dimensions(bbox_data_legende)
        target_width = plan_width * 0.5
        scale_factor = target_width / legend_width
        legend_center_x = legend_width / 2
        legend_center_y = legend_height / 2

        print(f'Scale factor: {scale_factor}, legende center({legend_center_x}, {legend_center_y})')

        # scale_base = APoint(legend_center_x, legend_center_y)
        time.sleep(2)
        target_base = APoint(leg_x, leg_y)
        for i, obj in enumerate(legend_objects):
            try:
                handle = obj.Handle
                print(f"  📍 Scaling object {i+1}/{len(legend_objects)} (Handle: {handle})")
                
                # Get object's center point for scaling
                try:
                    if hasattr(obj, 'InsertionPoint'):
                        base_pt = APoint(obj.InsertionPoint)
                        # center_x, center_y = obj.InsertionPoint[0], obj.InsertionPoint[1]
                    elif hasattr(obj, 'Center'):
                        # center_x, center_y = obj.Center[0], obj.Center[1]
                        base_pt = APoint(obj.Center)

                    else:
                        # Use bounding box center
                        bounds = obj.GetBoundingBox()
                        center_x = (bounds[0][0] + bounds[1][0]) / 2
                        center_y = (bounds[0][1] + bounds[1][1]) / 2
                        base_pt = APoint(center_x, center_y)
                    
                except:
                    # Fallback to legend center
                    center_x, center_y = legend_center_x, legend_center_y
                    base_pt = APoint(legend_center_x, legend_center_y)
                
                print(f'Center ({center_x},{center_y})')

                # Select and scale from object's own center
                destination_doc.SendCommand(f'_SELECT\n(handent "{handle}")\n\n')
                time.sleep(0.5)
                
                destination_doc.SendCommand(f'_SCALE\n\n{leg_x},{leg_y}\n{scale_factor}\n')
                time.sleep(1)
                try:
                    destination_doc.SendCommand(
                        f'(command "_MOVE" (handent "{handle}") "" "{center_x},{center_y}" "{coords_before[i]["center_x"]},{coords_before[i]["center_y"]}")\n'
                    )

                    time.sleep(1)
                except Exception as e:
                    print(f'Move unseccessful: {e}')
                    
                print(f"  ✅ Object {i+1} scaled from its center")
            
            except Exception as e:
                print(f"❌ Failed to scale object {i+1}: {e}")
            destination_doc.SendCommand('_ZOOM\nE\n')
    except Exception as e:
        print(f"💥 Legend insertion error: {e}")
        return False
    finally:
        pythoncom.CoUninitialize()


def inserer_tableau(source_file, bbox_data, destination_doc=None):
# Pick a very obvious point for testing
    pt = APoint(100, 100)

    # ⚠️ Use clearly valid dimensions
    rows = 4
    cols = 3
    row_height = 10
    col_width = 30
    pythoncom.CoInitialize()
    try:
        app =win32com.client.GetActiveObject("AutoCAD.Application")
        destination_doc = app.ActiveDocument
        model = destination_doc.ModelSpace
        destination_doc.Activate()
        
        try:
        #1. Open the original file 
            source_doc = app.Documents.Open(source_file)
            source_doc.Activate()
            time.sleep(1)

            acad_src = Autocad(create_if_not_exists = False)
            # source_doc = acad_src.doc
            source_ms = acad_src.doc.ModelSpace
        #2. Find the polyline with the most area 
            max_area = 0
            outer_polyline = None
            coords = []

            for obj in source_ms:
                try:
                    if obj.ObjectName == 'AcDbPolyline':
                        area = obj.Area  # This throws if not closed
                        if area > max_area:
                            max_area = area
                            outer_polyline = obj
                except Exception as e:
                    print(f"Skipping object: {e}")

            if outer_polyline:
                coords = outer_polyline.Coordinates
                print(f"Bornes: {coords}")
            else:
                print("No closed polyline with area found.")

            source_doc.Close(False)
            time.sleep(2)
            destination_doc.Activate()
            time.sleep(1)

            for i in range(0, len(coords), 2):
                x = coords[i]
                y = coords[i + 1]
                destination_doc.SendCommand(f'_CIRCLE\n{x},{y}\n2\n')
            
            # return coords if outer_polyline else None
        except Exception as e:
            print(f"Erreur lors de la recuperation des bornes: {e}")
            return None

        time.sleep(2)
        n_rows = (len(coords)//2) +1
        n_cols = 2
        row_height = 3.0
        col_width = 20.0
        table_x = bbox_data[3]
        table_y = bbox_data[6] + (bbox_data[6] - bbox_data[4]) 
        insert_pt = APoint(table_x, table_y)
        table = model.AddTable(insert_pt, n_rows, n_cols, row_height, col_width)

        # Set headers
        table.SetText(0, 0, "X")
        table.SetText(0, 1, "Y")

        # Fill coordinates
        point_index = 0
        for r in range(1, n_rows):
            if point_index * 2 + 1 >= len(coords):
                break
            x = coords[point_index * 2]
            y = coords[point_index * 2 + 1]
            table.SetText(r, 0, f"{x:.2f}")
            table.SetText(r, 1, f"{y:.2f}")
            point_index += 1

        print("✅ Table inserted successfully.")
    except Exception as e:
        print(f"💥 Table test error: {e}")
    finally:
        pythoncom.CoUninitialize()


#     pythoncom.CoInitialize()
#     try:
#         if destination_doc is None:
#             app = win32com.client.GetActiveObject("AutoCAD.Application")
#             destination_doc = app.ActiveDocument
#             print(f'📄 Active Document: {app.ActiveDocument}')
#             model = destination_doc.ModelSpace
#         time.sleep(3)
#         destination_doc.Activate()
        
#         table_x = bbox_data[3]
#         table_y = bbox_data[7] + (bbox_data[7] + bbox_data[1]) 
#         insert_pt = APoint(table_x, table_y)

#         #fetch rows number + data from original file
#         print(f"tableau coordonnees: {insert_pt}")




#         # destination_doc.SendCommand(f'_CIRCLE\n{table_x},{table_y}\n4\n')
#         table = model.AddTable(insert_pt, n_rows, n_cols, row_height, col_width)
#         table.TitleSuppressed = False
#         table.SetText(0, 0, "Tableau des coordonnées")

#         # Fill the cells
#         for row in range(1, 4):  # row 0 is usually title/header
#             for col in range(3):
#                 cell_text = f"Row{row} Col{col}"
#                 table.SetText(row, col, "blabla")

# #         # Position table at right side of bbox
# #         table_x = bbox_data['max_x'] + bbox_data['width'] * 0.1
# #         table_y = bbox_data['center_y']
        
# #         print(f"📊 Coordinate table position: ({table_x:.2f}, {table_y:.2f})")
# #         # Here you would insert your coordinate table
# #         # return inserer_component(table_source_path, (table_x, table_y, 0), destination_doc)
# #         return True
        
#     except Exception as e:
#         print(f"💥 Table insertion error: {e}")
#         return False
#     finally:
#         pythoncom.CoUninitialize()


# def inserer_boussole(bbox_data, destination_doc=None):
#     """Insert compass at center bottom of plan"""
#     if not bbox_data:
#         print("❌ No bbox data for compass insertion")
#         return False
    
#     try:
#         # Position compass at center, below plan
#         compass_x = bbox_data['center_x']
#         compass_y = bbox_data['min_y'] - abs(bbox_data['height']) * 0.3
        
#         print(f"🧭 Compass position: ({compass_x:.2f}, {compass_y:.2f})")
#         # return inserer_component(compass_source_path, (compass_x, compass_y, 0), destination_doc)
#         return True
        
#     except Exception as e:
#         print(f"💥 Compass insertion error: {e}")
#         return False


def creer_frame_a4(bbox_data, destination_doc=None):
    """Create A4 frame around all content"""
    if not bbox_data:
        print("❌ No bbox data for frame creation")
        return False
        
    try:
        acad = Autocad()
        model = acad.doc.ModelSpace
        
        # Frame size: 3 times the content size
        frame_width = 3 * bbox_data['width']
        frame_height = 3 * bbox_data['height']
        
        # Center frame around content
        center_x = bbox_data['center_x']
        center_y = bbox_data['center_y']
         
        min_x = center_x - frame_width / 2
        min_y = center_y - frame_height / 2
        max_x = center_x + frame_width / 2
        max_y = center_y + frame_height / 2
        
        frame_points = array.array('d', [
            min_x, min_y, 0.0,
            max_x, min_y, 0.0,
            max_x, max_y, 0.0,
            min_x, max_y, 0.0,
            min_x, min_y, 0.0
        ])
        
        frame = model.AddPolyline(frame_points)
        frame.Closed = True
        frame.color = 7  # White
        
        print(f"🖼️ A4 Frame created: {frame_width:.2f} x {frame_height:.2f}")
        return True
        
    except Exception as e:
        print(f"💥 Frame creation error: {e}")
        return False


def process_complete_workflow(plan_path, legend_path, destination_doc=None):
    """Complete workflow: plan + bbox + legend + frame"""
    try:
        print("🚀 Starting complete workflow...")
        
        # 1. Insert main plan
        if not inserer_plan(plan_path, destination_doc):
            print("❌ Plan insertion failed")
            return False
        
        # 2. Calculate bounding box
        bbox_data = calculer_bbox(destination_doc)
        if not bbox_data:
            print("❌ Bbox calculation failed")
            return False
        
        # 3. Insert legend
        if legend_path and not inserer_legende(legend_path, bbox_data, destination_doc):
            print("⚠️ Legend insertion failed")
        
        # 4. Insert other components (uncomment as needed)
        # inserer_tableau(bbox_data, destination_doc)
        # inserer_boussole(bbox_data, destination_doc)
        
        # 5. Create frame
        if not creer_frame_a4(bbox_data, destination_doc):
            print("⚠️ Frame creation failed")
        
        print("✅ Complete workflow finished!")
        return True
        
    except Exception as e:
        print(f"💥 Workflow error: {e}")
        return False


# Usage example:
# result = process_complete_workflow(
#     plan_path="C:/path/to/plan.dwg", 
#     legend_path="C:/path/to/legend.dwg"
# )