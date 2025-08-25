import ezdxf 
import re
import math
from ezdxf.math import BoundingBox, Matrix44

from operations.geometry import calculer_bbox, get_dimensions
from infra.com_manager import COMManager as CM
from shapely.geometry import Point, LineString
import array



def prepare_inserer_plan(source_dxf, bbox_data, layers):
    doc = ezdxf.readfile(source_dxf)
    # doc.header['$INSBASE'] = (0.0, 0.0, 0.0)
    model = doc.modelspace()

    # THIS IS THE KEY FIX - add version and setup=True
    # new_doc = ezdxf.new('R2013', setup=True)
    # new_model = new_doc.modelspace()

    base_point = (bbox_data[0], bbox_data[1])
    x, y, z = doc.header.get('$INSBASE', (0.0, 0.0, 0.0))    # entities_copied = 0

    print(f'INBASE: {x}, {y}, {z}')
    allowed_layers = [layer.lower() for layer in layers]
    
    # Copy layers from original document (important!)
    # for layer in doc.layers:
        # if layer.dxf.name not in new_doc.layers:
            # try:
                # new_doc.layers.add(
                    # name=layer.dxf.name,
                    # color=layer.dxf.color,
                    # linetype=getattr(layer.dxf, 'linetype', 'CONTINUOUS')
                # )
            # except:
                # If layer copy fails, create with default properties
                # new_doc.layers.add(layer.dxf.name)

    # for entity in model:
        # try:
            # Check if entity layer exists in new document
            # if hasattr(entity, 'dxf') and hasattr(entity.dxf, 'layer'):
                # if entity.dxf.layer not in new_doc.layers:
                    # new_doc.layers.add(entity.dxf.layer)
            
            # new_entity = entity.copy()
            # new_model.ad/d_entity(new_entity)
            # entities_copied += 1
            
        # except Exception as e:
            # print(f"⚠️ Skipped entity {entity.dxftype()}: {e}")
            # continue
    
    # Save to temporary DXF
    temp_dxf = source_dxf.replace('.dxf', '_temp_plan.dxf')
    # new_doc.saveas(temp_dxf)
    
    result = {
        'temp_dxf_path': temp_dxf,
        'base_point': base_point,
        'inbase': (x, y, z),
        'allowed_layers': allowed_layers,
        # 'entities_count': entities_copied
    }
    
    return {'success': True, 'data': result, 'message': 'plan data prepared successfully'}

def prepare_creer_frame_a4(bbox_data, echelle):
    print(f'🪜 echelle before processing:{echelle} ')

    scale_factor = int(echelle.split(":")[1].split(" ")[0])
    print(f"🖼️ Scale factor: {scale_factor}")
    frame_width = (210 * scale_factor) / 1000  # Convert mm to meters
    frame_height = (297 * scale_factor) / 1000
    
    center_x = (bbox_data[3] + bbox_data[0]) / 2
    center_y = (bbox_data[7] + bbox_data[1]) / 2
    print (f"🖼️ Frame center: ({center_x:.2f}, {center_y:.2f})")
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

    y = (center_y + max_y) / 2
    x1 = (min_x + center_x) / 2
    leg_x = (((min_x + x1) / 2) + min_x) / 2
    leg_y = (y + max_y) / 2

    x2 = (center_x + max_x) / 2
    # table_x = (((x2 + max_x) / 2) + max_x) / 2
    table_x = center_x  * (5/4)
    # table_x = (x2 / 2) + max_x / 2
    table_y = (((y + max_y) / 2) + max_y) / 2

    nord_x = center_x
    # nord_y = 
    nord_y =  max_y - (max_y / 8 )

    
    result = {
            'frame_coords': frame_points.tolist(),
            'frame_width': frame_width,
            'frame_height': frame_height,
            'nord_coords':[float(nord_x), float(nord_y)] , 
            'leg_coords': [float(leg_x), float(leg_y)],
            'table_coords': [float(table_x), float(table_y)]
        }
    
    return {'success': True, 'data': result, 'message': 'Frame created, legend and table insert points calculated successfuly'}
    
def prepare_inserer_legende(source_dxf, frame_width):
    """Calculate bbox, dimensions, scale factor, and create scaled legend DXF"""
    
    # Get bbox (you already have this)
    bbox_result = calculer_bbox(source_dxf)
    if not bbox_result['success']:
        return bbox_result
    
    bbox_data = bbox_result['data']
    
    # Calculate dimensions directly from bbox (no separate function needed)
    legende_width = bbox_data[3] - bbox_data[0]   # max_x - min_x
    legende_height = bbox_data[7] - bbox_data[1]  # max_y - min_y
    
    if legende_width <= 0:
        print("⚠️ Invalid legend width")
        return {'success': False, 'user_message': 'Failed to calculate legend dimensions'}
    
    # Calculate scale factor
    target_width = frame_width * 0.25

    print(f'frame_width: {frame_width}, target_width: {target_width}')
    scale_factor = target_width / legende_width

    doc = ezdxf.readfile(source_dxf)
    x, y, z = doc.header.get('$INSBASE', (0.0, 0.0, 0.0))    # entities_copied = 0

    print('scale factor: ', scale_factor)

    
    return {
        'success': True,
        'data': {
            'temp_dxf_path': "temp_scaled_dxf",
            'original_bbox': bbox_data,
            'original_width': legende_width,
            'original_height': legende_height,
            'scale_factor': scale_factor,
            'inbase': (x, y, z)
        },
        'message': f'Legend data prepared'
    }

def prepare_inserer_tableau(source_dxf, frame_width, bbox_data, table_data):
    """Analyze DXF for polylines and labels - pure Python, super fast"""
    
    MAX_VERTEX_SNAP_DIST = 7.0
    
    try:
        # Load DXF in memory (fast)
        doc = ezdxf.readfile(source_dxf)
        msp = doc.modelspace()
        
        print("🔍 Analyzing DXF for polylines and labels...")
        
        # --- Find largest polyline (much faster in DXF) ---
        max_area = 0
        outer_polyline_coords = None
        
        for entity in msp:
            if entity.dxftype() == 'LWPOLYLINE':
                try:
                    # Get coordinates from DXF polyline
                    coords = list(entity.get_points('xy'))
                    
                    # Calculate area using Shapely (faster than COM)
                    if len(coords) >= 3:
                        poly = LineString(coords)
                        area = abs(poly.length)  # or calculate actual area if closed
                        
                        if area > max_area:
                            max_area = area
                            outer_polyline_coords = coords
                            
                except Exception as e:
                    continue
        
        if not outer_polyline_coords:
            return {'success': False, 'user_message': 'Could not find the outer polyline'}
        
        print(f"✅ Found polyline with {len(outer_polyline_coords)} vertices")
        
        # --- Find borne labels (faster DXF iteration) ---
        borne_labels = []
        
        for entity in msp:
            try:
                if entity.dxftype() in ['TEXT', 'MTEXT']:
                    text = entity.dxf.text.strip()
                    if re.match(r'^B\d+$', text):
                        # Get insertion point from DXF
                        if entity.dxftype() == 'TEXT':
                            insertion_pt = (entity.dxf.insert.x, entity.dxf.insert.y)
                        else:  # MTEXT
                            insertion_pt = (entity.dxf.insert.x, entity.dxf.insert.y)
                        
                        borne_labels.append({
                            'Borne': text,
                            'insertion': insertion_pt
                        })
                        print(f"✅ Found label: {text} at {insertion_pt}")
            except:
                continue
        
        if not borne_labels:
            return {'success': False, 'user_message': 'Could not find bornes labels'}
        
        # --- Geometry calculations (pure Python - very fast) ---
        line = LineString(outer_polyline_coords)
        borne_results = []
        
        for label in borne_labels:
            print(f'🎟️ Processing label: {label["Borne"]}')
            
            label_pt = Point(label['insertion'])
            
            # Closest vertex calculation
            closest_vertex = min(outer_polyline_coords, 
                               key=lambda c: Point(c).distance(label_pt))
            closest_vertex_pt = Point(closest_vertex)
            
            dist_to_vertex = label_pt.distance(closest_vertex_pt)
            
            if dist_to_vertex <= MAX_VERTEX_SNAP_DIST:
                chosen_point = closest_vertex_pt
                print(f'🥇 Chosen point: vertex {chosen_point}')
            else:
                # Perpendicular projection
                proj_distance = line.project(label_pt)
                proj_point = line.interpolate(proj_distance)
                chosen_point = proj_point
                print(f'🥇 Chosen point: projection {chosen_point}')
            
            borne_results.append({
                'Borne': label['Borne'],
                'coords': (chosen_point.x, chosen_point.y)
            })
        
        # Sort by number in label
        borne_results.sort(key=lambda b: int(re.findall(r'\d+', b['Borne'])[0]))
        

        print("📑 Final borne coordinates:")
        for b in borne_results:
            print(f"{b['Borne']}: {b['coords']}")
        
        # --- Calculate table parameters ---
        n_rows = len(borne_results) + 2  # +2 for header
        n_cols = 3
        row_height = 8.0
        col_width = 50.0
        
        table_width = col_width * n_cols
        target_width = frame_width * 0.3
        scale_factor = target_width / table_width
        
        return {
            'success': True,
            'data': {
                'borne_results': borne_results,
                'table_coords': table_data,
                'n_rows': n_rows,
                'n_cols': n_cols,
                'row_height': row_height,
                'col_width': col_width,
                'scale_factor': scale_factor,
                'polyline_coords': outer_polyline_coords
            },
            'message': f'Table data prepared - {len(borne_results)} bornes processed'
        }
        
    except Exception as e:
        print(f"❌ Table preparation failed: {e}")
        return {'success': False, 'error': str(e)}

def prepare_inserer_croisillions(bornes_results, frame_width, frame_coords, echelle):
    # Extents from bornes
    min_x = float('inf'); min_y = float('inf')
    max_x = float('-inf'); max_y = float('-inf')
    echelle = int(echelle.split(":")[1].split(" ")[0])

    for borne in bornes_results:
        x, y = borne['coords']
        min_x = min(min_x, x); min_y = min(min_y, y)
        max_x = max(max_x, x); max_y = max(max_y, y)

    step = echelle // 10.0
    
    min_x = math.floor(min_x / 10.0) * 10.0
    min_y = math.floor(min_y / 10.0) * 10.0
    max_x = math.ceil (max_x / 10.0) * 10.0
    max_y = math.ceil (max_y / 10.0) * 10.0
    
    croisillons_coords = []
    x = min_x
    while x <= max_x:
        y = min_y
        while y <= max_y:
            croisillons_coords.append((int(round(x)), int(round(y))))
            y += step
        x += step
    labels = []
        
    
    # Where to place text (at frame edges)
    text_height = frame_width * 0.01

    frame_points = [(frame_coords[i], frame_coords[i+1]) for i in range(0, len(frame_coords), 3)]
    bl, br, tr, tl = frame_points[:4]  # bottom-left, bottom-right, top-right, top-left

    # Create LineStrings for each frame side
    bottom_edge = LineString([bl, br])
    right_edge = LineString([br, tr]) 
    top_edge = LineString([tr, tl])
    left_edge = LineString([tl, bl])

    x_projection = []
    y_projection = []

    x_group = {}
    y_group = {}
    
    for croisillion in croisillons_coords:
        x,y = croisillion
        if x not in x_group:
            x_group[x] = []
        x_group[x].append(croisillion)

        if y not in y_group:
            y_group[y] = []
        y_group[y].append(croisillion)
    
    for y_coords, points in y_group.items():
        for point in points:
            x,y = point
            proj_distance = bottom_edge.project(Point(x, y))
            proj_point = bottom_edge.interpolate(proj_distance)

            x_projection.append({
                'coords': (int(round(x)), int(round(y))),
                'insert_pt': (proj_point.x, proj_point.y),
                'insert_type': 'x_axis'
                })
    
    for x_coords, points in x_group.items():
        for point in points:
            x,y = point
            
            proj_distance = right_edge.project(Point(x, y))
            proj_point = right_edge.interpolate(proj_distance)

            y_projection.append({
                'coords': (int(round(x)), int(round(y))),
                'insert_pt': (proj_point.x, proj_point.y),
                'insert_type': 'y_axis'
                })

    cross_size = frame_width * 0.03
    return {
        'success': True,
        'data': {
            'croisillons_coords': croisillons_coords,
            'frame_width': frame_width, 
            # 'labels_coords': coordinate_labels,
            'x_projection': x_projection,
            'y_projection': y_projection,        # list[{text,pos,rotation}]
            'text_height': text_height,
            'half_size': cross_size / 2.0,
            'tick_size': cross_size / 4.0
        },
        'message': 'Croisillons data prepared'
    }


