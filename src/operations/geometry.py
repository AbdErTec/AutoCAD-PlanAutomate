
from ezdxf import bbox as ezbbox
import win32com.client
import array
import ezdxf
from ezdxf.math import BoundingBox, Vec3

def get_dimensions(bbox_data):
    print(f'bbox in get dimensions: {bbox_data}')
    if not bbox_data or len(bbox_data) < 8:
        return 0, 0
    width = bbox_data[3] - bbox_data[0]
    height = bbox_data[7] - bbox_data[1]
    return width, height

from infra.com_manager import COMManager as CM
def calculer_bbox(file_path):
    doc = ezdxf.readfile(file_path)
    msp = doc.modelspace()

    min_x = float('inf')
    min_y = float('inf') 
    max_x = float('-inf')
    max_y = float('-inf')

    entity_count = 0

    for entity in msp:
        entity_count += 1
        try:
            t = entity.dxftype()
            if t == 'LINE':
                s, e = entity.dxf.start, entity.dxf.end
                min_x = min(min_x, s.x, e.x)
                min_y = min(min_y, s.y, e.y)
                max_x = max(max_x, s.x, e.x)
                max_y = max(max_y, s.y, e.y)
            elif t == 'LWPOLYLINE':
                for x, y in entity.get_points('xy'):
                    min_x = min(min_x, x)
                    min_y = min(min_y, y)
                    max_x = max(max_x, x)
                    max_y = max(max_y, y)
            elif t == 'CIRCLE':
                c, r = entity.dxf.center, entity.dxf.radius
                min_x = min(min_x, c.x - r)
                min_y = min(min_y, c.y - r)
                max_x = max(max_x, c.x + r)
                max_y = max(max_y, c.y + r)
        except Exception as e:
            continue

    if min_x == float('inf'):
        return {'success': False, 'user_message': 'No valid entities found'}
    bbox = [(min_x), min_y, 0, max_x, min_y, 0, max_x, max_y, 0, min_x, max_y, 0, min_x, min_y, 0]
    return {'success': True, 'data': [float(item) for item in bbox], 'message': 'bbox calcule avec success'}

def ent_bbox(entity):
    """
    Axis-aligned bbox ((minx,miny),(maxx,maxy)) in WCS for ANY entity (incl. IMAGE).
    Returns None if bbox can’t be computed.
    """
    min_x = float('inf')
    min_y = float('inf') 
    max_x = float('-inf')
    max_y = float('-inf')
    
    try:

        t= entity.dxftype()
        if t == 'LINE':
            s, e = entity.dxf.start, entity.dxf.end
            min_x = min(min_x, s.x, e.x)
            min_y = min(min_y, s.y, e.y)
            max_x = max(max_x, s.x, e.x)
            max_y = max(max_y, s.y, e.y)
        elif t == 'LWPOLYLINE':
            for x, y in entity.get_points('xy'):
                min_x = min(min_x, x)
                min_y = min(min_y, y)
                max_x = max(max_x, x)
                max_y = max(max_y, y)
        elif t == 'CIRCLE':
            c, r = entity.dxf.center, entity.dxf.radius
            min_x = min(min_x, c.x - r)
            min_y = min(min_y, c.y - r)
            max_x = max(max_x, c.x + r)
            max_y = max(max_y, c.y + r)
        elif t in ("IMAGE", "RASTERIMAGE"):
            ext = ezbbox.extents([entity])
            if ext:
                mn, mx = ext
                min_x = float(mn.x)
                min_y = float(mn.y)
                max_x = float(mx.x)
                max_y = float(mx.y)

        return [(float(min_x), float(min_y)), (float(max_x), float(max_y))]
    except Exception as e:
        print(f'{t} error: {e}')
        return None
    
def bbox_overlap(bb1, bb2, tol=1e-9):
    (ax1, ay1), (ax2, ay2) = bb1
    (bx1, by1), (bx2, by2) = bb2
    return not (ax2 < bx1 - tol or bx2 < ax1 - tol or ay2 < by1 - tol or by2 < ay1 - tol)

