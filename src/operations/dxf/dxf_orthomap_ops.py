import ezdxf
from ezdxf.math import BoundingBox2d
from operations.geometry import get_dimensions, bbox_overlap, ent_bbox

def preparer_crop_ortophoto_map(source_dxf):
    doc = ezdxf.readfile(source_dxf)
    ms = doc.modelspace()

    images_data = []
    for entity in ms:
        print(f"Processing: {entity.dxftype()}")
        if entity.dxftype() == "IMAGE":
            layer = (entity.dxf.layer or '').lower()
            bbox = ent_bbox(entity)
            
            # Debug: Print image bbox
            print(f"Image bbox: {bbox}")
            
            if layer in ['orto', 'ortophoto', 'photo', 'aperçu sur fond haut']:
                images_data.append({
                    'image':{
                        'name':entity.dxf.layer,
                        'Type': 'ortophoto',
                        'entity': entity,
                        'handle': entity.dxf.handle,
                        'bbox': bbox,
                        'limites_data': None,
                        'clipping_data': None,
                        'done': None
                        }
                })
                print(f"📸 Found ortophoto: {entity.dxf.handle}")
            elif layer in ['carte', 'map', 'aperçu sur fond bas']:
                images_data.append({
                    'image':{
                        'name':entity.dxf.layer,
                        'Type': 'map',
                        'entity': entity,
                        'handle': entity.dxf.handle,
                        'bbox': bbox,
                        'limites_data': None,
                        'clipping_data': None,
                        'done': None
                    }
                })
                print(f"🗺️ Found carte: {entity.dxf.handle}")

    if not images_data:
        return {'success': False, 'user_message': 'No images found in source document'}

    # Collect all potential boundary entities first
    boundary_entities = []
    for ent in ms:
        if ent.dxftype() in ['LWPOLYLINE', 'POLYLINE', 'CIRCLE']:
            try:
                eb = ent_bbox(ent)
                if eb:  # Only add if bbox calculation succeeds
                    boundary_entities.append((ent, eb))
                    print(f"Found boundary entity: {ent.dxftype()} with bbox: {eb}")
            except Exception as e:
                print(f"Error calculating bbox for {ent.dxftype()}: {e}")
                continue

    print(f"Total boundary entities found: {len(boundary_entities)}")

    for img in images_data:
        img_data = img['image']
        bbox_image = img_data['bbox']
        print(f"\nProcessing image {img_data['handle']} with bbox: {bbox_image}")
        
        if not bbox_image:
            print("⚠️ Image has no valid bbox. Skipping…")
            continue
            
        selected = []
        
        # More lenient overlap detection
        for ent, eb in boundary_entities:
            try:
                # Check if bounding boxes overlap or are close
                if bbox_overlap(bbox_image, eb):
                    selected.append(ent)
                    print(f"✅ Entity {ent.dxftype()} overlaps with image")
                else:
                    # Also check if entity is contained within a larger area around the image
                    img_min, img_max = bbox_image
                    ent_min, ent_max = eb
                    
                    # Expand image bbox by 50% to catch nearby entities
                    img_w = img_max[0] - img_min[0]
                    img_h = img_max[1] - img_min[1]
                    expanded_min = (img_min[0] - img_w * 0.5, img_min[1] - img_h * 0.5)
                    expanded_max = (img_max[0] + img_w * 0.5, img_max[1] + img_h * 0.5)
                    expanded_bbox = (expanded_min, expanded_max)
                    
                    if bbox_overlap(expanded_bbox, eb):
                        selected.append(ent)
                        print(f"✅ Entity {ent.dxftype()} overlaps with expanded image area")
                        
            except Exception as e:
                print(f"Error checking overlap: {e}")
                continue

        print(f"Selected entities for image {img_data['handle']}: {len(selected)}")

        if not selected:
            print("⚠️ No overlapping entities for image. Skipping…")
            continue

        # Find largest area entity = parcel boundary
        max_area = 0.0
        limites = None
        
        for ent in selected:
            try:
                area = 0.0
                if ent.dxftype() == 'CIRCLE':
                    area = 3.14159 * (ent.dxf.radius ** 2)
                elif hasattr(ent, 'area'):
                    area = float(abs(ent.area))  # Use abs() to handle negative areas
                elif ent.dxftype() in ['LWPOLYLINE', 'POLYLINE']:
                    # Fallback: calculate area from bounding box
                    eb = ent_bbox(ent)
                    if eb:
                        min_pt, max_pt = eb
                        area = (max_pt[0] - min_pt[0]) * (max_pt[1] - min_pt[1])
                
                print(f"Entity {ent.dxftype()} area: {area}")
                
                if area > max_area:
                    max_area = area
                    limites = ent
                    
            except Exception as e:
                print(f"Error calculating area for {ent.dxftype()}: {e}")
                continue

        if not limites:
            print("❌ Could not find parcel boundary for an image")
            continue

        print(f"Selected parcel boundary: {limites.dxftype()} with area: {max_area}")

        try:
            lim_bb = ent_bbox(limites)
            if not lim_bb:
                print("❌ Could not get bbox for parcel boundary")
                continue
                
            min_pt, max_pt = lim_bb
            w = max_pt[0] - min_pt[0]
            h = max_pt[1] - min_pt[1]
            
            # Validate dimensions
            if w <= 0 or h <= 0:
                print(f"❌ Invalid parcel dimensions: w={w}, h={h}")
                continue
                
            # Reduce buffer to be more conservative
            buf_x = w * 0.3  # Reduced from 0.6
            buf_y = h * 0.3  # Reduced from 0.6

            clip_min_x = min_pt[0] - buf_x
            clip_min_y = min_pt[1] - buf_y
            clip_max_x = max_pt[0] + buf_x
            clip_max_y = max_pt[1] + buf_y

            clipping_points = [
                (clip_min_x, clip_min_y),
                (clip_max_x, clip_min_y),
                (clip_max_x, clip_max_y),
                (clip_min_x, clip_max_y),
            ]

            img_data['limites_data'] = {
                'limites_parcel_handle': limites.dxf.handle,
                'limites_parcel_bbox': lim_bb,
                'limites_parcels_origin': min_pt
            }
            img_data['clipping_data'] = {
                'clipping_points': clipping_points,
                'insertion_width': float(w),
                'insertion_height': float(h),
                'base_pt': (float(min_pt[0]), float(min_pt[1]))
            }
            
            print(f"✅ Successfully processed image {img_data['handle']}")
            print(f"   Parcel dimensions: {w:.2f} x {h:.2f}")
            print(f"   Clipping area: ({clip_min_x:.2f}, {clip_min_y:.2f}) to ({clip_max_x:.2f}, {clip_max_y:.2f})")
            
        except Exception as e:
            print(f"❌ Error processing parcel boundary: {e}")
            continue

    # Remove non-serializable objects before returning
    processed_images = 0
    for img in images_data:
        img_data = img.get('image', {})
        img_data.pop('entity', None)
        if img_data.get('clipping_data'):
            processed_images += 1

    print(f"\n📊 Summary: {processed_images}/{len(images_data)} images processed successfully")

    return {
        'success': True, 
        'data': images_data, 
        'message': f'Crop preparation completed successfully. Processed {processed_images}/{len(images_data)} images.'
    }