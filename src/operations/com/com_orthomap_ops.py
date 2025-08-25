from win32com.client import VARIANT
import pythoncom
import os
import win32com.client as win32com
import array
from infra.com_manager import COMManager as CM
from operations.com.com_block_ops import detect_placeholders
from operations.geometry import ent_bbox
import os, time, uuid
from pathlib import Path
from infra.com_manager import COMManager as CM
from operations.com.com_block_ops import detect_placeholders
import uuid


def acad_num(x, ndp=12):
    s = f"{float(x):.{ndp}f}".rstrip('0').rstrip('.')
    return s if s else "0"

def crop_orthophoto_map(source_dwg, prepared_data, block_ref, destination_doc=None):
    """
    Fast path per image:
      1) SOURCE: put image in pickfirst (by handle) → IMAGECLIP (Rect) with 2 pts
      2) SOURCE: pickfirst = image (+ limites) → COPYBASE @ base_pt
      3) DEST:   PASTECLIP @ placeholder_origin → SCALE Last about origin by uniform s
    """
    try:
        # === App + destination (sheet) ===
        if destination_doc is None:
            app, destination_doc = CM.get_app_and_doc()
            if not destination_doc:
                return {'success': False, 'user_message': 'No destination document'}
        else:
            app = CM.retry_com_operation(lambda: destination_doc.Application)

        CM.retry_com_operation(destination_doc.Activate)

        # === Detect placeholders once ===
        ph_res = detect_placeholders(block_ref, destination_doc=destination_doc)
        CM.check_success(ph_res, 'detect placeholders')
        placeholders = ph_res['placeholders']
        print(f"Detected {len(placeholders)} placeholders\n{placeholders}")

        def _norm(s): return (s or "").lower()
        def pick_placeholder(avail, img_type):
            target = 'orto_placeholder' if img_type == 'ortophoto' else 'map_placeholder'
            # exact or substring
            for i, ph in enumerate(avail):
                n = _norm(ph['name'])
                if n == target or target in n: return i, ph
            # tolerant tokens
            key = 'orto' if img_type == 'ortophoto' else 'map'
            for i, ph in enumerate(avail):
                if key in _norm(ph['name']): return i, ph
            return None, None

        # === Open source (images live here) ===
        src_doc = app.Documents.Open(source_dwg, False)
        try:
            used = set()

            for item in prepared_data:
                print(f"📂 Processing item: {item}")
                
                # Extract image data properly
                if 'image' in item:
                    img_data = item['image']
                else:
                    img_data = item  # Sometimes the structure is flatter
                
                t = _norm(img_data.get('Type', ''))
                print(f"--📂 Image type: {t}")
               
                if t not in ('ortophoto', 'carte', 'map'):
                    print(f"⚠️ Unknown image type: {t}; skip.")
                    continue

                try:
                    clip = img_data['clipping_data'] 
                    pts = clip['clipping_points'] 

                    if len(pts) < 4:
                        print(f"⚠️ Need 4 points for rectangular clip; skip {t}.")
                        continue

                    # Get opposite corners for rectangular clip
                    pt1 = (float(pts[0][0]), float(pts[0][1]))  # Bottom-left
                    pt2 = (float(pts[2][0]), float(pts[2][1]))  # Top-right
                    print(f"--📂 Clipping points: {pt1} -> {pt2}")

                    w = float(clip.get('insertion_width', 0))
                    h = float(clip.get('insertion_height', 0))
                    base_pt = clip.get('base_pt') or pts[0]
                    print(f"--📂 Insertion dimensions: {w} x {h} | Base point: {base_pt}")
                    
                    if w <= 0 or h <= 0:
                        print(f"⚠️ Invalid insertion dimensions for {t}; skip.")
                        continue
                
                except Exception as e:
                    print(f"⚠️ Invalid clipping data for {t}: {e}")
                    continue
                
                # Match a free placeholder
                img_type_for_placeholder = 'ortophoto' if t == 'ortophoto' else 'map'
                avail = [p for i, p in enumerate(placeholders) if i not in used]
                idx, ph = pick_placeholder(avail, img_type_for_placeholder)
                if ph is None:
                    print(f"❌ No placeholder for {t}; skip.")
                    continue
                    
                real_idx = next(i for i, p in enumerate(placeholders) if p is ph and i not in used)
                used.add(real_idx)

                ph_w = float(ph['width'])
                ph_h = float(ph['height']) 
                ph_org = ph['origin_pt']  # This should be the insertion point of placeholder
                
                # uniform scale to preserve photo aspect
                s = min(ph_w / w, ph_h / h)
                ip = f'{acad_num(ph_org[0])},{acad_num(ph_org[1])}'
                bp = f'{acad_num(base_pt[0])},{acad_num(base_pt[1])}'

                h_img = img_data.get('handle')
                limites_data = img_data.get('limites_data', {})
                h_lim = limites_data.get('limites_parcel_handle')
                
                if not h_img:
                    print(f"⚠️ Missing image handle; skip {t}.")
                    continue
                    
                print("--🤖 Executing commands...")

                # === SOURCE: Select image and clip ===
                CM.retry_com_operation(src_doc.Activate)
                
                # # Select the image first
                # CM.pumpCommand(src_doc, f'SELECT\n{h_img}\n\n')
                # print(f"🎯 Selected image: {h_img}")
                # CM.pump_sleep(0.5)
                
                # Use IMAGECLIP command for images
                # Try interactive IMAGECLIP
                CM.pumpCommand(src_doc, f'SELECT\n(handent "{h_img}")\n\n')
                CM.pump_sleep(0.2)
                CM.pumpCommand(src_doc, f'IMAGECLIP\nN\nR\n{acad_num(pt1[0])},{acad_num(pt1[1])}\n{acad_num(pt2[0])},{acad_num(pt2[1])}\n\n')
                # === SOURCE: Select image + limites and copy ===
                selection_handles = [h_img]
                if h_lim:
                    selection_handles.append(h_lim)
                
                selection_string = " ".join(f'(handent "{handle}")' for handle in selection_handles)
                CM.pumpCommand(src_doc, f'SELECT\n{selection_string}\n\n')
                #get dimensions of result 
                img_entity = CM.retry_com_operation(lambda: src_doc.HandleToObject(h_img))
                min_pt, max_pt = CM.retry_com_operation(img_entity.GetBoundingBox)
                # min_pt, max_pt = CM.retry_com_operation(img_data['entity'].GetBoundingBox)

                #get dimensions of result 
                w_result = max_pt[0] - min_pt[0]
                h_result = max_pt[1] - min_pt[1]
                base_point = f'{acad_num(min_pt[0])},{acad_num(min_pt[1])}'


                scale_x = ph_w / w_result
                scale_y = ph_h / h_result
                print(f'--📂 Image scaling: {scale_x} x {scale_y}')

                # === DEST: PASTECLIP ===   
                tmp_block =None
                try:

                    tmp_block = Path().home()/'Documents'/'PlanAutomate'/'temp'/f'{img_data['Type']}_{uuid.uuid4().hex[:6]}.dwg'
                except Exception as e:
                    print(f"⚠️ Failed to create temp block path: {e}")
                print(f'temp block path: {tmp_block}')
                try:
                    CM.pumpCommand(src_doc, f'-WBLOCK\n{tmp_block}\n{base_point}\n\n')
                    print(f"📦 WBLOCK {tmp_block} (picked: {'image+limites' if h_lim else 'image only'})")
                except Exception as e:
                    print(f"⚠️ WBLOCK failed: {e}")

                try:
                    # print(f'📦 INSERT {tmp_block} (picked: {"image+limites" if h_lim else "image only"})')
                    CM.pumpCommand(destination_doc, f'-INSERT\n{tmp_block}\n{ip}\n{scale_x}\n{scale_y}\n0\n')
                    print(f"📦 INSERT {tmp_block} in ({ph_org[0]},{ph_org[1]}) | (picked: {'image+limites' if h_lim else 'image only'})")
                except Exception as e:
                    print(f"⚠️ INSERT failed: {e}")
                # copy_cmd = f'COPYBASE\n{bp}\n\n'
                # CM.pumpCommand(src_doc, copy_cmd)
                # print(f"📋 COPYBASE @ {bp} (picked: {'image+limites' if h_lim else 'image only'})")

                # # === DEST: Paste and scale ===
                # CM.retry_com_operation(destination_doc.Activate)
                
                # paste_cmd = f'PASTECLIP\n{ip}\n'
                # CM.pumpCommand(destination_doc, paste_cmd)
                # print(f"📥 PASTECLIP @ {ip}")
                


                # Scale the pasted objects
                # scale_cmd = f'SCALE\nL\n{ip}\n{acad_num(s)}\n'
                # CM.pumpCommand(destination_doc, scale_cmd)
                # print(f"🔎 SCALE Last about {ip} by {s:.6f}")

            print("✅ Done: all images processed via clipboard flow.")
            
        except Exception as e:
            err = str(e)
            return {'success': False, 'error': err, 'user_message': CM.identify_error(err)}
        
        finally:
            try: 
                src_doc.Close(False)
            except Exception: 
                pass

        return {'success': True, 'message': 'Images inserted successfully'}

    except Exception as e:
        err = str(e)
        return {'success': False, 'error': err, 'user_message': CM.identify_error(err)}