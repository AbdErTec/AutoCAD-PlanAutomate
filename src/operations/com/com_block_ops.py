import os
import time
import pythoncom
from win32com.client import VARIANT
import array

import time
import pythoncom
from infra.com_manager import COMManager as CM

def inserer_block(block_path, insert_pt, frame_width, attr_values=None, block_name=None, scale_factor=1.0, rotation=0.0, destination_doc=None):
    """Insert a block from an external DWG file at specified position """
    try:
        
        if destination_doc is None:
            app, destination_doc = CM.get_app_and_doc()
            if not destination_doc:
                return None
        else:
            app = CM.retry_com_operation(lambda: destination_doc.Application)
        
        # Make sure destination is active
        CM.retry_com_operation(destination_doc.Activate)
        CM.pump_sleep(1)
        
        # Get insertion coordinates
        x, y = insert_pt
        print(f"🧱 Inserting block from {block_path} at ({x}, {y})")

        # If block_name not provided, use filename without extension
        if block_name is None:
            block_name = os.path.splitext(os.path.basename(block_path))[0]
        print('🧱 block name: ', block_name)
        
        # Create insertion point as VARIANT array
        insert_pt = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, 
                           array.array('d', [x, y, 0.0]))

        scale_factor = frame_width / 210

        # Insert the block
        # This will automatically define the block if it doesn't exist
        block_ref = CM.retry_com_operation(destination_doc.ModelSpace.InsertBlock,
            insert_pt,          # Insertion point
            block_path,         # Path to block file
            scale_factor,       # X scale
            scale_factor,       # Y scale  
            scale_factor,       # Z scale
            rotation           # Rotation in radians
        )
        
        CM.pump_sleep(1)

        print(f"✅ Block '{block_name}' inserted successfully")
        print(f"🎯 Block reference: {block_ref}\n type: {type(block_ref)}")

        # pumpCommand(destination_doc, "ATTDISP\n1\nnormal\n\n")
        
        # CM.pump_sleep(2)
        print("="*50)

        print(f"🔍 Debugging block attributes...")
        print(f"Block has attributes: {CM.retry_com_operation(lambda: block_ref.HasAttributes)}")
        
        if not CM.retry_com_operation(lambda: block_ref.HasAttributes):
            print("❌ Block has no attributes!")
            return {'success': False, 'user_message': "Block has no attributes!"}
        
        try:
            attributes = CM.retry_com_operation(block_ref.GetAttributes)
            print(f"📊 Found {len(attributes)} attributes")
            
            # List all available attributes
            print("\n📋 Available attributes:")
            for i, att in enumerate(attributes):
                print(f"  {i+1}. Tag: '{att.TagString}' | Current Value: '{att.TextString}'")
            
            # Show what we're trying to match
            if attr_values:
                print(f"\n🎯 Looking for these attributes:")
                for key, value in attr_values.items():
                    print(f"  '{key}' -> '{value}'")
                
                print(f"\n🔄 Matching results:")
                matches = 0
                for att in attributes:
                    tag = CM.retry_com_operation(lambda: att.TagString).upper().split('[')[1].split(']')[0].strip()
                    print(f"  '{tag}'")
                    # print(f'tag to match: {att['tag']}')
                    
                    if tag in attr_values:
                        matches += 1
                        print(f"  ✅ MATCH: '{tag}' will be set to '{attr_values[tag]}'")
                    else:
                        print(f"  ❌ NO MATCH: '{tag}' (not in provided values)")
                
                print(f"\n📈 Total matches: {matches}/{len(attributes)}")
                
                if matches == 0:
                    print("⚠️ WARNING: No attribute matches found!")
                    print("💡 Possible issues:")
                    print("   - Attribute names in block don't match your dictionary keys")
                    print("   - Case sensitivity issues")
                    print("   - Extra spaces in attribute names")
            
        except Exception as e:
            print(f"❌ Error getting attributes: {e}")
            error_msg = str(e)
            return {'success': False, 'error': error_msg, 'user_message': f"Error getting attributes: {e}"}

        print(f"📝 Starting attribute filling...")
        
        if not CM.retry_com_operation(lambda: block_ref.HasAttributes):
            print("❌ Block has no attributes to fill")
            return {'success': False, 'user_message': f"Block has no attributes to fill"}
        
        if not attr_values:
            print("❌ No attribute values provided")
            return {'success': False, 'user_message': f"No attributes provided"}
        
        try:
            attributes = CM.retry_com_operation(block_ref.GetAttributes)
            filled_count = 0
            
            print(f"🔄 Processing {len(attributes)} attributes...")
            
            updated_attrs = []
            for i, att in enumerate(attributes):
                tag = CM.retry_com_operation(lambda: att.TagString).upper().split('[')[1].split(']')[0].strip() 
                print(f"  Processing attribute {i+1}: '{tag}'")
                
                if tag in attr_values:
                    att.TextString = str(attr_values[tag])
                    att.Invisible = False
                    updated_attrs.append(att)

            time.sleep(0.3)
        
            for att in updated_attrs:
                CM.retry_com_operation(att.Update)
                print(f'updated {att}')
                filled_count += 1
            
            print(f"✅ Successfully filled {filled_count}/{len(attributes)} attributes")            # return filled_count > 0
            
        except Exception as e:
            print(f"❌ Error filling attributes: {e}")
            error_msg = str(e)
            user_msg = CM.identify_error(error_msg)
            return {'success': False, 'error':error_msg, 'user_message': user_msg}
        
        return {'success': True, 'data': block_ref, 'message':'Block inserted successfully'}
        
    except Exception as e:
        print(f"❌ Block insertion failed: {e}")
        error_msg = str(e)
        user_msg = CM.identify_error(error_msg)
        return {'success': False, 'error':error_msg, 'user_message': user_msg}
    
def detect_placeholders(block_ref, destination_doc=None):
    """
    Explodes the provided block_ref only, finds placeholder geometries by layer name,
    computes their bbox and origin (min point), and deletes the exploded temp objects.
    Returns: {'success': True, 'placeholders': [ {name, width, height, origin_pt}, ... ]}
    """
    try:
        if destination_doc is None:
            app, destination_doc = CM.get_app_and_doc()
            if not destination_doc:
                return {'success': False, 'user_message': 'No destination document'}

        CM.retry_com_operation(destination_doc.Activate)
        CM.pump_sleep(0.1)

        # explode ONLY the given block_ref
        exploded = CM.retry_com_operation(block_ref.Explode)
        if not exploded:
            return {'success': False, 'user_message': 'Explode returned no objects'}

        # find all layers that look like placeholders (in the exploded set only)
        placeholders = []
        for obj in exploded:
            try:
                layer_name = obj.Layer.strip().lower()
            except Exception:
                continue

            if 'placeholder' in layer_name:
                try:
                    min_pt, max_pt = obj.GetBoundingBox()
                    width = max_pt[0] - min_pt[0]
                    height = max_pt[1] - min_pt[1]

                    placeholders.append({
                        'name': layer_name,                 # e.g. 'orthophoto_placeholder' or 'map_placeholder'
                        'width': float(width),
                        'height': float(height),
                        'origin_pt': (float(min_pt[0]), float(min_pt[1]))
                    })
                except Exception:
                    continue

        # cleanup the exploded temps
        for obj in exploded:
            try:
                obj.Delete()
            except Exception:
                pass

        if not placeholders:
            return {'success': False, 'user_message': 'No placeholders found after explode'}

        return {'success': True, 'placeholders': placeholders, 'message': 'Placeholders detected successfully'}

    except Exception as e:
        return {'success': False, 'error': str(e), 'user_message': CM.identify_error(str(e))}