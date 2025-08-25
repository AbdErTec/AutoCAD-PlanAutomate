from infra.com_manager import COMManager as CM
import os
def inserer_qr(source_path, block_ref, destination_doc=None):
    """
    Insert and scale a QR code image to replace a placeholder in an exploded block.
    Improved version with better error handling and reliability.
    """
    try:
        # Initialize document and application
        if destination_doc is None:
            app, destination_doc = CM.get_app_and_doc()
            if not destination_doc:
                print("❌ Could not get destination document")
                return {'success': False, 'user_messsage': 'Could not get destination document'}
        else:
            app = CM.retry_com_operation(lambda: destination_doc.Application)

        # Validate source file exists
        if not os.path.exists(source_path):
            print(f"❌ QR code image not found: {source_path}")
            return {'success': False, 'user_messsage': 'Qr code image not found'}

        # Get block name with better error handling
        block_name_ref = getattr(block_ref, 'EffectiveName', None) or getattr(block_ref, 'Name', None)
        if not block_name_ref:
            print("❌ Could not determine block name")
            return {'success': False, 'user_messsage': 'could not determine block name'}
            
        print(f'🔍 Processing block: {block_name_ref}')

        # Validate block definition exists
        try:
            block_def = destination_doc.Blocks.Item(block_name_ref)
        except Exception as e:
            print(f"❌ Block definition not found: {block_name_ref}")
            error_msg = str(e)
            user_msg = CM.identify_error(error_msg)
            return {'success': False,'error': error_msg, 'user_messsage': user_msg}

        # Find all block references with the same name
        block_refs = []
        for obj in CM.retry_com_operation(lambda: destination_doc.ModelSpace):
            if hasattr(obj, 'Name') and obj.Name == block_name_ref:
                block_refs.append(obj)

        if not block_refs:
            print(f"❌ No block references found with name '{block_name_ref}'")
            return {'success': False, 'user_messsage': 'Block reference not found'}

        print(f"📦 Found {len(block_refs)} block reference(s)")

        # Process each block reference
        success_count = 0
        for i, block_ref_obj in enumerate(block_refs):
            try:
                print(f"\n🔄 Processing block reference {i+1}/{len(block_refs)}")
                before = destination_doc.ModelSpace.Count
                # Explode the block reference
                block_handle = CM.retry_com_operation(lambda: block_ref_obj.Handle)
                exploded_objects = CM.retry_com_operation(block_ref_obj.Explode)
                # CM.pumpCommand(destination_doc, f'_SELECT\n(handent "{block_handle}")\n\n')
                # CM.pump_sleep(0.5)
                # CM.pumpCommand(destination_doc, '-EXPLODE\n\n')
                # CM.pump_sleep(0.5)
                print(f"📤 Exploded block reference '{block_name_ref}'")
                after = destination_doc.ModelSpace.Count
                exploded_objects = [obj for i, obj in enumerate(CM.retry_com_operation(lambda: destination_doc.ModelSpace)) if i >= before]
                
                if not exploded_objects:
                    print("❌ No objects found after exploding block")
                    continue

                # Find placeholder object in exploded objects first
                placeholder = None
                placeholder_bounds = None
                
                for obj in exploded_objects:
                    if hasattr(obj, 'Layer') and obj.Layer.strip() == 'qr_placeholder':
                        try:
                            min_pt, max_pt = CM.retry_com_operation(obj.GetBoundingBox)
                            placeholder = obj
                            placeholder_bounds = {
                                'width': max_pt[0] - min_pt[0],
                                'height': max_pt[1] - min_pt[1],
                                'origin': min_pt
                            }
                            print(f"🎯 Found placeholder in exploded objects")
                            break
                        except Exception as e:
                            print(f"❌ Error getting placeholder bounds: {e}")
                            continue

                # If not found in exploded objects, search model space
                if placeholder is None:
                    print("🔍 Searching model space for placeholder...")
                    # CM.pump_sleep(0.5)
                    
                    placeholder_objs = []
                    try:
                        for obj in CM.retry_com_operation(lambda: destination_doc.ModelSpace):
                            if hasattr(obj, 'Layer') and obj.Layer.strip() == 'qr_placeholder':
                                placeholder_objs.append(obj)
                    except Exception as e:
                        print(f"❌ Error searching model space: {e}")
                        continue

                    if placeholder_objs:
                        placeholder = placeholder_objs[0]
                        try:
                            min_pt, max_pt = CM.retry_com_operation(placeholder.GetBoundingBox)
                            placeholder_bounds = {
                                'width': max_pt[0] - min_pt[0],
                                'height': max_pt[1] - min_pt[1],
                                'origin': min_pt
                            }
                            print(f"🎯 Found placeholder in model space")
                        except Exception as e:
                            print(f"❌ Error calculating placeholder bounds: {e}")
                            continue

                if placeholder is None or placeholder_bounds is None:
                    print("❌ No QR placeholder found")
                    continue

                print(f"📐 Placeholder dimensions: {placeholder_bounds['width']:.2f} x {placeholder_bounds['height']:.2f}")

                # Insert and scale the QR code image
                CM.retry_com_operation(destination_doc.Activate)
                CM.pump_sleep(1)

                # Get initial object count
                before_count = destination_doc.ModelSpace.Count

                # Insert the image using ATTACH command
                origin = placeholder_bounds['origin']
                print(f"🖼️ Inserting image at origin: ({origin[0]:.3f}, {origin[1]:.3f})")
                
                CM.pumpCommand(destination_doc, f'-attach\n{source_path}\n{origin[0]},{origin[1]}\n\nY\n\n')
                CM.pump_sleep(1)

                # Check if insertion was successful
                after_count = destination_doc.ModelSpace.Count
                if before_count >= after_count:
                    print("❌ Image insertion failed - no new objects detected")
                    continue

                # Get the newly inserted objects
                new_objects = [obj for j, obj in enumerate(CM.retry_com_operation(lambda: destination_doc.ModelSpace)) if j >= before_count]
                new_object_handles = [obj.Handle for obj in new_objects]
                
                print(f"🆕 Inserted {len(new_objects)} new object(s)")
             
                if not new_objects:
                    print("❌ No new objects found after insertion")
                    continue
                for obj in new_objects:
                    if obj.ObjectName == 'AcDbRasterImage':
                        print(f'Making background of {obj.Handle} transparent')
                        obj.Transparency = True
                        
                # Calculate scaling for the main object
                main_obj = new_objects[0]
                try:
                    img_min_pt, img_max_pt = CM.retry_com_operation(main_obj.GetBoundingBox)
                    img_width = img_max_pt[0] - img_min_pt[0]
                    img_height = img_max_pt[1] - img_min_pt[1]

                    print(f"📐 Image dimensions: {img_width:.2f} x {img_height:.2f}")

                    # Calculate scale factors
                    scale_x = placeholder_bounds['width'] / img_width if img_width > 0 else 1
                    scale_y = placeholder_bounds['height'] / img_height if img_height > 0 else 1
                    
                    # Use uniform scaling (maintain aspect ratio)
                    uniform_scale = min(scale_x, scale_y)
                    
                    print(f"📐 Scale factors: X={scale_x:.3f}, Y={scale_y:.3f}, Uniform={uniform_scale:.3f}")

                    # Only scale if necessary and scale factor is reasonable
                    if uniform_scale > 0.001 and abs(uniform_scale - 1.0) > 0.001:
                        # Select all new objects
                        selection_string = " ".join(f'(handent "{handle}")' for handle in new_object_handles)
                        
                        print("🎯 Selecting and scaling inserted objects...")
                        CM.pumpCommand(destination_doc, f'_SELECT\n{selection_string}\n\n')
                        CM.pump_sleep(0.5)

                        # Scale the objects
                        CM.pumpCommand(destination_doc, f'SCALE\n{origin[0]},{origin[1]}\n{uniform_scale}\n\n')
                        CM.pump_sleep(0.5)
                    else:
                        print("⚠️ Skipping scaling (factor too small or close to 1.0)")

                    print("✅ QR code inserted and scaled successfully")
                    success_count += 1

                    print('Deleting exploded block...')
                    for obj in exploded_objects:
                        print(f'🧹 deleting ({obj.Handle}, {obj.ObjectName})')
                        obj.Delete()
                        CM.pump_sleep(0.5)

                except Exception as e:
                    print(f"❌ Error during scaling: {e}")
                    continue

            except Exception as e:
                print(f"❌ Error processing block reference {i+1}: {e}")
                continue

        # Report final results
        if success_count > 0:
            print(f"\n✅ Successfully processed {success_count}/{len(block_refs)} block references")
            return {'success': True, 'data': 'QR_insere', 'user_messsage': 'Qr inserted successfully'}
        else:
            print(f"\n❌ Failed to process any block references")
            return {'success': False , 'user_messsage': 'Failed to process block reference'}

    except Exception as e:
        error_msg = str(e)
        user_msg = CM.identify_error(error_msg)

        return {'success': False, 'error':error_msg, 'user_message': user_msg}
