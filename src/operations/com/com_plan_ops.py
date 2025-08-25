from win32com.client import VARIANT
import pythoncom
from pythoncom import VT_VARIANT, VT_R8
import win32com.client
import math
import array
from infra.com_manager import COMManager as CM
import os 

NORD_ICON = os.path.abspath(os.path.join(os.path.dirname(__file__), "..",'..','..', "assets",'icons', "nord.png"))

def inserer_plan (prepared_data, source_dwg, destination_doc=None):

    try:
        if destination_doc is None:
            app, destination_doc = CM.get_app_and_doc()
            if not destination_doc:
                print("❌ Could not get destination document")
                return {'success': False, 'user_message': 'Could not get destination document'}
        else:
            app = CM.retry_com_operation(lambda: destination_doc.Application)
        
        ms = destination_doc.ModelSpace
        
        temp_dxf = prepared_data['temp_dxf_path']
        insert_pt = prepared_data['base_point']
        inbase = prepared_data['inbase']

        # CM.pumpCommand(destination_doc, f"BASE\n\n{insert_pt[0]},{insert_pt[1]}\n")
        # CM.pump_sleep(0.5)
        # CM.pumpCommand(destination_doc, f"-INSERT\n*{source_dwg}\n{insert_pt[0]},{insert_pt[1]}\n1\n0\n")
        CM.pumpCommand(destination_doc, f"-INSERT\n*{source_dwg}\n{inbase[0]},{inbase[1]}\n1\n0\n")
        CM.pump_sleep(0.5)
        
        # print('🔎 Getting layers...')
        # obj_count = 0
        # layers = []
        # try: 
        #     print(f'➕ source doc layers: {destination_doc.layers}')
        #     for layer in destination_doc.layers:
        #         if layer.LayerOn:
        #             obj_count += 1
        #             print(f'🔔 Layer {layer.Name} is visible')

        #             layers.append({
                
        #                 'name': layer.Name,
        #                 'visible': layer.LayerOn,
        #                 'frozen': layer.Freeze,
        #                 'locked': layer.Lock,
        #                 'color': layer.color,
        #                 # 'objects_count': count_objects,
        #                 # 'suggested_keep': suggest_layer_importance(layer.Name)

        #             })

            # print(f"✅ Got {obj_count} layers")
            # CM.pump_sleep(1)
            # print(f'🧱 layers in original plan: {layers}')

        # except Exception as e:
            # print(f"❌ Erreur lors de la lecture des layers: {e}")
            # return {'success':False, 'user_message': 'Could not read layers from the original plan'}
        


        print ('📜 Managing plan layers...')
        # allowed_layers = [layer.lower() for layer in prepared_data['layers']]
        for layer in destination_doc.Layers:
            if prepared_data['allowed_layers'] is None:
                print('No allowed layers provided. All layers will be visible.')
                break
            
            if layer.Name.lower() not in prepared_data['allowed_layers']:
                print(f'🔕 Turned off {layer.Name}')
                layer.layerOn = False
    
        return {'success': True, 'data': prepared_data, 'message': 'Plan inserted successfully'}


        
        # CM.pump_sleep(1)
    except Exception as e:
        print(f'❌ Fatal error in inserer_plan: {e}')
        error_msg = str(e)
        user_msg = CM.identify_error(error_msg)
        return {'success': False, 'error_message': error_msg, 'user_message': str(e)}

def creer_frame_a4(prepared_data, destination_doc = None):
    try:
        if destination_doc is None:
            app = win32com.client.GetActiveObject("AutoCAD.Application")
        
        destination_doc = CM.retry_com_operation(lambda: app.ActiveDocument)

        CM.retry_com_operation(destination_doc.Activate)
        CM.pump_sleep(0.1)

        print(f"🖼️ Creating frame: {prepared_data['frame_width']:.2f} x {prepared_data['frame_height']:.2f}")

        frame_width = prepared_data['frame_width']
        frame_height = prepared_data['frame_height']
        nord_coords = prepared_data['nord_coords']

        
        frame_points = array.array('d', prepared_data['frame_coords'])

        frame_variant = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, array.array('d', frame_points))
        
        frame = CM.retry_com_operation(destination_doc.ModelSpace.AddPolyline, frame_variant, operation_name='drawing frame')
        frame.Closed = True
        frame.color = 7  # White
        
        print(f"🖼️ A4 Frame created: {frame_width:.2f} x {frame_height:.2f}")
        before = destination_doc.ModelSpace.Count
        nord_img = None
        new_objects = None
        try:
            print('⬆️ inserting nord...')
            # CM.pumpCommand(destination_doc, f'-attach\n"{NORD_ICON}"\n{nord_coords[0]},{nord_coords[1]}\n\nY\n\n')
            CM.pumpCommand(destination_doc, 'IMAGEQUALITY\nHIGH\n')    # High quality
            nord_path = f'{NORD_ICON}'
            
            if not os.path.exists(NORD_ICON):
                print('icon doesnt exit')

            print(f'Nord path: {nord_path}')
            CM.pumpCommand(destination_doc, f'-ATTACH\n{nord_path}\n0,0\n\nY\n\n')

            new_objects = [obj for j, obj in enumerate(CM.retry_com_operation(lambda: destination_doc.ModelSpace)) if j >= before]
            if new_objects is None: 
                print('❌ failed to insert nord')
                return {'success': False, 'error_message': str(e), 'user_message': 'Couldn\'t insert nord image'}
            
            temp_nord = new_objects[0] if new_objects else None

            if temp_nord:
                min_nord, max_nord = temp_nord.GetBoundingBox()
                nord_width = max_nord[0] - min_nord[0]
                target_scale = frame_width * 0.07 
                scale_factor =  target_scale / nord_width
                # nord_x = nord_coords[0] - (nord_width * scale_factor)
                # Delete temp insertion
                temp_nord.Delete()
                before = destination_doc.ModelSpace.Count 
                # Insert at correct location with correct scale
                CM.pumpCommand(
                    destination_doc,
                    f'-ATTACH\n{nord_path}\n{float(nord_coords[0]-((nord_width*scale_factor)/2))},{float(nord_coords[1])}\n{scale_factor}\n0\n\n'
                )
    
                new_objects = [obj for j, obj in enumerate(CM.retry_com_operation(lambda: destination_doc.ModelSpace)) if j >= before]
                nord_img = new_objects[0] if new_objects else None
                if nord_img:
                    print(nord_img.Handle)
                    CM.pumpCommand(destination_doc, f'TRANSPARENCY\n(handent "{nord_img.Handle}")\n\nON\n')
                    CM.pumpCommand(destination_doc, 'IMAGEFRAME\n0\n')      # No frames
                      

                print(f'🖼️ Nord inserted in ({nord_coords}) and scaled by {scale_factor} successfully')
                
        except Exception as e:
            print(f'❌ failed to insert nord: {e}')
            return {'success': False, 'error_message': str(e), 'user_message': 'Failed to insert nord'}
        

        return {
            'success': True, 
            'data': {
                'frame_width': frame_width, 
                'frame_height': frame_height,
                'frame_coords': prepared_data['frame_coords'],
                'nord_coords': prepared_data['nord_coords'],
                'leg_coords': prepared_data['leg_coords'],
                'table_coords': prepared_data['table_coords']
                }, 
            'message': 'Frame created, legend and table insert points calculated successfuly'
            }

    except Exception as e:
        print(f'❌ failed to create a4 frame: {e}')
        error_msg = str(e)
        user_msg = CM.identify_error(error_msg)

        return {'success': False, 'error':error_msg, 'user_message': user_msg}

def inserer_legende(prepared_data, source_dwg, leg_data, destination_doc=None):
    """Insert pre-scaled legend - single operation"""
    try:
        if destination_doc is None:
            app, destination_doc = CM.get_app_and_doc()
        
        print('entered to inserer_legende')
        ms = destination_doc.ModelSpace
        # temp_dxf = prepared_data['temp_dxf_path']
        scale_factor = prepared_data['scale_factor']
        inbase = prepared_data['inbase']
        bbox = prepared_data['original_bbox']

        insert_pt = (leg_data[0] + scale_factor * (inbase[0] - bbox[0]), leg_data[1] + scale_factor * (inbase[1] - bbox[1]))
        
        print(f'🏷️ Inserting scaled legend at {insert_pt}')

        # Single INSERT - pre-scaled legend
        CM.pumpCommand(
            destination_doc, 
            f'-INSERT\n"{source_dwg}"\n{insert_pt[0]},{insert_pt[1]}\n{scale_factor}\n{scale_factor}\n0\n'
        )
        
        CM.pump_sleep(0.5)
        
        print(f"🏷️ Inserted scaled legend with scale factor {scale_factor}")
        
        return {
            'success': True,
            'data': prepared_data, 
            'message': 'Legend inserted successfully'
        }
        
    except Exception as e:
        return {'success': False, 'error': str(e)}
    
def inserer_tableau(prepared_data, destination_doc=None):
    """Create table using prepared data - minimal COM operations"""
    
    try:
        if destination_doc is None:
            app, destination_doc = CM.get_app_and_doc()
            if not destination_doc:
                return {'success': False, 'user_message': 'Could not get destination document'}
        
        
        borne_results = prepared_data['borne_results']
        table_coords = prepared_data['table_coords']
        n_rows = prepared_data['n_rows']
        n_cols = prepared_data['n_cols']
        row_height = prepared_data['row_height']
        col_width = prepared_data['col_width']
        scale_factor = prepared_data['scale_factor']
        # croisillons_coords = prepared_data['croisillons_coords']
        # half_size = prepared_data['half_size']
        
        model = destination_doc.ModelSpace
        CM.retry_com_operation(destination_doc.Activate)
        CM.pump_sleep(0.1)
        
        # --- Create table (single COM operation) ---
        table_x, table_y = table_coords[0], table_coords[1]
        pt = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8,
                    array.array('d', [table_x, table_y, 0.0]))
        
        table = CM.retry_com_operation(
            model.AddTable, 
            pt, n_rows, n_cols, row_height, col_width
        )
        CM.pump_sleep(0.3)
        
        if not table:
            return {'success': False, 'user_message': 'Could not create the table'}
        
        print("✅ Table created successfully")
        
        # --- Fill table (batch COM operations) ---
        CM.retry_com_operation(table.SetText, 0, 0, 'Table des coordonnees')
        CM.retry_com_operation(table.SetText, 1, 0, 'Borne')
        CM.retry_com_operation(table.SetText, 1, 1, 'X')
        CM.retry_com_operation(table.SetText, 1, 2, 'Y')
        
        # Fill data rows
        for row_idx, borne in enumerate(borne_results, start=2):
            table.SetText(row_idx, 0, borne['Borne'])
            table.SetText(row_idx, 1, f'{borne["coords"][0]:.2f}')
            table.SetText(row_idx, 2, f'{borne["coords"][1]:.2f}')
        
        # --- Scale table ---
        handle = table.Handle
        CM.pumpCommand(destination_doc, f'_SELECT\n(handent "{handle}")\n\n')
        CM.pump_sleep(0.2)
        CM.pumpCommand(destination_doc, 
                      f'_SCALE\n\n{table_x},{table_y}\n{scale_factor}\n')
        CM.pump_sleep(0.3)
        
        print(f"✅ Table scaled by factor: {scale_factor}")
        
        # Clean up
        CM.pumpCommand(destination_doc, '_ZOOM\nE\n\n')
        CM.pump_sleep(0.2)
        
        return {
            'success': True, 
            'data': {
                'bornes_results':borne_results,
                'borne_count': len(borne_results),
                'scale_factor': scale_factor
            },
            'message': 'Table created successfully'
        }
        
    except Exception as e:
        print(f"❌ Table insertion failed: {e}")
        return {'success': False, 'error': str(e)}
    
def inserer_croisillons(prepared_data, destination_doc=None):
    if destination_doc is None:
        app, destination_doc = CM.get_app_and_doc()
        if not destination_doc:
            return {'success': False, 'user_message': 'Could not get destination document'}
        
    # croisillions_coords = prepared_data['croisillons_coords']
    crosses   = prepared_data['croisillons_coords']
    # labels    = prepared_data['labels_coords']
    half_size = prepared_data['half_size']
    frame_width = prepared_data['frame_width']
    th        = prepared_data['text_height']
    x_projection = prepared_data['x_projection']
    y_projection = prepared_data['y_projection']
    # ticks     = prepared_data['ticks']

    projections = x_projection + y_projection
    # crosses
    for cross in crosses:
        x, y = cross
        CM.pumpCommand(
            destination_doc,
            f"LINE\n{x-half_size},{y}\n{x+half_size},{y}\n\n"
            f"LINE\n{x},{y-half_size}\n{x},{y+half_size}\n\n"
        )
        # CM.pump_sleep(0.1)

    for proj in projections:
        if proj is not None:
            txt = None
            rot = None
            insert_pt = None
            i_p = proj['insert_pt']
            
            if proj['insert_type'] == 'x_axis':
                rot = 90
                txt = int(round(proj['insert_pt'][0]))
                insert_pt = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, 
                                array.array('d', [i_p[0], i_p[1] + half_size*3, 0.0]))
                # CM.pumpCommand(destination_doc, f"LINE\n{i_p[0]},{i_p[1]}\n{i_p[0]},{i_p[1] + half_size}\n\n")
                line_obj = destination_doc.ModelSpace.AddLine(
                    VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, 
                            array.array('d', [i_p[0], i_p[1], 0.0])),
                    VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, 
                            array.array('d', [i_p[0], i_p[1] + half_size, 0.0]))
                )
            elif proj['insert_type'] == 'y_axis':
                rot = 0
                txt = int(round(proj['insert_pt'][1]))

                insert_pt = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, 
                                array.array('d', [i_p[0] - half_size*5, i_p[1], 0.0]))
                # CM.pumpCommand(destination_doc, f"LINE\n{i_p[0]},{i_p[1]}\n{i_p[0] - half_size},{i_p[1]}\n\n")
                line_obj = destination_doc.ModelSpace.AddLine(
                    VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, 
                            array.array('d', [i_p[0], i_p[1], 0.0])),
                    VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, 
                            array.array('d', [i_p[0] - half_size, i_p[1], 0.0]))
                )

        text_obj = CM.retry_com_operation(
        destination_doc.ModelSpace.AddText,
        txt,
        insert_pt,
        th,
        operation_name='adding text label'
    )
    
    # Set rotation after creation if needed
        text_obj.Rotation = math.radians(rot)

    print(f"✅ Croisillons and labels added successfully")
    return {
        'success': True,
        'message': 'Croisillons and labels added successfully'
    }

    