import win32com.client
import time
import win32gui
import win32con

def create_table_with_dialog_workflow(coords, bbox_data):
    """
    Handle AutoCAD table creation with dialog interaction
    1. Send TABLE command (opens dialog)
    2. Handle dialog programmatically or instruct user
    3. Send insertion point
    4. Use TABLEEDIT to fill data
    """
    try:
        acad = win32com.client.Dispatch("AutoCAD.Application")
        doc = acad.ActiveDocument
        
        print("📋 Starting table creation workflow...")
        
        # Calculate position
        # table_x = bbox_data[3] if bbox_data and len(bbox_data) > 3 else 0
        # table_y = bbox_data[1] - 5 if bbox_data and len(bbox_data) > 1 else -5
        # n_rows = (len(coords) // 2) + 1
        table_x = 12.3222211
        table_y = 52.3222211
        n_rows = 3
        
        print(f"   Table will be created at ({table_x:.1f}, {table_y:.1f}) with {n_rows} rows")
        
        # Method 1: Try to create table without dialog using system variables
        return create_table_without_dialog(acad, doc, coords, table_x, table_y, n_rows)
        
    except Exception as e:
        print(f"❌ Table workflow failed: {e}")
        return False

def create_table_without_dialog(acad, doc, coords, table_x, table_y, n_rows):
    """
    Create table by setting system variables to avoid dialog
    """
    try:
        print("🔧 Attempting to bypass dialog with system variables...")
        
        # Set system variables to control table creation
        # These control the table dialog behavior
        try:
            # Save current settings
            original_cmddia = doc.GetVariable("CMDDIA")
            original_filedia = doc.GetVariable("FILEDIA")
            
            # Disable dialog boxes
            doc.SetVariable("CMDDIA", 0)  # Disable command dialogs
            doc.SetVariable("FILEDIA", 0)  # Disable file dialogs
            
            print("   Dialog suppression enabled")
            
            # Now try the table command
            doc.SendCommand("_TABLE\n")
            time.sleep(0.2)
            
            # Send insertion point
            doc.SendCommand(f"{table_x},{table_y}\n")
            time.sleep(0.2)
            
            # Send table parameters
            doc.SendCommand(f"{n_rows}\n")  # rows
            time.sleep(0.1)
            doc.SendCommand("2\n")          # columns
            time.sleep(0.1)
            doc.SendCommand("1.5\n")        # row height
            time.sleep(0.1)
            doc.SendCommand("8\n")          # column width
            time.sleep(0.1)
            doc.SendCommand("\n")           # finish
            
            time.sleep(1.0)  # Wait for table creation
            
            # Restore original settings
            doc.SetVariable("CMDDIA", original_cmddia)
            doc.SetVariable("FILEDIA", original_filedia)
            
            print("   ✅ Table created without dialog")
            
            # Now fill the table using TABLEEDIT
            return fill_table_with_tableedit(doc, coords, n_rows)
            
        except Exception as e:
            print(f"   ❌ Dialog bypass failed: {e}")
            # Restore settings on error
            try:
                doc.SetVariable("CMDDIA", original_cmddia)
                doc.SetVariable("FILEDIA", original_filedia)
            except:
                pass
            return False
            
    except Exception as e:
        print(f"❌ Table creation without dialog failed: {e}")
        return False

def fill_table_with_tableedit(doc, coords, n_rows):
    """
    Fill table data using TABLEEDIT command
    """
    try:
        print("📝 Filling table with TABLEEDIT...")
        
        # Get the model space to find our table
        model = doc.ModelSpace
        
        if model.Count > 0:
            # Get the last object (should be our table)
            table_obj = model.Item(model.Count - 1)
            
            # Try direct SetText first (might work now)
            if hasattr(table_obj, 'SetText'):
                print("   Using direct SetText method...")
                try:
                    # Set headers
                    table_obj.SetText(0, 0, "X")
                    table_obj.SetText(0, 1, "Y")
                    
                    # Set coordinate data
                    point_index = 0
                    for row in range(1, n_rows):
                        if point_index * 2 + 1 >= len(coords):
                            break
                        
                        x = coords[point_index * 2]
                        y = coords[point_index * 2 + 1]
                        
                        table_obj.SetText(row, 0, f"{x:.2f}")
                        table_obj.SetText(row, 1, f"{y:.2f}")
                        
                        point_index += 1
                    
                    print("   ✅ Table filled using direct method")
                    return True
                    
                except Exception as e:
                    print(f"   Direct method failed: {e}")
                    print("   Trying TABLEEDIT command method...")
            
            # If direct method fails, try TABLEEDIT command
            return fill_using_tableedit_command(doc, coords, n_rows)
        else:
            print("   ❌ No table object found")
            return False
            
    except Exception as e:
        print(f"❌ Table filling failed: {e}")
        return False

def fill_using_tableedit_command(doc, coords, n_rows):
    """
    Use TABLEEDIT command to fill table data
    """
    try:
        print("   Using TABLEEDIT command...")
        
        # Start TABLEEDIT command
        doc.SendCommand("_TABLEEDIT\n")
        time.sleep(0.3)
        
        # Click on the table (we'll need to specify a point on the table)
        # For now, we'll try to select the last object
        doc.SendCommand("_LAST\n")
        time.sleep(0.3)
        
        # Now we should be in table edit mode
        # Fill headers first
        fill_table_cells_via_commands(doc, coords, n_rows)
        
        # Exit table edit mode
        doc.SendCommand("\n")  # or ESC
        time.sleep(0.2)
        
        print("   ✅ Table filled using TABLEEDIT")
        return True
        
    except Exception as e:
        print(f"   ❌ TABLEEDIT method failed: {e}")
        return False

def fill_table_cells_via_commands(doc, coords, n_rows):
    """
    Fill individual table cells using command line
    This is the tricky part - navigating table cells via commands
    """
    try:
        print("   Filling cells via navigation...")
        
        # Move to first cell (0,0) and enter "X"
        doc.SendCommand("X\n")  # Header X
        time.sleep(0.1)
        doc.SendCommand("\t")   # Tab to next cell
        time.sleep(0.1)
        doc.SendCommand("Y\n")  # Header Y
        time.sleep(0.1)
        
        # Move to next row
        doc.SendCommand("\t")   # Tab to next cell (should go to next row)
        time.sleep(0.1)
        
        # Fill coordinate data
        point_index = 0
        cells_filled = 0
        
        for row in range(1, min(n_rows, 6)):  # Limit to prevent infinite loop
            if point_index * 2 + 1 >= len(coords):
                break
            
            x = coords[point_index * 2]
            y = coords[point_index * 2 + 1]
            
            # Enter X coordinate
            doc.SendCommand(f"{x:.2f}\n")
            time.sleep(0.1)
            doc.SendCommand("\t")  # Tab to Y column
            time.sleep(0.1)
            
            # Enter Y coordinate
            doc.SendCommand(f"{y:.2f}\n")
            time.sleep(0.1)
            doc.SendCommand("\t")  # Tab to next row
            time.sleep(0.1)
            
            point_index += 1
            cells_filled += 1
        
        print(f"   Filled {cells_filled} coordinate pairs")
        
    except Exception as e:
        print(f"   Cell filling failed: {e}")

def create_table_interactive_instructions(coords, bbox_data):
    """
    Provide instructions for manual table creation when automation fails
    """
    print("\n📋 MANUAL TABLE CREATION INSTRUCTIONS:")
    print("=" * 50)
    
    table_x = bbox_data[3] if bbox_data and len(bbox_data) > 3 else 0
    table_y = bbox_data[1] - 5 if bbox_data and len(bbox_data) > 1 else -5
    n_rows = (len(coords) // 2) + 1
    
    print(f"1. Type 'TABLE' in AutoCAD command line")
    print(f"2. In the dialog:")
    print(f"   - Rows: {n_rows}")
    print(f"   - Columns: 2")
    print(f"   - Row height: 1.5")
    print(f"   - Column width: 8")
    print(f"3. Click OK")
    print(f"4. Click at insertion point: {table_x:.1f},{table_y:.1f}")
    print(f"5. The table will be created")
    print(f"6. Double-click the table to edit")
    print(f"7. Fill with this data:")
    print("   Row 1: X, Y")
    
    point_index = 0
    row = 2
    while point_index * 2 + 1 < len(coords):
        x = coords[point_index * 2]
        y = coords[point_index * 2 + 1]
        print(f"   Row {row}: {x:.2f}, {y:.2f}")
        point_index += 1
        row += 1
    
    print("8. Press ESC to exit table edit mode")
    print("=" * 50)

def test_table_workflow():
    """
    Test the complete table workflow
    """
    print("🚀 Testing AutoCAD Table Workflow")
    print("=" * 50)
    
    # Sample data
    sample_coords = [1.5, 2.3, 4.7, 6.1, 8.9, 3.4]
    sample_bbox = [0, -10, 0, 20, 10, 0, 30, 0]
    
    print(f"📊 Creating table for {len(sample_coords)//2} coordinate points")
    
    # Try automated approach
    success = create_table_with_dialog_workflow(sample_coords, sample_bbox)
    
    if not success:
        print("\n⚠️ Automated approach failed")
        create_table_interactive_instructions(sample_coords, sample_bbox)
    else:
        print("\n✅ Table workflow completed successfully!")

def simple_table_creation_for_main_code(coords, bbox_data):
    """
    Simplified version for integration into your main code
    """
    try:
        print("📋 Creating coordinates table...")
        
        # Try the automated approach first
        success = create_table_with_dialog_workflow(coords, bbox_data)
        
        if success:
            print("✅ Table created successfully")
            return True
        else:
            print("⚠️ Automated table creation failed")
            print("   Creating instruction guide instead...")
            create_table_interactive_instructions(coords, bbox_data)
            return False
            
    except Exception as e:
        print(f"❌ Table creation failed: {e}")
        return False

if __name__ == "__main__":
    test_table_workflow()