import os
import subprocess
from pathlib import Path
import shutil

class Converter:
    def __init__(self, oda_exe_path=None):
        if oda_exe_path is None:
            oda_exe_path = Path(__file__).resolve().parent.parent.parent / "assets" / "bin" / "ODAFileConverter" / "ODAFileConverter.exe"
        self.oda_exe_path = str(oda_exe_path)
        self.temp_dir = Path.home() / "Documents" / "PlanAutomate" / "temp"
        self.temp_dir.mkdir(parents=True, exist_ok=True)

        
    def dwg_to_dxf(self, dwg_path, version="ACAD2013"):
            try:
                if not os.path.exists(dwg_path):
                    return {"success": False, "user_message": f"File not found: {dwg_path}"}
    
                dwg_path = Path(dwg_path)
                dxf_path = self.temp_dir / (dwg_path.stem + ".dxf")
                
                temp_input_dir = self.temp_dir / "input"
                temp_input_dir.mkdir(exist_ok=True)
                
                temp_dwg_path = temp_input_dir / dwg_path.name
                shutil.copy2(dwg_path, temp_dwg_path)

                if dxf_path.exists():
                    dxf_path.unlink()
                
                cmd = [
                self.oda_exe_path,
                str(temp_input_dir),   
                str(self.temp_dir),     
                version,               
                "DXF",               
                "0",                   
                "1"                    
                ]

                result = subprocess.run(
                    cmd, 
                    capture_output=True, 
                    text=True, 
                    timeout=60,
                    creationflags=subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0
                )
                
                if any(temp_input_dir.iterdir()):
                    print('cleaning up...')
                    shutil.rmtree(temp_input_dir)
                
                if result.returncode == 0 and dxf_path.exists():
                    print('doc converted to dxf')
                    return {"success": True, "data": str(dxf_path), "message": "Converted successfully"}
                else:
                    return {"success": False, "user_message": f"Conversion failed: {result.stderr}"}
                    
            except subprocess.TimeoutExpired:
                return {"success": False, "user_message": "Conversion timed out"}
            except Exception as e:
                return {"success": False, "user_message": f"Conversion error: {str(e)}"}