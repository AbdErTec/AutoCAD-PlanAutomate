# app/worker.py
from PyQt5.QtCore import QThread, pyqtSignal
import traceback
import pythoncom
from pathlib import Path
import win32com.client
from datetime import datetime
from infra.helpers import *


from infra.com_manager import COMManager as CM

from handbooks.handbook_models import (
    PlanHandbook,
    OrthoMapHandbook,
    CartoucheHandbook,
)

class AutoCADWorker(QThread):
    progress_updated = pyqtSignal(int)       
    status_updated = pyqtSignal(str)         
    process_finished = pyqtSignal(bool, str) 

    def __init__(self, file_paths, selected_layers, selected_echelle, client_data, images_paths):
        super().__init__()
        self.file_paths = file_paths
        self.images_paths = images_paths
        self.selected_layers = selected_layers
        self.selected_echelle = selected_echelle
        self.client_data = client_data

        self.ctx = {
            "file_paths": file_paths,
            "images_paths": images_paths,
            "selected_layers": selected_layers,
            "selected_echelle": selected_echelle,
            "client_data": client_data,
        }
        self._total_steps = 0
        self._done_steps = 0

    def _build_handbooks(self, destination_doc=None):
        print('📚 Building handbooks...')
        plan = PlanHandbook(destination_doc, self.ctx, self.file_paths, self.selected_layers, self.selected_echelle)
        print('* Plan handbok built')
        ortho = OrthoMapHandbook(destination_doc, self.ctx, self.file_paths, self.client_data)
        print('* Ortho handbok built')
        cart = CartoucheHandbook(destination_doc, self.ctx, self.file_paths, self.client_data, self.images_paths)
        print('* Cartouche handbok built')
        return [plan, ortho, cart]

    def _count_total_steps(self, handbooks):
        self._total_steps = sum(len(getattr(hb, "handbook", [])) for hb in handbooks) or 1

    def _checkpoint_callback(self):
        def _cb(step_name, result, ctx):
            self._done_steps += 1
            pct = int(self._done_steps * 100 / max(self._total_steps, 1))
            self.progress_updated.emit(pct) 
            self.status_updated.emit(f"Step done: {step_name} ({pct}%)")
        return _cb

    def run(self):
        try:
            pythoncom.CoInitialize()
            print("== 🤖 ACAD worker starting... ==")
            
            try:
                app, destination_doc, created_new = ensure_app_and_doc()  
                print(f"✅ Using document: {destination_doc.Name} (created_new={created_new})")
            except Exception as e:
                print(f"❌ Failed to get/create AutoCAD doc: {e}")
                self.error.emit("Impossible d’ouvrir AutoCAD ou créer un document.")
                return

            app.Visible = True

            if not destination_doc:
                self.error.emit("Le document AutoCAD ciblé n’est plus ouvert.")
                return
            
            CM.retry_com_operation(destination_doc.Activate)
            self.progress_updated.emit(0)
            self.status_updated.emit("Préparation de l'environnement...")

            handbooks = self._build_handbooks(destination_doc)
            self._count_total_steps(handbooks)
            cb = self._checkpoint_callback() if self._total_steps else None
            
            for hb in handbooks:
                hb.emitter = self   
                if hasattr(hb, "on_step_done") and callable(self._checkpoint_callback()):
                    hb.on_step_done = cb
                hb.execute_handbook()

                interesting = {k: self.ctx.get(k) for k in [
                    "bbox_data",
                    "frame_width", "frame_height", "frame_coords",
                    "leg_coords", "table_coords",
                    "bornes_results",
                    "orthomap_insert_pt", "cartouche_insert_pt",
                    "orthomap_ref", "cartouche_ref", "qr_ref"
                ] if k in self.ctx}
                print(f"[ctx after {hb.__class__.__name__}] {interesting}")
            self.progress_updated.emit(90)
            self.status_updated.emit("On y est presque...")
            doc_path = f'Plan {self.client_data["plan"]}_{datetime.now().strftime("%Y%m%d__%H%M%S")}_{self.client_data["declaration_par"]}.dwg'
            destination_doc.SaveAs(doc_path)
            print(f"✅ Plan saved to {doc_path}")
            self.status_updated.emit(f"Plan saved to {doc_path}")
            CM.pump_sleep(1.5)
            self.progress_updated.emit(100)
            self.status_updated.emit("Création du plan terminé.")
            print("\n✅ Dummy pipeline finished OK.")

            self.process_finished.emit(True, "Création du plan terminé avec succès.")

        except Exception as e:
            print(f"❌ COM failed: {CM.identify_error(str(e))}")
            traceback.print_exc()
            self.process_finished.emit(False, str(e)) 
            return
        finally:
            pythoncom.CoUninitialize()