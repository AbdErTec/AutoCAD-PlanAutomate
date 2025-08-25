from PyQt5.QtCore import QThread, pyqtSignal
import traceback

from pa_workflow.com_utils import get_echelle_and_layers as g_e_a_l

class AutoCADPrepWorker(QThread):
    done = pyqtSignal(dict, str)   
    error = pyqtSignal(str)
    status_updated = pyqtSignal(str)

    def __init__(self, file_paths, echelles):
        super().__init__()
        self.file_paths = file_paths
        self.echelles = echelles

    def run(self):
        try:
            self.status_updated.emit('Conversion DWG → DXF et lecture des calques...')
            result = g_e_a_l(self.file_paths[0], echelles=self.echelles)

            if not result.get('success'):
                self.error.emit(result.get("user_message", "Erreur inconnue"))
                return

            self.status_updated.emit('Lecture terminée')
            self.done.emit(result["data"], "Prep completed (DXF only)")

        except Exception as e:
            print(f"❌ Prep worker error: {e}")
            traceback.print_exc()
            self.error.emit(str(e))
