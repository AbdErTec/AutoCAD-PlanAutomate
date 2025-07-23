from pyautocad import Autocad
import sys
import pythoncom
import comtypes
from comtypes.automation import VARIANT, VT_ARRAY, VT_DISPATCH

import threading

import utils
import Gui
if __name__ == "__main__":
    pythoncom.CoInitialize()  # initialize COM threading

    acad = Autocad(create_if_not_exists=True)
    import time
    doc = acad.doc
    time.sleep(2)
    ms = doc.ModelSpace
    acad.prompt("Hello, Autocad from Python\n")
    print(f"Nom du fichier: {acad.doc.Name}")
    form_data = Gui.start()
    if not form_data:
        print("User cancelled or didn't submit form.")
        sys.exit()


    # sys.exit(app.exec())
    sys.exit(0)   
