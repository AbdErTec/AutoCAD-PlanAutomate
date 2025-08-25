from win32com.client import VARIANT
import win32clipboard
from infra.com_manager import COMManager as CM
from operations.geometry import calculer_bbox, get_dimensions
from infra.converter import Converter
import win32com.client
import pythoncom
import ezdxf
from pathlib import Path

    
def get_echelle_and_layers(source_path, echelles, marge=20):
    try:
        import os
        import ezdxf

        con = Converter()
        dxf_plan = con.dwg_to_dxf(source_path)
        dxf_path = dxf_plan['data'] if isinstance(dxf_plan, dict) else dxf_plan

        if not dxf_path or not os.path.exists(dxf_path):
            return {'success': False, 'user_message': f"DXF introuvable: {dxf_path}"}

        layers = []
        try:
            doc = ezdxf.readfile(dxf_path)
            for l in doc.layers:
                try:
                    is_off = bool(l.is_off())    
                except Exception:
                    is_off = int(getattr(l.dxf, "color", 7)) < 0

                try:
                    is_frozen = bool(l.is_frozen)
                except Exception:
                    flags = int(getattr(l.dxf, "flags", 0))
                    is_frozen = bool(flags & 1)

                try:
                    is_locked = bool(l.is_locked)
                except Exception:
                    flags = int(getattr(l.dxf, "flags", 0))
                    is_locked = bool(flags & 4)
                color_val = int(getattr(l.dxf, "color", 7))
                true_color = None
                try:
                    if hasattr(l.dxf, "true_color") and l.dxf.true_color is not None:
                        true_color = l.dxf.true_color  
                except Exception:
                    pass

                if not is_off:
                    layers.append({
                        'name': l.dxf.name,
                        'visible': not is_off,
                        'frozen': is_frozen,
                        'locked': is_locked,
                        'color': color_val,
                        'true_color': true_color,
                    })

            print(f"✅ Got {len(layers)} layers from DXF")
        except Exception as e:
            print(f"❌ Erreur lors de la lecture des layers depuis DXF: {e}")
            return {'success': False, 'user_message': "Lecture des calques DXF échouée"}

        print('🔎 Calculating echelles...') 
        bbox_plan_result = calculer_bbox(dxf_path)
        if not bbox_plan_result:
            print("❌ Error calculating bounding box")
            return {'success': False, 'user_message': 'Error calculating bounding box'}

        bbox_plan = bbox_plan_result['data']
        print(f'📐 Plan bounding box (DXF): {bbox_plan} | type: {type(bbox_plan)}')

        try:
            a4_width_mm = 210.0 - (2 * marge)
            a4_height_mm = 297.0 - (2 * marge)

            plan_width, plan_height = get_dimensions(bbox_plan)
            print(f'📐 Plan dimensions (width, height): {plan_width:.2f}, {plan_height:.2f}')

            echelles_standards = [echelle.split('/')[1].strip() for echelle in echelles.split('-')]
            print(f'Echelles reçues: {echelles_standards}')

            echelles_analyses = []
            for echelle in echelles_standards:
                echelle = int(echelle)
                contenu_largeur_m = (a4_width_mm * echelle) / 1000.0
                contenu_hauteur_m = (a4_height_mm * echelle) / 1000.0
                inclus = plan_width <= contenu_largeur_m and plan_height <= contenu_hauteur_m

                if inclus:
                    occupation_width = (contenu_largeur_m / plan_width) * 100.0 if plan_width else 0.0
                    occupation_height = (contenu_hauteur_m / plan_height) * 100.0 if plan_height else 0.0
                    occupation_max = max(occupation_width, occupation_height)

                    if 50 <= occupation_max <= 80:
                        score_qualite = 100 - abs(occupation_max - 65)  # ideal ~65%
                    elif 30 <= occupation_max < 50:
                        score_qualite = 80 - (50 - occupation_max)
                    elif 80 < occupation_max <= 95:
                        score_qualite = 90 - (occupation_max - 80)
                    else:
                        score_qualite = max(0, 60 - abs(occupation_max - 65))

                    echelles_analyses.append({
                        'echelle': echelle,
                        'echelle_str': f"1:{echelle}",
                        'occupation_pct': occupation_max,
                        'score_qualite': score_qualite,
                        'contenu_largeur_m': contenu_largeur_m,
                        'contenu_hauteur_m': contenu_hauteur_m,
                        'inclus': True
                    })

            echelles_analyses.sort(key=lambda x: x['score_qualite'], reverse=True)
            print(f"✅ echelles analyses: {echelles_analyses}")
        except Exception as e:
            print(f"❌ Error calculating echelles: {e}")
            return {'success': False, 'user_message': 'Could not calculate echelles'}

        try:
            echelles_finales = []
            for i, ech in enumerate(echelles_analyses):
                label = ech['echelle_str']
                if i == 0:
                    label += " Recommandée"
                elif ech['score_qualite'] > 80:
                    label += " Excellente"
                elif ech['score_qualite'] > 60:
                    label += " Bonne"
                else:
                    label += " Acceptable"
                label += f" ({ech['occupation_pct']:.0f}% du A4)"

                echelles_finales.append({
                    'label': label,
                    'value': ech['echelle'],
                    'is_recommended': i == 0,
                    'occupation': ech['occupation_pct'],
                    'score': ech['score_qualite']
                })
        except Exception as e:
            print(f"❌ Error preparing scale labels: {e}")
            return {'success': False, 'user_message': 'Scale labeling failed'}

        result = {
            'echelles_finales': echelles_finales,
            'layers': layers
        }
        return {'success': True, 'data': result, 'message': 'Prep completed (DXF-only)'}
    except Exception as e:
        print(f"❌ Fatal error in get_echelle_and_layers_from_dxf: {e}")
        import traceback; traceback.print_exc()
        return {'success': False, 'user_message': str(e)}

def rezoom(destination_doc=None):

    if destination_doc is None:
        app, destination_doc = CM.get_app_and_doc()
        if not destination_doc:
            return None
    else:
        app = CM.retry_com_operation(lambda: destination_doc.Application)
    try:    
        # Make sure destination is active
        CM.retry_com_operation(destination_doc.Activate)
        CM.pump_sleep(0.5)

        print('🔎rezooming...')
        CM.pumpCommand(destination_doc, '_AI_SELALL\n\n')
        CM.pump_sleep(0.5)
        CM.pumpCommand(destination_doc, '_ZOOM\nE\n\n')

        return {'success': True, 'data': None, 'message': 'Rezoom successfuly'}
    
    except Exception as e:
        print(f'Couldnt zoom on the document: {e}')
        import traceback
        traceback.print_exec()
        return {'success': False, 'error_message': str(e), 'user_message': CM.identify_error(str(e))}
