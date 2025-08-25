from pa_workflow.steps.signals import set_progress, set_status
from infra.com_manager import COMManager as CM
from operations.geometry import calculer_bbox

from operations.com import com_plan_ops as com_plan
from operations.dxf import dxf_plan_ops as dxf_plan

def step_calculer_bbox_plan(file_path, emitter):
    
    set_status(emitter, "Calcul du bbox du plan")
    set_progress(emitter, 5)
    
    bbox_data_plan = calculer_bbox(file_path=file_path)
    CM.pump_sleep(2)
    CM.check_success(bbox_data_plan, 'calcul bbox plan')
    if len(bbox_data_plan['data']) < 4:
        raise Exception("bbox_data_plan trop court")
    return bbox_data_plan

def step_inserer_plan(source_dwg, source_dxf, bbox_data, layers, emitter):
    set_status(emitter, "Insertion du plan")
    set_progress(emitter, 7)
    
    prepared_data = dxf_plan.prepare_inserer_plan(source_dxf=source_dxf, bbox_data=bbox_data, layers=layers)
    CM.check_success(prepared_data, "plan_preparation")

    data = prepared_data['data']
    # return prepared_data
    
    plan_result = com_plan.inserer_plan(prepared_data=data, source_dwg= source_dwg)
    # CM.pump_sleep(2)
    CM.check_success(plan_result, 'inserer_plan')
    return plan_result

def step_creer_frame_a4(bbox_data, echelle, emitter):
    set_status(emitter, "Création du frame")
    set_progress(emitter, 15)
    prepared_data = dxf_plan.prepare_creer_frame_a4(bbox_data=bbox_data,echelle=echelle)
    CM.check_success(prepared_data, 'frame_prep')

    data = prepared_data['data']
    
    if not isinstance(data, dict) or 'leg_coords' not in prepared_data['data'] or 'table_coords' not in prepared_data['data']:
        raise Exception('frame data not prepared')
    

    frame_result = com_plan.creer_frame_a4(prepared_data=data)
    CM.check_success(frame_result, 'frame_insert')
    return frame_result
    
def step_inserer_legende(source_dwg, source_dxf, frame_width, bbox_data, leg_coords, emitter):
    set_status(emitter, "Insertion de la légende")
    set_progress(emitter, 20)
    prepared_data = dxf_plan.prepare_inserer_legende(source_dxf=source_dxf, frame_width=frame_width)
    CM.check_success(prepared_data, "legend_prep")

    data = prepared_data['data']

    legende_result = com_plan.inserer_legende(prepared_data=data, source_dwg= source_dwg, leg_data=leg_coords)
    print(f'legend result: {legende_result}')
    CM.pump_sleep(2)
    CM.check_success(legende_result, 'inserer legende')
    return legende_result

def step_inserer_tab(source_dxf, frame_width, bbox_data, table_coords, emitter):
    set_status(emitter, "Insertion des coordonnées")
    set_progress(emitter, 25)

    prepared_data = dxf_plan.prepare_inserer_tableau(source_dxf=source_dxf, frame_width=frame_width, bbox_data=bbox_data,table_data=table_coords)
    CM.check_success(prepared_data, 'bornes_prep')

    data = prepared_data['data']

    tab_result = com_plan.inserer_tableau(prepared_data=data)
    CM.pump_sleep(2)
    CM.check_success(tab_result, 'inserer tab')
    return tab_result

def step_inserer_croisillions(bornes_results, frame_width, frame_coords, echelle,emitter):
    set_status(emitter, "Insertion des croisillions")
    set_progress(emitter, 28)

    prepared_data = dxf_plan.prepare_inserer_croisillions(bornes_results=bornes_results, frame_width=frame_width, frame_coords=frame_coords, echelle=echelle)
    CM.check_success(prepared_data, 'croisillons_prep')

    data = prepared_data['data']

    croisillons_result = com_plan.inserer_croisillons(prepared_data=data)
    CM.pump_sleep(2)
    CM.check_success(croisillons_result, 'inserer croisillons')
    return croisillons_result

def step_calculer_insert_pts(bbox_data, frame_coords, frame_width, emitter):
    set_status(emitter, "Calcul des points d’insertion")
    offset = 0.1 * (bbox_data[3] - bbox_data[0])
    orthomap_insert_pt = (frame_coords[3] + offset, frame_coords[1])
    cartouche_insert_pt = (frame_coords[0] - offset - frame_width, frame_coords[1])
    return {
        "success": True,
        "data": {
            "orthomap_insert_pt": orthomap_insert_pt,
            "cartouche_insert_pt": cartouche_insert_pt,
        },
        "message": "ok",
    }
