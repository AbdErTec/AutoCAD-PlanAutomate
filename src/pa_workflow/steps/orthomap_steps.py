from pa_workflow.steps.signals import set_status, set_progress

from infra.com_manager import COMManager as CM

from operations.com import com_block_ops as com_block
from operations.com import com_orthomap_ops as com_om 
from operations.dxf import dxf_orthomap_ops as dxf_om

def step_inserer_block_orthoMap(block_path, insert_pt, frame_width, attr_values, emitter):
    print("step_inserer_block_orthoMap")
    set_status(emitter, "Insertion de l’orthophoto + extrait de carte")
    set_progress(emitter, 30)

    orthomap_result = com_block.inserer_block(block_path=block_path,insert_pt= insert_pt, frame_width=frame_width, attr_values=attr_values)
    CM.check_success(orthomap_result, 'inserer block orthoMap ')
    return orthomap_result

def step_inserer_orthophoto_carte(source_dwg, source_dxf, block_ref, emitter):
    prepared_data = dxf_om.preparer_crop_ortophoto_map(source_dxf=source_dxf) 
    CM.check_success(prepared_data, 'orthomap_prep')
    data = prepared_data['data']
    orthomap_insert_result = com_om.crop_orthophoto_map(prepared_data=data, source_dwg=source_dwg, block_ref=block_ref)
    CM.pump_sleep(1)
    CM.check_success(orthomap_insert_result, 'inserer orthophoto+carte')
    return orthomap_insert_result
