from pa_workflow.steps.signals import set_status, set_progress
from infra.com_manager import COMManager as CM
from operations.geometry import calculer_bbox

from operations.com import com_block_ops as com_block

from operations.com import com_cartouche_ops as com_cartouche


def step_inserer_block_cartouche(block_path, insert_pt, frame_width, attr_values, emitter):
    set_status(emitter, "Insertion de la cartouche")
    set_progress(emitter, 60)

    cartouche_result = com_block.inserer_block(block_path=block_path, insert_pt=insert_pt, frame_width=frame_width, attr_values=attr_values)
    CM.check_success(cartouche_result, 'inserer block cartouche')
    return cartouche_result

def step_inserer_qr(source_path, block_ref, emitter):
    set_progress(emitter, 80)
    qr_result = com_cartouche.inserer_qr(source_path=source_path,block_ref= block_ref)
    CM.pump_sleep(2)
    CM.check_success(qr_result, 'inserer qr ')
    return qr_result
