import os
from pa_workflow.steps.plan_steps import *
from pa_workflow.steps.orthomap_steps import *
from pa_workflow.steps.cartouche_steps import *
from infra.converter import Converter

# from pa_workflow.steps.dummy_steps import *
from handbooks.handbook import Handbook
from infra.com_manager import COMManager as CM
from infra.helpers import resource_path

# ORTHOMAP_BLOCK_PATH = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..', 'assets', 'blocks', 'OrthoMap.dwg'))
# CARTOUCHE_BLOCK_PATH = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), '..','..', 'assets','blocks', 'cartouche.dwg'))
ORTHOMAP_BLOCK_PATH = resource_path("assets", "blocks", "OrthoMap.dwg")
CARTOUCHE_BLOCK_PATH = resource_path("assets", "blocks", "cartouche.dwg")

class PlanHandbook(Handbook):
    def __init__(self, doc, ctx, file_paths, selected_layers, selected_echelle):
        super().__init__(doc, ctx)
        self.file_paths = file_paths
        self.selected_layers = selected_layers
        self.selected_echelle = selected_echelle
        con = Converter()
        self.dxf_plan_path = con.dwg_to_dxf(self.file_paths[0])  
        self.dxf_legende_path = con.dwg_to_dxf(self.file_paths[1])  
        self.handbook = [

            ('calcul_bbox_plan',  step_calculer_bbox_plan, {'file_path': self.dxf_plan_path}),
            
            ('inserer_plan',  step_inserer_plan, {'source_dwg':self.file_paths[0], 'source_dxf': self.dxf_plan_path, 'bbox_data': '__bbox_data__', 'layers': self.selected_layers}),
            
            ('creer_frame_a4',  step_creer_frame_a4, {'bbox_data': '__bbox_data__', 'echelle': self.selected_echelle}),
            
            ('inserer_legende',  step_inserer_legende, {'source_dwg':self.file_paths[1],'source_dxf': self.dxf_legende_path, 'bbox_data': '__bbox_data__', 'frame_width':'__frame_width__', 'leg_coords': '__leg_coords__'}),
            
            ('inserer_tab',  step_inserer_tab, {'source_dxf': self.dxf_plan_path, 'bbox_data': '__bbox_data__', 'frame_width':'__frame_width__', 'table_coords': '__table_coords__'}),
        
            ('inserer_croisillions',  step_inserer_croisillions, {'bornes_results': '__bornes_results__', 'frame_width': '__frame_width__', 'frame_coords': '__frame_coords__', 'echelle': self.selected_echelle}),
            
            ('calculer_insert_pts',  step_calculer_insert_pts, {'bbox_data': '__bbox_data__', 'frame_coords': '__frame_coords__', 'frame_width': '__frame_width__'}),
        ]

    def execute_handbook(self):
        print("="*25,'Executing plan Handbook','='*25)
        super().execute_handbook()
        print("="*50)
        
    
    CTX_MAP = {
        "calcul_bbox_plan": {
            "bbox_data": "data",
        },
        
        "creer_frame_a4": {
            "leg_coords":   "data.leg_coords",
            "table_coords": "data.table_coords",
            "frame_width":  "data.frame_width",
            "frame_height": "data.frame_height",
            "frame_coords": "data.frame_coords",
        },
        
        "inserer_tab": {
            "bornes_results": "data.bornes_results",
        },
        
        "calculer_insert_pts": {
            "orthomap_insert_pt": "data.orthomap_insert_pt",   # top-level in result
            "cartouche_insert_pt": "data.cartouche_insert_pt",
        },
    }
    
    
    def clean(self):
        try:
            CM.pumpCommand(self.doc, '_ZOOM\nE\n\n')
            CM.pump_sleep(0.5)
            print("Cleanup: zoomed to extents")
        except Exception as e:
            print(f"Cleanup failed: {e}")


class OrthoMapHandbook(Handbook):
    def __init__(self, doc, ctx, file_paths, client_data):
        super().__init__(doc, ctx)
        self.file_paths = file_paths
        self.client_data = client_data
        con = Converter()
        self.dxf_input = con.dwg_to_dxf(self.file_paths[2])
        self.handbook = [

            ('inserer_block_orthoMap',  step_inserer_block_orthoMap, {'block_path': ORTHOMAP_BLOCK_PATH, 'insert_pt': '__orthomap_insert_pt__', 'frame_width': '__frame_width__', 'attr_values': {
                    "APERCU_SUR_FOND_HAUT": self.client_data['apercu_fond_haut'],
                    "APERCU_SUR_FOND_BAS": self.client_data['apercu_fond_bas']
                }}),
            
            ('inserer_orthophoto+carte',  step_inserer_orthophoto_carte, {'source_dwg':self.file_paths[2],'source_dxf': self.dxf_input, 'block_ref': '__orthomap_ref__'}),
        ]

    def execute_handbook(self):
        print("="*25,'Executing orthomap Handbook','='*25)
        super().execute_handbook()
        print("="*50)


    CTX_MAP = {
        "inserer_block_orthoMap": {
            "orthomap_ref": "data", 
        },
    }

    def _convert_container_to_block(self,destination_doc, block_data):

        if not block_data or not block_data.get('handle'):
            return None

        try:
            return CM.retry_com_operation(destination_doc.HandleToObject, block_data['handle'])

        except Exception as e:
            print(f"Error deserializing block ref: {e}")
            return None

    def restore_block_refs_from_context(self, destination_doc):
        if 'orthomap_ref' in self.ctx and isinstance(self.ctx['orthomap_ref'], dict):
            self.ctx['orthomap_ref'] = self._convert_container_to_block(destination_doc, self.ctx['orthomap_ref'])

        if 'cartouche_ref' in self.ctx and isinstance(self.ctx['cartouche_ref'], dict):
            self.ctx['cartouche_ref'] = self._convert_container_to_block(destination_doc, self.ctx['cartouche_ref'])

        return self.ctx


    def clean(self):
        try:
            CM.pumpCommand(self.doc, '_ZOOM\nE\n\n')
            CM.pump_sleep(0.5)
            print("Cleanup: zoomed to extents")
        except Exception as e:
            print(f"Cleanup failed: {e}")


class CartoucheHandbook(Handbook):
    def __init__(self, doc, ctx, file_paths, client_data, images_paths):
        super().__init__(doc, ctx)
        self.file_paths = file_paths
        self.client_data = client_data
        self.images_paths = images_paths
  
        values_cartouche = {
            "REGION": self.client_data['region'],
            'REGION_AR': self.client_data['region_ar'],
            'PROVINCE/PREFECTURE': self.client_data['province'],
            'PROVINCE/PREFECTURE_AR': self.client_data['province_ar'],
            'COMMUNE': self.client_data['commune'],
            'COMMUNE_AR': self.client_data['commune_ar'],
            'SITUATION': self.client_data['situation'],
            'SITUATION_AR': self.client_data['situation_ar'],
            'PLAN': self.client_data['plan'],
            'CONTENANCE': self.client_data['contenance'],
            'DEMANDE_PAR': self.client_data['demande_par'],
            'PROPRIETE_DTE': self.client_data['propriete_dte'],
            'REFERENCE_FONCIERE': self.client_data['reference_fonciere'],
            'OBSERVATIONS': self.client_data['observations'],
            'DECLARATION_DE': self.client_data['declaration_par'],
            'CIN': self.client_data['cin'],
            'LEVE_LE': self.client_data['leve_le'],
            'AGENT_LEVEUR': self.client_data['agent_leveur'],
            'ECHELLE': self.client_data['echelle'],
            'DATE': self.client_data['date'],
            'NUMERO_DOSSIER': self.client_data['numero_dossier'],
            'NIVELLEMENT': self.client_data['nivellement'],
            'COODONNEES': self.client_data['coordonnees'],
            'FICHIER': self.client_data['fichier'],
            'XREF': self.client_data['xref']
        }
     
        self.handbook = [
            ('inserer_block_cartouche',  step_inserer_block_cartouche, {'block_path': CARTOUCHE_BLOCK_PATH, 'insert_pt': '__cartouche_insert_pt__', 'frame_width': '__frame_width__','attr_values': values_cartouche}),
            
            ('inserer_qr',  step_inserer_qr, {'source_path': self.images_paths[1], 'block_ref': '__cartouche_ref__'})
        ]

    def execute_handbook(self):
        print("="*25,'Executing cartouche Handbook','='*25)
        super().execute_handbook()
        print("="*50)

    CTX_MAP = {
        "inserer_block_cartouche": {"cartouche_ref": "data"},
        "inserer_qr":              {"qr_ref": "data"},
    }


    def clean(self):
        try:
            CM.pumpCommand(self.doc, '_ZOOM\nE\n\n')
            CM.pump_sleep(0.5)
            print("Cleanup: zoomed to extents")
        except Exception as e:
            print(f"Cleanup failed: {e}")
            