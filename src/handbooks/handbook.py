from abc import abstractmethod, ABC
from infra.com_manager import COMManager as CM
from pathlib import Path
import os
import inspect
import ezdxf
import win32com.client as win32
from infra.serializers import to_container

import traceback

class Handbook(ABC):
    def __init__(self, doc, ctx):
        self.doc = doc or None
        self.ctx = ctx or {}
        self.handbook = []

    def execute_handbook(self):
        for step_name, func, params in self.handbook:
            args = self._resolve_args(params)

            try:
                if 'emitter' in inspect.signature(func).parameters:
                    args['emitter'] = getattr(self, 'emitter', None)
            except Exception:
                pass

            print(f"🛠️ args for {step_name}: {args}")

            try:
                args["emitter"] = getattr(self, "emitter", None)   
                result = func(**args) 
                if not result.get("success", True):
                    raise Exception(result.get("user_message") or result.get("error") or "Step failed")

                print(f"📫 result from {step_name}: {result}")
                self._update_context(step_name, result)
                print(f"✅ Step '{step_name}' completed")

            except Exception as e:
                print(f"❌ Step '{step_name}' failed: {e}")
                traceback.print_exc()
                raise

        print("=" * 50)    

    def _unwrap_result(self, val):
        if isinstance(val, dict) and "success" in val:
            if not val.get("success", True):
                raise Exception(val.get("user_message") or val.get("error") or "Upstream step failed")
            return val.get("data")
        if isinstance(val, Path):
            return str(val)
        return val

    def _resolve_args(self, template_args):
        resolved = {}
        for k, v in template_args.items():
            if isinstance(v, str) and v.startswith("__") and v.endswith("__"):
                key = v.strip("_")
                if key in self.ctx:
                    resolved[k] = self._unwrap_result(self.ctx[key])
                else:
                    have = ", ".join(sorted(self.ctx.keys()))
                    raise Exception(
                        f"Argument manquant pour l'étape: {k} (needs ctx['{key}']). "
                        f"Current ctx keys: [{have}]"
                    )
            else:
                resolved[k] = self._unwrap_result(v)
        return resolved

    def _update_context(self, step_name, result):
        mapping = getattr(self, "CTX_MAP", {}).get(step_name)
        if not mapping:
            return
        for ctx_key, path in mapping.items():
            val = self._get_by_path(result, path)
            self.ctx[ctx_key] = val

    def _get_by_path(self, obj, path):
        cur = obj
        for part in path.split("."):
            if isinstance(cur, dict) and part in cur:
                cur = cur[part]
            else:
                return None
        return cur
    
    def clean(self):
        pass

    def has_next(self):
        return self.current_step_index < len(self.steps_map)