from pathlib import Path
from infra.com_manager import COMManager as CM

def serialize_block_ref(block_ref):
    if block_ref is None:
        return None
    try:
        return {
            "handle":    CM.retry_com_operation(lambda: block_ref.Handle),
            "name":      CM.retry_com_operation(lambda: getattr(block_ref, "Name", "")),
            "insert_pt": list(CM.retry_com_operation(lambda: block_ref.InsertionPoint)),
            "rotation":  CM.retry_com_operation(lambda: block_ref.Rotation),
            "x_scale":   CM.retry_com_operation(lambda: block_ref.XScaleFactor),
            "y_scale":   CM.retry_com_operation(lambda: block_ref.YScaleFactor),
            "z_scale":   CM.retry_com_operation(lambda: block_ref.ZScaleFactor),
            "object_id": str(CM.retry_com_operation(lambda: getattr(block_ref, "ObjectID", ""))),
        }
    except Exception:
        try:
            return {"handle": CM.retry_com_operation(lambda: block_ref.Handle)}
        except Exception:
            return {"handle": None}

def to_container(obj):
    if hasattr(obj, "tolist"):
        return obj.tolist()

    if isinstance(obj, Path):
        return str(obj)

    if isinstance(obj, dict):
        return {k: to_container(v) for k, v in obj.items()}

    if isinstance(obj, (list, tuple, set)):
        return [to_container(v) for v in obj]

    if hasattr(obj, "Handle"):
        return serialize_block_ref(obj)

    return obj
