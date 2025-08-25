
def _emit(emitter, signal_name, *args):
    """Safely emit a Qt signal if emitter exists."""
    if emitter is None:
        return
    sig = getattr(emitter, signal_name, None)
    if sig is not None:
        sig.emit(*args)

def set_status(emitter, message: str):
    _emit(emitter, "status_updated", message)

def set_progress(emitter, value: int):
    _emit(emitter, "progress_updated", value)
