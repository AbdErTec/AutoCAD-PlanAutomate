import time
import pythoncom
import win32com.client
import time
import pythoncom

class COMManager:

    @staticmethod
    def get_app_and_doc(retries=5, delay=1):
        for attempt in range(1, retries + 1):
                try:
                    app = win32com.client.GetActiveObject("AutoCAD.Application")
                    doc = COMManager.retry_com_operation(lambda: app.ActiveDocument)
                    if doc: 
                        return app, doc
                except Exception as e:
                    print(f"❌ Attempt {attempt}/{retries} – Cannot connect to AutoCAD: {e}")
                
                if attempt < retries:
                    time.sleep(delay)  

        print(f"❌ Failed to get active AutoCAD document after {retries} attempts")
        return None, None
    @staticmethod
    def is_autocad_responsive(doc):
        try:
            _ = doc.Name  
            return True
        except pythoncom.com_error:
            return False

    @staticmethod
    def wait_for_autocad(doc, timeout=30, interval=2):
        start = time.time()
        while time.time() - start < timeout:
            if COMManager.is_autocad_responsive(doc):
                return True
            print("⏳ AutoCAD not responding, waiting...")
            time.sleep(interval)
        return False

    @staticmethod
    def check_success(result, operation_name = "Operation"):
        if not result.get('success', False):
            error_msg = result.get('user_message', f"{operation_name} a echoue")
            raise Exception(error_msg)

    @staticmethod
    def identify_error(error_msg):

        if not error_msg:
            return "Unknown error occurred"
        
        error_lower = str(error_msg).lower()
        
        # COM Connection Issues (Most Common)
        if any(pattern in error_lower for pattern in [
            "call was rejected by callee",
            "call was cancelled by the message filter"
        ]):
            return "AutoCAD is busy processing another command. Please wait and try again."
        
        if any(pattern in error_lower for pattern in [
            "disconnected from its clients",
            "disconnected from its client"
        ]):
            return "Connection to AutoCAD was lost. Please try again or restart AutoCAD if the problem persists."
        
        if "rpc server is unavailable" in error_lower:
            return "AutoCAD is not responding. Please restart the plugin or AutoCAD if needed."
        
        if any(pattern in error_lower for pattern in [
            "automation object has been deleted",
            "object has been erased"
        ]):
            return "An AutoCAD object was deleted while being used. Please try the operation again."
        
        if any(pattern in error_lower for pattern in [
            "open.activate",
            "failed to activate document"
        ]):
            return "Cannot open or activate the drawing file. Make sure the file exists and is not corrupted."
        
        if any(pattern in error_lower for pattern in [
            "open.close",
            "failed to close document"
        ]):
            return "Error closing the drawing. The file might still be in use or corrupted."
        
        if "application.activedocument" in error_lower:
            return "No drawing is currently open in AutoCAD. Please open a drawing first."
        
        if any(pattern in error_lower for pattern in [
            "sendcommand",
            "send command failed",
            "<unknown>sendcommand"
        ]):
            return "AutoCAD command failed. This might be due to invalid input or AutoCAD being in an unexpected state."
        
        return "An unexpected error occurred. Please restart the plugin or AutoCAD if needed."


    @staticmethod
    def is_com_retryable_error(error_msg):
        retry_patterns = [
            "call was rejected by callee",
            "disconnected from its clients", 
            "rpc server is unavailable",
        ]
        return any(pattern in error_msg.lower() for pattern in retry_patterns)


    @staticmethod
    def retry_com_operation(operation_func, *args, max_retries=10, delay=3, operation_name="COM operation", **kwargs):
        for attempt in range(1, max_retries + 1):
            try:
                result = operation_func(*args, **kwargs)
                COMManager.pump_sleep(2)
                if attempt > 1:
                    print(f"✅ {operation_name} succeeded on attempt {attempt}")

                return result
            except Exception as e:
                error_msg = str(e)
                if COMManager.is_com_retryable_error(error_msg) and attempt < max_retries:
                    print(f"⚠️ {operation_name} failed (attempt {attempt}/{max_retries}): {error_msg}")
                    COMManager.pump_sleep(delay)
                else:
                    raise e
        raise Exception(f"Failed {operation_name} after {max_retries} attempts")


    @staticmethod
    def pump_sleep(seconds, retries=5, delay=2):
        for attempt in range(retries):
            try:
                start_time = time.time()
                while time.time() - start_time < seconds:
                    pythoncom.PumpWaitingMessages()
                    time.sleep(0.5)
                return  # ✅ Success, exit the function
            except Exception as e:
                msg = str(e).lower()
                if 'call was rejected by callee' in msg or 'disconnected from its clients' in msg:
                    print(f"⚠️ COM error on attempt {attempt+1}: {e}")
                    time.sleep(delay)
                else:
                    raise e

        raise RuntimeError("🚨 pump_sleep failed after multiple retries due to COM errors.")

    @staticmethod
    def pumpCommand(doc, command, retries=10, delay=2):
        for attempt in range(1, retries + 1):
            try:
                doc.SendCommand(command)
                COMManager.pump_sleep(2)
                if attempt > 1:
                    print(f"✅ {command} succeeded on attempt {attempt}")
                break

            except Exception as e:
                error_msg = str(e)
                if COMManager.is_com_retryable_error(error_msg) and attempt < retries:
                    print(f"⚠️ {command} failed (attempt {attempt}/{retries}): {error_msg}")
                    COMManager.pump_sleep(delay)
                else:
                    raise e
        else:
            raise Exception(f"Failed {command} after {retries} attempts")