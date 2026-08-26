import os
import logging
from pathlib import Path
from pywinauto import Application, Desktop

# ==========================================
# CONFIGURACIÓN DE TRAZABILIDAD (LOGS)
# ==========================================
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - [%(name)s] %(message)s'
)
logger = logging.getLogger("ExcelBot")

# ==========================================
# 1. FILE EXPLORER (Motor Win32 Híbrido)
# ==========================================
class FileExplorer:
    def __init__(self):
        self.desktop = Desktop(backend="win32")

    def handle_open_dialog(self, target_path: Path):
        logger.info("Buscando ventana 'Abrir' con motor Win32...")
        dialog = self.desktop.window(class_name="#32770", title_re=".*(Abrir|Open).*")
        dialog.wait("ready", timeout=20)
        
        logger.info(f"Inyectando ruta origen: {target_path.resolve()}")
        # REGLA UNIVERSAL: Atrapa el primer cuadro de edición de texto de la ventana
        edit_box = dialog.child_window(class_name="Edit", found_index=0)
        edit_box.set_edit_text(str(target_path.resolve()))
        
        logger.info("Clic en botón Abrir...")
        open_btn = dialog.child_window(control_id=1)
        open_btn.click()

    def handle_save_as_dialog(self, dest_path: Path):
        dest_dir = dest_path.resolve().parent
        dest_dir.mkdir(parents=True, exist_ok=True)

        logger.info("Buscando ventana 'Guardar como' con motor Win32...")
        dialog = self.desktop.window(class_name="#32770", title_re=".*(Guardar como|Save As).*")
        dialog.wait("ready", timeout=20)
        
        logger.info(f"Inyectando ruta destino: {dest_path.resolve()}")
        # Aplicamos la misma regla universal para el guardado
        edit_box = dialog.child_window(class_name="Edit", found_index=0)
        edit_box.set_edit_text(str(dest_path.resolve()))
        
        logger.info("Clic en botón Guardar...")
        save_btn = dialog.child_window(control_id=1)
        save_btn.click()
        
        overwrite_dialog = self.desktop.window(class_name="#32770", title_re=".*(Confirmar|Confirm).*")
        if overwrite_dialog.exists(timeout=3):
            logger.warning("Confirmando sobrescritura...")
            yes_btn = overwrite_dialog.child_window(title_re=".*(Sí|Yes).*")
            yes_btn.click()

# ==========================================
# 2. EXCEL MANAGER
# ==========================================
class ExcelManager:
    def __init__(self):
        self.app = Application(backend="uia")
        self.main_window = None

    def open_file(self):
        logger.info("Inicializando Microsoft Excel...")
        os.system("start excel")
        
        self.app.connect(title_re=".*Excel.*", timeout=20)
        self.main_window = self.app.window(title_re=".*Excel.*", control_type="Window")
        self.main_window.wait("ready", timeout=30)
        
        logger.info("Dando foco a Excel y enviando atajo (Ctrl + F12)...")
        self.main_window.set_focus()
        self.main_window.type_keys('^{F12}')

    def save_as(self):
        logger.info("Dando foco a Excel y enviando atajo (F12)...")
        self.main_window.set_focus()
        self.main_window.type_keys('{F12}')

# ==========================================
# ORQUESTADOR
# ==========================================
def main():
    input_path = Path(".data/input/origen.xlsx")
    output_path = Path(".data/output/destino.xlsx")
    
    excel_manager = ExcelManager()
    explorer = FileExplorer()
    
    logger.info("--- INICIANDO CASO DE PRUEBA 01: APERTURA ---")
    excel_manager.open_file()
    explorer.handle_open_dialog(input_path)
    
    logger.info("--- INICIANDO CASO DE PRUEBA 02: GUARDADO SEGURO ---")
    # EL PARCHE: Refrescamos la ventana de Excel porque al abrir el archivo la interfaz cambió.
    # Además, usamos "visible" en lugar de "ready" para que sea más dinámico.
    excel_manager.main_window = excel_manager.app.window(title_re=".*Excel.*", control_type="Window")
    excel_manager.main_window.wait("visible", timeout=15) 
    
    excel_manager.save_as()
    explorer.handle_save_as_dialog(output_path)
    
    logger.info("¡Flujo completado con éxito!")

if __name__ == "__main__":
    main()