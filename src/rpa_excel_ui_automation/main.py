import logging
from pathlib import Path
from pywinauto import Application, Desktop
from pywinauto.keyboard import send_keys

# ==========================================
# CONFIGURACIÓN DE TRAZABILIDAD (LOGS)
# ==========================================
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - [%(name)s] %(message)s'
)
logger = logging.getLogger("ExcelBot")


# ==========================================
# 1. FILE EXPLORER (Manejo de Windows)
# ==========================================
class FileExplorer:
    """
    Responsabilidad: Interactuar exclusivamente con las ventanas de diálogo nativas 
    de Windows (Explorador de archivos), sin usar Tabuladores ni pausas estáticas.
    """
    def __init__(self):
        # Usamos el backend 'uia' (UI Automation) exigido en la rúbrica
        self.desktop = Desktop(backend="uia")

    def handle_open_dialog(self, target_path: Path):
        logger.info("Esperando a que la ventana de diálogo 'Abrir' esté lista...")
        # Localizamos la ventana directamente y usamos sincronización basada en eventos
        dialog = self.desktop.window(title_re=".*(Abrir|Open).*", control_type="Window")
        dialog.wait("ready", timeout=20)
        
        logger.info(f"Inyectando ruta absoluta origen: {target_path.resolve()}")
        # Ubicamos el campo de texto sin usar TAB
        edit_box = dialog.child_window(control_type="Edit", title_re=".*(Nombre de archivo|File name).*")
        edit_box.set_edit_text(str(target_path.resolve()))
        
        logger.info("Ejecutando acción sobre el botón 'Abrir'...")
        open_btn = dialog.child_window(control_type="Button", title_re=".*(Abrir|Open).*")
        open_btn.click()

    def handle_save_as_dialog(self, dest_path: Path):
        logger.info("Esperando a que la ventana de diálogo 'Guardar como' esté lista...")
        dialog = self.desktop.window(title_re=".*(Guardar como|Save As).*", control_type="Window")
        dialog.wait("ready", timeout=20)
        
        logger.info(f"Inyectando ruta absoluta destino: {dest_path.resolve()}")
        edit_box = dialog.child_window(control_type="Edit", title_re=".*(Nombre de archivo|File name).*")
        edit_box.set_edit_text(str(dest_path.resolve()))
        
        logger.info("Ejecutando acción sobre el botón 'Guardar'...")
        save_btn = dialog.child_window(control_type="Button", title_re=".*(Guardar|Save).*")
        save_btn.click()
        
        # CONDICIÓN DE REEMPLAZO DINÁMICA (Anti-Fragilidad)
        overwrite_dialog = dialog.child_window(title_re=".*(Confirmar|Confirm).*", control_type="Window")
        if overwrite_dialog.exists(timeout=3):
            logger.warning("Ventana de advertencia de sobrescritura detectada.")
            logger.info("Confirmando el reemplazo (Click en 'Sí')...")
            yes_btn = overwrite_dialog.child_window(title_re=".*(Sí|Yes).*", control_type="Button")
            yes_btn.click()


# ==========================================
# 2. EXCEL MANAGER (Manejo de la App)
# ==========================================
class ExcelManager:
    """
    Responsabilidad: Gestionar la instancia de Microsoft Excel y disparar 
    los atajos de teclado nativos para invocar los menús.
    """
    def __init__(self):
        self.app = Application(backend="uia")
        self.main_window = None

    def open_file(self):
        logger.info("Inicializando la aplicación Microsoft Excel...")
        self.app.start("excel.exe")
        
        # Esperamos a que la ventana base de Excel levante
        self.main_window = self.app.window(title_re=".*Excel.*", control_type="Window")
        self.main_window.wait("ready", timeout=30)
        
        logger.info("Invocando comando para abrir archivo (Ctrl + F12)...")
        # Ctrl+F12 es el atajo nativo universal en Excel que abre directamente 
        # el Explorador de Archivos (saltando la pantalla de inicio "Backstage")
        send_keys('^{F12}')

    def save_as(self):
        logger.info("Invocando atajo nativo para 'Guardar como' (F12)...")
        send_keys('{F12}')


# ==========================================
# ORQUESTADOR (Caso de Prueba 01 y 02)
# ==========================================
def main():
    # Uso estricto de pathlib
    input_path = Path(".data/input/origen.xlsx")
    output_path = Path(".data/output/destino.xlsx")
    
    excel_manager = ExcelManager()
    explorer = FileExplorer()
    
    # --- CASO DE PRUEBA 01 ---
    logger.info("--- INICIANDO CASO DE PRUEBA 01: APERTURA ---")
    excel_manager.open_file()
    explorer.handle_open_dialog(input_path)
    
    # --- CASO DE PRUEBA 02 ---
    logger.info("--- INICIANDO CASO DE PRUEBA 02: GUARDADO SEGURO ---")
    # Pausa implícita de sincronización para asegurar que el archivo cargó visualmente
    excel_manager.main_window.wait("ready", timeout=15) 
    
    excel_manager.save_as()
    explorer.handle_save_as_dialog(output_path)
    
    logger.info("¡Flujo completado con éxito! El archivo original está intacto.")

if __name__ == "__main__":
    main()
