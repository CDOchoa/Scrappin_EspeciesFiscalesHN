# app.py - Versión Streamlit (Adaptada desde Tkinter)

import sys
import threading
import pandas as pd
import os
import streamlit as st
import traceback
from datetime import datetime
import time

# Asegúrate de importar SARValidator y cargar_api_keys_remotas_seguras desde core_processor
try:
    # Intentar primero con la versión de Streamlit
    from core_processor_streamlit import SARValidator, cargar_api_keys_remotas_seguras, API_KEYS
except ImportError:
    try:
        # Fallback a la versión original
        from core_processor import SARValidator, cargar_api_keys_remotas_seguras, API_KEYS
        st.warning("⚠️ Usando versión original de core_processor. Se recomienda usar core_processor_streamlit.py")
    except ImportError as e:
        st.error(f"Error crítico al importar core_processor: {e}")
        st.stop()


# ----------------------------------------------------------------------
# CONFIGURACIÓN INICIAL DE STREAMLIT
# ----------------------------------------------------------------------
st.set_page_config(
    page_title="SAR Document Validator",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ----------------------------------------------------------------------
# INICIALIZACIÓN DE SESSION STATE
# ----------------------------------------------------------------------
if 'df' not in st.session_state:
    st.session_state.df = None
if 'processor' not in st.session_state:
    st.session_state.processor = None
if 'processing_thread' not in st.session_state:
    st.session_state.processing_thread = None
if 'is_running' not in st.session_state:
    st.session_state.is_running = False
if 'stop_requested' not in st.session_state:
    st.session_state.stop_requested = threading.Event()
if 'total_rows' not in st.session_state:
    st.session_state.total_rows = 0
if 'completed_count' not in st.session_state:
    st.session_state.completed_count = 0
if 'failed_count' not in st.session_state:
    st.session_state.failed_count = 0
if 'pending_count' not in st.session_state:
    st.session_state.pending_count = 0
if 'log_messages' not in st.session_state:
    st.session_state.log_messages = []
if 'progress_value' not in st.session_state:
    st.session_state.progress_value = 0
if 'api_keys_loaded' not in st.session_state:
    st.session_state.api_keys_loaded = False
if 'api_keys_count' not in st.session_state:
    st.session_state.api_keys_count = 0


# ----------------------------------------------------------------------
# FUNCIONES AUXILIARES
# ----------------------------------------------------------------------

def log_message(message):
    """Agrega un mensaje al log de la aplicación."""
    timestamp = datetime.now().strftime("%H:%M:%S")
    st.session_state.log_messages.append(f"[{timestamp}] {message}")


def update_status_counts(total, completed, failed):
    """Actualiza las variables de estado."""
    st.session_state.total_rows = total
    st.session_state.completed_count = completed
    st.session_state.failed_count = failed
    st.session_state.pending_count = total - completed - failed


def update_progress(current_index, total, status_message, detail_status):
    """Callback para actualizar el progreso desde el hilo secundario."""
    if total == 0:
        return
    
    progress_val = int(((current_index + 1) / total) * 100)
    st.session_state.progress_value = progress_val
    
    log_message(f"Fila {current_index + 1}/{total} - {status_message}")
    
    # Actualizar contadores si la fila ya ha terminado de procesar
    if detail_status in ["Éxito", "Fallido", "Error"]:
        completed = st.session_state.completed_count
        failed = st.session_state.failed_count
        
        if detail_status == "Éxito":
            completed += 1
        else:  # Fallido o Error
            failed += 1
        
        update_status_counts(st.session_state.total_rows, completed, failed)


def check_api_keys():
    """Comprueba y carga las API Keys."""
    try:
        if cargar_api_keys_remotas_seguras():
            count = len(API_KEYS)
            st.session_state.api_keys_loaded = True
            st.session_state.api_keys_count = count
            return True
        else:
            st.session_state.api_keys_loaded = False
            st.session_state.api_keys_count = 0
            return False
    except Exception as e:
        st.session_state.api_keys_loaded = False
        st.session_state.api_keys_count = 0
        st.error(f"❌ ERROR CRÍTICO al cargar Keys: {e}")
        return False


def run_processing(excel_path, output_path, mode, headless):
    """Método que se ejecuta en el hilo secundario para el procesamiento."""
    try:
        log_message("Iniciando navegador Selenium...")
        
        st.session_state.processor = SARValidator(output_path, mode, headless)
        
        if not st.session_state.processor.initialize_driver():
            raise Exception("Fallo al inicializar el navegador.")
        
        log_message(f"Procesando {len(st.session_state.df)} registros. Modo: {mode}")
        
        # El core_processor necesita la columna 'Estado_Proceso' para saber qué guardar
        df_result = st.session_state.processor.procesar_dataframe(
            st.session_state.df, 
            update_progress
        )
        
        # Actualizar el DataFrame con los resultados del core
        if isinstance(df_result, pd.DataFrame):
            st.session_state.df = df_result
        
        log_message("Proceso principal finalizado.")

    except Exception as e:
        log_message(f"🔴 ERROR CRÍTICO: {e}")
        log_message(traceback.format_exc())
        st.error(f"🔴 ERROR CRÍTICO en el proceso: {e}")
    
    finally:
        if st.session_state.processor:
            st.session_state.processor.close_driver()
        
        reset_ui_after_completion()


def reset_ui_after_completion():
    """Restablece el estado después de que el hilo termina."""
    st.session_state.is_running = False
    
    # Comprobar estado final para actualizar el progreso al 100%
    if st.session_state.df is not None:
        total = len(st.session_state.df)
        completed = len(st.session_state.df[
            st.session_state.df['Estado_Proceso'].str.contains('Válido|Data OK|PDF OK', na=False, case=False)
        ])
        failed = len(st.session_state.df[
            st.session_state.df['Estado_Proceso'].str.contains('Error|Fallido', na=False, case=False)
        ])
        update_status_counts(total, completed, failed)
        st.session_state.progress_value = 100
        log_message("✅ PROCESO COMPLETADO (o detenido) Y RECURSOS LIBERADOS.")


def start_processing(excel_path, output_path, mode, headless):
    """Prepara e inicia el procesamiento en un hilo separado."""
    if not os.path.exists(excel_path):
        st.error("Por favor, seleccione un archivo Excel válido.")
        return
    if not os.path.exists(output_path):
        st.error("Por favor, seleccione una carpeta de salida válida.")
        return

    # Configurar estado
    st.session_state.is_running = True
    st.session_state.stop_requested.clear()
    update_status_counts(0, 0, 0)
    st.session_state.log_messages = []
    st.session_state.progress_value = 0

    # Cargar DataFrame inicial
    try:
        st.session_state.df = pd.read_excel(excel_path)
        st.session_state.total_rows = len(st.session_state.df)
        # Inicializar columna de estado para seguimiento de pendientes
        st.session_state.df['Estado_Proceso'] = 'Pendiente'
        update_status_counts(st.session_state.total_rows, 0, 0)
    except Exception as e:
        st.error(f"Fallo al cargar el archivo Excel: {e}")
        st.session_state.is_running = False
        return
    
    # Inicializar y empezar el hilo
    st.session_state.processing_thread = threading.Thread(
        target=run_processing,
        args=(excel_path, output_path, mode, headless)
    )
    st.session_state.processing_thread.start()


def stop_processing():
    """Solicita detener el proceso y espera a que el hilo termine."""
    if not st.session_state.is_running:
        return
    if not st.session_state.processing_thread:
        return
    if not st.session_state.processing_thread.is_alive():
        return
    
    log_message("🛑 Solicitud de detención recibida. Esperando a que termine el registro actual...")
    st.session_state.stop_requested.set()


# ----------------------------------------------------------------------
# INTERFAZ PRINCIPAL DE STREAMLIT
# ----------------------------------------------------------------------

# Título principal
st.title("📄✅ SAR Document Validator")

# Verificar API Keys al inicio
if not st.session_state.api_keys_loaded:
    check_api_keys()

# ----------------------------------------------------------------------
# BARRA LATERAL - CONFIGURACIÓN
# ----------------------------------------------------------------------
with st.sidebar:
    st.header("🛠️ Configuración")
    
    # Estado de API Keys
    st.subheader("🔑 Estado de API Keys")
    if st.session_state.api_keys_loaded:
        st.success(f"✅ API Keys cargadas: {st.session_state.api_keys_count} activas.")
    else:
        st.error("❌ ERROR: Fallo al cargar API Keys. Verifique 'client_secrets.json' y permisos.")
    
    if st.button("🔄 Recargar API Keys"):
        check_api_keys()
        st.rerun()
    
    st.divider()
    
    # Archivo Excel
    st.subheader("📁 Archivo Excel (Entrada)")
    uploaded_file = st.file_uploader(
        "Seleccionar archivo Excel",
        type=['xlsx', 'xls'],
        disabled=st.session_state.is_running
    )
    
    excel_path = None
    if uploaded_file is not None:
        # Guardar el archivo temporalmente
        temp_dir = "temp_uploads"
        os.makedirs(temp_dir, exist_ok=True)
        excel_path = os.path.join(temp_dir, uploaded_file.name)
        with open(excel_path, "wb") as f:
            f.write(uploaded_file.getbuffer())
        st.success(f"Archivo cargado: {uploaded_file.name}")
    
    st.divider()
    
    # Carpeta de salida
    st.subheader("📂 Carpeta de Resultados")
    output_path = st.text_input(
        "Ruta de carpeta de salida",
        value="./resultados",
        disabled=st.session_state.is_running
    )
    
    if st.button("📁 Crear carpeta si no existe", disabled=st.session_state.is_running):
        try:
            os.makedirs(output_path, exist_ok=True)
            st.success(f"Carpeta lista: {output_path}")
        except Exception as e:
            st.error(f"Error al crear carpeta: {e}")
    
    st.divider()
    
    # Modo de salida
    st.subheader("⚙️ Modo de Salida")
    mode = st.radio(
        "Seleccionar modo",
        ["EXCEL_DATA", "PDF"],
        format_func=lambda x: "Solo Excel (Datos)" if x == "EXCEL_DATA" else "Excel + PDF (Captura)",
        disabled=st.session_state.is_running
    )
    
    # Modo headless
    headless = st.checkbox(
        "Modo Invisible (Recomendado)",
        value=True,
        disabled=st.session_state.is_running
    )

# ----------------------------------------------------------------------
# ÁREA PRINCIPAL - CONTROL Y PROGRESO
# ----------------------------------------------------------------------

col1, col2, col3 = st.columns([2, 2, 3])

with col1:
    if st.button(
        "🚀 Ejecutar Proceso" if not st.session_state.is_running else "⏸️ EN EJECUCIÓN...",
        disabled=st.session_state.is_running or not st.session_state.api_keys_loaded,
        type="primary",
        use_container_width=True
    ):
        if excel_path:
            start_processing(excel_path, output_path, mode, headless)
            st.rerun()
        else:
            st.error("Por favor, cargue un archivo Excel primero.")

with col2:
    if st.button(
        "🛑 Detener Proceso",
        disabled=not st.session_state.is_running,
        use_container_width=True
    ):
        stop_processing()

with col3:
    # Botón de descarga de pendientes/errores
    if st.session_state.df is not None:
        df_errors = st.session_state.df[
            st.session_state.df['Estado_Proceso'].str.contains(
                'Pendiente|Error|Fallido|Incierto', 
                na=False, 
                case=False
            )
        ].copy()
        
        if not df_errors.empty:
            # Preparar DataFrame para descarga
            original_cols = [c for c in st.session_state.df.columns 
                           if c not in ['Estado_Proceso', 'Detalle_Validacion', 
                                      'original_index', 'RTN_EXTRAIDO', 
                                      'NUM_DOCUMENTO_BUSQUEDA', 'FECHA_DOCUMENTO_BUSQUEDA']]
            export_cols = [c for c in original_cols if c in df_errors.columns] + \
                         ['Estado_Proceso', 'Detalle_Validacion']
            df_export = df_errors[export_cols]
            
            # Convertir a Excel para descarga
            from io import BytesIO
            buffer = BytesIO()
            df_export.to_excel(buffer, index=False)
            buffer.seek(0)
            
            st.download_button(
                label="⬇️ Descargar Pendientes/Errores",
                data=buffer,
                file_name=f"SAR_Pendientes_Errores_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

st.divider()

# ----------------------------------------------------------------------
# BARRA DE PROGRESO Y CONTADORES
# ----------------------------------------------------------------------

# Barra de progreso
st.progress(st.session_state.progress_value / 100)

# Contadores
col1, col2, col3, col4 = st.columns(4)
with col1:
    st.metric("Total", st.session_state.total_rows)
with col2:
    st.metric("Pendientes", st.session_state.pending_count)
with col3:
    st.metric("Completados", st.session_state.completed_count)
with col4:
    st.metric("Fallidos", st.session_state.failed_count)

st.divider()

# ----------------------------------------------------------------------
# LOGS
# ----------------------------------------------------------------------

st.subheader("📋 Últimas Acciones")

# Contenedor para logs con scroll
log_container = st.container(height=300)
with log_container:
    if st.session_state.log_messages:
        # Mostrar los últimos 50 mensajes
        for message in st.session_state.log_messages[-50:]:
            st.text(message)
    else:
        st.info("No hay mensajes de log aún.")

# ----------------------------------------------------------------------
# AUTO-REFRESH MIENTRAS ESTÁ EN EJECUCIÓN
# ----------------------------------------------------------------------

if st.session_state.is_running:
    time.sleep(2)  # Esperar 2 segundos antes de refrescar
    st.rerun()


# ----------------------------------------------------------------------
# PUNTO DE ENTRADA PRINCIPAL
# ----------------------------------------------------------------------
if __name__ == '__main__':
    # Modificación crítica para evitar múltiples ventanas en Windows
    if sys.platform.startswith('win'):
        import multiprocessing
        multiprocessing.freeze_support()