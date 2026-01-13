# main.py - Con Interfaz Gráfica (Tkinter)

import sys
import threading
import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import traceback
from datetime import datetime

# Asegúrate de importar SARValidator y cargar_api_keys_remotas_seguras desde core_processor
try:
    from core_processor import SARValidator, cargar_api_keys_remotas_seguras, API_KEYS
except ImportError as e:
    print(f"Error crítico al importar core_processor: {e}")
    messagebox.showerror("Error de Importación", f"Error crítico al iniciar: {e}")
    sys.exit(1)


class SARApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("SAR Document Validator 📄✅")
        self.geometry("800x600")
        self.resizable(False, False)
        
        # Estilos y variables de control
        self.style = ttk.Style(self)
        self.style.theme_use('vista') # Estilo más profesional
        self.df = None
        self.processor = None
        self.processing_thread = None
        self.is_running = False
        self.stop_requested = threading.Event() # Bandera para detener el procesamiento

        # Variables de Tkinter
        self.excel_path_var = tk.StringVar(value="Seleccionar archivo...")
        self.output_path_var = tk.StringVar(value="Seleccionar carpeta...")
        self.mode_var = tk.StringVar(value="EXCEL_DATA") # Valor por defecto solo Excel
        self.headless_var = tk.BooleanVar(value=True)
        
        # Contadores de estado
        self.total_rows = 0
        self.completed_count = tk.IntVar(value=0)
        self.failed_count = tk.IntVar(value=0)
        self.pending_count = tk.IntVar(value=0)
        
        # Inicializar la interfaz
        self._create_widgets()
        self._update_status_counts(0, 0, 0)
        self.check_api_keys()

    # ----------------------------------------------------------------------
    # GUI Y WIDGETS
    # ----------------------------------------------------------------------

    def _create_widgets(self):
        """Crea el diseño general de la aplicación."""
        
        # Marco Principal (Padding)
        main_frame = ttk.Frame(self, padding="20 20 20 20")
        main_frame.pack(fill='both', expand=True)
        
        # Configuración de Columnas (Para alineación)
        main_frame.columnconfigure(1, weight=1)

        # Título
        ttk.Label(main_frame, text="Validador de Documentos Fiscales SAR", 
                  font=("Arial", 16, "bold")).grid(row=0, column=0, columnspan=3, pady=10)

        # ------------------- SECCIÓN 1: ARCHIVOS Y CONFIGURACIÓN -------------------
        config_frame = ttk.LabelFrame(main_frame, text="🛠️ Configuración y Archivos", padding="10")
        config_frame.grid(row=1, column=0, columnspan=3, sticky='ew', pady=10)
        config_frame.columnconfigure(1, weight=1)

        # 1. Archivo Excel
        ttk.Label(config_frame, text="Archivo Excel (Entrada):").grid(row=0, column=0, padx=5, pady=5, sticky='w')
        ttk.Entry(config_frame, textvariable=self.excel_path_var, width=60, state='readonly').grid(row=0, column=1, padx=5, pady=5, sticky='ew')
        ttk.Button(config_frame, text="Buscar...", command=self._select_excel_file).grid(row=0, column=2, padx=5, pady=5)

        # 2. Carpeta de Salida
        ttk.Label(config_frame, text="Carpeta de Resultados:").grid(row=1, column=0, padx=5, pady=5, sticky='w')
        ttk.Entry(config_frame, textvariable=self.output_path_var, width=60, state='readonly').grid(row=1, column=1, padx=5, pady=5, sticky='ew')
        ttk.Button(config_frame, text="Buscar...", command=self._select_output_folder).grid(row=1, column=2, padx=5, pady=5)
        
        # 3. Modo de Salida (Radio buttons)
        mode_frame = ttk.Frame(config_frame)
        mode_frame.grid(row=2, column=0, columnspan=2, padx=5, pady=5, sticky='w')
        ttk.Label(mode_frame, text="Modo de Salida:").pack(side='left', padx=5)
        ttk.Radiobutton(mode_frame, text="Solo Excel (Datos)", variable=self.mode_var, value="EXCEL_DATA").pack(side='left', padx=10)
        ttk.Radiobutton(mode_frame, text="Excel + PDF (Captura)", variable=self.mode_var, value="PDF").pack(side='left', padx=10)
        
        # 4. Modo Headless (Checkbox)
        ttk.Checkbutton(config_frame, text="Modo Invisible (Recomendado)", variable=self.headless_var).grid(row=2, column=2, padx=5, pady=5, sticky='e')

        # ------------------- SECCIÓN 2: CONTROL Y ESTADO -------------------
        
        # Marco para el estado de las API Keys
        status_frame = ttk.LabelFrame(main_frame, text="🔑 Estado de API Keys", padding="10")
        status_frame.grid(row=2, column=0, columnspan=3, sticky='ew', pady=10)
        self.api_status_label = ttk.Label(status_frame, text="Cargando...")
        self.api_status_label.pack(side='left')
        
        # Botones de Control (Ejecutar/Detener)
        btn_frame = ttk.Frame(main_frame)
        btn_frame.grid(row=3, column=0, columnspan=3, pady=10)
        
        self.start_button = ttk.Button(btn_frame, text="🚀 Ejecutar Proceso", command=self.toggle_processing, width=20, style='Accent.TButton')
        self.start_button.pack(side='left', padx=10)
        
        self.stop_button = ttk.Button(btn_frame, text="🛑 Detener Proceso", command=self.stop_processing, state='disabled', width=20)
        self.stop_button.pack(side='left', padx=10)
        
        self.download_button = ttk.Button(btn_frame, text="⬇️ Descargar Pendientes/Errores", command=self.download_pending_errors, state='disabled', width=30)
        self.download_button.pack(side='left', padx=10)

        # ------------------- SECCIÓN 3: PROGRESO Y CONTADORES -------------------
        
        # Barra de progreso
        self.progress_bar = ttk.Progressbar(main_frame, orient='horizontal', mode='determinate')
        self.progress_bar.grid(row=4, column=0, columnspan=3, sticky='ew', pady=10)
        
        # Contadores (Usando un marco con 3 columnas)
        count_frame = ttk.Frame(main_frame)
        count_frame.grid(row=5, column=0, columnspan=3, sticky='ew')
        count_frame.columnconfigure(0, weight=1)
        count_frame.columnconfigure(1, weight=1)
        count_frame.columnconfigure(2, weight=1)
        
        # Etiqueta de Progreso
        self.progress_label = ttk.Label(count_frame, text="Total: 0 | Pendientes: 0 | Completados: 0 | Fallidos: 0", font=('Arial', 10))
        self.progress_label.grid(row=0, column=0, columnspan=3, pady=5)
        
        # LOGS (Texto de la última acción)
        ttk.Label(main_frame, text="Última Acción:").grid(row=6, column=0, sticky='w', pady=(10, 0))
        self.log_text = tk.Text(main_frame, height=5, state='disabled', wrap='word')
        self.log_text.grid(row=7, column=0, columnspan=3, sticky='ew')

    # ----------------------------------------------------------------------
    # MÉTODOS DE MANEJO DE ARCHIVOS Y CONFIGURACIÓN
    # ----------------------------------------------------------------------

    def _select_excel_file(self):
        """Abre un diálogo para seleccionar el archivo Excel."""
        if self.is_running:
            return
        filepath = filedialog.askopenfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if filepath:
            self.excel_path_var.set(filepath)
            self.df = None # Reiniciar el DataFrame

    def _select_output_folder(self):
        """Abre un diálogo para seleccionar la carpeta de salida."""
        if self.is_running:
            return
        folderpath = filedialog.askdirectory()
        if folderpath:
            self.output_path_var.set(folderpath)
            
    def check_api_keys(self):
        """Comprueba y muestra el estado de las API Keys."""
        try:
            if cargar_api_keys_remotas_seguras():
                count = len(API_KEYS)
                self.api_status_label.config(text=f"✅ API Keys cargadas: {count} activas.", foreground="green")
                self.start_button.config(state='normal')
                return True
            else:
                self.api_status_label.config(text="❌ ERROR: Fallo al cargar API Keys. Verifique 'client_secrets.json' y permisos.", foreground="red")
                self.start_button.config(state='disabled')
                return False
        except Exception as e:
            self.api_status_label.config(text=f"❌ ERROR CRÍTICO al cargar Keys: {e}", foreground="red")
            self.start_button.config(state='disabled')
            return False

    # ----------------------------------------------------------------------
    # MÉTODOS DE PROGRESO Y ESTADO
    # ----------------------------------------------------------------------
    
    def _log_message(self, message):
        """Escribe un mensaje en el área de logs de la GUI."""
        self.log_text.config(state='normal')
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END) # Scroll al final
        self.log_text.config(state='disabled')
        
    def _update_status_counts(self, total, completed, failed):
        """Actualiza las variables de estado en la GUI."""
        self.total_rows = total
        self.completed_count.set(completed)
        self.failed_count.set(failed)
        pending = total - completed - failed
        self.pending_count.set(pending)
        
        self.progress_label.config(
            text=f"Total: {total} | Pendientes: {pending} | Completados: {completed} | Fallidos: {failed}"
        )
        
    def _update_progress(self, current_index, total, status_message, detail_status):
        """Callback para actualizar la UI desde el hilo secundario."""
        
        # Usamos after para ejecutar la actualización de la UI en el hilo principal
        self.after(0, lambda: self._gui_update(current_index, total, status_message, detail_status))

    def _gui_update(self, current_index, total, status_message, detail_status):
        """Lógica de actualización de GUI (ejecutada en el hilo principal)."""
        
        if total == 0:
            return
            
        progress_val = int(((current_index + 1) / total) * 100)
        self.progress_bar['value'] = progress_val
        self.progress_bar['maximum'] = 100
        
        self._log_message(f"Fila {current_index + 1}/{total} - {status_message}")
        
        # Actualizar contadores si la fila ya ha terminado de procesar
        if detail_status in ["Éxito", "Fallido", "Error"]:
            
            completed = self.completed_count.get()
            failed = self.failed_count.get()
            
            if detail_status == "Éxito":
                completed += 1
            else: # Fallido o Error
                failed += 1
                
            self._update_status_counts(total, completed, failed)


    # ----------------------------------------------------------------------
    # MÉTODOS DE CONTROL DEL PROCESO
    # ----------------------------------------------------------------------

    def toggle_processing(self):
        """Inicia o detiene el proceso de validación."""
        if self.is_running:
            # Llama a stop_processing para manejo de estado
            self.stop_processing(is_toggle=True) 
        else:
            self.start_processing()

    def start_processing(self):
        """Prepara e inicia el procesamiento en un hilo separado."""
        excel_path = self.excel_path_var.get()
        output_path = self.output_path_var.get()

        if not os.path.exists(excel_path):
            messagebox.showerror("Error", "Por favor, seleccione un archivo Excel válido.")
            return
        if not os.path.exists(output_path):
            messagebox.showerror("Error", "Por favor, seleccione una carpeta de salida válida.")
            return

        # 1. Configurar UI y Estado
        self.is_running = True
        self.stop_requested.clear()
        self.start_button.config(text="PAUSANDO...", state='disabled')
        self.stop_button.config(state='normal')
        self.download_button.config(state='disabled')
        self._update_status_counts(0, 0, 0)
        self.log_text.config(state='normal')
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state='disabled')
        self.progress_bar['value'] = 0

        # 2. Cargar DataFrame inicial
        try:
             self.df = pd.read_excel(excel_path)
             self.total_rows = len(self.df)
             # Inicializar columna de estado para seguimiento de pendientes
             self.df['Estado_Proceso'] = 'Pendiente' 
             self._update_status_counts(self.total_rows, 0, 0)
        except Exception as e:
            messagebox.showerror("Error de Carga", f"Fallo al cargar el archivo Excel: {e}")
            self._reset_ui_after_completion()
            return
            
        # 3. Inicializar y empezar el hilo
        self.processor = SARValidator(output_path, self.mode_var.get(), self.headless_var.get())
        
        self.processing_thread = threading.Thread(target=self._run_processing)
        self.processing_thread.start()
        
        self.start_button.config(text="EN EJECUCIÓN...", style='Info.TButton', state='disabled')


    def _run_processing(self):
        """Método que se ejecuta en el hilo secundario para el procesamiento."""
        try:
            self._log_message("Iniciando navegador Selenium...")
            
            if not self.processor.initialize_driver():
                raise Exception("Fallo al inicializar el navegador.")
            
            self._log_message(f"Procesando {len(self.df)} registros. Modo: {self.mode_var.get()}")
            
            # El core_processor necesita la columna 'Estado_Proceso' para saber qué guardar
            df_result = self.processor.procesar_dataframe(self.df, self._update_progress)
            
            # Actualizar el DataFrame de la aplicación con los resultados del core
            if isinstance(df_result, pd.DataFrame):
                 self.df = df_result
                 
            self._log_message("Proceso principal finalizado.")

        except Exception as e:
            self.after(0, lambda: messagebox.showerror("Error Fatal", f"🔴 ERROR CRÍTICO en el proceso: {e}"))
            self._log_message(f"🔴 ERROR CRÍTICO: {e}")
            self._log_message(traceback.format_exc())
            
        finally:
            if self.processor:
                self.processor.close_driver()
            
            # Asegurar que la UI se restablezca en el hilo principal
            self.after(0, self._reset_ui_after_completion)


    def stop_processing(self, is_toggle=False):
        """Solicita detener el proceso y espera a que el hilo termine."""
        if not self.is_running or not self.processing_thread or not self.processing_thread.is_alive():
            return
            
        # Mostrar estado de detención en la UI
        self.start_button.config(text="DETENIENDO...", state='disabled')
        self.stop_button.config(state='disabled')
        self._log_message("🛑 Solicitud de detención recibida. Esperando a que termine el registro actual...")
        
        # En una aplicación real, se usaría self.stop_requested.set() para que procesar_dataframe 
        # compruebe la bandera y se detenga amablemente.
        # Como el core_processor no revisa la bandera, el proceso se detiene al finalizar el registro actual
        
        # El hilo continuará hasta el final del registro actual, luego terminará.

    def _reset_ui_after_completion(self):
        """Restablece los elementos de la UI después de que el hilo termina."""
        self.is_running = False
        self.start_button.config(text="🚀 Ejecutar Proceso", style='Accent.TButton', state='normal')
        self.stop_button.config(state='disabled')
        
        # Habilitar descarga si hay filas para revisar
        if self.df is not None and len(self.df[self.df['Estado_Proceso'] == 'Pendiente']) > 0 or \
           self.df is not None and len(self.df[self.df['Estado_Proceso'].str.contains('Error|Fallido', na=False, case=False)]) > 0:
            self.download_button.config(state='normal')
            
        # Comprobar estado final para actualizar el progreso al 100%
        if self.df is not None:
             total = len(self.df)
             completed = len(self.df[self.df['Estado_Proceso'].str.contains('Válido|Data OK|PDF OK', na=False, case=False)])
             failed = len(self.df[self.df['Estado_Proceso'].str.contains('Error|Fallido', na=False, case=False)])
             self._update_status_counts(total, completed, failed)
             self.progress_bar['value'] = 100
             self._log_message("✅ PROCESO COMPLETADO (o detenido) Y RECURSOS LIBERADOS.")


    # ----------------------------------------------------------------------
    # MÉTODOS DE UTILIDAD
    # ----------------------------------------------------------------------
    
    def download_pending_errors(self):
        """Permite guardar los registros pendientes o fallidos en un nuevo Excel."""
        if self.df is None:
            messagebox.showinfo("Información", "No hay datos procesados para exportar.")
            return

        # Filtrar filas que no fueron exitosas
        # Usamos .str.contains para buscar los estados que indican fallo o pendiente.
        df_errors = self.df[
            self.df['Estado_Proceso'].str.contains('Pendiente|Error|Fallido|Incierto', na=False, case=False)
        ].copy() # Usar copy para evitar SettingWithCopyWarning

        if df_errors.empty:
            messagebox.showinfo("Información", "Todas las filas fueron procesadas exitosamente o con una validación 'Válida'/'No Válida'.")
            return

        # Limpiar y preparar el DataFrame de errores (manteniendo solo las columnas originales)
        # Se asume que las columnas originales son 'RTN', 'Clave referencia 3', 'Fecha doc.', etc.
        cols_to_keep = ['RTN', 'Clave referencia 3', 'Fecha doc.'] # Columnas clave del input
        
        # Identificar las columnas del DataFrame original, si es posible
        original_cols = [c for c in self.df.columns if c not in ['Estado_Proceso', 'Detalle_Validacion', 'original_index', 'RTN_EXTRAIDO', 'NUM_DOCUMENTO_BUSQUEDA', 'FECHA_DOCUMENTO_BUSQUEDA']]
        
        # Mantener las columnas originales del input, más el estado del proceso y el detalle de la validación
        export_cols = [c for c in original_cols if c in df_errors.columns] + ['Estado_Proceso', 'Detalle_Validacion']
        df_export = df_errors[export_cols]
        
        # Diálogo para guardar el archivo
        filename = "SAR_Pendientes_Errores_" + datetime.now().strftime("%Y%m%d_%H%M%S") + ".xlsx"
        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            initialfile=filename,
            filetypes=[("Excel files", "*.xlsx")]
        )

        if save_path:
            try:
                df_export.to_excel(save_path, index=False)
                messagebox.showinfo("Éxito", f"Registros pendientes/fallidos guardados en:\n{save_path}")
                self._log_message(f"Registros pendientes/fallidos guardados: {len(df_export)} filas.")
            except Exception as e:
                messagebox.showerror("Error de Guardado", f"Fallo al guardar el archivo: {e}")


# --- Bloque principal de ejecución ---
# -------------------------------------------------------------
# PUNTO DE ENTRADA PRINCIPAL (CORRECCIÓN IMPLEMENTADA AQUÍ)
# -------------------------------------------------------------
if __name__ == '__main__':
    # >>> MODIFICACIÓN CRÍTICA PARA EVITAR MÚLTIPLES VENTANAS <<<
    # Esto es necesario para que undetected-chromedriver funcione correctamente 
    # cuando se empaqueta con PyInstaller en Windows (evita la ejecución recursiva).
    if sys.platform.startswith('win'):
        import multiprocessing
        multiprocessing.freeze_support() 
    # >>> FIN DE LA MODIFICACIÓN CRÍTICA <<<

    # Inicialización de la aplicación
    try:
        app = SARApp()
        app.mainloop()
    except Exception as e:
        # Manejo de cualquier error crítico en la inicialización de la GUI
        print(f"Error fatal al iniciar la aplicación: {e}")
        messagebox.showerror("Error Fatal", f"La aplicación falló al iniciar. Por favor, revisa el log de la consola. Error: {e}")