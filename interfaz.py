import os
import re
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, filedialog, Canvas
import pandas as pd
from datetime import datetime
import sys
import json
import sqlite3 # Importar la librería sqlite3 directamente
import webbrowser
from collections import defaultdict
import threading
import traceback
from openpyxl import Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment, Font, Border, Side, PatternFill
from openpyxl.formatting.rule import CellIsRule
import subprocess
import requests
import sys
import ctypes
import psutil

try:
    import winsound
except ImportError:
    print("Librería 'winsound' no encontrada. No se reproducirán sonidos (solo disponible en Windows).")
    winsound = None

__version__ = "1.2.0" # IMPORTANTE: Esta es la versión de tu script local

# Reemplaza 'tu-usuario' y 'tu-repositorio' con los tuyos
URL_VERSION = "https://raw.githubusercontent.com/ZombPool/P-11-Sistema-verificaci-n-de-datos/main/version.txt"
URL_SCRIPT = "https://github.com/ZombPool/P-11-Sistema-verificaci-n-de-datos/releases/download/1.2.0/interfaz.exe"
# --- Dependencias Requeridas 
# Intenta importar xlrd para archivos .xls (Geometría antigua)
try:
    import xlrd
except ImportError:
    messagebox.showwarning("Dependencia Faltante", 
                           "La librería 'xlrd' no está instalada. Es necesaria para leer archivos .xls.\n\n"
                           "Para instalarla, ejecuta: pip install xlrd")

# Intenta importar openpyxl para archivos .xlsx (Reportes y resultados nuevos)
try:
    import openpyxl
except ImportError:
    messagebox.showwarning("Dependencia Faltante", 
                           "La librería 'openpyxl' no está instalada. Es necesaria para leer y escribir archivos .xlsx.\n\n"
                           "Para instalarla, ejecuta: pip install openpyxl")

# Para un mejor aspecto visual, usaremos ttkbootstrap.
try:
    import ttkbootstrap as ttk
    from ttkbootstrap.constants import *
except ImportError:
    messagebox.showerror("Librería Faltante", 
                         "Por favor, instala la librería 'ttkbootstrap' para la interfaz gráfica:\n\npip install ttkbootstrap")
    sys.exit()

def is_osk_running():
    """Verifica si el teclado en pantalla (osk.exe) ya está en ejecución."""
    for proc in psutil.process_iter(['name']):
        if proc.info['name'] == "osk.exe":
            return True
    return False

def open_keyboard(event=None):
    """Abre el teclado en pantalla pidiendo elevación de administrador."""
    if not is_osk_running():
        try:
            # Este método le pide a Windows que ejecute 'osk.exe' usando el verbo 'runas',
            # que significa "Ejecutar como Administrador".
            # Esto mostrará la ventana de confirmación (UAC) si es necesario.
            ctypes.windll.shell32.ShellExecuteW(
                None,       # hwnd
                "runas",    # lpOperation
                "osk.exe",  # lpFile
                None,       # lpParameters
                None,       # lpDirectory
                1           # nShowCmd
            )
        except Exception as e:
            # Esto puede ocurrir si el usuario hace clic en "No" en la ventana de permisos.
            # El código de error para "operación cancelada por el usuario" es 1223.
            if hasattr(e, 'winerror') and e.winerror == 1223:
                print("El usuario canceló la solicitud de permisos para abrir el teclado.")
            else:
                messagebox.showerror("Error de Teclado", f"No se pudo iniciar el teclado virtual: {e}")

# --- CLASES DE ANÁLISIS INTEGRADAS ---

class AnalisisILRL:
    def __init__(self, master=None):
        self.master = master

    def abrir_archivo(self, ruta_archivo):
        try:
            if sys.platform == 'win32':
                os.startfile(ruta_archivo)
            elif sys.platform == 'darwin':
                subprocess.run(['open', ruta_archivo], check=True)
            else:
                subprocess.run(['xdg-open', ruta_archivo], check=True)
        except Exception as e:
            print(f"Advertencia: No se pudo abrir el archivo automáticamente: {e}")

    def extraer_clave(self, archivo):
        base = os.path.splitext(os.path.basename(archivo))[0]
        # Patrón para nombres cortos (ej. JMO-2512000030001)
        patron_nuevo = r'J(?:R)?MO-(\d{9})(\d{4})'
        m_nuevo = re.match(patron_nuevo, base, re.IGNORECASE)
        if m_nuevo:
            return f"{m_nuevo.group(1)}-{m_nuevo.group(2)}"
        # Patrón para nombres largos (ej. JMO-251200003-SC-LC-0001)
        patron_anterior = r'J(?:R)?MO-(\d{9})-(?:LC|SC|SCLC|LCSC|SC-LC|LC-SC)(?:-[KF])?-(\d{4})'
        m_anterior = re.match(patron_anterior, base, re.IGNORECASE)
        if m_anterior:
            return f"{m_anterior.group(1)}-{m_anterior.group(2)}"
        return None

    def leer_resultado_y_fecha(self, ruta):
        try:
            df = pd.read_excel(ruta, header=None)
            inicio_filas_datos = 12
            col7_vals = df.iloc[inicio_filas_datos:, 7].dropna().astype(str).str.upper()
            col8_vals = df.iloc[inicio_filas_datos:, 8].dropna().astype(str).str.upper()
            count_col7 = col7_vals.isin(['PASS', 'FAIL']).sum()
            count_col8 = col8_vals.isin(['PASS', 'FAIL']).sum()
            col_resultado = 8 if count_col8 >= count_col7 else 7
            col_fecha = 10 if col_resultado == 8 else 9
            resultados_raw = df.iloc[inicio_filas_datos:, col_resultado].dropna().astype(str).str.upper().tolist()
            fechas_raw = df.iloc[inicio_filas_datos:, col_fecha].dropna().tolist()
            if not resultados_raw:
                return None, None, None
            resultado_final = 'APROBADO' if all(r == 'PASS' for r in resultados_raw) else 'RECHAZADO'
            fechas_datetime = [pd.to_datetime(f) for f in fechas_raw if f]
            ultima_fecha = max(fechas_datetime).strftime("%d/%m/%Y %H:%M") if fechas_datetime else ''
            detalle_str = ", ".join([f"C{i+1} {r}" for i, r in enumerate(resultados_raw)])
            return resultado_final, ultima_fecha, detalle_str
        except Exception:
            return None, None, None

    # --- MODIFICACIÓN PRINCIPAL: Acepta lista de carpetas ---
    def analizar_carpetas_ilrl(self, lista_carpetas, total_esperado, update_progress_callback=None):
        archivos_excel = []
        
        # 1. Recolectar archivos de TODAS las carpetas encontradas
        for carpeta in lista_carpetas:
            if os.path.isdir(carpeta):
                files = [os.path.join(dirpath, f) for dirpath, _, filenames in os.walk(carpeta) for f in filenames if f.endswith('.xlsx')]
                archivos_excel.extend(files)

        if not archivos_excel:
            return {}, [], {}, ["No se encontraron archivos .xlsx en las rutas de la O.T."]

        agrupados = defaultdict(list)
        errores_lectura = []
        
        # 2. Procesar todos los archivos encontrados (mismo lógica de antes)
        total_archivos = len(archivos_excel)
        for i, ruta_completa in enumerate(archivos_excel):
            archivo_nombre = os.path.basename(ruta_completa)
            clave = self.extraer_clave(archivo_nombre)
            if not clave:
                errores_lectura.append(f"'{archivo_nombre}': No se pudo extraer clave. Ignorado.")
                continue
            resultado, fecha, detalle = self.leer_resultado_y_fecha(ruta_completa)
            if resultado is None:
                errores_lectura.append(f"'{archivo_nombre}': Error al leer. Ignorado.")
                continue
            
            # Guardamos la ruta para referencia
            agrupados[clave].append({'archivo': archivo_nombre, 'ruta': ruta_completa, 'resultado': resultado, 'fecha': fecha, 'detalle': detalle})
            
            if update_progress_callback:
                update_progress_callback((i + 1) / total_archivos * 100)
        
        resultados_finales = {}
        rechazados = []
        
        # 3. Consolidar: Si hay duplicados (ej. en carpeta 1 y carpeta 2), toma el más reciente
        for clave, items in agrupados.items():
            items_filtrados = [i for i in items if i['fecha']]
            if not items_filtrados:
                mas_reciente = items[-1]
            else:
                mas_reciente = max(items_filtrados, key=lambda x: datetime.strptime(x['fecha'], "%d/%m/%Y %H:%M"))
            
            resultados_finales[clave] = {'resultado': mas_reciente['resultado'], 'fecha': mas_reciente['fecha'], 'archivos': [i['archivo'] for i in items]}
            
            if mas_reciente['resultado'] == 'RECHAZADO':
                rechazados.append([clave, 'RECHAZADO', mas_reciente['detalle'], mas_reciente['archivo'], mas_reciente['fecha']])
        
        return resultados_finales, rechazados, agrupados, errores_lectura

    def generar_reporte_excel_ilrl(self, resultados, total_esperado, rechazados, ruta_destino_base, agrupados):
        carpeta_analisis = os.path.join(ruta_destino_base, "ANALISIS DE O.T")
        os.makedirs(carpeta_analisis, exist_ok=True)
        nombre_archivo_reporte = "Reporte_Analisis_ILRL.xlsx"
        nueva_ruta_destino = os.path.join(carpeta_analisis, nombre_archivo_reporte)
        
        if os.path.exists(nueva_ruta_destino):
            try:
                os.remove(nueva_ruta_destino)
            except Exception as e:
                return None, f"No se pudo eliminar el archivo anterior. Asegúrate de que no esté abierto.\nError: {e}"
        
        wb = Workbook()
        ws1 = wb.active
        ws1.title = "Resultados OT"
        ws1.append(['Cable', 'Resultado Final', 'Última Fecha de Medición', 'Archivos Analizados', 'Rutas Completas'])
        
        def sort_key_numeric_cable_id(cable_str):
            parts = cable_str.split('-')
            return int(parts[1]) if len(parts) >= 2 and parts[1].isdigit() else cable_str
        
        for cable in sorted(resultados.keys(), key=sort_key_numeric_cable_id):
            info = resultados[cable]
            ws1.append([cable, info['resultado'], info['fecha'], ", ".join(info['archivos']), ", ".join([i['ruta'] for i in agrupados[cable]])])
        
        if ws1.max_row > 1:
            tabla1 = Table(displayName="TablaResultados", ref=f"A1:E{ws1.max_row}")
            tabla1.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
            ws1.add_table(tabla1)
        
        ws2 = wb.create_sheet("Cantidades de cables")
        orden_trabajo_prefijo = ""
        if resultados:
            first_key = next(iter(resultados.keys()))
            match = re.match(r'(\d{9})-\d{4}', first_key)
            if match:
                orden_trabajo_prefijo = match.group(1) + "-"
        
        esperados_formato_completo_set = {f"{orden_trabajo_prefijo}{str(i).zfill(4)}" for i in range(1, total_esperado + 1)} if orden_trabajo_prefijo else set()
        claves_encontradas_set = set(resultados.keys())
        faltantes = sorted(list(esperados_formato_completo_set - claves_encontradas_set))
        
        ws2.append(['Total Esperado', total_esperado])
        ws2.append(['Encontrados', len(claves_encontradas_set)])
        ws2.append(['Faltantes', len(faltantes)])
        ws2.append([])
        ws2.append(['Listado de Faltantes (sugerido)'])
        for f in faltantes:
            ws2.append([f])
        
        ws3 = wb.create_sheet("Cables rechazados")
        ws3.append(['Cable', 'Estado', 'Detalle', 'Archivo Fuente', 'Fecha de Medición'])
        for r in rechazados:
            ws3.append(r)
            
        for ws in [ws1, ws2, ws3]:
            for col in ws.columns:
                max_length = max(len(str(cell.value)) for cell in col if cell.value) if any(c.value for c in col) else 0
                ws.column_dimensions[get_column_letter(col[0].column)].width = max_length + 2
        
        try:
            wb.save(nueva_ruta_destino)
            return nueva_ruta_destino, None
        except PermissionError:
            return None, f"No se pudo guardar el archivo. Asegúrate de que no esté abierto:\n{nueva_ruta_destino}"
        except Exception as e:
            return None, f"Error inesperado al guardar el archivo: {e}"

    # --- MÉTODO PROCESAR ACTUALIZADO ---
    def procesar_ilrl(self, ot_number, config, total_esperado, update_progress_callback=None):
        # 1. Identificar rutas
        rutas_posibles = []
        if config.get('ruta_base_ilrl'):
            rutas_posibles.append(os.path.join(config['ruta_base_ilrl'], ot_number))
        if config.get('ruta_base_ilrl_2'):
            rutas_posibles.append(os.path.join(config['ruta_base_ilrl_2'], ot_number))
            
        rutas_validas = [r for r in rutas_posibles if os.path.isdir(r)]
        
        if not rutas_validas:
             return None, f"No se encontró la carpeta de la OT '{ot_number}' en ninguna de las rutas configuradas."

        # 2. Analizar múltiples carpetas
        resultados, rechazados, agrupados, errores_lectura = self.analizar_carpetas_ilrl(rutas_validas, total_esperado, update_progress_callback)
        
        if not resultados and not errores_lectura:
            return None, "No se encontraron datos válidos para el análisis IL/RL."
        
        # Usamos la primera ruta válida como base para guardar el reporte
        ruta_reporte, error_guardado = self.generar_reporte_excel_ilrl(resultados, total_esperado, rechazados, rutas_validas[0], agrupados)
        
        if error_guardado:
            return None, error_guardado
        return ruta_reporte, "\n".join(errores_lectura) if errores_lectura else None

class AnalisisGEO:
    def __init__(self, master=None):
        self.master = master

    def abrir_archivo(self, ruta_archivo):
        try:
            if sys.platform == 'win32':
                os.startfile(ruta_archivo)
            elif sys.platform == 'darwin':
                subprocess.run(['open', ruta_archivo], check=True)
            else:
                subprocess.run(['xdg-open', ruta_archivo], check=True)
        except Exception as e:
            raise RuntimeError(f"No se pudo abrir el archivo automáticamente: {e}")

    # ... (Mantén los métodos 'ajustar_columnas' y 'aplicar_estilos' igual que antes) ...
    def ajustar_columnas(self, ws):
        column_widths = {'A': 22, 'B': 15, 'C': 40, 'D': 25}
        for col, width in column_widths.items():
            ws.column_dimensions[col].width = width

    def aplicar_estilos(self, ws):
        header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True)
        border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
        center_alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.border = border
            cell.alignment = center_alignment
        for row in ws.iter_rows(min_row=2, max_col=ws.max_column, max_row=ws.max_row):
            for cell in row:
                cell.border = border
                if cell.column_letter == 'B':
                    cell.alignment = center_alignment
                    if cell.value == "ACEPTADO":
                        cell.fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
                        cell.font = Font(color="006100")
                    elif cell.value == "RECHAZADO":
                        cell.fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
                        cell.font = Font(color="9C0006")
                else:
                    cell.alignment = Alignment(wrap_text=True, vertical='center')

    # --- LÓGICA MULTI-ARCHIVO ---
    def analizar_archivos_geo_multi(self, lista_archivos, total_esperado, update_progress_callback=None, order_id_to_filter=None, mode="Duplex"):
        if not lista_archivos:
            return {}, [], 0, 0, ["No se proporcionaron archivos para analizar."]

        errores_lectura = []
        puntas_requeridas = ['1', '2'] if mode == "Simplex" else ['1', '2', '3', '4']
        
        # Estructura maestra: { serial_base: { punta_id: { resultado: ..., es_retrabajo: ..., timestamp: ... } } }
        datos_globales = {} 
        seriales_encontrados_set = set()

        total_files = len(lista_archivos)
        
        # 1. Leer y Fusionar todos los archivos
        for i, file_path in enumerate(lista_archivos):
            try:
                df = pd.read_excel(file_path, header=None, skiprows=12)
                
                # Función auxiliar regex (misma que en verif. individual)
                def parse_serial(s):
                    if not isinstance(s, str): return None, None
                    s = s.strip().upper()
                    match = re.search(r'(J(?:R)?MO-?\d{13}|\d{13})(-?([1-4R][1-4]?))?', s)
                    if match:
                         base = re.sub(r'[^0-9]', '', match.group(1))
                         punta = match.group(3) if match.group(3) else "N/A"
                         return base, punta
                    # Fallback simple
                    match_simple = re.search(r'(\d{13})', s)
                    if match_simple: return match_simple.group(1), "N/A"
                    return None, None

                for _, row in df.iterrows():
                    if len(row) < 7: continue
                    clave_base, punta = parse_serial(str(row[0]))
                    
                    if clave_base and punta != "N/A":
                        # Filtrar por O.T. si se requiere (evitar leer basura de otros cables si el archivo está mezclado)
                        if order_id_to_filter:
                            ot_clean = re.sub(r'[^0-9]', '', order_id_to_filter)
                            if not clave_base.startswith(ot_clean):
                                continue

                        seriales_encontrados_set.add(clave_base)
                        
                        resultado = str(row[6]).strip().upper() if pd.notna(row[6]) else "SIN_DATO"
                        # Intentar leer fecha
                        ts = None
                        if pd.notna(row[3]) and pd.notna(row[4]):
                             ts = pd.to_datetime(f"{row[3]} {row[4]}", errors='coerce')
                        
                        punta_limpia = punta.replace('R', '')
                        es_retrabajo = 'R' in punta
                        
                        if clave_base not in datos_globales:
                            datos_globales[clave_base] = {}
                        
                        # --- LÓGICA DE PRIORIDAD (Fusión) ---
                        # Si la punta no existe, o si es un Retrabajo y lo que teníamos NO era retrabajo -> Guardar/Sobrescribir
                        if punta_limpia not in datos_globales[clave_base]:
                            datos_globales[clave_base][punta_limpia] = {'res': resultado, 'rework': es_retrabajo, 'ts': ts}
                        else:
                            existente = datos_globales[clave_base][punta_limpia]
                            if es_retrabajo and not existente['rework']:
                                datos_globales[clave_base][punta_limpia] = {'res': resultado, 'rework': es_retrabajo, 'ts': ts}
                            # Si ambos son R o ambos normales, quedarse con el más reciente (si hay fecha)
                            elif es_retrabajo == existente['rework'] and ts and existente['ts'] and ts > existente['ts']:
                                datos_globales[clave_base][punta_limpia] = {'res': resultado, 'rework': es_retrabajo, 'ts': ts}

            except Exception as e:
                errores_lectura.append(f"Error leyendo {os.path.basename(file_path)}: {e}")
            
            if update_progress_callback:
                update_progress_callback((i + 1) / total_files * 80)

        # 2. Generar Reporte Final
        reporte_data = []
        
        # Calcular faltantes generales
        orden_trabajo_prefijo = ""
        if order_id_to_filter:
             # Formato JMO-XXXX... -> XXXX...-
             match = re.search(r'(\d{9})', order_id_to_filter)
             if match: orden_trabajo_prefijo = match.group(1) + "-"
        
        series_esperadas_set = {f"{orden_trabajo_prefijo}{str(i).zfill(4)}" for i in range(1, total_esperado + 1)} if orden_trabajo_prefijo else set()
        
        # Ojo: seriales_encontrados_set tiene el formato numérico (2512...). Necesitamos asegurarnos de comparar peras con peras.
        # series_esperadas_set también es "2512...-0001".
        # Vamos a formatear las encontradas para match (esto asume que clave_base ya viene como 2512000010001)
        claves_encontradas_formateadas = set()
        for cb in seriales_encontrados_set:
            # Insertar guion si es necesario para matchear formato esperado? 
            # El formato esperado es OT-Consecutivo (ej. 251200003-0001). 
            # El clave_base es 13 dígitos (ej. 2512000030001).
            if len(cb) == 13:
                cf = cb[:9] + "-" + cb[9:]
                claves_encontradas_formateadas.add(cf)
            else:
                claves_encontradas_formateadas.add(cb)

        faltantes_claves = sorted(list(series_esperadas_set - claves_encontradas_formateadas))

        # 3. Evaluar cada cable encontrado
        processed_count = 0
        total_items = len(datos_globales)

        for clave_base, puntas_data in datos_globales.items():
            # Formato visual OT-Consec
            clave_visual = clave_base[:9] + "-" + clave_base[9:] if len(clave_base) == 13 else clave_base
            
            estado_final = "ACEPTADO"
            detalles_puntas = []
            max_ts = pd.NaT

            for p_num in puntas_requeridas:
                if p_num in puntas_data:
                    data = puntas_data[p_num]
                    res_punta = "PASS" if data['res'] == 'PASS' else "FAIL"
                    tipo = "(R)" if data['rework'] else ""
                    detalles_puntas.append(f"P{p_num}{tipo}: {res_punta}")
                    
                    if data['res'] != 'PASS': estado_final = "RECHAZADO"
                    if data['ts'] and (pd.isna(max_ts) or data['ts'] > max_ts): max_ts = data['ts']
                else:
                    detalles_puntas.append(f"P{p_num}: FALTANTE")
                    estado_final = "RECHAZADO"

            reporte_data.append({
                'Número de Serie (OT-Cable)': clave_visual,
                'Estado Final': estado_final,
                'Detalle Puntas': ', '.join(detalles_puntas),
                'Última Medición': max_ts
            })
            
            processed_count += 1
            if update_progress_callback and total_items > 0:
                 update_progress_callback(80 + (processed_count / total_items) * 20)

        return reporte_data, faltantes_claves, len(datos_globales), len(faltantes_claves), errores_lectura

    def generar_reporte_excel_geo(self, reporte_data, faltantes_claves, total_esperado, total_encontrados, total_faltantes, ruta_destino_base, order_id=""):
        carpeta_analisis = os.path.join(ruta_destino_base, "ANALISIS DE O.T")
        os.makedirs(carpeta_analisis, exist_ok=True)
        nombre_archivo_reporte = f"AnalisisGEO_{order_id}.xlsx" if order_id and order_id != "Desconocida" else "AnalisisGEO_Reporte.xlsx"
        nueva_ruta_destino = os.path.join(carpeta_analisis, nombre_archivo_reporte)
        
        if os.path.exists(nueva_ruta_destino):
            try:
                os.remove(nueva_ruta_destino)
            except Exception as e:
                return None, f"No se pudo eliminar el archivo anterior.\nError: {e}"
        
        wb = Workbook()
        ws1 = wb.active
        ws1.title = "Resultados GEO"
        df_reporte = pd.DataFrame(reporte_data)
        if not df_reporte.empty:
            # Ordenar por consecutivo
            df_reporte['Orden_Cable'] = df_reporte['Número de Serie (OT-Cable)'].apply(lambda x: int(x.split('-')[1]) if '-' in x and x.split('-')[1].isdigit() else 0)
            df_reporte.sort_values(by='Orden_Cable', inplace=True)
            df_reporte.drop(columns=['Orden_Cable'], inplace=True)
            
            ws1.append(df_reporte.columns.tolist())
            for row in df_reporte.to_numpy().tolist():
                ws1.append(row)
            self.aplicar_estilos(ws1)
            self.ajustar_columnas(ws1)
            if ws1.max_row > 1:
                tabla1 = Table(displayName="TablaResultadosGEO", ref=f"A1:{get_column_letter(ws1.max_column)}{ws1.max_row}")
                tabla1.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
                ws1.add_table(tabla1)
        else:
            ws1.append(['Número de Serie (OT-Cable)', 'Estado Final', 'Detalle Puntas', 'Última Medición'])

        ws2 = wb.create_sheet("Resumen GEO")
        ws2.append(['Métrica', 'Valor'])
        ws2.append(['Total esperado', total_esperado])
        ws2.append(['Total encontrados', total_encontrados])
        ws2.append(['Total faltantes', total_faltantes])
        ws2.append([])
        ws2.append(['Cables Faltantes'])
        for f in faltantes_claves:
            ws2.append([f])
        
        for col in ws2.columns:
            max_length = max(len(str(cell.value)) for cell in col if cell.value) if any(c.value for c in col) else 0
            ws2.column_dimensions[get_column_letter(col[0].column)].width = max_length + 2
            
        try:
            wb.save(nueva_ruta_destino)
            return nueva_ruta_destino, None
        except PermissionError:
            return None, f"No se pudo guardar el archivo.\n{nueva_ruta_destino}"
        except Exception as e:
            return None, f"Error inesperado al guardar: {e}"

    # --- MÉTODO PROCESAR ACTUALIZADO ---
    def procesar_geo(self, ot_number, config, total_esperado, update_progress_callback=None, mode="Duplex"):
        # 1. Buscar archivos en TODAS las rutas configuradas
        rutas_base = []
        if config.get('ruta_base_geo'): rutas_base.append(config['ruta_base_geo'])
        if config.get('ruta_base_geo_2'): rutas_base.append(config['ruta_base_geo_2'])
        
        archivos_candidatos = []
        for ruta_base in rutas_base:
            if os.path.isdir(ruta_base):
                encontrados = [os.path.join(ruta_base, f) for f in os.listdir(ruta_base) 
                               if f.lower().endswith(('.xlsx', '.xls')) 
                               and not f.startswith('~$') 
                               and ot_number in f] # Filtro simple por nombre
                # Ojo: si hay varios archivos de la misma OT en una carpeta (ej. versiones viejas),
                # aquí los añadimos todos. La función de análisis ya se encarga de priorizar por fecha/retrabajo.
                archivos_candidatos.extend(encontrados)
        
        if not archivos_candidatos:
            return None, f"No se encontraron archivos para la OT '{ot_number}' en las rutas configuradas."

        # 2. Analizar y Fusionar
        reporte_data, faltantes, total_encontrados, total_faltantes, errores = self.analizar_archivos_geo_multi(
            archivos_candidatos, total_esperado, update_progress_callback, order_id_to_filter=ot_number, mode=mode
        )

        if not reporte_data and errores:
             return None, f"Errores de lectura: {errores[0]}"
        if not reporte_data:
             return None, "No se encontraron datos válidos."

        # 3. Generar Excel
        # Usamos la primera carpeta configurada como destino principal
        ruta_destino_base = config['ruta_base_geo'] if os.path.isdir(config['ruta_base_geo']) else os.path.dirname(archivos_candidatos[0])
        
        ruta_reporte, error_guardado = self.generar_reporte_excel_geo(
            reporte_data, faltantes, total_esperado, total_encontrados, total_faltantes, ruta_destino_base, order_id=ot_number
        )
        
        if error_guardado: return None, error_guardado
        return ruta_reporte, "\n".join(errores) if errores else None


class App(ttk.Window):
    def __init__(self):
        super().__init__(themename="litera")
        self.title("FibraTrace - Sistema de Trazabilidad")
        self.geometry("1200x800")

        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)

        self.cable_mode = tk.StringVar(value="Duplex")

        # --- Configuración Centralizada ---
        self.config_file = "config.json"
        self.password = "admin123" # Contraseña para funciones de administrador
        self.config = self.load_config()
        self.init_database()
        
        self.create_sidebar()
        self.create_main_content_area()

        self.pages = {}
        self.create_pages()
        self.show_page("Dashboard")

        self.bind_class("TEntry", "<FocusIn>", open_keyboard)

    # ... (dentro de la clase App) ...

    def check_for_updates(self):
        """Verifica si hay una nueva versión disponible en GitHub."""
        try:
            response = requests.get(URL_VERSION, timeout=5)
            if response.status_code == 200:
                remote_version = response.text.strip()
                
                # Comparamos las versiones (una simple comparación de strings funciona para "1.0.1" > "1.0.0")
                if remote_version > __version__:
                    if messagebox.askyesno("Actualización Disponible", 
                                          f"Hay una nueva versión ({remote_version}) disponible.\n"
                                          f"Tu versión actual es {__version__}.\n\n"
                                          "¿Deseas descargarla y reiniciar la aplicación ahora?"):
                        self.apply_update()
                else:
                    messagebox.showinfo("Sin Actualizaciones", "Ya estás utilizando la versión más reciente.")
            else:
                messagebox.showerror("Error de Verificación", f"No se pudo verificar la versión (Código: {response.status_code}).")
        except requests.RequestException as e:
            messagebox.showerror("Error de Conexión", f"No se pudo conectar a GitHub para buscar actualizaciones.\n\nError: {e}")

    def apply_update(self):
        """Descarga el nuevo ejecutable y ejecuta el actualizador."""
        try:
            # Mostramos un mensaje de que la descarga está en proceso
            self.page_title_label.config(text="Descargando actualización...")
            self.update_idletasks()

            response = requests.get(URL_SCRIPT, timeout=60) # Aumentamos a 60s por si la conexión es lenta
            if response.status_code == 200:
                # --- CORRECCIÓN 1: El archivo temporal ahora es un .exe ---
                nuevo_script_path = "interfaz_new.exe"
                with open(nuevo_script_path, "wb") as f:
                    f.write(response.content)

                # Obtenemos el nombre del ejecutable actual (ej. "interfaz.exe")
                script_actual = os.path.basename(sys.argv[0])

                # --- CORRECCIÓN 2: Ejecutamos updater.exe directamente ---
                # Esto asume que updater.exe está en la misma carpeta que interfaz.exe
                subprocess.Popen(['updater.exe', script_actual, nuevo_script_path])

                # Cerramos la aplicación actual para que el updater pueda trabajar
                self.destroy()

            else:
                messagebox.showerror("Error de Descarga", f"No se pudo descargar la actualización (Código: {response.status_code}).")
                self.page_title_label.config(text=self.pages["Dashboard"].__class__.__name__) # Restaura el título a la página actual

        except requests.RequestException as e:
            messagebox.showerror("Error de Conexión", f"Ocurrió un error al descargar el archivo.\n\nError: {e}")
            self.page_title_label.config(text=self.pages["Dashboard"].__class__.__name__)
        except Exception as e:
            messagebox.showerror("Error Inesperado", f"Ocurrió un error al aplicar la actualización:\n\n{e}")
            self.page_title_label.config(text=self.pages["Dashboard"].__class__.__name__)

    # Dentro de la clase App
    def load_config(self):
        """Carga la configuración desde config.json o crea un archivo por defecto."""
        default_config = {
            "ruta_base_ilrl": "C:\\Ruta\\Por\\Defecto\\ILRL_SC_LC",
            "ruta_base_ilrl_2": "", 
            "ruta_base_geo": "C:\\Ruta\\Por\\Defecto\\Geometria_SC_LC",
            "ruta_base_geo_2": "",
            "ruta_base_ilrl_mpo": "C:\\Ruta\\Por\\Defecto\\ILRL_MPO",
            "ruta_base_geo_mpo": "C:\\Ruta\\Por\\Defecto\\Geometria_MPO",
            "ruta_base_polaridad_mpo": "C:\\Ruta\\Por\\Defecto\\Polaridad_MPO",
            
            # --- NUEVO: Switches para MPO (Por defecto todo activado) ---
            "check_mpo_ilrl": True,
            "check_mpo_geo": True,
            "check_mpo_polaridad": True,
            # ------------------------------------------------------------
            
            "db_path": os.path.join(os.path.expanduser('~'), 'Documents', 'FibraTraceData', 'verifications.db')
        }
        if os.path.exists(self.config_file):
            with open(self.config_file, 'r') as f:
                config = json.load(f)
                for key, value in default_config.items():
                    config.setdefault(key, value)
                return config
        else:
            self.save_config(default_config)
            return default_config

    def save_config(self, config_data):
        """Guarda la configuración actual en config.json."""
        with open(self.config_file, 'w') as f:
            json.dump(config_data, f, indent=4)
        self.config = config_data # Actualizar la configuración en memoria

    def init_database(self):
        """Inicializa la base de datos y crea/actualiza las tablas necesarias."""
        try:
            db_dir = os.path.dirname(self.config['db_path'])
            os.makedirs(db_dir, exist_ok=True)
            conn = sqlite3.connect(self.config['db_path'])
            cursor = conn.cursor()

            # --- Tabla de Verificaciones ---
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS cable_verifications (
                    id INTEGER PRIMARY KEY AUTOINCREMENT, entry_date TEXT NOT NULL,
                    serial_number TEXT NOT NULL, ot_number TEXT NOT NULL,
                    overall_status TEXT NOT NULL, ilrl_status TEXT,
                    ilrl_details TEXT, geo_status TEXT, geo_details TEXT
                )
            """)
            # Añadir columnas de MPO si no existen
            cursor.execute("PRAGMA table_info(cable_verifications)")
            columns = [info[1] for info in cursor.fetchall()]
            if 'polaridad_status' not in columns:
                cursor.execute("ALTER TABLE cable_verifications ADD COLUMN polaridad_status TEXT")
            if 'polaridad_details' not in columns:
                cursor.execute("ALTER TABLE cable_verifications ADD COLUMN polaridad_details TEXT")

            # --- Tabla de Configuraciones de OT para MPO ---
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS ot_configurations (
                    ot_number TEXT PRIMARY KEY, drawing_number TEXT, link TEXT,
                    num_conectores_a INTEGER, fibers_per_connector_a INTEGER,
                    num_conectores_b INTEGER, fibers_per_connector_b INTEGER,
                    ilrl_ot_header TEXT, ilrl_serie_header TEXT,
                    ilrl_fecha_header TEXT, ilrl_hora_header TEXT,
                    ilrl_estado_header TEXT, ilrl_conector_header TEXT,
                    last_modified TEXT
                )
            """)
            
            conn.commit()
            conn.close()
        except Exception as e:
            messagebox.showerror("Error de Base de Datos", f"No se pudo inicializar la base de datos: {e}")

    # En la clase App, reemplaza este método completo:
    def create_sidebar(self):
        sidebar_frame = ttk.Frame(self, style='secondary.TFrame', padding=10)
        sidebar_frame.grid(row=0, column=0, sticky="ns")

        ttk.Label(sidebar_frame, text="FibraTrace", font=("Helvetica", 18, "bold"), style='inverse-secondary.TLabel').pack(pady=10)
        ttk.Separator(sidebar_frame).pack(fill='x', pady=10)

        ttk.Button(sidebar_frame, text="Dashboard", command=lambda: self.show_page("Dashboard"), style='primary.TButton').pack(fill='x', pady=5)

        ttk.Label(sidebar_frame, text="PRODUCTOS SC & LC", style='inverse-secondary.TLabel', font=("Helvetica", 10, "bold")).pack(pady=(20, 5), anchor='w')
        
        # --- BLOQUE AÑADIDO: Interruptor Simplex/Duplex ---
        switch_frame = ttk.Frame(sidebar_frame, style='secondary.TFrame')
        switch_frame.pack(fill='x', pady=5, padx=5)
        
        simplex_label = ttk.Label(switch_frame, text="Simplex", style='inverse-secondary.TLabel')
        simplex_label.pack(side='left')
        
        self.mode_switch = ttk.Checkbutton(
            switch_frame,
            bootstyle="primary-round-toggle",
            variable=self.cable_mode,
            onvalue="Duplex",
            offvalue="Simplex"
        )
        self.mode_switch.pack(side='left', padx=5)
        
        duplex_label = ttk.Label(switch_frame, text="Duplex", style='inverse-secondary.TLabel')
        duplex_label.pack(side='left')
        # ----------------------------------------------------
        ttk.Button(sidebar_frame, text="Verificación Individual", command=lambda: self.show_page("Verificacion_LC_SC"), style='primary.TButton').pack(fill='x', pady=5)
        ttk.Button(sidebar_frame, text="Análisis de O.T.", command=lambda: self.show_page("Reportes_LC_SC"), style='primary.TButton').pack(fill='x', pady=5)
        ttk.Button(sidebar_frame, text="Registro WH", command=lambda: self.show_page("RegistroWH"), style='primary.TButton').pack(fill='x', pady=5)

        ttk.Label(sidebar_frame, text="PRODUCTOS MPO", style='inverse-secondary.TLabel', font=("Helvetica", 10, "bold")).pack(pady=(20, 5), anchor='w')
        ttk.Button(sidebar_frame, text="Verificación Individual", command=lambda: self.show_page("Verificacion_MPO"), style='primary.TButton').pack(fill='x', pady=5)
        ttk.Button(sidebar_frame, text="Análisis de O.T.", command=lambda: self.show_page("Reportes_MPO"), style='primary.TButton').pack(fill='x', pady=5)
        ttk.Button(sidebar_frame, text="Registro WH", command=lambda: self.show_page("RegistroWHMPO"), style='primary.TButton').pack(fill='x', pady=5)
        
        ttk.Label(sidebar_frame, text="HERRAMIENTAS", style='inverse-secondary.TLabel', font=("Helvetica", 10, "bold")).pack(pady=(20, 5), anchor='w')
        ttk.Button(sidebar_frame, text="Configurar Rutas", command=self.open_settings_window, style='info.TButton').pack(fill='x', pady=5)
        ttk.Button(sidebar_frame, text="Ver Registros", command=lambda: self.show_page("Registros"), style='info.TButton').pack(fill='x', pady=5)
        ttk.Button(sidebar_frame, text="Diagnóstico DB", command=self.show_db_diagnostics, style='info.TButton').pack(fill='x', pady=5)
        ttk.Button(sidebar_frame, text="Buscar Actualizaciones", command=self.check_for_updates, style='info.outline.TButton').pack(fill='x', pady=5)

    def create_main_content_area(self):
        self.main_frame = ttk.Frame(self, padding=20)
        self.main_frame.grid(row=0, column=1, sticky="nsew")
        self.main_frame.grid_rowconfigure(1, weight=1)
        self.main_frame.grid_columnconfigure(0, weight=1)
        
        header_frame = ttk.Frame(self.main_frame)
        header_frame.grid(row=0, column=0, sticky="ew", pady=(0, 20))
        self.page_title_label = ttk.Label(header_frame, text="Dashboard", font=("Helvetica", 24, "bold"))
        self.page_title_label.pack(side='left')

    def create_pages(self):
        self.pages["Dashboard"] = DashboardPage(self.main_frame)
        self.pages["Verificacion_LC_SC"] = VerificacionLC_SC_Page(self.main_frame, self)
        self.pages["Reportes_LC_SC"] = Reportes_LC_SC_Page(self.main_frame, self)
        self.pages["Registros"] = RecordsPage(self.main_frame, self)
        self.pages["Verificacion_MPO"] = VerificacionMPO_Page(self.main_frame, self)
        self.pages["Reportes_MPO"] = AnalisisMPOPage(self.main_frame, self)
        self.pages["RegistroWH"] = RegistroWH_Page(self.main_frame, self)
        self.pages["RegistroWHMPO"] = RegistroWHMPO_Page(self.main_frame, self)


    def show_page(self, page_name):
        title_map = {
            "Verificacion_LC_SC": "Verificación Individual (LC/SC)",
            "Reportes_LC_SC": "Análisis de O.T. (LC/SC)",
            "Verificacion_MPO": "Verificación Individual (MPO)",
            "Reportes_MPO": "Análisis de O.T. (MPO)"
        }
        display_title = title_map.get(page_name, page_name.replace("_", " "))
        self.page_title_label.config(text=display_title)
        
        for page in self.pages.values():
            page.grid_forget()
        
        page = self.pages[page_name]
        page.grid(row=1, column=0, sticky="nsew")
        if isinstance(page, RecordsPage):
            page.load_records()

    def request_password(self, callback):
        password_ingresada = simpledialog.askstring("Contraseña Requerida", 
                                                    "Ingrese la contraseña de administrador:", 
                                                    show='*', parent=self)
        if password_ingresada == self.password:
            callback()
        else:
            messagebox.showerror("Acceso Denegado", "Contraseña incorrecta.")
        
    def open_settings_window(self):
        self.request_password(lambda: SettingsWindow(self))

    def show_db_diagnostics(self):
        db_path = self.config['db_path']
        size = os.path.getsize(db_path) if os.path.exists(db_path) else 0
        messagebox.showinfo("Diagnóstico de Base de Datos",
                            f"La base de datos se está guardando en:\n\n{db_path}\n\n"
                            f"Tamaño del archivo: {size} bytes")

    def guardar_ot_configuration(self, ot_data):
        """Guarda o actualiza la configuración de una OT en la base de datos."""
        try:
            conn = sqlite3.connect(self.config['db_path'])
            cursor = conn.cursor()
            cursor.execute("""
                INSERT INTO ot_configurations (
                    ot_number, drawing_number, link, num_conectores_a,
                    fibers_per_connector_a, num_conectores_b, fibers_per_connector_b,
                    ilrl_ot_header, ilrl_serie_header, ilrl_fecha_header,
                    ilrl_hora_header, ilrl_estado_header, ilrl_conector_header, last_modified
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(ot_number) DO UPDATE SET
                    drawing_number=excluded.drawing_number, link=excluded.link,
                    num_conectores_a=excluded.num_conectores_a, fibers_per_connector_a=excluded.fibers_per_connector_a,
                    num_conectores_b=excluded.num_conectores_b, fibers_per_connector_b=excluded.fibers_per_connector_b,
                    ilrl_ot_header=excluded.ilrl_ot_header, ilrl_serie_header=excluded.ilrl_serie_header,
                    ilrl_fecha_header=excluded.ilrl_fecha_header, ilrl_hora_header=excluded.ilrl_hora_header,
                    ilrl_estado_header=excluded.ilrl_estado_header, ilrl_conector_header=excluded.ilrl_conector_header,
                    last_modified=excluded.last_modified
            """, (
                ot_data['ot_number'], ot_data['drawing_number'], ot_data['link'],
                ot_data['num_conectores_a'], ot_data['fibers_per_connector_a'],
                ot_data['num_conectores_b'], ot_data['fibers_per_connector_b'],
                ot_data.get('ilrl_ot_header', 'Work number'), ot_data.get('ilrl_serie_header', 'Serial number'),
                ot_data.get('ilrl_fecha_header', 'Date'), ot_data.get('ilrl_hora_header', 'Time'),
                ot_data.get('ilrl_estado_header', 'Alarm Status'), ot_data.get('ilrl_conector_header', 'connector label'),
                datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            ))
            conn.commit()
            conn.close()
            return True
        except Exception as e:
            messagebox.showerror("Error de Base de Datos", f"No se pudo guardar la configuración: {e}")
            return False

# --- PÁGINAS DE LA APLICACIÓN ---

class DashboardPage(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent, padding=10)

        # --- Contenido existente ---
        ttk.Label(self, text="Bienvenido al Sistema de Trazabilidad FibraTrace.", font=("Helvetica", 16)).pack(pady=20)
        ttk.Label(self, text="Selecciona una opción del menú de la izquierda para comenzar.", font=("Helvetica", 12)).pack(pady=10)

        # --- CÓDIGO AÑADIDO ---
        # Creamos una etiqueta para mostrar la versión del programa.
        # Usamos la variable global __version__ que ya tienes definida.
        version_label = ttk.Label(
            self,
            text=f"Versión {__version__}",
            font=("Helvetica", 10),
            style='secondary.TLabel'  # Un estilo sutil (texto gris) de ttkbootstrap
        )
        # Usamos .pack() con side='bottom' para anclar la etiqueta a la parte inferior.
        version_label.pack(side='bottom', pady=10, anchor='se')
        # --------------------

class VerificacionLC_SC_Page(ttk.Frame):
    def __init__(self, parent, app_instance):
        super().__init__(parent, padding=10)
        self.app = app_instance 
        self.last_ilrl_result = None 
        self.last_geo_result = None
        self.create_widgets()

    def create_widgets(self):
        container = ttk.Frame(self, style='TFrame', padding=20)
        container.pack(expand=True, fill='both')
        
        input_frame = ttk.Frame(container)
        input_frame.pack(fill='x', pady=10)
        
        ttk.Label(input_frame, text="Número de OT:", font=("Helvetica", 11, "bold")).grid(row=0, column=0, sticky='w', padx=5, pady=5)
        self.ot_entry = ttk.Entry(input_frame, width=30, font=("Helvetica", 11))
        self.ot_entry.grid(row=0, column=1, sticky='ew', padx=5, pady=5)
        
        ttk.Label(input_frame, text="Número de Serie (13 dígitos):", font=("Helvetica", 11, "bold")).grid(row=1, column=0, sticky='w', padx=5, pady=5)
        self.serie_entry = ttk.Entry(input_frame, width=30, font=("Helvetica", 11))
        self.serie_entry.grid(row=1, column=1, sticky='ew', padx=5, pady=5)
        self.serie_entry.bind("<KeyRelease>", self.verificar_cable_automatico)
        
        input_frame.columnconfigure(1, weight=1)

        verify_button = ttk.Button(container, text="Verificar Cable (Manual)", command=self.verificar_cable, style='success.TButton', padding=10)
        verify_button.pack(pady=20)

        self.result_text = tk.Text(container, height=15, width=80, wrap="word", font=("Courier New", 10), state=tk.DISABLED, relief="flat", bg="#f0f0f0")
        self.result_text.pack(fill='both', expand=True, pady=10)

        self.result_text.tag_configure("header", font=("Helvetica", 14, "bold"), foreground="#0056b3")
        self.result_text.tag_configure("bold", font=("Courier New", 10, "bold"))
        self.result_text.tag_configure("info", font=("Courier New", 10, "italic"), foreground="grey")
        self.result_text.tag_configure("APROBADO", foreground="#28a745")
        self.result_text.tag_configure("RECHAZADO", foreground="#dc3545")
        self.result_text.tag_configure("NO ENCONTRADO", foreground="#fd7e14")
        self.result_text.tag_configure("ERROR", foreground="#dc3545", font=("Courier New", 10, "bold"))
        self.result_text.tag_configure("ilrl_link", foreground="#0056b3", underline=True)
        self.result_text.tag_bind("ilrl_link", "<Button-1>", lambda e: self.show_details_window("ilrl"))
        self.result_text.tag_bind("ilrl_link", "<Enter>", lambda e: self.result_text.config(cursor="hand2"))
        self.result_text.tag_bind("ilrl_link", "<Leave>", lambda e: self.result_text.config(cursor=""))
        self.result_text.tag_configure("geo_link", foreground="#0056b3", underline=True)
        self.result_text.tag_bind("geo_link", "<Button-1>", lambda e: self.show_details_window("geo"))
        self.result_text.tag_bind("geo_link", "<Enter>", lambda e: self.result_text.config(cursor="hand2"))
        self.result_text.tag_bind("geo_link", "<Leave>", lambda e: self.result_text.config(cursor=""))

        self.result_text.tag_configure("ilrl_file_link", foreground="#4682B4", underline=True)
        self.result_text.tag_bind("ilrl_file_link", "<Button-1>", lambda e: self.open_file_location("ilrl"))
        self.result_text.tag_bind("ilrl_file_link", "<Enter>", lambda e: self.result_text.config(cursor="hand2"))
        self.result_text.tag_bind("ilrl_file_link", "<Leave>", lambda e: self.result_text.config(cursor=""))
        
        self.result_text.tag_configure("geo_file_link", foreground="#4682B4", underline=True)
        self.result_text.tag_bind("geo_file_link", "<Button-1>", lambda e: self.open_file_location("geo"))
        self.result_text.tag_bind("geo_file_link", "<Enter>", lambda e: self.result_text.config(cursor="hand2"))
        self.result_text.tag_bind("geo_file_link", "<Leave>", lambda e: self.result_text.config(cursor=""))

        self.result_text.tag_configure("ERROR", foreground="#dc3545", font=("Courier New", 10, "bold"))
        self.result_text.tag_configure("final_status_large", font=("Courier New", 14, "bold"))
        self.result_text.tag_configure("ilrl_link", foreground="#0056b3", underline=True)

        self.show_waiting_message()

    def show_waiting_message(self):
        self.result_text.config(state=tk.NORMAL)
        self.result_text.delete(1.0, tk.END)
        self.result_text.insert(tk.END, "Esperando un número de serie valido (13 digitos)", "info")
        self.result_text.config(state=tk.DISABLED)

    def _log_verification(self, log_data):
        try:
            conn = sqlite3.connect(self.app.config['db_path'])
            cursor = conn.cursor()
            cursor.execute("""
                INSERT INTO cable_verifications (
                    entry_date, serial_number, ot_number, overall_status,
                    ilrl_status, ilrl_details, geo_status, geo_details
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                log_data['serial_number'],
                log_data['ot_number'],
                log_data['overall_status'],
                log_data['ilrl_status'],
                json.dumps(log_data['ilrl_details']),
                log_data['geo_status'],
                json.dumps(log_data['geo_details'])
            ))
            conn.commit()
            conn.close()
        except Exception as e:
            messagebox.showerror("Error de Base de Datos", f"No se pudo registrar la verificación: {e}")

    # En la clase VerificacionLC_SC_Page, reemplaza este método:

    def verificar_cable_automatico(self, event=None):
        serie_raw = self.serie_entry.get().strip()
        ot_numero = self.ot_entry.get().strip()
        
        # --- MODIFICACIÓN: Regex ajustado para aceptar JMO o JRMO ---
        if re.match(r'J(R)?MO\d{13}', serie_raw, re.IGNORECASE):
            # Si escanea un código completo, extraemos números y rellenamos OT si falta
            numeros_serie = re.sub(r'[^0-9]', '', serie_raw)
            if not ot_numero:
                 self.ot_entry.insert(0, f"JMO-{numeros_serie[:9]}")
            self.verificar_cable()
            
        elif len(serie_raw) == 13 and serie_raw.isdigit():
            # Caso manual de 13 dígitos
            self.verificar_cable()
        else:
            self.show_waiting_message()

    # En la clase VerificacionLC_SC_Page, reemplaza este método:

    def verificar_cable(self, event=None):
        ot_numero = self.ot_entry.get().strip().upper()
        serie_raw = self.serie_entry.get().strip()
        
        self.result_text.config(state=tk.NORMAL)
        self.result_text.delete(1.0, tk.END)

        if not ot_numero or not serie_raw:
            self.result_text.insert(tk.END, "ERROR: Por favor ingrese OT y Número de Serie.", "ERROR")
            self.result_text.config(state=tk.DISABLED)
            return

        # 1. Normalización numérica
        serie_numerica = re.sub(r'[^0-9]', '', serie_raw)
        
        if len(serie_numerica) != 13:
            messagebox.showerror("Formato Inválido", "El número de serie debe contener 13 dígitos.")
            self.result_text.config(state=tk.DISABLED)
            return

        # 2. Construcción del Serial con Prefijo Correcto (JMO o JRMO)
        prefijo_serie = "JRMO-" if "JRMO" in serie_raw.upper() else "JMO-"
        serie_cable = f"{prefijo_serie}{serie_numerica}"
        
        # 3. Validación de coincidencia OT
        # Quitamos JMO/JRMO de la OT input para comparar solo números
        ot_parte_input = re.sub(r'[^0-9]', '', ot_numero)
        serie_ot_parte = serie_numerica[:9]

        if ot_parte_input != serie_ot_parte:
            messagebox.showerror("Error de Coincidencia", "La OT del número de serie no corresponde a la OT trabajada.")
            self.result_text.config(state=tk.DISABLED)
            return
        
        current_mode = self.app.cable_mode.get()
        self.result_text.insert(tk.END, f"Verificando cable {serie_cable} en OT {ot_numero} (Modo: {current_mode})...\n", "header")
        self.result_text.insert(tk.END, "-"*60 + "\n\n")

        # Pasamos el serie_cable (que puede ser JRMO-...) a las funciones
        self.last_ilrl_result = self.buscar_y_procesar_ilrl(ot_numero, serie_cable, current_mode)
        self.last_geo_result = self.buscar_y_procesar_geo(ot_numero, serie_cable, current_mode)
        
        self.mostrar_resultado("IL/RL", self.last_ilrl_result)
        self.mostrar_resultado("Geometría", self.last_geo_result)

        final_status = "NO ENCONTRADO"
        if self.last_ilrl_result['status'] not in ['NO ENCONTRADO', 'ERROR'] or self.last_geo_result['status'] not in ['NO ENCONTRADO', 'ERROR']:
            if self.last_ilrl_result['status'] == 'APROBADO' and self.last_geo_result['status'] == 'APROBADO':
                final_status = 'APROBADO'
            else:
                final_status = 'RECHAZADO'
        
        self.result_text.insert(tk.END, "\n" + "-"*60 + "\n")
        self.result_text.insert(tk.END, "ESTADO FINAL: ", ("bold", "final_status_large"))
        self.result_text.insert(tk.END, f"{final_status}\n", (final_status, "final_status_large"))
        
        if winsound:
            try:
                if final_status == "APROBADO":
                    winsound.Beep(1200, 200)
                elif final_status == "RECHAZADO":
                    winsound.Beep(400, 500)
            except Exception as e:
                print(f"No se pudo reproducir el sonido: {e}")
        
        self.result_text.config(state=tk.DISABLED)
        
        log_data = {
            'serial_number': serie_cable,
            'ot_number': ot_numero,
            'overall_status': final_status,
            'ilrl_status': self.last_ilrl_result['status'],
            'ilrl_details': self.last_ilrl_result,
            'geo_status': self.last_geo_result['status'],
            'geo_details': self.last_geo_result
        }
        self._log_verification(log_data)

    # En la clase VerificacionLC_SC_Page, reemplaza este método:

    def mostrar_resultado(self, tipo, resultado):
        link_tag = "ilrl_link" if tipo == "IL/RL" else "geo_link"
        file_link_tag = "ilrl_file_link" if tipo == "IL/RL" else "geo_file_link" # <-- Nuevo tag

        self.result_text.insert(tk.END, f"Análisis {tipo}:\n", "bold")
        self.result_text.insert(tk.END, f"  Estado: ")
        
        # Aplicar el tag de ESTADO (para detalles)
        self.result_text.insert(tk.END, f"{resultado['status']}", (resultado['status'], link_tag))
        
        # --- INICIO DE LA MODIFICACIÓN ---
        details = resultado['details']
        if "Archivo:" in details:
            details_text, file_text = details.rsplit("Archivo:", 1)
            file_text = "Archivo:" + file_text
            
            self.result_text.insert(tk.END, f"\n  Detalles: {details_text.strip()}")
            # Aplicar el tag de ARCHIVO (para explorador)
            self.result_text.insert(tk.END, f"\n  {file_text.strip()}", (file_link_tag)) 
            self.result_text.insert(tk.END, "\n\n")
        else:
            self.result_text.insert(tk.END, f"\n  Detalles: {details}\n\n")
        # --- FIN DE LA MODIFICACIÓN ---

    def buscar_y_procesar_ilrl(self, ot, serie, mode):
        # Lista de rutas a buscar (Primaria y Secundaria)
        rutas_base = [self.app.config['ruta_base_ilrl']]
        if self.app.config.get('ruta_base_ilrl_2'):
            rutas_base.append(self.app.config['ruta_base_ilrl_2'])

        serie_terminacion = serie[-4:]
        candidatos = []

        # Buscar en todas las rutas configuradas
        for ruta_base in rutas_base:
            if not os.path.isdir(ruta_base): continue
            
            ruta_ot = os.path.join(ruta_base, ot)
            if os.path.isdir(ruta_ot):
                # Buscar archivos que coincidan con la terminación
                for root, _, files in os.walk(ruta_ot):
                    for f in files:
                        if f.endswith('.xlsx') and not f.startswith('~$') and serie_terminacion in f:
                            full_path = os.path.join(root, f)
                            candidatos.append(full_path)

        if not candidatos:
            return {'status': 'NO ENCONTRADO', 'details': f'Ningún archivo con terminación "{serie_terminacion}" en las rutas configuradas.', 'raw_data': []}
        
        # Si hay duplicados o archivos en ambas rutas, tomamos el más reciente
        archivo_a_procesar = max(candidatos, key=os.path.getmtime)
        
        return self.procesar_archivo_ilrl(archivo_a_procesar, mode)

    def buscar_y_procesar_geo(self, ot, serie, mode):
        # 1. Identificar rutas configuradas
        rutas_base = [self.app.config['ruta_base_geo']]
        if self.app.config.get('ruta_base_geo_2'):
            rutas_base.append(self.app.config['ruta_base_geo_2'])

        archivos_candidatos = []

        # 2. Buscar archivos de la OT en TODAS las rutas
        for ruta_base in rutas_base:
            if not os.path.isdir(ruta_base): continue
            
            # Buscamos archivos que contengan el número de OT en el nombre
            encontrados = [os.path.join(ruta_base, f) for f in os.listdir(ruta_base) 
                           if f.lower().endswith(('.xlsx', '.xls')) 
                           and not f.startswith('~$') 
                           and ot in f]
            
            if encontrados:
                # De cada ruta, nos interesa el archivo más reciente de esa OT 
                # (Por si hay versiones viejas en la misma carpeta)
                archivo_reciente_ruta = max(encontrados, key=os.path.getmtime)
                archivos_candidatos.append(archivo_reciente_ruta)

        if not archivos_candidatos:
            return {'status': 'NO ENCONTRADO', 'details': f'Ningún archivo para la OT "{ot}" en las rutas configuradas.', 'raw_data': []}

        # 3. Procesar TODOS los archivos encontrados y fusionar los datos
        return self.procesar_multiples_archivos_geo(archivos_candidatos, serie, mode)

    # En la clase VerificacionLC_SC_Page, reemplaza estos dos métodos:

    def procesar_archivo_ilrl(self, ruta, mode): # Ahora recibe 'mode'
        try:
            expected_count = 2 if mode == "Simplex" else 4 # Usa el 'mode' recibido
            
            df = pd.read_excel(ruta, header=None)
            rows = df.iloc[12:]
            # ... (el resto del código de esta función no cambia)
            col7_vals = rows[7].dropna().astype(str).str.upper()
            col8_vals = rows[8].dropna().astype(str).str.upper()
            count_col7 = col7_vals.isin(['PASS', 'FAIL']).sum()
            count_col8 = col8_vals.isin(['PASS', 'FAIL']).sum()
            result_col = 8 if count_col8 >= count_col7 else 7
            results = rows[result_col].dropna().astype(str).str.upper().tolist()
            valid_results = [r for r in results if r in ['PASS', 'FAIL']]
            if not valid_results: 
                return {'status': 'RECHAZADO', 'details': 'No se encontraron mediciones.', 'raw_data': []}
            all_pass = all(r == 'PASS' for r in valid_results)
            count_ok = len(valid_results) == expected_count
            status = 'APROBADO' if all_pass and count_ok else 'RECHAZADO'
            details = f"{len(valid_results)}/{expected_count} mediciones encontradas. "
            if not count_ok:
                details += f"Se esperaban {expected_count} para modo {mode}. "
            if not all_pass:
                details += f"{valid_results.count('FAIL')} mediciones con FALLA. "
            details += f"Archivo: {os.path.basename(ruta)}"
            raw_data = [{'linea': i + 1, 'resultado': res} for i, res in enumerate(valid_results)]
            return {'status': status, 'details': details, 'raw_data': raw_data, 'file_path': ruta}
        except Exception as e:
            return {'status': 'ERROR', 'details': f'Fallo al procesar {os.path.basename(ruta)}: {e}', 'raw_data': []}

    # En la clase VerificacionLC_SC_Page, reemplaza este método:

    def procesar_archivo_geo(self, ruta, serie_objetivo, mode):
        try:
            puntas_requeridas = ['1', '2'] if mode == "Simplex" else ['1', '2', '3', '4']

            df = pd.read_excel(ruta, header=None, skiprows=12)
            
            # Función auxiliar interna actualizada para JRMO
            def normalize_serial_geo(s):
                if not isinstance(s, str): return "", ""
                s_upper = s.strip().upper()
                # --- MODIFICACIÓN DE REGEX: Acepta JMO o JRMO ---
                match = re.search(r'(J(?:R)?MO-?\d{13}|\d{13})(-?([1-4R][1-4]?))?', s_upper)
                if match:
                    # Extraemos SOLO los números del serial base para comparar
                    base_serial_raw = match.group(1)
                    base_serial_numeric = re.sub(r'[^0-9]', '', base_serial_raw)
                    
                    punta = match.group(3) if match.group(3) else "N/A"
                    return base_serial_numeric, punta
                
                # Fallback: intentar extraer 13 digitos si el regex falla
                fallback_match = re.search(r'(\d{13})', s_upper)
                if fallback_match:
                     return fallback_match.group(1), "N/A"
                     
                return "", "N/A"

            # --- MODIFICACIÓN: Usamos solo la parte numérica del input para buscar ---
            serie_objetivo_norm = re.sub(r'[^0-9]', '', serie_objetivo)
            # -----------------------------------------------------------------------

            punta_results = []
            
            for _, row in df.iterrows():
                # raw_serial ahora será solo números (los 13 dígitos) gracias a la función de arriba
                raw_serial_numeric, punta = normalize_serial_geo(str(row[0]))
                
                # Comparamos números con números
                if serie_objetivo_norm == raw_serial_numeric and punta != "N/A":
                    result = str(row[6]).upper() if len(row) > 6 and pd.notna(row[6]) else "SIN_DATO"
                    punta_results.append({'punta': punta, 'resultado': result, 'fuente': str(row[0])})
            
            if not punta_results:
                return {'status': 'NO ENCONTRADO', 'details': 'No hay mediciones para este serial.', 'raw_data': []}

            found_puntas = {p['punta'].replace('R',''): p['resultado'] for p in punta_results}
            status = 'APROBADO'
            missing_puntas = []
            
            for p_req in puntas_requeridas:
                if p_req not in found_puntas:
                    status = 'RECHAZADO'
                    missing_puntas.append(p_req)
                elif found_puntas[p_req] != 'PASS':
                    status = 'RECHAZADO'
            
            pass_count = sum(1 for p in found_puntas.values() if p == 'PASS')
            details = f"{pass_count}/{len(puntas_requeridas)} puntas requeridas OK."
            
            if missing_puntas:
                details += f" Faltan: {', '.join(missing_puntas)}."
            details += f" Archivo: {os.path.basename(ruta)}"
            
            return {'status': status, 'details': details, 'raw_data': punta_results, 'file_path': ruta}

        except Exception as e:
            return {'status': 'ERROR', 'details': f'Fallo al procesar {os.path.basename(ruta)}: {e}', 'raw_data': []}

    def show_details_window(self, analysis_type):
        if analysis_type == "ilrl":
            data = self.last_ilrl_result
            title = "Detalles de Análisis IL/RL"
        elif analysis_type == "geo":
            data = self.last_geo_result
            title = "Detalles de Análisis de Geometría"
        else:
            return
        if not data or not data.get('raw_data'):
            messagebox.showinfo(title, "No hay datos detallados para mostrar.")
            return
        
        # --- LÍNEA AÑADIDA: Inyectamos el N/S desde el campo de texto ---
        data['serial_number'] = self.serie_entry.get().strip()
        # -----------------------------------------------------------
        
        DetailsWindow(self, title, data, analysis_type)

    # En la clase VerificacionLC_SC_Page, reemplaza este método:

    # En la clase VerificacionLC_SC_Page, reemplaza este método:

    def open_file_location(self, analysis_type):
        """Abre la carpeta que contiene el archivo de reporte."""
        if analysis_type == 'ilrl':
            data = self.last_ilrl_result
        elif analysis_type == 'geo':
            data = self.last_geo_result
        else:
            return

        if not data or not data.get('file_path'):
            messagebox.showinfo("Error", "No se encontró la ruta del archivo.", parent=self)
            return
            
        file_path = data['file_path']
        if not os.path.exists(file_path):
            messagebox.showerror("Error", f"La ruta del archivo no existe:\n{file_path}", parent=self)
            return
            
        try:
            # --- INICIO DE LA CORRECCIÓN ---
            # 1. Obtener el DIRECTORIO que contiene el archivo
            folder_path = os.path.dirname(file_path)
            
            # 2. Asegurarse de que la ruta sea absoluta
            path_absoluto = os.path.abspath(folder_path)

            if not os.path.isdir(path_absoluto):
                messagebox.showerror("Error", f"La ruta de la carpeta no existe:\n{path_absoluto}", parent=self)
                return

            # 3. Usar os.startfile() - es la forma más robusta en Windows
            os.startfile(path_absoluto)
            # --- FIN DE LA CORRECCIÓN ---
            
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo abrir la carpeta del archivo:\n{e}", parent=self)
    
    def procesar_multiples_archivos_geo(self, lista_archivos, serie_objetivo, mode):
        try:
            puntas_requeridas = ['1', '2'] if mode == "Simplex" else ['1', '2', '3', '4']
            
            # Regex para extraer datos del nombre en Excel (igual que antes)
            def normalize_serial_geo(s):
                if not isinstance(s, str): return "", ""
                s_upper = s.strip().upper()
                match = re.search(r'(J(?:R)?MO-?\d{13}|\d{13})(-?([1-4R][1-4]?))?', s_upper)
                if match:
                    base_serial_raw = match.group(1)
                    base_serial_numeric = re.sub(r'[^0-9]', '', base_serial_raw)
                    punta = match.group(3) if match.group(3) else "N/A"
                    return base_serial_numeric, punta
                
                fallback_match = re.search(r'(\d{13})', s_upper)
                if fallback_match:
                     return fallback_match.group(1), "N/A"
                return "", "N/A"

            serie_objetivo_norm = re.sub(r'[^0-9]', '', serie_objetivo)
            
            # Contenedores para la fusión de datos
            todas_mediciones = [] # Lista cruda para logs
            puntas_encontradas_map = {} # Diccionario para lógica de prioridad (Punta -> Info)

            archivos_usados = set()

            # --- ITERAR SOBRE CADA ARCHIVO ENCONTRADO (JWS1-1, JWS1-2, etc.) ---
            for ruta in lista_archivos:
                try:
                    df = pd.read_excel(ruta, header=None, skiprows=12)
                    
                    for _, row in df.iterrows():
                        raw_serial_numeric, punta = normalize_serial_geo(str(row[0]))
                        
                        # Si coincide el número de serie y la punta es válida
                        if serie_objetivo_norm == raw_serial_numeric and punta != "N/A":
                            
                            resultado = str(row[6]).upper() if len(row) > 6 and pd.notna(row[6]) else "SIN_DATO"
                            fecha_raw = row[3] if len(row) > 3 else None # Intentar capturar fecha si existe columna
                            
                            # Normalizamos punta (quitar R para comparar prioridades)
                            punta_limpia = punta.replace('R', '')
                            es_retrabajo = 'R' in punta

                            datos_medicion = {
                                'punta_original': punta,
                                'punta_limpia': punta_limpia,
                                'resultado': resultado,
                                'fuente': str(row[0]),
                                'archivo': os.path.basename(ruta),
                                'es_retrabajo': es_retrabajo
                            }
                            
                            todas_mediciones.append(datos_medicion)
                            archivos_usados.add(os.path.basename(ruta))

                            # --- LÓGICA DE FUSIÓN (Merge) ---
                            # Si ya tenemos esta punta, ¿la reemplazamos?
                            # Prioridad: Retrabajo (R) > Normal
                            # Si ambos son R o ambos Normales, el último leído (suponiendo orden de archivos)
                            # Idealmente usaríamos fecha, pero aquí asumimos que R mata a normal.
                            
                            if punta_limpia not in puntas_encontradas_map:
                                puntas_encontradas_map[punta_limpia] = datos_medicion
                            else:
                                existente = puntas_encontradas_map[punta_limpia]
                                # Si el nuevo es Retrabajo y el viejo no, reemplazamos
                                if es_retrabajo and not existente['es_retrabajo']:
                                    puntas_encontradas_map[punta_limpia] = datos_medicion
                                # Si ambos son iguales en prioridad, actualizamos (asumiendo lectura secuencial o fechas en el futuro)
                                elif es_retrabajo == existente['es_retrabajo']:
                                    puntas_encontradas_map[punta_limpia] = datos_medicion

                except Exception as e:
                    print(f"Error leyendo archivo {ruta}: {e}")
                    # Continuamos con el siguiente archivo si uno falla

            # --- EVALUACIÓN FINAL ---
            if not puntas_encontradas_map:
                return {'status': 'NO ENCONTRADO', 'details': 'No hay mediciones para este serial en los archivos revisados.', 'raw_data': []}

            status = 'APROBADO'
            missing_puntas = []
            
            for p_req in puntas_requeridas:
                if p_req not in puntas_encontradas_map:
                    status = 'RECHAZADO'
                    missing_puntas.append(p_req)
                elif puntas_encontradas_map[p_req]['resultado'] != 'PASS':
                    status = 'RECHAZADO'
            
            pass_count = sum(1 for p in puntas_requeridas if p in puntas_encontradas_map and puntas_encontradas_map[p]['resultado'] == 'PASS')
            
            details = f"{pass_count}/{len(puntas_requeridas)} puntas OK."
            if missing_puntas:
                details += f" Faltan: {', '.join(missing_puntas)}."
            
            # Listar archivos fuente
            archivos_str = ", ".join(list(archivos_usados))
            details += f" Fuentes: {archivos_str}"
            
            # Formatear raw_data para mostrar en la ventana de detalles
            raw_data_formatted = [{'punta': m['punta_original'], 'resultado': m['resultado'], 'fuente': f"{m['fuente']} ({m['archivo']})"} for m in todas_mediciones]

            # Usamos la ruta del primer archivo encontrado para "abrir ubicación" si el usuario hace click
            ruta_principal = lista_archivos[0] if lista_archivos else ""

            return {'status': status, 'details': details, 'raw_data': raw_data_formatted, 'file_path': ruta_principal}

        except Exception as e:
            return {'status': 'ERROR', 'details': f'Fallo al procesar múltiples archivos: {e}', 'raw_data': []}

class RegistroWHMPO_Page(ttk.Frame):
    def __init__(self, parent, app_instance):
        super().__init__(parent, padding=20)
        self.app = app_instance

        # --- Creamos los widgets para esta página ---
        container = ttk.Frame(self, style='TFrame')
        container.pack(expand=True, fill='both', pady=20)

        # Un texto descriptivo
        ttk.Label(
            container,
            text="Módulo de Registro en Almacén (MPO)",
            font=("Helvetica", 16, "bold")
        ).pack(pady=10)

        ttk.Label(
            container,
            text="Presiona el botón para abrir el archivo de registro de series para empaque de productos MPO.",
            font=("Helvetica", 11),
            wraplength=500
        ).pack(pady=10)

        # El botón que abrirá el archivo de Excel
        open_button = ttk.Button(
            container,
            text="Abrir Registro WH (MPO)", # Texto del botón actualizado
            command=self.abrir_registro_mpo,
            style='success.TButton',
            padding=15
        )
        open_button.pack(pady=30)

    def abrir_registro_mpo(self):
        """
        Busca y abre el archivo de registro para MPO.
        Asumiremos un nombre de archivo como 'MPORegistroWH.xlsm'.
        """
        # Puedes cambiar 'MPORegistroWH.xlsm' por el nombre real de tu archivo de MPO
        file_name = "MPORegistroWH.xlsm"
        try:
            base_path = os.path.dirname(sys.argv[0])
            file_path = os.path.join(base_path, file_name)

            if os.path.exists(file_path):
                os.startfile(file_path)
            else:
                messagebox.showerror(
                    "Archivo no Encontrado",
                    f"No se pudo encontrar el archivo '{file_name}'.\n\nAsegúrate de que esté en la misma carpeta que el programa."
                )
        except Exception as e:
            messagebox.showerror("Error", f"Ocurrió un error al intentar abrir el archivo:\n\n{e}")

# Coloca esta nueva clase junto a las otras clases de páginas (ej. después de DetailsWindow)

class RegistroWH_Page(ttk.Frame):
    def __init__(self, parent, app_instance):
        super().__init__(parent, padding=20)
        self.app = app_instance

        # --- Creamos los widgets para esta página ---
        container = ttk.Frame(self, style='TFrame')
        container.pack(expand=True, fill='both', pady=20)

        # Un texto descriptivo
        ttk.Label(
            container,
            text="Módulo de Registro en Almacén (WH)",
            font=("Helvetica", 16, "bold")
        ).pack(pady=10)

        ttk.Label(
            container,
            text="Presiona el botón para abrir el archivo de registro de series para empaque.",
            font=("Helvetica", 11),
            wraplength=500  # Ajusta el texto si la ventana es estrecha
        ).pack(pady=10)

        # El botón que abrirá el archivo de Excel
        open_button = ttk.Button(
            container,
            text="Abrir Registro WH (MP1RegistroWH.xlsm)",
            command=self.abrir_registro, # Llama a la función de abajo
            style='success.TButton',
            padding=15
        )
        open_button.pack(pady=30)

    def abrir_registro(self):
        """
        Encuentra y abre el archivo MP1RegistroWH.xlsm que está en la misma
        carpeta que el ejecutable del programa.
        """
        try:
            # sys.argv[0] nos da la ruta del script o del .exe
            # os.path.dirname() nos da la carpeta que lo contiene
            base_path = os.path.dirname(sys.argv[0])

            # Construimos la ruta completa al archivo de Excel
            file_path = os.path.join(base_path, "MP1RegistroWH.xlsm")

            if os.path.exists(file_path):
                # os.startfile es la forma recomendada en Windows para abrir un archivo
                # con su programa predeterminado (en este caso, Excel).
                os.startfile(file_path)
            else:
                # Si no encuentra el archivo, avisa al usuario.
                messagebox.showerror(
                    "Archivo no Encontrado",
                    f"No se pudo encontrar el archivo 'MP1RegistroWH.xlsm'.\n\nAsegúrate de que esté en la misma carpeta que el programa."
                )
        except Exception as e:
            messagebox.showerror("Error", f"Ocurrió un error al intentar abrir el archivo:\n\n{e}")

# --- PÁGINA PARA REPORTES DE O.T. (LC/SC) ---
class Reportes_LC_SC_Page(ttk.Frame):
    def __init__(self, parent, app_instance):
        super().__init__(parent, padding=20)
        self.app = app_instance
        self.analisis_ilrl = AnalisisILRL()
        self.analisis_geo = AnalisisGEO()
        self.create_widgets()

    def create_widgets(self):
        container = ttk.Frame(self, style='TFrame')
        container.pack(expand=True, fill='both')

        config_frame = ttk.LabelFrame(container, text="Parámetros del Análisis", padding=15)
        config_frame.pack(fill='x', pady=(0, 20))
        config_frame.columnconfigure(1, weight=1)

        # ILRL Folder
        ttk.Label(config_frame, text="Carpeta O.T. (para ILRL):", font="-weight bold").grid(row=0, column=0, sticky='w', pady=5, padx=5)
        self.folder_path_var = tk.StringVar()
        folder_entry = ttk.Entry(config_frame, textvariable=self.folder_path_var, state="readonly", width=70)
        folder_entry.grid(row=0, column=1, sticky='ew', pady=5, padx=5)
        ttk.Button(config_frame, text="Seleccionar Carpeta...", command=self.select_folder, style='outline.TButton').grid(row=0, column=2, sticky='w', pady=5, padx=5)
        
        # GEO File
        ttk.Label(config_frame, text="Archivo Geometría:", font="-weight bold").grid(row=1, column=0, sticky='w', pady=5, padx=5)
        self.geo_file_path_var = tk.StringVar()
        geo_file_entry = ttk.Entry(config_frame, textvariable=self.geo_file_path_var, state="readonly", width=70)
        geo_file_entry.grid(row=1, column=1, sticky='ew', pady=5, padx=5)
        ttk.Button(config_frame, text="Seleccionar Archivo...", command=self.select_geo_file, style='outline.TButton').grid(row=1, column=2, sticky='w', pady=5, padx=5)

        ttk.Separator(config_frame, orient='horizontal').grid(row=2, column=0, columnspan=3, sticky='ew', pady=10)

        # OT and Quantity
        ttk.Label(config_frame, text="Número de O.T.:", font="-weight bold").grid(row=3, column=0, sticky='w', pady=5, padx=5)
        self.ot_var = tk.StringVar()
        ttk.Entry(config_frame, textvariable=self.ot_var).grid(row=3, column=1, sticky='w', pady=5, padx=5)

        ttk.Label(config_frame, text="Total de cables esperados:", font="-weight bold").grid(row=4, column=0, sticky='w', pady=5, padx=5)
        self.total_cables_var = tk.StringVar()
        ttk.Entry(config_frame, textvariable=self.total_cables_var).grid(row=4, column=1, sticky='w', pady=5, padx=5)

        action_frame = ttk.Frame(container)
        action_frame.pack(fill='x', pady=10)
        ttk.Button(action_frame, text="Generar Reporte IL/RL", command=lambda: self.run_analysis("ilrl"), style='success.TButton', padding=10).pack(side='left', padx=10, expand=True)
        ttk.Button(action_frame, text="Generar Reporte Geometría", command=lambda: self.run_analysis("geo"), style='success.TButton', padding=10).pack(side='left', padx=10, expand=True)

        result_frame = ttk.LabelFrame(container, text="Estado del Análisis", padding=15)
        result_frame.pack(fill='both', expand=True, pady=(20, 0))
        self.progress_bar = ttk.Progressbar(result_frame, mode='determinate')
        self.progress_bar.pack(fill='x', pady=10)
        self.result_label = ttk.Label(result_frame, text="Listo para iniciar el análisis.", wraplength=700)
        self.result_label.pack(fill='x', pady=10)

    def select_folder(self):
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.folder_path_var.set(folder_selected)
            match = re.search(r'JMO-(\d{9})', folder_selected, re.IGNORECASE)
            if match:
                self.ot_var.set(match.group(1))
    
    def select_geo_file(self):
        file_selected = filedialog.askopenfilename(
            title="Seleccionar archivo de Geometría",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if file_selected:
            self.geo_file_path_var.set(file_selected)
            # Intentar autodetectar la OT del nombre del archivo también
            match = re.search(r'JMO-(\d{9})', os.path.basename(file_selected), re.IGNORECASE)
            if match and not self.ot_var.get():
                self.ot_var.set(match.group(1))

    def run_analysis(self, analysis_type):
        ot = self.ot_var.get().strip()
        total_str = self.total_cables_var.get().strip()
        
        # Validación
        if not ot:
            messagebox.showerror("Error", "Por favor, ingresa el número de O.T.", parent=self)
            return
        if not total_str.isdigit() or int(total_str) <= 0:
            messagebox.showerror("Error", "Por favor, ingresa un número válido para el total de cables.", parent=self)
            return
            
        total = int(total_str)
        
        self.progress_bar["value"] = 0
        self.result_label.config(text=f"Iniciando análisis de {analysis_type.upper()}...")
        
        # Ahora pasamos la configuración completa (self.app.config) a los threads
        if analysis_type == "ilrl":
            thread = threading.Thread(target=self._run_ilrl_thread, args=(ot, self.app.config, total), daemon=True)
        else: # geo
            thread = threading.Thread(target=self._run_geo_thread, args=(ot, self.app.config, total), daemon=True)
        
        thread.start()

    def _run_ilrl_thread(self, ot, config, total):
        try:
            # Llamamos al nuevo método procesar_ilrl que usa rutas automáticas
            ruta, errores = self.analisis_ilrl.procesar_ilrl(ot, config, total, self.update_progress)
            self.after(0, self.show_result, ruta, errores, "IL/RL")
        except Exception:
            self.after(0, self.show_result, None, f"Error crítico en el análisis IL/RL:\n{traceback.format_exc()}", "IL/RL")

    def _run_geo_thread(self, ot, config, total):
        try:
            current_mode = self.app.cable_mode.get()
            # Llamamos al nuevo método procesar_geo que busca en múltiples rutas
            ruta, errores = self.analisis_geo.procesar_geo(ot, config, total, self.update_progress, mode=current_mode)
            self.after(0, self.show_result, ruta, errores, "Geometría")
        except Exception:
            self.after(0, self.show_result, None, f"Error crítico en el análisis de Geometría:\n{traceback.format_exc()}", "Geometría")

    def update_progress(self, value):
        self.progress_bar["value"] = value

    def show_result(self, ruta_generada, errores, tipo_analisis):
        self.progress_bar["value"] = 100
        if ruta_generada:
            self.result_label.config(text=f"Reporte de {tipo_analisis} generado con éxito.")
            if messagebox.askyesno("Éxito", f"Reporte de {tipo_analisis} generado con éxito.\n\n"
                                          f"Guardado en:\n{ruta_generada}\n\n"
                                          "¿Deseas abrir el archivo ahora?", parent=self):
                self.analisis_ilrl.abrir_archivo(ruta_generada)
        else:
            self.result_label.config(text=f"El análisis de {tipo_analisis} falló.")
            messagebox.showerror("Análisis Fallido", f"No se pudo generar el reporte de {tipo_analisis}.\n\n"
                                                   f"Detalles:\n{errores}", parent=self)
        if errores and ruta_generada:
             messagebox.showwarning("Advertencia", f"El reporte se generó, pero se encontraron los siguientes problemas:\n\n{errores}", parent=self)

# --- PÁGINA PARA ANÁLISIS DE O.T. (MPO) ---
class AnalisisMPOPage(ttk.Frame):
    def __init__(self, parent, app_instance):
        super().__init__(parent, padding=20)
        self.app = app_instance
        self.create_widgets()

    def create_widgets(self):
        container = ttk.Frame(self, style='TFrame')
        container.pack(expand=True, fill='both')

        # --- Frame de Configuración ---
        config_frame = ttk.LabelFrame(container, text="Parámetros del Análisis MPO", padding=15)
        config_frame.pack(fill='x', pady=(0, 20))
        config_frame.columnconfigure(1, weight=1)

        # ILRL File
        ttk.Label(config_frame, text="Archivo IL/RL MPO:", font="-weight bold").grid(row=0, column=0, sticky='w', pady=5, padx=5)
        self.ilrl_file_var = tk.StringVar()
        ttk.Entry(config_frame, textvariable=self.ilrl_file_var, state="readonly").grid(row=0, column=1, sticky='ew', pady=5, padx=5)
        ttk.Button(config_frame, text="Seleccionar Archivo...", command=self.select_ilrl_file, style='outline.TButton').grid(row=0, column=2, sticky='w', pady=5, padx=5)

        # GEO File
        ttk.Label(config_frame, text="Archivo Geometría MPO:", font="-weight bold").grid(row=1, column=0, sticky='w', pady=5, padx=5)
        self.geo_file_var = tk.StringVar()
        ttk.Entry(config_frame, textvariable=self.geo_file_var, state="readonly").grid(row=1, column=1, sticky='ew', pady=5, padx=5)
        ttk.Button(config_frame, text="Seleccionar Archivo...", command=self.select_geo_file, style='outline.TButton').grid(row=1, column=2, sticky='w', pady=5, padx=5)

        # Polarity Folder
        ttk.Label(config_frame, text="Carpeta Polaridad MPO:", font="-weight bold").grid(row=2, column=0, sticky='w', pady=5, padx=5)
        self.polaridad_folder_var = tk.StringVar()
        ttk.Entry(config_frame, textvariable=self.polaridad_folder_var, state="readonly").grid(row=2, column=1, sticky='ew', pady=5, padx=5)
        ttk.Button(config_frame, text="Seleccionar Carpeta...", command=self.select_polaridad_folder, style='outline.TButton').grid(row=2, column=2, sticky='w', pady=5, padx=5)
        
        ttk.Separator(config_frame, orient='horizontal').grid(row=3, column=0, columnspan=3, sticky='ew', pady=10)

        # OT and Quantity
        ttk.Label(config_frame, text="Número de O.T. (JMO-):", font="-weight bold").grid(row=4, column=0, sticky='w', pady=5, padx=5)
        self.ot_var = tk.StringVar()
        ttk.Entry(config_frame, textvariable=self.ot_var).grid(row=4, column=1, sticky='w', pady=5, padx=5)

        ttk.Label(config_frame, text="Total de cables esperados:", font="-weight bold").grid(row=5, column=0, sticky='w', pady=5, padx=5)
        self.total_cables_var = tk.StringVar()
        ttk.Entry(config_frame, textvariable=self.total_cables_var).grid(row=5, column=1, sticky='w', pady=5, padx=5)

        # --- Frame de Acciones ---
        action_frame = ttk.Frame(container)
        action_frame.pack(fill='x', pady=10)
        
        ttk.Button(action_frame, text="Generar Reporte IL/RL", command=lambda: self.run_analysis("ilrl"), style='info.TButton', padding=10).pack(side='left', padx=10, expand=True)
        ttk.Button(action_frame, text="Generar Reporte Geometría", command=lambda: self.run_analysis("geo"), style='info.TButton', padding=10).pack(side='left', padx=10, expand=True)
        ttk.Button(action_frame, text="Generar Reporte Polaridad", command=lambda: self.run_analysis("polaridad"), style='info.TButton', padding=10).pack(side='left', padx=10, expand=True)

        # --- Frame de Progreso y Resultados ---
        result_frame = ttk.LabelFrame(container, text="Estado del Análisis", padding=15)
        result_frame.pack(fill='both', expand=True, pady=(20, 0))

        self.progress_bar = ttk.Progressbar(result_frame, mode='determinate')
        self.progress_bar.pack(fill='x', pady=10)
        
        self.result_label = ttk.Label(result_frame, text="Listo para iniciar el análisis MPO.", wraplength=700)
        self.result_label.pack(fill='x', pady=10)

    def _select_file(self, var, title):
        file_selected = filedialog.askopenfilename(title=title, filetypes=[("Excel files", "*.xlsx *.xls")])
        if file_selected:
            var.set(file_selected)
            match = re.search(r'(JMO-\d{9})', os.path.basename(file_selected), re.IGNORECASE)
            if match and not self.ot_var.get():
                self.ot_var.set(match.group(1).upper())

    def select_ilrl_file(self):
        self._select_file(self.ilrl_file_var, "Seleccionar archivo de IL/RL MPO")

    def select_geo_file(self):
        self._select_file(self.geo_file_var, "Seleccionar archivo de Geometría MPO")
        
    def select_polaridad_folder(self):
        folder_selected = filedialog.askdirectory(title="Seleccionar carpeta de Polaridad (ej. JMO-250800005)")
        if folder_selected:
            self.polaridad_folder_var.set(folder_selected)
            match = re.search(r'JMO-(\d{9})', os.path.basename(folder_selected), re.IGNORECASE)
            if match and not self.ot_var.get():
                # Corrección para consistencia: sugerir sin prefijo
                self.ot_var.set(match.group(1))

    def run_analysis(self, analysis_type):
        ot_number = self.ot_var.get().strip().upper()
        if not ot_number.startswith("JMO-"):
            ot_number = f"JMO-{ot_number}"
            
        total_str = self.total_cables_var.get()

        # Validaciones comunes
        if not ot_number:
            messagebox.showerror("Error", "Por favor, ingresa el número de O.T.", parent=self)
            return
        if not total_str.isdigit() or int(total_str) <= 0:
            messagebox.showerror("Error", "Por favor, ingresa un número válido para el total de cables.", parent=self)
            return
        total = int(total_str)
        
        self.progress_bar["value"] = 0
        self.result_label.config(text=f"Iniciando análisis de {analysis_type.upper()} para MPO...")
        
        # Validaciones y arranque de hilos específicos
        if analysis_type == "ilrl":
            file_path = self.ilrl_file_var.get()
            if not file_path:
                messagebox.showerror("Error", "Por favor, selecciona el archivo de IL/RL para MPO.", parent=self)
                return
            thread = threading.Thread(target=self._run_ilrl_mpo_thread, args=(file_path, ot_number, total), daemon=True)
            thread.start()
        elif analysis_type == "geo":
            file_path = self.geo_file_var.get()
            if not file_path:
                messagebox.showerror("Error", "Por favor, selecciona el archivo de Geometría para MPO.", parent=self)
                return
            thread = threading.Thread(target=self._run_geo_mpo_thread, args=(file_path, ot_number, total), daemon=True)
            thread.start()
        elif analysis_type == "polaridad":
            folder_path = self.polaridad_folder_var.get()
            if not folder_path:
                messagebox.showerror("Error", "Por favor, selecciona la carpeta de Polaridad para MPO.", parent=self)
                return
            thread = threading.Thread(target=self._run_polaridad_mpo_thread, args=(folder_path, ot_number, total), daemon=True)
            thread.start()

    def _cargar_ot_configuration(self, ot_number):
        try:
            conn = sqlite3.connect(self.app.config['db_path'])
            conn.row_factory = sqlite3.Row
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM ot_configurations WHERE ot_number = ?", (ot_number,))
            row = cursor.fetchone()
            conn.close()
            return dict(row) if row else None
        except Exception as e:
            messagebox.showerror("Error de Base de Datos", f"No se pudo cargar la configuración de la OT: {e}", parent=self)
            return None

    def _run_ilrl_mpo_thread(self, file_path, ot_number, total_cables):
        try:
            self.update_progress(5)
            ot_config = self._cargar_ot_configuration(ot_number)
            if not ot_config:
                self.after(0, self.show_result, None, f"No se encontró configuración para la OT {ot_number}. Por favor, configúrela primero.", "IL/RL MPO")
                return

            self.update_progress(15)
            df = pd.read_excel(file_path, sheet_name="Results")
            h = {k: ot_config.get(v, d) for k, v, d in [('serie', 'ilrl_serie_header', 'Serial number'), ('estado', 'ilrl_estado_header', 'Alarm Status'), ('conector', 'ilrl_conector_header', 'connector label'), ('fecha', 'ilrl_fecha_header', 'Date'), ('hora', 'ilrl_hora_header', 'Time')]}
            
            if any(header not in df.columns for header in h.values()):
                missing = [k for k, v in h.items() if v not in df.columns]
                self.after(0, self.show_result, None, f"El archivo no contiene los encabezados esperados: {', '.join(missing)}.", "IL/RL MPO")
                return
            
            self.update_progress(30)
            df[h['serie']] = df[h['serie']].astype(str)
            df['timestamp'] = pd.to_datetime(df[h['fecha']].astype(str) + ' ' + df[h['hora']].astype(str), errors='coerce')
            all_series = df[h['serie']].dropna().unique()

            reporte_data = []
            rechazados = []
            
            num_conectores_a = ot_config.get('num_conectores_a', 1)
            fibras_por_conector_a = ot_config.get('fibers_per_connector_a', 12)
            num_conectores_b = ot_config.get('num_conectores_b', 1)
            fibras_por_conector_b = ot_config.get('fibers_per_connector_b', 12)
            total_fibras_esperadas = (num_conectores_a * fibras_por_conector_a) + (num_conectores_b * fibras_por_conector_b)

            for i, serie in enumerate(all_series):
                df_cable = df[df[h['serie']] == serie].copy()
                
                estado_final = "APROBADO"
                detalles_list = []

                # --- LÓGICA DE SELECCIÓN DE ÚLTIMAS MEDICIONES ---
                df_cable.sort_values(by='timestamp', ascending=False, inplace=True)

                # Filtrar las últimas N mediciones para cada lado
                df_lado_a = df_cable[df_cable[h['conector']] == 'A'].head(fibras_por_conector_a * num_conectores_a)
                df_lado_b = df_cable[df_cable[h['conector']] == 'B'].head(fibras_por_conector_b * num_conectores_b)
                
                # Combinar los dataframes filtrados para el análisis final
                df_final_cable = pd.concat([df_lado_a, df_lado_b])
                
                # Normalizar la columna de estado en el DF final
                df_final_cable[h['estado']] = df_final_cable[h['estado']].str.strip().str.upper()

                # 1. Validar fallas en el conjunto final
                fails = len(df_final_cable[df_final_cable[h['estado']] != 'PASS'])
                if fails > 0:
                    estado_final = "RECHAZADO"
                    detalles_list.append(f"{fails} medicion(es) con FALLA")
                else:
                    detalles_list.append(f"{len(df_final_cable)} medicion(es) con PASS")
                
                # 2. Validar Lado A
                fibras_a_reales = len(df_lado_a)
                fibras_a_esperadas = num_conectores_a * fibras_por_conector_a
                detalles_list.append(f"Fibras Lado A: {fibras_a_reales}/{fibras_a_esperadas}")
                if fibras_a_reales != fibras_a_esperadas:
                    estado_final = "RECHAZADO"
                
                # 3. Validar Lado B
                fibras_b_reales = len(df_lado_b)
                fibras_b_esperadas = num_conectores_b * fibras_por_conector_b
                detalles_list.append(f"Fibras Lado B: {fibras_b_reales}/{fibras_b_esperadas}")
                if fibras_b_reales != fibras_b_esperadas:
                    estado_final = "RECHAZADO"
                
                # 4. Validar total de fibras
                total_mediciones = len(df_final_cable)
                detalles_list.append(f"Total Fibras: {total_mediciones}/{total_fibras_esperadas}")
                if total_mediciones != total_fibras_esperadas:
                    estado_final = "RECHAZADO"
                
                ultima_medicion = df_final_cable['timestamp'].max()
                ultima_medicion_str = ultima_medicion.strftime('%d/%m/%Y %H:%M:%S') if pd.notna(ultima_medicion) else "N/A"
                
                detalle_str = ". ".join(detalles_list)
                
                reporte_data.append({
                    'Número de Serie': serie, 
                    'Estado Final': estado_final, 
                    'Detalle': detalle_str,
                    'Última Medición': ultima_medicion_str
                })
                if estado_final == "RECHAZADO":
                    rechazados.append([serie, estado_final, detalle_str, ultima_medicion_str])
                
                self.update_progress(30 + int((i + 1) / len(all_series) * 65))

            ruta_reporte, error = self._generar_reporte_excel_mpo(
                reporte_data, total_cables, rechazados, file_path, ot_number, "ILRL_MPO"
            )
            self.after(0, self.show_result, ruta_reporte, error, "IL/RL MPO")
            
        except Exception as e:
            self.after(0, self.show_result, None, f"Error crítico en el análisis IL/RL MPO:\n{traceback.format_exc()}", "IL/RL MPO")

    def _run_polaridad_mpo_thread(self, folder_path, ot_number, total_cables):
        try:
            self.update_progress(10)
            
            mediciones = {} # {serie: {'fecha': dt, 'estado': 'APROBADO/RECHAZADO', 'archivo': ...}}
            
            for estado_carpeta, estado_resultado in [('PASS', 'APROBADO'), ('FAIL', 'RECHAZADO')]:
                subcarpeta = os.path.join(folder_path, estado_carpeta)
                if not os.path.isdir(subcarpeta):
                    continue
                
                for filename in os.listdir(subcarpeta):
                    if filename.lower().endswith('.xlsx') and not filename.startswith('~$'):
                        match_serie = re.search(r'(JMO\d{13})', filename, re.IGNORECASE)
                        match_fecha = re.search(r'(\d{4}-\d{2}-\d{2} \d{2}-\d{2}-\d{2})', filename)
                        
                        if match_serie and match_fecha:
                            serie = match_serie.group(1).upper()
                            fecha_str = match_fecha.group(1)
                            fecha_dt = datetime.strptime(fecha_str, '%Y-%m-%d %H-%M-%S')

                            if serie not in mediciones or fecha_dt > mediciones[serie]['fecha']:
                                mediciones[serie] = {
                                    'fecha': fecha_dt,
                                    'estado': estado_resultado,
                                    'archivo': filename
                                }
            
            self.update_progress(60)

            reporte_data = []
            rechazados = []
            for serie, data in mediciones.items():
                reporte_data.append({
                    'Número de Serie': serie,
                    'Estado Final': data['estado'],
                    'Último Archivo': data['archivo'],
                    'Fecha Medición': data['fecha'].strftime('%Y-%m-%d %H:%M:%S')
                })
                if data['estado'] == 'RECHAZADO':
                    rechazados.append([serie, 'RECHAZADO', data['archivo']])

            self.update_progress(90)
            
            ruta_reporte, error = self._generar_reporte_excel_mpo(
                reporte_data, total_cables, rechazados, folder_path, ot_number, "Polaridad_MPO"
            )
            self.after(0, self.show_result, ruta_reporte, error, "Polaridad MPO")

        except Exception as e:
            self.after(0, self.show_result, None, f"Error crítico en el análisis de Polaridad MPO:\n{traceback.format_exc()}", "Polaridad MPO")

    def update_progress(self, value):
        self.progress_bar["value"] = value

    def show_result(self, ruta_generada, errores, tipo_analisis):
        self.progress_bar["value"] = 100
        if ruta_generada:
            self.result_label.config(text=f"Reporte de {tipo_analisis} generado con éxito.")
            if messagebox.askyesno("Éxito", f"Reporte de {tipo_analisis} generado con éxito.\n\n"
                                          f"Guardado en:\n{ruta_generada}\n\n"
                                          "¿Deseas abrir el archivo ahora?", parent=self):
                self.abrir_archivo(ruta_generada)
        else:
            self.result_label.config(text=f"El análisis de {tipo_analisis} falló.")
            messagebox.showerror("Análisis Fallido", f"No se pudo generar el reporte de {tipo_analisis}.\n\n"
                                                   f"Detalles:\n{errores}", parent=self)
        if errores and ruta_generada:
             messagebox.showwarning("Advertencia", f"El reporte se generó, pero se encontraron los siguientes problemas:\n\n{errores}", parent=self)

    def abrir_archivo(self, ruta_archivo):
        try:
            if sys.platform == 'win32':
                os.startfile(ruta_archivo)
            elif sys.platform == 'darwin':
                subprocess.run(['open', ruta_archivo], check=True)
            else:
                subprocess.run(['xdg-open', ruta_archivo], check=True)
        except Exception as e:
            print(f"Advertencia: No se pudo abrir el archivo automáticamente: {e}")

# --- PÁGINA PARA ANÁLISIS DE O.T. (MPO) ---
class AnalisisMPOPage(ttk.Frame):
    def __init__(self, parent, app_instance):
        super().__init__(parent, padding=20)
        self.app = app_instance
        self.create_widgets()

    def create_widgets(self):
        container = ttk.Frame(self, style='TFrame')
        container.pack(expand=True, fill='both')

        config_frame = ttk.LabelFrame(container, text="Parámetros del Análisis MPO", padding=15)
        config_frame.pack(fill='x', pady=(0, 20))
        config_frame.columnconfigure(1, weight=1)

        # ILRL File
        ttk.Label(config_frame, text="Archivo IL/RL MPO:", font="-weight bold").grid(row=0, column=0, sticky='w', pady=5, padx=5)
        self.ilrl_file_var = tk.StringVar()
        ttk.Entry(config_frame, textvariable=self.ilrl_file_var, state="readonly").grid(row=0, column=1, sticky='ew', pady=5, padx=5)
        ttk.Button(config_frame, text="Seleccionar Archivo...", command=self.select_ilrl_file, style='outline.TButton').grid(row=0, column=2, sticky='w', pady=5, padx=5)

        # GEO File
        ttk.Label(config_frame, text="Archivo Geometría MPO:", font="-weight bold").grid(row=1, column=0, sticky='w', pady=5, padx=5)
        self.geo_file_var = tk.StringVar()
        ttk.Entry(config_frame, textvariable=self.geo_file_var, state="readonly").grid(row=1, column=1, sticky='ew', pady=5, padx=5)
        ttk.Button(config_frame, text="Seleccionar Archivo...", command=self.select_geo_file, style='outline.TButton').grid(row=1, column=2, sticky='w', pady=5, padx=5)

        # Polarity Folder
        ttk.Label(config_frame, text="Carpeta Polaridad MPO:", font="-weight bold").grid(row=2, column=0, sticky='w', pady=5, padx=5)
        self.polaridad_folder_var = tk.StringVar()
        ttk.Entry(config_frame, textvariable=self.polaridad_folder_var, state="readonly").grid(row=2, column=1, sticky='ew', pady=5, padx=5)
        ttk.Button(config_frame, text="Seleccionar Carpeta...", command=self.select_polaridad_folder, style='outline.TButton').grid(row=2, column=2, sticky='w', pady=5, padx=5)
        
        ttk.Separator(config_frame, orient='horizontal').grid(row=3, column=0, columnspan=3, sticky='ew', pady=10)

        # OT and Quantity
        ttk.Label(config_frame, text="Número de O.T. (sin JMO-):", font="-weight bold").grid(row=4, column=0, sticky='w', pady=5, padx=5)
        self.ot_var = tk.StringVar()
        ttk.Entry(config_frame, textvariable=self.ot_var).grid(row=4, column=1, sticky='w', pady=5, padx=5)

        ttk.Label(config_frame, text="Total de cables esperados:", font="-weight bold").grid(row=5, column=0, sticky='w', pady=5, padx=5)
        self.total_cables_var = tk.StringVar()
        ttk.Entry(config_frame, textvariable=self.total_cables_var).grid(row=5, column=1, sticky='w', pady=5, padx=5)

        # --- Frame de Acciones ---
        action_frame = ttk.Frame(container)
        action_frame.pack(fill='x', pady=10)
        
        ttk.Button(action_frame, text="Generar Reporte IL/RL", command=lambda: self.run_analysis("ilrl"), style='info.TButton', padding=10).pack(side='left', padx=10, expand=True)
        ttk.Button(action_frame, text="Generar Reporte Geometría", command=lambda: self.run_analysis("geo"), style='info.TButton', padding=10).pack(side='left', padx=10, expand=True)
        ttk.Button(action_frame, text="Generar Reporte Polaridad", command=lambda: self.run_analysis("polaridad"), style='info.TButton', padding=10).pack(side='left', padx=10, expand=True)

        # --- Frame de Progreso y Resultados ---
        result_frame = ttk.LabelFrame(container, text="Estado del Análisis", padding=15)
        result_frame.pack(fill='both', expand=True, pady=(20, 0))

        self.progress_bar = ttk.Progressbar(result_frame, mode='determinate')
        self.progress_bar.pack(fill='x', pady=10)
        
        self.result_label = ttk.Label(result_frame, text="Listo para iniciar el análisis MPO.", wraplength=700)
        self.result_label.pack(fill='x', pady=10)

    def _select_file(self, var, title):
        file_selected = filedialog.askopenfilename(title=title, filetypes=[("Excel files", "*.xlsx *.xls")])
        if file_selected:
            var.set(file_selected)
            # --- CORRECCIÓN: Acepta JMO o JRMO ---
            match = re.search(r'(J(?:R)?MO-\d{9})', os.path.basename(file_selected), re.IGNORECASE)
            if match and not self.ot_var.get():
                self.ot_var.set(match.group(1).upper())

    def select_ilrl_file(self):
        self._select_file(self.ilrl_file_var, "Seleccionar archivo de IL/RL MPO")

    def select_geo_file(self):
        self._select_file(self.geo_file_var, "Seleccionar archivo de Geometría MPO")
        
    def select_polaridad_folder(self):
        folder_selected = filedialog.askdirectory(title="Seleccionar carpeta de Polaridad (ej. JMO-250800005)")
        if folder_selected:
            self.polaridad_folder_var.set(folder_selected)
            match = re.search(r'JMO-(\d{9})', os.path.basename(folder_selected), re.IGNORECASE)
            if match and not self.ot_var.get():
                # Corrección para consistencia: sugerir sin prefijo
                self.ot_var.set(match.group(1))
    
    def run_analysis(self, analysis_type):
        ot_number_raw = self.ot_var.get().strip()
        total_str = self.total_cables_var.get()
        
        # Validaciones comunes
        if not ot_number_raw:
            messagebox.showerror("Error", "Por favor, ingresa el número de O.T.", parent=self)
            return
        if not total_str.isdigit() or int(total_str) <= 0:
            messagebox.showerror("Error", "Por favor, ingresa un número válido para el total de cables.", parent=self)
            return
        
        ot_number = f"JMO-{ot_number_raw}" if not ot_number_raw.startswith("JMO-") else ot_number_raw
        total = int(total_str)
        
        self.progress_bar["value"] = 0
        self.result_label.config(text=f"Iniciando análisis de {analysis_type.upper()} para MPO...")
        
        # Validaciones y arranque de hilos específicos
        if analysis_type == "ilrl":
            file_path = self.ilrl_file_var.get()
            if not file_path:
                messagebox.showerror("Error", "Por favor, selecciona el archivo de IL/RL para MPO.", parent=self)
                return
            thread = threading.Thread(target=self._run_ilrl_mpo_thread, args=(file_path, ot_number, total), daemon=True)
            thread.start()
        elif analysis_type == "geo":
            file_path = self.geo_file_var.get()
            if not file_path:
                messagebox.showerror("Error", "Por favor, selecciona el archivo de Geometría para MPO.", parent=self)
                return
            thread = threading.Thread(target=self._run_geo_mpo_thread, args=(file_path, ot_number, total), daemon=True)
            thread.start()
        elif analysis_type == "polaridad":
            folder_path = self.polaridad_folder_var.get()
            if not folder_path:
                messagebox.showerror("Error", "Por favor, selecciona la carpeta de Polaridad para MPO.", parent=self)
                return
            thread = threading.Thread(target=self._run_polaridad_mpo_thread, args=(folder_path, ot_number, total), daemon=True)
            thread.start()

    def _cargar_ot_configuration(self, ot_number):
        try:
            conn = sqlite3.connect(self.app.config['db_path'])
            conn.row_factory = sqlite3.Row
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM ot_configurations WHERE ot_number = ?", (ot_number,))
            row = cursor.fetchone()
            conn.close()
            return dict(row) if row else None
        except Exception as e:
            messagebox.showerror("Error de Base de Datos", f"No se pudo cargar la configuración de la OT: {e}", parent=self)
            return None

# En la clase AnalisisMPOPage, reemplaza este método completo:

    def _run_ilrl_mpo_thread(self, file_path, ot_number, total_cables):
        try:
            self.update_progress(5)
            ot_config = self._cargar_ot_configuration(ot_number)
            if not ot_config:
                self.after(0, self.show_result, None, f"No se encontró configuración para la OT {ot_number}. Por favor, configúrela primero.", "IL/RL MPO")
                return

            self.update_progress(15)
            df = pd.read_excel(file_path, sheet_name="Results")
            h = {k: ot_config.get(v, d) for k, v, d in [('serie', 'ilrl_serie_header', 'Serial number'), ('estado', 'ilrl_estado_header', 'Alarm Status'), ('conector', 'ilrl_conector_header', 'connector label'), ('fecha', 'ilrl_fecha_header', 'Date'), ('hora', 'ilrl_hora_header', 'Time')]}
            
            if any(header not in df.columns for header in h.values()):
                missing = [k for k, v in h.items() if v not in df.columns]
                self.after(0, self.show_result, None, f"El archivo no contiene los encabezados esperados: {', '.join(missing)}.", "IL/RL MPO")
                return
            
            self.update_progress(30)
            df[h['serie']] = df[h['serie']].astype(str) # Esta es la columna con el N/S completo (ej. ...0028-F)

            # --- INICIO DE LA NUEVA LÓGICA DE PRIORIZACIÓN DE RETRABAJOS (-F) ---
            
            # 1. *** LÍNEA CORREGIDA ***
            #    Añadimos dayfirst=True para que entienda '29/09/2025' como Día/Mes/Año.
            date_series = pd.to_datetime(df[h['fecha']], dayfirst=True, errors='coerce').dt.strftime('%d/%m/%Y')
            
            time_series = df[h['hora']].astype(str).str.replace('a. m.', 'AM', regex=False).str.replace('p. m.', 'PM', regex=False).str.strip()
            full_datetime_str = date_series + ' ' + time_series
            df['timestamp'] = pd.to_datetime(full_datetime_str, format='%d/%m/%Y %I:%M:%S %p', errors='coerce')
            
            # Esta línea ahora solo borrará las filas *realmente* corruptas (como las del cable 0001)
            df.dropna(subset=['timestamp'], inplace=True)

            # 2. Crear una columna 'base_serie' que contenga solo los 13 dígitos
            df['base_serie'] = df[h['serie']].str.extract(r'(\d{13})')
            df.dropna(subset=['base_serie'], inplace=True) # Descartar filas que no tengan un N/S de 13 dígitos

            # 3. Obtener la lista de cables base únicos
            all_base_series = df['base_serie'].dropna().unique()

            reporte_data = []
            rechazados = []
            
            num_conectores_a = ot_config.get('num_conectores_a', 1)
            fibras_por_conector_a = ot_config.get('fibers_per_connector_a', 12)
            num_conectores_b = ot_config.get('num_conectores_b', 1)
            fibras_por_conector_b = ot_config.get('fibers_per_connector_b', 12)
            total_fibras_esperadas = (num_conectores_a * fibras_por_conector_a) + (num_conectores_b * fibras_por_conector_b)

            # 4. Iterar sobre los cables base, no sobre los N/S completos
            for i, base_serie in enumerate(all_base_series):
                # Obtener todas las mediciones para este cable base (ej. original y -F)
                df_cable_group = df[df['base_serie'] == base_serie]

                # 5. Encontrar qué N/S completo (original o -F) tiene la fecha más reciente
                latest_timestamp = pd.NaT
                best_sub_group_df = None
                best_full_serial = None

                # Iterar sobre las variaciones (ej. "...0028" y "...0028-F")
                for full_serial in df_cable_group[h['serie']].unique():
                    sub_group_df = df_cable_group[df_cable_group[h['serie']] == full_serial]
                    current_max_timestamp = sub_group_df['timestamp'].max() # Fecha más reciente de esta variación

                    if best_sub_group_df is None or current_max_timestamp > latest_timestamp:
                        latest_timestamp = current_max_timestamp
                        best_sub_group_df = sub_group_df
                        best_full_serial = full_serial # Guardar el N/S "ganador"

                # 6. 'best_sub_group_df' ahora contiene solo las filas del N/S más reciente
                if best_sub_group_df is None:
                    continue 
                
                df_cable = best_sub_group_df.copy()
                serie_para_reporte = best_full_serial # Usaremos este N/S para el reporte

                # --- FIN DE LA NUEVA LÓGICA ---

                # 7. El resto de la lógica de análisis se aplica solo al grupo "ganador"
                estado_final = "APROBADO"
                detalles_list = []

                df_cable.sort_values(by='timestamp', ascending=False, inplace=True)

                df_lado_a = df_cable[df_cable[h['conector']] == 'A'].head(fibras_por_conector_a * num_conectores_a)
                df_lado_b = df_cable[df_cable[h['conector']] == 'B'].head(fibras_por_conector_b * num_conectores_b)
                
                df_final_cable = pd.concat([df_lado_a, df_lado_b])
                
                if df_final_cable.empty: continue

                df_final_cable.loc[:, h['estado']] = df_final_cable[h['estado']].str.strip().str.upper()

                fails = len(df_final_cable[df_final_cable[h['estado']] != 'PASS'])
                if fails > 0:
                    estado_final = "RECHAZADO"
                    detalles_list.append(f"{fails} medicion(es) con FALLA")
                else:
                    detalles_list.append(f"{len(df_final_cable)} medicion(es) con PASS")
                
                fibras_a_reales = len(df_lado_a)
                fibras_a_esperadas = num_conectores_a * fibras_por_conector_a
                detalles_list.append(f"Fibras Lado A: {fibras_a_reales}/{fibras_a_esperadas}")
                if fibras_a_reales != fibras_a_esperadas: estado_final = "RECHAZADO"
                
                fibras_b_reales = len(df_lado_b)
                fibras_b_esperadas = num_conectores_b * fibras_por_conector_b
                detalles_list.append(f"Fibras Lado B: {fibras_b_reales}/{fibras_b_esperadas}")
                if fibras_b_reales != fibras_b_esperadas: estado_final = "RECHAZADO"
                
                total_mediciones = len(df_final_cable)
                detalles_list.append(f"Total Fibras: {total_mediciones}/{total_fibras_esperadas}")
                if total_mediciones != total_fibras_esperadas: estado_final = "RECHAZADO"
                
                ultima_medicion_str = ""
                fechas_por_punta = []
                if not df_lado_a.empty:
                    max_fecha_a = df_lado_a['timestamp'].max()
                    fechas_por_punta.append(f"Lado A: {max_fecha_a.strftime('%d/%m/%Y %H:%M:%S') if pd.notna(max_fecha_a) else 'N/A'}")
                if not df_lado_b.empty:
                    max_fecha_b = df_lado_b['timestamp'].max()
                    fechas_por_punta.append(f"Lado B: {max_fecha_b.strftime('%d/%m/%Y %H:%M:%S') if pd.notna(max_fecha_b) else 'N/A'}")
                ultima_medicion_str = ". ".join(fechas_por_punta)
                
                detalle_str = ". ".join(detalles_list)
                
                # 8. Guardar en el reporte usando el N/S "ganador"
                reporte_data.append({
                    'Número de Serie': serie_para_reporte, 
                    'Estado Final': estado_final, 
                    'Detalle': detalle_str,
                    'Última Medición': ultima_medicion_str
                })
                if estado_final == "RECHAZADO":
                    rechazados.append([serie_para_reporte, estado_final, detalle_str, ultima_medicion_str])
                
                self.update_progress(30 + int((i + 1) / len(all_base_series) * 65))

            ruta_reporte, error = self._generar_reporte_excel_mpo(
                reporte_data, total_cables, rechazados, file_path, ot_number, "ILRL_MPO"
            )
            self.after(0, self.show_result, ruta_reporte, error, "IL/RL MPO")
            
        except Exception as e:
            self.after(0, self.show_result, None, f"Error crítico en el análisis IL/RL MPO:\n{traceback.format_exc()}", "IL/RL MPO")
            
    def _run_geo_mpo_thread(self, file_path, ot_number, total_cables):
        try:
            self.update_progress(5)
            ot_config = self._cargar_ot_configuration(ot_number)
            if not ot_config:
                self.after(0, self.show_result, None, f"No se encontró configuración para la OT {ot_number}.", "Geometría MPO")
                return

            self.update_progress(15)
            df = pd.read_excel(file_path, sheet_name="MT12", header=None)
            header_row_index = next((i for i, row in df.iterrows() if str(row.iloc[0]).strip().lower() == 'name'), -1)
            
            if header_row_index == -1:
                self.after(0, self.show_result, None, "No se encontró encabezado 'Name' en la hoja 'MT12'.", "Geometría MPO")
                return
            
            df.columns = df.iloc[header_row_index].str.strip().str.lower()
            df = df.iloc[header_row_index + 1:].rename(columns={'pass/fail': 'result'})
            self.update_progress(30)

            # --- FUNCIÓN MEJORADA PARA RECONOCER RETRABAJOS Y JRMO ---
            def parse_geo_name_mpo(name_str):
                # --- CORRECCIÓN: Regex ajustada para J(R)MO ---
                match = re.match(r'(J(?:R)?MO\d{13})-(R?\d+)(-[FK])?', str(name_str).upper())
                if match:
                    # Retorna: base del serial (ej. JRMO...001), punta (ej. "1" o "R1"), y si es fanout
                    return match.group(1), match.group(2), match.group(3)
                return None, None, None

            df[['serie_base', 'lado', 'retrabajo_info']] = df['name'].apply(lambda x: pd.Series(parse_geo_name_mpo(x)))
            
            # Limpiar filas que no se pudieron parsear
            df.dropna(subset=['serie_base', 'lado'], inplace=True)

            if 'date & time' in df.columns:
                df['timestamp'] = pd.to_datetime(df['date & time'], errors='coerce')
            else:
                 df['timestamp'] = pd.to_datetime(df['date'].astype(str) + ' ' + df['time'].astype(str), errors='coerce')

            all_series = df['serie_base'].dropna().unique()
            reporte_data = []
            rechazados = []
            
            num_lados_esperados = ot_config.get('num_conectores_a', 1) + ot_config.get('num_conectores_b', 1)
            lados_esperados = [str(i + 1) for i in range(num_lados_esperados)]

            for i, serie in enumerate(all_series):
                df_cable = df[df['serie_base'] == serie].copy().sort_values(by='timestamp', ascending=False)
                
                # --- INICIO DE LA NUEVA LÓGICA DE PRIORIZACIÓN DE RETRABAJOS ---
                originales = {}
                retrabajos = {}
                for _, medicion in df_cable.iterrows():
                    punta_original = str(medicion['lado'])
                    punta_normalizada = punta_original.replace('R', '')

                    if 'R' in punta_original:
                        if punta_normalizada not in retrabajos:
                            retrabajos[punta_normalizada] = {'Punta_Original': punta_original, 'Resultado': medicion['result'], 'Fecha': medicion['timestamp'], 'Nombre_Completo': medicion['name']}
                    else:
                        if punta_normalizada not in originales:
                            originales[punta_normalizada] = {'Punta_Original': punta_original, 'Resultado': medicion['result'], 'Fecha': medicion['timestamp'], 'Nombre_Completo': medicion['name']}

                mediciones_definitivas = {}
                for punta, medicion in retrabajos.items():
                    mediciones_definitivas[punta] = medicion
                
                for punta, medicion in originales.items():
                    if punta not in mediciones_definitivas:
                        mediciones_definitivas[punta] = medicion
                # --- FIN DE LA NUEVA LÓGICA DE PRIORIZACIÓN ---

                estado_final = "APROBADO"
                detalles_list = []
                fechas_list = []

                for lado_esperado in lados_esperados:
                    if lado_esperado in mediciones_definitivas:
                        medicion = mediciones_definitivas[lado_esperado]
                        estado_lado = str(medicion['Resultado']).upper()
                        nombre_original = medicion['Nombre_Completo']
                        fecha_lado = medicion['Fecha']

                        if estado_lado != 'PASS':
                            estado_final = "RECHAZADO"
                        
                        detalles_list.append(f"Lado {medicion['Punta_Original']}: {estado_lado}")
                        fechas_list.append(f"Lado {medicion['Punta_Original']}: {fecha_lado.strftime('%d/%m/%Y %H:%M:%S') if pd.notna(fecha_lado) else 'N/A'}")
                    else:
                        estado_final = "RECHAZADO"
                        detalles_list.append(f"Lado {lado_esperado}: FALTANTE")

                detalle_str = ". ".join(detalles_list)
                fechas_str = ". ".join(fechas_list)
                
                reporte_data.append({
                    'Número de Serie': serie, 
                    'Estado Final': estado_final, 
                    'Detalle': detalle_str,
                    'Última Medición': fechas_str
                })
                if estado_final == "RECHAZADO":
                    rechazados.append([serie, estado_final, detalle_str, fechas_str])

                self.update_progress(30 + int((i + 1) / len(all_series) * 65))

            ruta_reporte, error = self._generar_reporte_excel_mpo(
                reporte_data, total_cables, rechazados, file_path, ot_number, "GEO_MPO"
            )
            self.after(0, self.show_result, ruta_reporte, error, "Geometría MPO")

        except Exception as e:
            self.after(0, self.show_result, None, f"Error crítico en el análisis de Geometría MPO:\n{traceback.format_exc()}", "Geometría MPO")

    def _run_polaridad_mpo_thread(self, folder_path, ot_number, total_cables):
        try:
            self.update_progress(10)
            
            mediciones = {} 
            
            for estado_carpeta, estado_resultado in [('PASS', 'APROBADO'), ('FAIL', 'RECHAZADO')]:
                subcarpeta = os.path.join(folder_path, estado_carpeta)
                if not os.path.isdir(subcarpeta):
                    continue
                
                for filename in os.listdir(subcarpeta):
                    if filename.lower().endswith('.xlsx') and not filename.startswith('~$'):
                        # --- CORRECCIÓN: Regex ajustada para J(R)MO ---
                        match_serie = re.search(r'(J(?:R)?MO\d{13})', filename, re.IGNORECASE)
                        match_fecha = re.search(r'(\d{4}-\d{2}-\d{2} \d{2}-\d{2}-\d{2})', filename)
                        
                        if match_serie and match_fecha:
                            serie = match_serie.group(1).upper()
                            fecha_str = match_fecha.group(1)
                            fecha_dt = datetime.strptime(fecha_str, '%Y-%m-%d %H-%M-%S')

                            if serie not in mediciones or fecha_dt > mediciones[serie]['fecha']:
                                mediciones[serie] = {
                                    'fecha': fecha_dt,
                                    'estado': estado_resultado,
                                    'archivo': filename
                                }
            
            self.update_progress(60)

            reporte_data = []
            rechazados = []
            for serie, data in mediciones.items():
                reporte_data.append({
                    'Número de Serie': serie,
                    'Estado Final': data['estado'],
                    'Último Archivo': data['archivo'],
                    'Fecha Medición': data['fecha'].strftime('%Y-%m-%d %H:%M:%S')
                })
                if data['estado'] == 'RECHAZADO':
                    rechazados.append([serie, 'RECHAZADO', data['archivo']])

            self.update_progress(90)
            
            ruta_reporte, error = self._generar_reporte_excel_mpo(
                reporte_data, total_cables, rechazados, folder_path, ot_number, "Polaridad_MPO"
            )
            self.after(0, self.show_result, ruta_reporte, error, "Polaridad MPO")

        except Exception as e:
            self.after(0, self.show_result, None, f"Error crítico en el análisis de Polaridad MPO:\n{traceback.format_exc()}", "Polaridad MPO")

    def _generar_reporte_excel_mpo(self, reporte_data, total_esperado, rechazados, ruta_base, ot_number, tipo_reporte):
        # Para archivos, el reporte se guarda junto al archivo. Para carpetas, dentro de la carpeta.
        if os.path.isfile(ruta_base):
            carpeta_destino = os.path.dirname(ruta_base)
        else:
            carpeta_destino = ruta_base
            
        carpeta_analisis = os.path.join(carpeta_destino, "ANALISIS DE O.T")
        os.makedirs(carpeta_analisis, exist_ok=True)
        
        ot_sufijo = ot_number.replace("JMO-", "")
        nombre_archivo = f"Reporte_Analisis_{tipo_reporte}_{ot_sufijo}.xlsx"
        ruta_destino = os.path.join(carpeta_analisis, nombre_archivo)
        
        try:
            if os.path.exists(ruta_destino):
                os.remove(ruta_destino)

            wb = Workbook()
            
            # --- Estilos reutilizables ---
            header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
            header_font = Font(color="FFFFFF", bold=True)
            center_alignment = Alignment(horizontal='center', vertical='center')
            
            # --- Hoja 1: Resultados OT ---
            ws1 = wb.active
            ws1.title = "Resultados OT"
            df_reporte = pd.DataFrame(reporte_data)
            
            headers = []
            if not df_reporte.empty:
                headers = list(df_reporte.columns)
                ws1.append(headers)
                if 'Número de Serie' in df_reporte.columns:
                     df_reporte.sort_values(by='Número de Serie', inplace=True)
                for r in df_reporte.itertuples(index=False):
                    ws1.append(list(r))
                
                # Aplicar tabla con filtros
                tabla_resultados = Table(displayName="ResultadosOT", ref=f"A1:{get_column_letter(ws1.max_column)}{ws1.max_row}")
                style = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False,
                                       showLastColumn=False, showRowStripes=True, showColumnStripes=False)
                tabla_resultados.tableStyleInfo = style
                ws1.add_table(tabla_resultados)

                # Formato condicional para APROBADO/RECHAZADO
                green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
                red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
                
                ws1.conditional_formatting.add(f"B2:B{ws1.max_row}", CellIsRule(operator='equal', formula=['APROBADO'], fill=green_fill))
                ws1.conditional_formatting.add(f"B2:B{ws1.max_row}", CellIsRule(operator='equal', formula=['RECHAZADO'], fill=red_fill))

            # --- Hoja 2: Resumen y Faltantes ---
            ws2 = wb.create_sheet("Resumen")
            encontrados = len(reporte_data)
            faltantes = total_esperado - encontrados
            ws2.append(['Métrica', 'Valor'])
            ws2.append(['Total Esperado', total_esperado])
            ws2.append(['Total Encontrados', encontrados])
            ws2.append(['Total Faltantes', faltantes])
            for cell in ws2["A"]: cell.font = Font(bold=True) # Negrita para la primera columna
            
            # --- Hoja 3: Rechazados ---
            if rechazados:
                ws3 = wb.create_sheet("Rechazados")
                if headers:
                    ws3.append(headers)
                rechazados.sort(key=lambda x: x[0])
                for fila in rechazados:
                    ws3.append(fila)
                
                tabla_rechazados = Table(displayName="Rechazados", ref=f"A1:{get_column_letter(ws3.max_column)}{ws3.max_row}")
                style_rechazados = TableStyleInfo(name="TableStyleMedium4", showFirstColumn=False,
                                                  showLastColumn=False, showRowStripes=True, showColumnStripes=False)
                tabla_rechazados.tableStyleInfo = style_rechazados
                ws3.add_table(tabla_rechazados)

            # Auto-ajustar columnas y aplicar estilos a encabezados en todas las hojas
            for ws_name in wb.sheetnames:
                ws = wb[ws_name]
                for col in ws.columns:
                    max_length = 0
                    column_letter = get_column_letter(col[0].column)
                    for cell in col:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                    adjusted_width = (max_length + 2) if max_length < 50 else 50
                    ws.column_dimensions[column_letter].width = adjusted_width
                
                # Estilo de encabezado
                if ws.max_row > 0:
                    for cell in ws[1]:
                        cell.fill = header_fill
                        cell.font = header_font
                        cell.alignment = center_alignment

            wb.save(ruta_destino)
            return ruta_destino, None
        
        except PermissionError:
            error_msg = (f"No se pudo guardar el reporte porque el archivo está siendo usado por otro proceso.\n\n"
                         f"Por favor, cierra el archivo:\n{nombre_archivo}\n\ny vuelve a intentarlo.")
            return None, error_msg
        except Exception as e:
            return None, f"No se pudo guardar el reporte:\n{traceback.format_exc()}"
    
    def update_progress(self, value):
        self.progress_bar["value"] = value

    def show_result(self, ruta_generada, errores, tipo_analisis):
        self.progress_bar["value"] = 100
        if ruta_generada:
            self.result_label.config(text=f"Reporte de {tipo_analisis} generado con éxito.")
            if messagebox.askyesno("Éxito", f"Reporte de {tipo_analisis} generado con éxito.\n\n"
                                          f"Guardado en:\n{ruta_generada}\n\n"
                                          "¿Deseas abrir el archivo ahora?", parent=self):
                self.abrir_archivo(ruta_generada)
        else:
            self.result_label.config(text=f"El análisis de {tipo_analisis} falló.")
            messagebox.showerror("Análisis Fallido", f"No se pudo generar el reporte de {tipo_analisis}.\n\n"
                                                   f"Detalles:\n{errores}", parent=self)
        if errores and ruta_generada:
             messagebox.showwarning("Advertencia", f"El reporte se generó, pero se encontraron los siguientes problemas:\n\n{errores}", parent=self)

    def abrir_archivo(self, ruta_archivo):
        try:
            if sys.platform == 'win32':
                os.startfile(ruta_archivo)
            elif sys.platform == 'darwin':
                subprocess.run(['open', ruta_archivo], check=True)
            else:
                subprocess.run(['xdg-open', ruta_archivo], check=True)
        except Exception as e:
            print(f"Advertencia: No se pudo abrir el archivo automáticamente: {e}")

class RecordsPage(ttk.Frame):
    def __init__(self, parent, app_instance):
        super().__init__(parent, padding=10)
        self.app = app_instance
        self.create_widgets()

    def create_widgets(self):
        filter_frame = ttk.Frame(self)
        filter_frame.pack(fill='x', pady=5)
        ttk.Label(filter_frame, text="Filtrar por OT o Serie:").pack(side='left', padx=5)
        self.filter_entry = ttk.Entry(filter_frame, width=30)
        self.filter_entry.pack(side='left', padx=5)
        self.filter_entry.bind("<KeyRelease>", lambda e: self.load_records())

        cols = ("ID", "Fecha", "No. Serie", "OT", "Estado Final", "ILRL", "GEO", "Polaridad")
        self.tree = ttk.Treeview(self, columns=cols, show="headings", style='primary.Treeview')
        for col in cols:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=120, anchor='w')
        
        self.tree.column("ID", width=40, stretch=False)
        self.tree.pack(fill='both', expand=True, pady=10)

        self.tree.tag_configure('APROBADO', foreground='green')
        self.tree.tag_configure('RECHAZADO', foreground='red')
        self.tree.tag_configure('NO ENCONTRADO', foreground='orange')
        
    def load_records(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        filtro = self.filter_entry.get().strip()
        try:
            conn = sqlite3.connect(self.app.config['db_path'])
            cursor = conn.cursor()
            query = "SELECT id, entry_date, serial_number, ot_number, overall_status, ilrl_status, geo_status, polaridad_status FROM cable_verifications"
            params = []
            if filtro:
                query += " WHERE serial_number LIKE ? OR ot_number LIKE ?"
                params.extend([f'%{filtro}%', f'%{filtro}%'])
            query += " ORDER BY id DESC"
            
            cursor.execute(query, params)
            for row in cursor.fetchall():
                row_list = list(row)
                while len(row_list) < 8:
                    row_list.append('N/A')
                self.tree.insert("", "end", values=tuple(row_list), tags=(row[4],))
            conn.close()
        except Exception as e:
            messagebox.showerror("Error de Base de Datos", f"No se pudieron cargar los registros: {e}")

class SettingsWindow(tk.Toplevel):
    def __init__(self, app_instance):
        super().__init__(app_instance)
        self.app = app_instance
        self.title("Configurar Rutas")
        self.geometry("800x500") # Hacemos la ventana más alta
        self.transient(self.app)
        self.grab_set()

        self.bind_class("TEntry", "<FocusIn>", open_keyboard)

        frame = ttk.Frame(self, padding=20)
        frame.pack(fill='both', expand=True)
        frame.columnconfigure(1, weight=1)

        # --- SECCIÓN LC/SC ---
        ttk.Label(frame, text="Rutas SC/LC", font=("Helvetica", 10, "bold")).grid(row=0, column=0, columnspan=3, sticky='w', pady=(0,5))

        # ILRL Primaria
        ttk.Label(frame, text="Ruta ILRL (Principal):").grid(row=1, column=0, sticky='w', pady=2)
        self.ilrl_path = tk.StringVar(value=self.app.config['ruta_base_ilrl'])
        ttk.Entry(frame, textvariable=self.ilrl_path, width=60).grid(row=1, column=1, sticky='ew', padx=5)
        ttk.Button(frame, text="...", command=lambda: self.browse_folder(self.ilrl_path)).grid(row=1, column=2, padx=5)

        # ILRL Secundaria (NUEVO)
        ttk.Label(frame, text="Ruta ILRL (Secundaria):").grid(row=2, column=0, sticky='w', pady=2)
        self.ilrl_path_2 = tk.StringVar(value=self.app.config.get('ruta_base_ilrl_2', ''))
        ttk.Entry(frame, textvariable=self.ilrl_path_2, width=60).grid(row=2, column=1, sticky='ew', padx=5)
        ttk.Button(frame, text="...", command=lambda: self.browse_folder(self.ilrl_path_2)).grid(row=2, column=2, padx=5)

        # Geo Primaria
        ttk.Label(frame, text="Ruta Geometría (Principal):").grid(row=3, column=0, sticky='w', pady=2)
        self.geo_path = tk.StringVar(value=self.app.config['ruta_base_geo'])
        ttk.Entry(frame, textvariable=self.geo_path, width=60).grid(row=3, column=1, sticky='ew', padx=5)
        ttk.Button(frame, text="...", command=lambda: self.browse_folder(self.geo_path)).grid(row=3, column=2, padx=5)

        # Geo Secundaria (NUEVO)
        ttk.Label(frame, text="Ruta Geometría (Secundaria):").grid(row=4, column=0, sticky='w', pady=2)
        self.geo_path_2 = tk.StringVar(value=self.app.config.get('ruta_base_geo_2', ''))
        ttk.Entry(frame, textvariable=self.geo_path_2, width=60).grid(row=4, column=1, sticky='ew', padx=5)
        ttk.Button(frame, text="...", command=lambda: self.browse_folder(self.geo_path_2)).grid(row=4, column=2, padx=5)
        
        ttk.Separator(frame).grid(row=5, column=0, columnspan=3, pady=10, sticky='ew')

        # --- SECCIÓN MPO ---
        ttk.Label(frame, text="Rutas MPO", font=("Helvetica", 10, "bold")).grid(row=6, column=0, columnspan=3, sticky='w', pady=(0,5))

        ttk.Label(frame, text="Ruta ILRL (MPO):").grid(row=7, column=0, sticky='w', pady=2)
        self.ilrl_mpo_path = tk.StringVar(value=self.app.config['ruta_base_ilrl_mpo'])
        ttk.Entry(frame, textvariable=self.ilrl_mpo_path, width=60).grid(row=7, column=1, sticky='ew', padx=5)
        ttk.Button(frame, text="...", command=lambda: self.browse_folder(self.ilrl_mpo_path)).grid(row=7, column=2, padx=5)

        ttk.Label(frame, text="Ruta Geometría (MPO):").grid(row=8, column=0, sticky='w', pady=2)
        self.geo_mpo_path = tk.StringVar(value=self.app.config['ruta_base_geo_mpo'])
        ttk.Entry(frame, textvariable=self.geo_mpo_path, width=60).grid(row=8, column=1, sticky='ew', padx=5)
        ttk.Button(frame, text="...", command=lambda: self.browse_folder(self.geo_mpo_path)).grid(row=8, column=2, padx=5)

        ttk.Label(frame, text="Ruta Polaridad (MPO):").grid(row=9, column=0, sticky='w', pady=2)
        self.polaridad_mpo_path = tk.StringVar(value=self.app.config['ruta_base_polaridad_mpo'])
        ttk.Entry(frame, textvariable=self.polaridad_mpo_path, width=60).grid(row=9, column=1, sticky='ew', padx=5)
        ttk.Button(frame, text="...", command=lambda: self.browse_folder(self.polaridad_mpo_path)).grid(row=9, column=2, padx=5)
        
        ttk.Separator(frame).grid(row=10, column=0, columnspan=3, pady=10, sticky='ew')
        
        # Database Path
        ttk.Label(frame, text="Ruta Base de Datos:").grid(row=11, column=0, sticky='w', pady=2)
        self.db_path = tk.StringVar(value=self.app.config['db_path'])
        ttk.Entry(frame, textvariable=self.db_path, width=60).grid(row=11, column=1, sticky='ew', padx=5)
        ttk.Button(frame, text="...", command=self.browse_db_file).grid(row=11, column=2, padx=5)
        
        btn_frame = ttk.Frame(frame)
        btn_frame.grid(row=12, column=0, columnspan=3, pady=20)
        ttk.Button(btn_frame, text="Guardar", command=self.save_and_close, style='success.TButton').pack(side='left', padx=10)
        ttk.Button(btn_frame, text="Cancelar", command=self.destroy, style='danger.TButton').pack(side='left', padx=10)
        
    def browse_folder(self, path_var):
        initial = path_var.get() if os.path.isdir(path_var.get()) else "/"
        directory = filedialog.askdirectory(initialdir=initial)
        if directory:
            path_var.set(directory)

    def browse_db_file(self):
        initial_dir = os.path.dirname(self.db_path.get()) if os.path.exists(self.db_path.get()) else "/"
        filepath = filedialog.asksaveasfilename(
            defaultextension=".db",
            filetypes=[("Database files", "*.db"), ("All files", "*.*")],
            initialdir=initial_dir,
            initialfile=os.path.basename(self.db_path.get())
        )
        if filepath:
            self.db_path.set(filepath)

    def save_and_close(self):
        new_config = {
            "ruta_base_ilrl": self.ilrl_path.get(),
            "ruta_base_ilrl_2": self.ilrl_path_2.get(), # Nuevo
            "ruta_base_geo": self.geo_path.get(),
            "ruta_base_geo_2": self.geo_path_2.get(),   # Nuevo
            "ruta_base_ilrl_mpo": self.ilrl_mpo_path.get(),
            "ruta_base_geo_mpo": self.geo_mpo_path.get(),
            "ruta_base_polaridad_mpo": self.polaridad_mpo_path.get(),
            "db_path": self.db_path.get()
        }
        self.app.save_config(new_config)
        messagebox.showinfo("Guardado", "La configuración se ha guardado correctamente.")
        self.destroy()

class MPOConfigWindow(tk.Toplevel):
    def __init__(self, app_instance, parent_page):
        super().__init__(app_instance)
        self.app = app_instance
        self.parent_page = parent_page
        self.title("Configuración de Mediciones MPO")
        self.geometry("350x300")
        self.transient(self.app)
        self.grab_set()

        # Título
        lbl_title = ttk.Label(self, text="Seleccionar Mediciones Activas", font=("Helvetica", 12, "bold"))
        lbl_title.pack(pady=15)

        frame_switches = ttk.Frame(self)
        frame_switches.pack(fill="both", expand=True, padx=40)

        # Variables booleanas cargadas desde config
        self.var_ilrl = tk.BooleanVar(value=self.app.config.get('check_mpo_ilrl', True))
        self.var_geo = tk.BooleanVar(value=self.app.config.get('check_mpo_geo', True))
        self.var_pol = tk.BooleanVar(value=self.app.config.get('check_mpo_polaridad', True))

        # Switches (Checkbuttons) con estilo 'round-toggle' si usas ttkbootstrap
        chk_ilrl = ttk.Checkbutton(frame_switches, text="Medición IL/RL", variable=self.var_ilrl, bootstyle="round-toggle")
        chk_ilrl.pack(anchor="w", pady=10, fill='x')
        
        chk_geo = ttk.Checkbutton(frame_switches, text="Medición Geometría", variable=self.var_geo, bootstyle="round-toggle")
        chk_geo.pack(anchor="w", pady=10, fill='x')

        chk_pol = ttk.Checkbutton(frame_switches, text="Medición Polaridad", variable=self.var_pol, bootstyle="round-toggle")
        chk_pol.pack(anchor="w", pady=10, fill='x')

        # Botones de Acción
        btn_frame = ttk.Frame(self, padding=20)
        btn_frame.pack(side="bottom", fill="x")

        ttk.Button(btn_frame, text="Guardar Cambios", command=self.save_config, style="success.TButton").pack(side="right", padx=5)
        ttk.Button(btn_frame, text="Cancelar", command=self.destroy, style="danger.TButton").pack(side="right", padx=5)

    def save_config(self):
        # Actualizamos solo las llaves relevantes en la configuración
        new_settings = self.app.config.copy()
        new_settings["check_mpo_ilrl"] = self.var_ilrl.get()
        new_settings["check_mpo_geo"] = self.var_geo.get()
        new_settings["check_mpo_polaridad"] = self.var_pol.get()
        
        self.app.save_config(new_settings)
        messagebox.showinfo("Configuración", "Preferencias actualizadas correctamente.", parent=self)
        self.destroy()

class VerificacionMPO_Page(ttk.Frame):
    def __init__(self, parent, app_instance):
        super().__init__(parent, padding=10)
        self.app = app_instance
        self.last_ilrl_result = None
        self.last_geo_result = None
        self.last_polaridad_result = None
        self.create_widgets()
        
    # En la clase VerificacionMPO_Page, reemplaza este método:
    def create_widgets(self):
        container = ttk.Frame(self, style='TFrame', padding=20)
        container.pack(expand=True, fill='both')
        top_frame = ttk.Frame(container)
        top_frame.pack(fill='x', pady=10)
        input_frame = ttk.Frame(top_frame)
        input_frame.pack(side='left', fill='x', expand=True)
        ttk.Label(input_frame, text="Número de OT:", font=("Helvetica", 11, "bold")).grid(row=0, column=0, sticky='w', padx=5, pady=5)
        self.ot_entry = ttk.Entry(input_frame, width=30, font=("Helvetica", 11))
        self.ot_entry.grid(row=0, column=1, sticky='ew', padx=5, pady=5)
        ttk.Label(input_frame, text="Número de Serie (13 dígitos):", font=("Helvetica", 11, "bold")).grid(row=1, column=0, sticky='w', padx=5, pady=5)
        self.serie_entry = ttk.Entry(input_frame, width=30, font=("Helvetica", 11))
        self.serie_entry.grid(row=1, column=1, sticky='ew', padx=5, pady=5)
        self.serie_entry.bind("<KeyRelease>", self.verificar_cable_automatico)

        # --- LÍNEA AÑADIDA ---
        # Conectamos nuestra nueva función al evento <FocusIn> (cuando se selecciona el campo)
        self.serie_entry.bind("<FocusIn>", self.on_serie_focus_in)
        # --------------------
        input_frame.columnconfigure(1, weight=1)
        action_buttons_frame = ttk.Frame(top_frame)
        action_buttons_frame.pack(side='left', padx=20)
        
        ot_details_button = ttk.Button(action_buttons_frame, text="Ver Detalles OT", command=self.open_ot_details_window, style='primary.Outline.TButton', padding=10)
        ot_details_button.pack(pady=5, fill='x')
        
        ot_config_button = ttk.Button(action_buttons_frame, text="Configurar OT", command=self.open_ot_config_window, style='info.TButton', padding=10)
        ot_config_button.pack(pady=5, fill='x')

        # --- NUEVO BOTÓN: Configuración de Búsqueda ---
        btn_search_config = ttk.Button(action_buttons_frame, text="⚙ Config. Búsqueda", command=self.open_search_config, style='secondary.TButton', padding=10)
        btn_search_config.pack(pady=5, fill='x')
        # ----------------------------------------------
        verify_button = ttk.Button(action_buttons_frame, text="Verificar (Manual)", command=self.verificar_cable, style='success.TButton', padding=10)
        verify_button.pack(pady=5, fill='x')

        self.result_text = tk.Text(container, height=15, width=80, wrap="word", font=("Courier New", 10), state=tk.DISABLED, relief="flat", bg="#f0f0f0")
        self.result_text.pack(fill='both', expand=True, pady=10)
        self.setup_text_tags()
        self.show_waiting_message()

    def open_search_config(self):
        MPOConfigWindow(self.app, self)
    
    def on_serie_focus_in(self, event=None):
        """
        Este método se activa cuando el campo de texto del número de serie
        recibe el foco (al hacer clic o al escanear).
        Selecciona todo el texto que haya en él.
        """
        self.serie_entry.select_range(0, tk.END)

    # En la clase VerificacionMPO_Page
    # En la clase VerificacionMPO_Page, reemplaza este método:

    def setup_text_tags(self):
        tags = {
            "header": {"font": ("Helvetica", 14, "bold"), "foreground": "#0056b3"},
            "bold": {"font": ("Courier New", 10, "bold")},
            "info": {"font": ("Courier New", 10, "italic"), "foreground": "grey"},
            "APROBADO": {"foreground": "#28a745"}, "PASS": {"foreground": "#28a745"},
            "RECHAZADO": {"foreground": "#dc3545"}, "FAIL": {"foreground": "#dc3545"},
            "NO ENCONTRADO": {"foreground": "#fd7e14"},
            "DATOS INCOMPLETOS": {"foreground": "#FF8C00", "font": ("Courier New", 10, "bold")},
            "DESACTIVADO": {"foreground": "#999999", "font": ("Courier New", 10, "italic")},
            "ERROR": {"foreground": "#dc3545", "font": ("Courier New", 10, "bold")},
            "final_status_large": {"font": ("Courier New", 14, "bold")}
        }
        for tag, options in tags.items():
            self.result_text.tag_configure(tag, **options)
            
        # --- BLOQUE MODIFICADO ---
        for test_type in ["ilrl", "geo", "polaridad"]:
            # 1. Hipervínculo para el ESTADO (abre detalles)
            link_tag = f"{test_type}_link"
            self.result_text.tag_configure(link_tag, foreground="#0056b3", underline=True)
            self.result_text.tag_bind(link_tag, "<Button-1>", lambda e, t=test_type: self.show_details_window(t))
            self.result_text.tag_bind(link_tag, "<Enter>", lambda e: self.result_text.config(cursor="hand2"))
            self.result_text.tag_bind(link_tag, "<Leave>", lambda e: self.result_text.config(cursor=""))

            # 2. Hipervínculo para el ARCHIVO (abre explorador)
            file_link_tag = f"{test_type}_file_link"
            self.result_text.tag_configure(file_link_tag, foreground="#4682B4", underline=True) # Un azul diferente (SteelBlue)
            self.result_text.tag_bind(file_link_tag, "<Button-1>", lambda e, t=test_type: self.open_file_location(t))
            self.result_text.tag_bind(file_link_tag, "<Enter>", lambda e: self.result_text.config(cursor="hand2"))
            self.result_text.tag_bind(file_link_tag, "<Leave>", lambda e: self.result_text.config(cursor=""))
        # --- FIN DEL BLOQUE ---

    def show_waiting_message(self):
        self.result_text.config(state=tk.NORMAL)
        self.result_text.delete(1.0, tk.END)
        self.result_text.insert(tk.END, "Esperando un número de serie válido (13 dígitos)...", "info")
        self.result_text.config(state=tk.DISABLED)
    
    def open_ot_details_window(self):
        ot_input = self.ot_entry.get().strip().upper()
        if not ot_input:
            messagebox.showwarning("Falta OT", "Por favor, ingrese un número de O.T. para ver sus detalles.", parent=self)
            return
        if not ot_input.startswith('JMO-'):
            ot_input = f"JMO-{ot_input}"
        ot_data = self._cargar_ot_configuration(ot_input)
        if not ot_data:
            messagebox.showinfo("No Encontrado", f"No se encontró ninguna configuración para la O.T.: {ot_input}", parent=self)
            return
        OTDetailsWindow(self, ot_data)

    def open_ot_config_window(self):
        current_ot = self.ot_entry.get().strip().upper()
        if not current_ot:
            messagebox.showwarning("Falta OT", "Ingrese un número de OT antes de configurar.", parent=self)
            return
        if not current_ot.startswith("JMO-"):
            current_ot = f"JMO-{current_ot}"
        OTConfigWindow(self, self.app, current_ot)
    
    # En la clase VerificacionMPO_Page, reemplaza este método:

    def verificar_cable_automatico(self, event=None):
        serie_raw = self.serie_entry.get().strip()
        
        # --- MODIFICACIÓN: Regex ajustado para aceptar JMO o JRMO ---
        # Explicación: J(R)?MO significa que la 'R' es opcional.
        if re.match(r'J(R)?MO\d{13}', serie_raw, re.IGNORECASE):
            # Extraemos solo los números
            numeros_serie = re.sub(r'[^0-9]', '', serie_raw)
            
            if not self.ot_entry.get().strip():
                self.ot_entry.delete(0, tk.END)
                self.ot_entry.insert(0, f"JMO-{numeros_serie[:9]}")
            
            self.verificar_cable()
            
        elif len(serie_raw) == 13 and serie_raw.isdigit():
            if not self.ot_entry.get().strip():
                self.ot_entry.delete(0, tk.END)
                self.ot_entry.insert(0, f"JMO-{serie_raw[:9]}")
            self.verificar_cable()
        elif len(serie_raw) > 0:
            self.show_waiting_message()

    # En la clase VerificacionMPO_Page, reemplaza este método:

    def verificar_cable(self, event=None):
        ot_raw = self.ot_entry.get().strip().upper()
        serie_raw = self.serie_entry.get().strip()

        if not ot_raw or not serie_raw:
            self.show_waiting_message()
            return

        # 1. Normalización numérica
        serie_numerica = re.sub(r'[^0-9]', '', serie_raw)

        if len(serie_numerica) != 13:
            messagebox.showerror("Formato Inválido", "El número de serie escaneado o escrito no contiene 13 dígitos.")
            self.show_waiting_message()
            return

        # Validación de coincidencia OT
        ot_from_serie = serie_numerica[:9]
        ot_from_input = re.sub(r'[^0-9]', '', ot_raw)

        if ot_from_serie != ot_from_input:
            messagebox.showerror(
                "Error de Coincidencia",
                "La OT del número de serie no corresponde a la OT trabajada."
            )
            return

        # Estandarizamos OT
        ot = f"JMO-{ot_raw}" if not ot_raw.startswith("JMO-") else ot_raw
        
        # Detectar prefijo JMO o JRMO
        prefijo_serie = "JRMO-" if "JRMO" in serie_raw.upper() else "JMO-"
        serie = f"{prefijo_serie}{serie_numerica}"

        ot_config = self._cargar_ot_configuration(ot)
        if not ot_config:
            messagebox.showwarning("Configuración Faltante", f"No se encontró configuración para la OT {ot}. Por favor, configurela primero.", parent=self)
            return
            
        self.result_text.config(state=tk.NORMAL)
        self.result_text.delete(1.0, tk.END)
        self.result_text.insert(tk.END, f"Verificando cable MPO {serie} en OT {ot}...\n", "header")
        self.result_text.insert(tk.END, "-"*70 + "\n\n")
        self.update_idletasks()
        
        # --- LÓGICA DE SWITCHES ---
        
        # 1. IL/RL
        if self.app.config.get('check_mpo_ilrl', True):
            self.last_ilrl_result = self.buscar_y_procesar_ilrl_mpo(ot, serie, ot_config)
        else:
            self.last_ilrl_result = {'status': 'DESACTIVADO', 'details': 'Medición desactivada por el usuario.', 'raw_data': []}

        # 2. Geometría
        if self.app.config.get('check_mpo_geo', True):
            self.last_geo_result = self.buscar_y_procesar_geo_mpo(ot, serie, ot_config)
        else:
            self.last_geo_result = {'status': 'DESACTIVADO', 'details': 'Medición desactivada por el usuario.', 'raw_data': []}

        # 3. Polaridad
        if self.app.config.get('check_mpo_polaridad', True):
            self.last_polaridad_result = self.buscar_y_procesar_polaridad_mpo(ot, serie)
        else:
            self.last_polaridad_result = {'status': 'DESACTIVADO', 'details': 'Medición desactivada por el usuario.', 'raw_data': {}}
        
        # Mostrar Resultados
        self.mostrar_resultado_mpo("IL/RL", self.last_ilrl_result)
        self.mostrar_resultado_mpo("Geometría", self.last_geo_result)
        self.mostrar_resultado_mpo("Polaridad", self.last_polaridad_result)
        
        # Lógica de Semáforo Final
        # Si está desactivado, cuenta como True para no bloquear el aprobado general
        ilrl_pass = self.last_ilrl_result['status'] in ["APROBADO", "DESACTIVADO"]
        geo_pass = self.last_geo_result['status'] in ["APROBADO", "DESACTIVADO"]
        polaridad_pass = self.last_polaridad_result['status'] in ["PASS", "APROBADO", "DESACTIVADO"]
        
        final_status = "APROBADO" if ilrl_pass and geo_pass and polaridad_pass else "RECHAZADO"
        
        self.result_text.insert(tk.END, "\n" + "-"*70 + "\n")
        self.result_text.insert(tk.END, "ESTADO FINAL: ", ("bold", "final_status_large"))
        self.result_text.insert(tk.END, f"{final_status}\n", (final_status, "final_status_large"))
        
        if winsound:
            try:
                if final_status == "APROBADO":
                    winsound.Beep(1200, 200)
                elif final_status == "RECHAZADO":
                    winsound.Beep(400, 500)
            except Exception as e:
                print(f"No se pudo reproducir el sonido: {e}")

        self.result_text.config(state=tk.DISABLED)
        
        # Guardar en log (Base de datos)
        log_data = {'serial_number': serie, 'ot_number': ot, 'overall_status': final_status,
                    'ilrl_status': self.last_ilrl_result['status'], 'ilrl_details': json.dumps(self.last_ilrl_result, default=str),
                    'geo_status': self.last_geo_result['status'], 'geo_details': json.dumps(self.last_geo_result, default=str),
                    'polaridad_status': self.last_polaridad_result['status'], 'polaridad_details': json.dumps(self.last_polaridad_result, default=str)}
        self._log_mpo_verification(log_data)

    def _log_mpo_verification(self, log_data):
        try:
            conn = sqlite3.connect(self.app.config['db_path'])
            cursor = conn.cursor()
            cursor.execute("""
                INSERT INTO cable_verifications (
                    entry_date, serial_number, ot_number, overall_status,
                    ilrl_status, ilrl_details, geo_status, geo_details,
                    polaridad_status, polaridad_details
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                log_data['serial_number'], log_data['ot_number'], log_data['overall_status'],
                log_data['ilrl_status'], log_data['ilrl_details'], log_data['geo_status'],
                log_data['geo_details'], log_data['polaridad_status'], log_data['polaridad_details']
            ))
            conn.commit()
            conn.close()
        except Exception as e:
            messagebox.showerror("Error de Base de Datos", f"No se pudo registrar la verificación MPO: {e}", parent=self)

    # En la clase VerificacionMPO_Page, reemplaza este método:

    # En la clase VerificacionMPO_Page, reemplaza este método:

    def mostrar_resultado_mpo(self, tipo, resultado):
        """Muestra una línea de resultado en el widget de texto."""
        # Mapeo para los tags de DETALLES (Estado)
        tipo_map = {
            "IL/RL": "ilrl_link",
            "Geometría": "geo_link",
            "Polaridad": "polaridad_link"
        }
        # Mapeo para los tags de ARCHIVO (Explorador)
        file_tipo_map = {
            "IL/RL": "ilrl_file_link",
            "Geometría": "geo_file_link",
            "Polaridad": "polaridad_file_link"
        }
        
        link_tag = tipo_map.get(tipo, "") 
        file_link_tag = file_tipo_map.get(tipo, "") # <-- Nuevo tag de archivo

        self.result_text.insert(tk.END, f"{tipo}:\n", "bold")
        self.result_text.insert(tk.END, f"  Estado: ")

        status = resultado.get('status', 'ERROR')
        details = resultado.get('details', 'No hay detalles disponibles.')

        # Aplicar el tag de ESTADO (para detalles) solo al 'status'
        self.result_text.insert(tk.END, f"{status}", (status, link_tag))
        
        # --- INICIO DE LA MODIFICACIÓN: Aplicar el tag de ARCHIVO ---
        if "Archivo:" in details:
            details_text, file_text = details.rsplit("Archivo:", 1)
            file_text = "Archivo:" + file_text
            
            self.result_text.insert(tk.END, f"\n  Detalles: {details_text.strip()}")
            # Aplicar el tag de ARCHIVO (para explorador) solo al texto del archivo
            self.result_text.insert(tk.END, f"\n  {file_text.strip()}", (file_link_tag)) 
            self.result_text.insert(tk.END, "\n\n")
        else:
            self.result_text.insert(tk.END, f"\n  Detalles: {details}\n\n")
        # --- FIN DE LA MODIFICACIÓN ---

    def _cargar_ot_configuration(self, ot_number):
        try:
            conn = sqlite3.connect(self.app.config['db_path'])
            conn.row_factory = sqlite3.Row
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM ot_configurations WHERE ot_number = ?", (ot_number,))
            row = cursor.fetchone()
            conn.close()
            return dict(row) if row else None
        except Exception as e:
            messagebox.showerror("Error de Base de Datos", f"No se pudo cargar la configuración de la OT: {e}", parent=self)
            return None

    # En la clase VerificacionMPO_Page, reemplaza este método completo:

    # En la clase VerificacionMPO_Page, reemplaza este método completo:

    # En la clase VerificacionMPO_Page, reemplaza este método completo:

    def buscar_y_procesar_ilrl_mpo(self, ot, serie, config):
        ruta_ot_ilrl = os.path.join(self.app.config['ruta_base_ilrl_mpo'], ot)
        if not os.path.isdir(ruta_ot_ilrl): 
            return {'status': 'NO ENCONTRADO', 'details': f'Carpeta de OT no encontrada en ILRL MPO.', 'raw_data': []}
        
        archivos_encontrados = [os.path.join(ruta_ot_ilrl, f) for f in os.listdir(ruta_ot_ilrl) if f.lower().endswith(('.xlsx', '.xls')) and not f.startswith('~$') and ot.lower() in f.lower()]
        if not archivos_encontrados: 
            return {'status': 'NO ENCONTRADO', 'details': f'Ningún archivo ILRL para la OT "{ot}".', 'raw_data': []}
        
        archivo_a_procesar = max(archivos_encontrados, key=os.path.getmtime)
        
        try:
            df = pd.read_excel(archivo_a_procesar, sheet_name="Results")
            h = {k: config.get(v, d) for k, v, d in [('serie', 'ilrl_serie_header', 'Serial number'), ('estado', 'ilrl_estado_header', 'Alarm Status'), ('conector', 'ilrl_conector_header', 'connector label'), ('fecha', 'ilrl_fecha_header', 'Date'), ('hora', 'ilrl_hora_header', 'Time')]}

            if any(header not in df.columns for header in h.values()):
                return {'status': 'ERROR', 'details': f"Faltan encabezados en {os.path.basename(archivo_a_procesar)}", 'raw_data': []}

            # --- CORRECCIÓN 1: Crear la columna base_serie ANTES de usarla ---
            df[h['serie']] = df[h['serie']].astype(str)
            df['base_serie'] = df[h['serie']].str.extract(r'(\d{13})')
            # -----------------------------------------------------------------

            serie_numerica_buscada = re.sub(r'[^0-9]', '', serie)
            
            # Ahora sí podemos filtrar porque 'base_serie' ya existe
            df_cable_group_all = df[df['base_serie'] == serie_numerica_buscada].copy()

            date_series = pd.to_datetime(df[h['fecha']], dayfirst=True, errors='coerce').dt.strftime('%d/%m/%Y')
            time_series = df[h['hora']].astype(str).str.replace('a. m.', 'AM', regex=False).str.replace('p. m.', 'PM', regex=False).str.strip()
            full_datetime_str = date_series + ' ' + time_series
            df['timestamp'] = pd.to_datetime(full_datetime_str, format='%d/%m/%Y %I:%M:%S %p', errors='coerce')
            
            # Volvemos a filtrar el grupo con timestamp ya calculado para validaciones
            df_cable_group_all = df[df['base_serie'] == serie_numerica_buscada].copy()
            
            if df_cable_group_all.empty:
                return {'status': 'NO ENCONTRADO', 'details': f'Serie no encontrada en {os.path.basename(archivo_a_procesar)}', 'raw_data': []}
            
            filas_invalidas = df_cable_group_all[df_cable_group_all['timestamp'].isna()]

            if not filas_invalidas.empty:
                details_str = f"El cable tiene {len(filas_invalidas)} mediciones sin fecha/hora válidas. Revise las columnas '{h['fecha']}' y '{h['hora']}'."
                raw_data_list = [{'conector': row[h['conector']], 'resultado': row[h['estado']], 'fecha_original': str(row[h['fecha']]), 'hora_original': str(row[h['hora']])} for _, row in filas_invalidas.iterrows()]
                return {'status': 'DATOS INCOMPLETOS', 'details': details_str, 'raw_data': raw_data_list, 'file_path': archivo_a_procesar, 'error_type': 'fechas_invalidas'}

            latest_timestamp = pd.NaT
            best_sub_group_df = None
            serie_para_reporte = "" 

            for full_serial in df_cable_group_all[h['serie']].unique():
                sub_group_df = df_cable_group_all[df_cable_group_all[h['serie']] == full_serial]
                current_max_timestamp = sub_group_df['timestamp'].max() 

                if best_sub_group_df is None or current_max_timestamp > latest_timestamp:
                    latest_timestamp = current_max_timestamp
                    best_sub_group_df = sub_group_df
                    serie_para_reporte = full_serial 
            
            if best_sub_group_df is None:
                return {'status': 'NO ENCONTRADO', 'details': 'No se encontraron mediciones válidas.', 'raw_data': []}
                
            df_cable = best_sub_group_df.copy()

            num_conectores_a = config.get('num_conectores_a', 1)
            fibras_por_conector_a = config.get('fibers_per_connector_a', 12)
            num_conectores_b = config.get('num_conectores_b', 1)
            fibras_por_conector_b = config.get('fibers_per_connector_b', 12)
            total_fibras_esperadas = (num_conectores_a * fibras_por_conector_a) + (num_conectores_b * fibras_por_conector_b)
            
            df_cable.sort_values(by='timestamp', ascending=False, inplace=True)
            df_lado_a = df_cable[df_cable[h['conector']] == 'A'].head(fibras_por_conector_a * num_conectores_a)
            df_lado_b = df_cable[df_cable[h['conector']] == 'B'].head(fibras_por_conector_b * num_conectores_b)
            
            df_final_cable = pd.concat([df_lado_a, df_lado_b])

            if df_final_cable.empty:
                return {'status': 'NO ENCONTRADO', 'details': 'El grupo de mediciones más reciente está vacío.', 'raw_data': []}

            raw_data = []
            overall_pass = True
            pass_count = 0
            
            df_final_cable.loc[:, h['estado']] = df_final_cable[h['estado']].str.strip().str.upper()
            
            mediciones_lado_a = []
            estado_lado_a = "APROBADO"
            for _, row in df_lado_a.iterrows():
                resultado = str(row[h['estado']]).strip().upper()
                mediciones_lado_a.append({'fibra': len(mediciones_lado_a)+1, 'resultado': resultado})
                if resultado == 'PASS': pass_count += 1
                else: estado_lado_a = "RECHAZADO"; overall_pass = False
            if not mediciones_lado_a: estado_lado_a = "NO ENCONTRADO"
            raw_data.append({'conector': 'A', 'estado': estado_lado_a, 'mediciones': mediciones_lado_a})
            
            mediciones_lado_b = []
            estado_lado_b = "APROBADO"
            for _, row in df_lado_b.iterrows():
                resultado = str(row[h['estado']]).strip().upper()
                mediciones_lado_b.append({'fibra': len(mediciones_lado_b)+1, 'resultado': resultado})
                if resultado == 'PASS': pass_count += 1
                else: estado_lado_b = "RECHAZADO"; overall_pass = False
            if not mediciones_lado_b: estado_lado_b = "NO ENCONTRADO"
            raw_data.append({'conector': 'B', 'estado': estado_lado_b, 'mediciones': mediciones_lado_b})

            total_mediciones = len(df_final_cable)
            if total_mediciones != total_fibras_esperadas: overall_pass = False

            status = 'APROBADO' if overall_pass else 'RECHAZADO'
            
            if total_mediciones != total_fibras_esperadas:
                details = f"Incompleto: {pass_count}/{total_mediciones} mediciones OK (Esperadas: {total_fibras_esperadas}). Archivo: {os.path.basename(archivo_a_procesar)}"
            else:
                details = f"{pass_count}/{total_mediciones} mediciones OK. Archivo: {os.path.basename(archivo_a_procesar)}"
            
            return {'status': status, 'details': details, 'raw_data': raw_data, 'file_path': archivo_a_procesar, 'serial_number': serie_para_reporte}

        except Exception as e: 
            return {'status': 'ERROR', 'details': f'Fallo al procesar {os.path.basename(archivo_a_procesar)}: {traceback.format_exc()}', 'raw_data': []}

    def buscar_y_procesar_geo_mpo(self, ot, serie, config):
        ruta_base_geo = self.app.config['ruta_base_geo_mpo']
        if not os.path.isdir(ruta_base_geo): 
            return {'status': 'NO ENCONTRADO', 'details': 'Carpeta de Geometría MPO no encontrada.', 'raw_data': []}
        
        ot_numerica = ot.replace('JMO-', '')
        archivos_encontrados = [os.path.join(ruta_base_geo, f) for f in os.listdir(ruta_base_geo) if ot_numerica in f and f.lower().endswith(('.xlsx', '.xls')) and not f.startswith('~$')]
        
        if not archivos_encontrados: 
            return {'status': 'NO ENCONTRADO', 'details': f'Ningún archivo de Geometría para la OT "{ot}".', 'raw_data': []}
        
        archivo_a_procesar = max(archivos_encontrados, key=os.path.getmtime)

        try:
            full_df = pd.read_excel(archivo_a_procesar, sheet_name="MT12", header=None)
            header_row_index = next((i for i, row in full_df.iterrows() if str(row.iloc[0]).strip().lower() == 'name'), -1)
            
            if header_row_index == -1: 
                return {'status': 'ERROR', 'details': "No se encontró encabezado 'Name' en la hoja 'MT12'.", 'raw_data': []}
            
            df = full_df.iloc[header_row_index:].copy()
            df.columns = df.iloc[0].str.strip().str.lower()
            df = df[1:].rename(columns={'pass/fail': 'result'})
            
            df['temp_serie_numerica'] = df['name'].astype(str).str.extract(r'(\d{13})')
            serie_numerica_buscada = re.sub(r'[^0-9]', '', serie)
            
            # --- CORRECCIÓN 2: Corregido el nombre de la variable (df_cable_cable -> df_cable) ---
            df_cable = df[df['temp_serie_numerica'] == serie_numerica_buscada].copy()
            # -------------------------------------------------------------------------------------

            if df_cable.empty: 
                return {'status': 'NO ENCONTRADO', 'details': f'Serie no encontrada en {os.path.basename(archivo_a_procesar)}', 'raw_data': []}
            
            def parse_geo_name_mpo(name_str):
                match = re.match(r'(JMO\d{13})-(R?\d+)(-[FK])?', str(name_str).upper())
                if match: return match.group(1), match.group(2)
                return None, None
            
            parsed_data = df_cable['name'].apply(lambda x: pd.Series(parse_geo_name_mpo(x)))
            if parsed_data.empty or parsed_data.shape[1] < 2:
                 return {'status': 'ERROR', 'details': f'No se pudieron parsear los nombres de los conectores.', 'raw_data': []}

            df_cable[['serie_base', 'lado']] = parsed_data

            if 'date & time' in df_cable.columns:
                df_cable['timestamp'] = pd.to_datetime(df_cable['date & time'], dayfirst=True, errors='coerce')
            else:
                df_cable['timestamp'] = pd.to_datetime(df_cable['date'].astype(str) + ' ' + df_cable['time'].astype(str), dayfirst=True, errors='coerce')
            
            df_cable = df_cable.sort_values(by='timestamp', ascending=False)
            
            originales = {}
            retrabajos = {}
            for _, medicion in df_cable.iterrows():
                punta_original = str(medicion['lado'])
                punta_normalizada = punta_original.replace('R', '')
                if 'R' in punta_original:
                    if punta_normalizada not in retrabajos:
                        retrabajos[punta_normalizada] = {'Punta_Original': punta_original, 'Resultado': medicion['result'], 'Fuente': medicion['name']}
                else:
                    if punta_normalizada not in originales:
                        originales[punta_normalizada] = {'Punta_Original': punta_original, 'Resultado': medicion['result'], 'Fuente': medicion['name']}
            
            mediciones_definitivas = {}
            for punta, medicion in retrabajos.items():
                mediciones_definitivas[punta] = medicion
            for punta, medicion in originales.items():
                if punta not in mediciones_definitivas:
                    mediciones_definitivas[punta] = medicion

            # --- LÍNEA DE DEPURACIÓN ---
            print("\n*** MEDICIONES DEFINITIVAS ELEGIDAS ***")
            print(mediciones_definitivas)
            print("************************************\n")
            # ------------------------------------

            status = 'APROBADO'
            total_conectores_esperados = config.get('num_conectores_a', 0) + config.get('num_conectores_b', 0)
            lados_esperados = [str(i + 1) for i in range(total_conectores_esperados)]
            pass_count = 0
            
            for lado_req in lados_esperados:
                if lado_req in mediciones_definitivas:
                    if str(mediciones_definitivas[lado_req]['Resultado']).upper() == 'PASS':
                        pass_count += 1
                    else:
                        status = 'RECHAZADO'
                else:
                    status = 'RECHAZADO'
            
            if len(mediciones_definitivas) < total_conectores_esperados:
                status = 'RECHAZADO'
                details = f"Faltan mediciones. Encontradas: {len(mediciones_definitivas)}, Esperadas: {total_conectores_esperados}."
            else:
                details = f"{pass_count}/{total_conectores_esperados} mediciones OK."

            details += f" Archivo: {os.path.basename(archivo_a_procesar)}"
            raw_data = [{'conector': v['Punta_Original'], 'resultado': v['Resultado'], 'serie_completo': v['Fuente']} for v in mediciones_definitivas.values()]
            
            # --- LÍNEA MODIFICADA: Añadimos serial_number al retorno ---
            # Como todas las filas de df_cable pertenecen al mismo cable base, tomamos el primero
            primer_serial_completo = df_cable.iloc[0]['name'] if not df_cable.empty else serie
            
            return {'status': status, 'details': details, 'raw_data': raw_data, 'file_path': archivo_a_procesar, 'serial_number': primer_serial_completo}

        except Exception as e:
            return {'status': 'ERROR', 'details': f'Fallo al procesar {os.path.basename(archivo_a_procesar)}: {traceback.format_exc()}', 'raw_data': []}
        
    def buscar_y_procesar_polaridad_mpo(self, ot, serie):
        ruta_base_polaridad = self.app.config['ruta_base_polaridad_mpo']
        if not os.path.isdir(ruta_base_polaridad): return {'status': 'NO ENCONTRADO', 'details': 'Carpeta de Polaridad MPO no encontrada.', 'raw_data': {}}
        serie_sin_prefijo = re.sub(r'[^0-9]', '', serie)
        archivos_encontrados = [os.path.join(root, f) for root, _, files in os.walk(ruta_base_polaridad) for f in files if serie_sin_prefijo in f and f.lower().endswith('.xlsx') and not f.startswith('~$')]
        if not archivos_encontrados: return {'status': 'NO ENCONTRADO', 'details': f'Ningún archivo de Polaridad para la serie "{serie}".', 'raw_data': {}}
        archivo_a_procesar = max(archivos_encontrados, key=os.path.getmtime)
        try:
            df = pd.read_excel(archivo_a_procesar, sheet_name=0, header=None)
            serial_number_raw = df.iloc[3, 1] if df.shape[0] > 3 and df.shape[1] > 1 else None
            status_raw = df.iloc[12, 1] if df.shape[0] > 12 and df.shape[1] > 1 else None
            file_name = os.path.basename(archivo_a_procesar)
            date_match = re.search(r'(\d{4}-\d{2}-\d{2} \d{2}-\d{2}-\d{2})', file_name)
            date_raw = date_match.group(1).replace('-', ' ') if date_match else "N/A"
            if not serial_number_raw or not status_raw: return {'status': 'ERROR', 'details': f'Datos clave no encontrados en {file_name}', 'raw_data': {}}
            status = str(status_raw).strip().upper()
            details = f"Resultado de {file_name}. Fecha: {date_raw}"
            raw_data = {'serial_number': str(serial_number_raw), 'status': status, 'test_date': date_raw, 'file_name': file_name}
            return {'status': status, 'details': details, 'raw_data': raw_data, 'file_path': archivo_a_procesar}
        except Exception as e: return {'status': 'ERROR', 'details': f'Fallo al procesar {os.path.basename(archivo_a_procesar)}: {e}', 'raw_data': {}}

    def show_details_window(self, analysis_type):
        data = getattr(self, f"last_{analysis_type}_result", None)
        title = f"Detalles de Análisis {analysis_type.upper()} (MPO)"
        if not data or not data.get('raw_data'):
            messagebox.showinfo(title, "No hay datos detallados para mostrar.")
            return
        DetailsWindow(self, title, data, analysis_type)
    
    # En la clase VerificacionMPO_Page, añade esta nueva función:

    # En la clase VerificacionMPO_Page, reemplaza este método:

    # En la clase VerificacionMPO_Page, reemplaza este método:

    # En la clase VerificacionMPO_Page, reemplaza este método:

    # En la clase VerificacionMPO_Page, reemplaza este método:

    def open_file_location(self, analysis_type):
        """Abre la carpeta que contiene el archivo de reporte."""
        data = getattr(self, f"last_{analysis_type}_result", None)
        if not data or not data.get('file_path'):
            messagebox.showinfo("Error", "No se encontró la ruta del archivo.", parent=self)
            return
            
        file_path = data['file_path']
        if not os.path.exists(file_path):
            messagebox.showerror("Error", f"La ruta del archivo no existe:\n{file_path}", parent=self)
            return
            
        try:
            # --- INICIO DE LA CORRECCIÓN ---
            # 1. Obtener el DIRECTORIO que contiene el archivo
            folder_path = os.path.dirname(file_path)
            
            # 2. Asegurarse de que la ruta sea absoluta
            path_absoluto = os.path.abspath(folder_path)

            if not os.path.isdir(path_absoluto):
                messagebox.showerror("Error", f"La ruta de la carpeta no existe:\n{path_absoluto}", parent=self)
                return

            # 3. Usar os.startfile() - es la forma más robusta en Windows
            #    para abrir una carpeta en el explorador.
            os.startfile(path_absoluto)
            # --- FIN DE LA CORRECCIÓN ---
            
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo abrir la carpeta del archivo:\n{e}", parent=self)

class OTConfigWindow(tk.Toplevel):
    def __init__(self, parent_page, app_instance, ot_number):
        super().__init__(parent_page)
        self.parent_page = parent_page
        self.app = app_instance
        self.ot_number = ot_number
        self.title(f"Configuración para OT: {ot_number}")
        self.geometry("800x700")
        self.transient(parent_page)
        self.grab_set()
        main_frame = ttk.Frame(self)
        main_frame.pack(fill=tk.BOTH, expand=True)
        canvas = tk.Canvas(main_frame, highlightthickness=0)
        yscrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=yscrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        yscrollbar.pack(side="right", fill="y")
        def _on_canvas_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
            canvas.itemconfigure(0, width=event.width)
        canvas.bind("<Configure>", _on_canvas_configure)
        frame = ttk.Frame(scrollable_frame, padding=20)
        frame.pack(fill='x', expand=True)
        
        # Diccionario de variables con nombres en INGLÉS para consistencia
        self.vars = {
            "drawing_number": tk.StringVar(), "link": tk.StringVar(),
            "num_conectores_a": tk.StringVar(value="1"), "fibers_per_connector_a": tk.StringVar(value="12"),
            "num_conectores_b": tk.StringVar(value="1"), "fibers_per_connector_b": tk.StringVar(value="12"),
            "ilrl_ot_header": tk.StringVar(value='Work number'), "ilrl_serie_header": tk.StringVar(value='Serial number'),
            "ilrl_fecha_header": tk.StringVar(value='Date'), "ilrl_hora_header": tk.StringVar(value='Time'),
            "ilrl_estado_header": tk.StringVar(value='Alarm Status'), "ilrl_conector_header": tk.StringVar(value='connector label')
        }
        self.load_existing_config()
        self.create_config_ui(frame)
        self.after(100, self.draw_mpo_cable_config)

    def create_config_ui(self, frame):
        ot_frame = ttk.LabelFrame(frame, text="Configuración General de OT", padding=10)
        ot_frame.pack(fill='x', expand=True, pady=5)
        ot_frame.columnconfigure(1, weight=1)
        ttk.Label(ot_frame, text="Número de Dibujo:").grid(row=0, column=0, sticky='w', padx=5, pady=2)
        ttk.Entry(ot_frame, textvariable=self.vars["drawing_number"]).grid(row=0, column=1, sticky='ew', padx=5, pady=2)
        ttk.Label(ot_frame, text="Link a Dibujo:").grid(row=1, column=0, sticky='w', padx=5, pady=2)
        ttk.Entry(ot_frame, textvariable=self.vars["link"]).grid(row=1, column=1, sticky='ew', padx=5, pady=2)
        
        cable_frame = ttk.LabelFrame(frame, text="Configuración del Cable MPO", padding=10)
        cable_frame.pack(fill='x', expand=True, pady=5)
        cable_frame.columnconfigure(1, weight=1)
        cable_frame.columnconfigure(3, weight=1)
        
        # Lado A
        ttk.Label(cable_frame, text="Lado A - Conectores:").grid(row=0, column=0, sticky='w', padx=5, pady=2)
        ttk.Entry(cable_frame, textvariable=self.vars["num_conectores_a"]).grid(row=0, column=1, sticky='ew', padx=5, pady=2)
        ttk.Label(cable_frame, text="Lado A - Fibras/Conector:").grid(row=0, column=2, sticky='w', padx=5, pady=2)
        ttk.Entry(cable_frame, textvariable=self.vars["fibers_per_connector_a"]).grid(row=0, column=3, sticky='ew', padx=5, pady=2)
        
        # Lado B
        ttk.Label(cable_frame, text="Lado B - Conectores:").grid(row=1, column=0, sticky='w', padx=5, pady=2)
        ttk.Entry(cable_frame, textvariable=self.vars["num_conectores_b"]).grid(row=1, column=1, sticky='ew', padx=5, pady=2)
        ttk.Label(cable_frame, text="Lado B - Fibras/Conector:").grid(row=1, column=2, sticky='w', padx=5, pady=2)
        ttk.Entry(cable_frame, textvariable=self.vars["fibers_per_connector_b"]).grid(row=1, column=3, sticky='ew', padx=5, pady=2)
        
        header_frame = ttk.LabelFrame(frame, text="Encabezados de Reporte ILRL", padding=10)
        header_frame.pack(fill='x', expand=True, pady=5)
        header_frame.columnconfigure(1, weight=1)
        headers = [("O.T.:", "ilrl_ot_header"), ("Nro. de Serie:", "ilrl_serie_header"), ("Fecha:", "ilrl_fecha_header"), ("Hora:", "ilrl_hora_header"), ("Estado:", "ilrl_estado_header"), ("Conector:", "ilrl_conector_header")]
        for i, (label, key) in enumerate(headers):
            ttk.Label(header_frame, text=f"Encabezado de {label}").grid(row=i, column=0, sticky=tk.W, pady=2, padx=5)
            ttk.Entry(header_frame, textvariable=self.vars[key]).grid(row=i, column=1, sticky=tk.EW, padx=5)
            
        canvas_frame = ttk.LabelFrame(frame, text="Vista Previa", padding=10)
        canvas_frame.pack(fill='both', expand=True, pady=5)
        self.cable_canvas = Canvas(canvas_frame, bg="white", height=200)
        self.cable_canvas.pack(fill='x', expand=True)
        
        ttk.Button(frame, text="Guardar Configuración", command=self.save_config, style="success.TButton").pack(pady=20)
        
        for key in ["num_conectores_a", "fibers_per_connector_a", "num_conectores_b", "fibers_per_connector_b"]:
            self.vars[key].trace_add("write", self.draw_mpo_cable_config)

    def load_existing_config(self):
        ot_data = self.parent_page._cargar_ot_configuration(self.ot_number)
        if ot_data:
            for key, var in self.vars.items():
                var.set(ot_data.get(key, var.get()))

    def draw_mpo_cable_config(self, *args):
        self.cable_canvas.delete("all")
        try:
            num_ca, fibras_ca, num_cb, fibras_cb = [int(self.vars[k].get()) for k in ["num_conectores_a", "fibers_per_connector_a", "num_conectores_b", "fibers_per_connector_b"]]
        except (ValueError, tk.TclError): return
        w, h = self.cable_canvas.winfo_width(), self.cable_canvas.winfo_height()
        if w < 50 or h < 50: return
        cy = h / 2
        self.cable_canvas.create_line(50, cy, w - 50, cy, width=10, fill="grey")
        total_h_a = num_ca * 20 + (num_ca - 1) * 5
        start_y_a = cy - total_h_a / 2
        for i in range(num_ca):
            y0 = start_y_a + i * 25
            self.cable_canvas.create_rectangle(20, y0, 50, y0 + 20, fill="black", outline="black")
            self.cable_canvas.create_text(35, y0 + 10, text=f"A{i+1}", fill="white", font=("Helvetica", 8, "bold"))
        total_h_b = num_cb * 20 + (num_cb - 1) * 5
        start_y_b = cy - total_h_b / 2
        for i in range(num_cb):
            y0 = start_y_b + i * 25
            self.cable_canvas.create_rectangle(w - 50, y0, w - 20, y0 + 20, fill="black", outline="black")
            self.cable_canvas.create_text(w - 35, y0 + 10, text=f"B{i+1}", fill="white", font=("Helvetica", 8, "bold"))
        self.cable_canvas.create_text(10, h - 10, anchor='sw', text=f"Total Fibras A: {num_ca * fibras_ca}", font=("Helvetica", 9))
        self.cable_canvas.create_text(w - 10, h - 10, anchor='se', text=f"Total Fibras B: {num_cb * fibras_cb}", font=("Helvetica", 9))

    def save_config(self):
        try:
            ot_data = {key: var.get() for key, var in self.vars.items()}
            ot_data['ot_number'] = self.ot_number
            for key in ["num_conectores_a", "fibers_per_connector_a", "num_conectores_b", "fibers_per_connector_b"]:
                ot_data[key] = int(ot_data[key])
            if self.app.guardar_ot_configuration(ot_data):
                messagebox.showinfo("Éxito", "Configuración guardada correctamente.", parent=self)
                self.destroy()
        except ValueError: 
            messagebox.showerror("Error de Entrada", "Los campos de conectores y fibras deben ser números enteros.", parent=self)
        except Exception as e: 
            messagebox.showerror("Error", f"No se pudo guardar la configuración: {e}", parent=self)

class OTDetailsWindow(tk.Toplevel):
    def __init__(self, parent, ot_data):
        super().__init__(parent)
        self.title(f"Detalles de OT: {ot_data['ot_number']}")
        self.geometry("600x450")
        self.transient(parent); self.grab_set()
        frame = ttk.Frame(self, padding=20)
        frame.pack(fill=tk.BOTH, expand=True)
        ttk.Label(frame, text=f"Detalles para {ot_data['ot_number']}", font=("Helvetica", 16, "bold")).pack(anchor='w', pady=(0, 10))
        ttk.Label(frame, text=f"Número de Dibujo: {ot_data['drawing_number']}").pack(anchor='w')
        if ot_data.get('link'):
            link_label = ttk.Label(frame, text=f"Link: {ot_data['link']}", foreground="blue", cursor="hand2")
            link_label.pack(anchor='w')
            link_label.bind("<Button-1>", lambda e: webbrowser.open_new(ot_data['link']))
        ttk.Separator(frame, orient='horizontal').pack(fill='x', pady=10)
        ttk.Label(frame, text="Configuración del Cable:", font=("Helvetica", 12, "bold")).pack(anchor='w', pady=(5,5))
        ttk.Label(frame, text=f"  Lado A: {ot_data['num_conectores_a']} conector(es) de {ot_data['fibers_per_connector_a']} fibras").pack(anchor='w')
        ttk.Label(frame, text=f"  Lado B: {ot_data['num_conectores_b']} conector(es) de {ot_data['fibers_per_connector_b']} fibras").pack(anchor='w')
        ttk.Separator(frame, orient='horizontal').pack(fill='x', pady=10)
        ttk.Label(frame, text="Encabezados de Reporte ILRL:", font=("Helvetica", 12, "bold")).pack(anchor='w', pady=(5,5))
        for key, label in [('ilrl_ot_header', 'OT'), ('ilrl_serie_header', 'Serie'), ('ilrl_fecha_header', 'Fecha'), ('ilrl_hora_header', 'Hora'), ('ilrl_estado_header', 'Estado'), ('ilrl_conector_header', 'Conector')]:
            ttk.Label(frame, text=f"  {label}: '{ot_data[key]}'").pack(anchor='w')
        ttk.Button(frame, text="Cerrar", command=self.destroy, style='secondary.TButton').pack(pady=20)

class DetailsWindow(tk.Toplevel):
    # En la clase DetailsWindow, reemplaza el método __init__:

    def __init__(self, parent, title, data, analysis_type):
        super().__init__(parent)
        self.title(title); self.geometry("750x400"); self.transient(parent); self.grab_set() # Hice la ventana un poco más ancha
        frame = ttk.Frame(self, padding=20)
        frame.pack(fill=tk.BOTH, expand=True)
        ttk.Label(frame, text=f"Detalles para: {data.get('details', 'N/A').split(' Archivo:')[0]}", wraplength=700).pack(anchor='w', pady=(0, 5))
        ttk.Label(frame, text=f"Archivo: {os.path.basename(data.get('file_path', 'N/A'))}", wraplength=700).pack(anchor='w', pady=(0, 10))
        tree_frame = ttk.Frame(frame)
        tree_frame.pack(fill=tk.BOTH, expand=True)
        
        cols, values_list = self.get_details_data(data, analysis_type)
        if not cols: return
        
        tree = ttk.Treeview(tree_frame, columns=cols, show="headings")
        
        # --- BLOQUE MODIFICADO: Ajuste dinámico de columnas ---
        for col in cols:
            tree.heading(col, text=col)
            if "Fuente" in col or "Serie" in col:
                tree.column(col, anchor='w', width=220)
            elif "Resultado" in col or "Conector" in col or "Punta" in col or "Línea" in col or "Fibra" in col:
                tree.column(col, anchor='center', width=80)
            else:
                tree.column(col, anchor='w', width=120)
        # --- FIN DEL BLOQUE ---
        
        tree.tag_configure('PASS', foreground='green'); tree.tag_configure('APROBADO', foreground='green')
        tree.tag_configure('FAIL', foreground='red'); tree.tag_configure('RECHAZADO', foreground='red')
        
        for values in values_list:
            tag_val = ""
            if len(values) > 1: tag_val = values[1] if analysis_type != 'polaridad' else (values[1] if values[0] == 'status' else '')
            tag = 'PASS' if str(tag_val) in ['PASS', 'APROBADO'] else 'FAIL' if str(tag_val) in ['FAIL', 'RECHAZADO'] else ''
            tree.insert("", "end", values=values, tags=(tag,))
        
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        vsb.pack(side='right', fill='y')
        tree.pack(side='left', fill='both', expand=True)
        ttk.Button(frame, text="Cerrar", command=self.destroy, style='secondary.TButton').pack(pady=20)

    # En la clase DetailsWindow, reemplaza este método:

    # En la clase DetailsWindow, reemplaza el método get_details_data:

    def get_details_data(self, data, analysis_type):
        raw_data = data.get('raw_data', [])
        # Extraer el serial number del nivel superior del diccionario 'data'
        serial_num = data.get('serial_number', 'N/A')

        if analysis_type == "ilrl":
            if data.get('error_type') == 'fechas_invalidas':
                cols = ("Conector", "Resultado", "Fecha (Original)", "Hora (Original)", "Número de Serie")
                values_list = [(
                    item.get('conector'), 
                    item.get('resultado'), 
                    item.get('fecha_original'), 
                    item.get('hora_original'),
                    serial_num # Añadir el N/S
                ) for item in raw_data]
                return cols, values_list

            # SC/LC (simple)
            if raw_data and isinstance(raw_data[0], dict) and 'linea' in raw_data[0]:
                cols = ("Línea", "Resultado", "Número de Serie")
                values_list = [(item.get('linea'), item.get('resultado'), serial_num) for item in raw_data]
                return cols, values_list
            # MPO (normal)
            else:
                cols = ("Conector", "Fibra", "Resultado", "Número de Serie")
                values_list = []
                for conn in raw_data:
                    for fib in conn.get('mediciones', []):
                        values_list.append((conn.get('conector'), fib.get('fibra'), fib.get('resultado'), serial_num))
                return cols, values_list
        
        elif analysis_type == "geo":
            # SC/LC
            if raw_data and isinstance(raw_data[0], dict) and 'punta' in raw_data[0]:
                cols = ("Punta", "Resultado", "Número de Serie (Fuente)")
                values_list = [(item.get('punta'), item.get('resultado'), item.get('fuente')) for item in raw_data]
                return cols, values_list
            # MPO
            else:
                cols = ("Conector", "Resultado", "Número de Serie (Fuente)")
                values_list = [(item.get('conector'), item.get('resultado'), item.get('serie_completo')) for item in raw_data]
                return cols, values_list
        
        elif analysis_type == "polaridad":
            return ("Propiedad", "Valor"), list(data.get('raw_data', {}).items())
        
        return (), []

if __name__ == "__main__":
    app = App()
    app.mainloop()

