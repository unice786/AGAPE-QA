import os
import re
import subprocess
import tkinter as tk
from tkinter import filedialog, scrolledtext

try:
   import openpyxl
except ImportError:
   openpyxl = None

try:
   import pypdf
except ImportError:
   pypdf = None

# ------------------------------------------------------------
#  CEREBRO DE EXTRACCIÓN ADAPTADO A MATRICES Y CAPAS FLOTANTES
# ------------------------------------------------------------
def analizar_y_aislar_metadatos_excel(texto_completo, nombre_fichero=None):
   """
   Estrategia Especializada: Detecta metadatos estructurados tanto en celdas
   como en texto extraído de formas de dibujo/objetos OLE flotantes.
   """
   # === 1. CONTROL DE CÓDIGO INTELIGENTE ===
   codigo_fichero = None
   if nombre_fichero:
       match_fn = re.search(r'([A-Z]\.[A-Z]{1,4}\.\d{3})', nombre_fichero.upper())
       if match_fn:
           codigo_fichero = match_fn.group(1)

   texto_comprimido = re.sub(r'[\s|]', '', texto_completo).upper()
   texto_comprimido = (texto_comprimido
                       .replace('Ó', 'O').replace('É', 'E')
                       .replace('Á', 'A').replace('Í', 'I').replace('Ú', 'U'))
   
   texto_comprimido_seguro = texto_comprimido.replace("FECHADEREVISION", "FECHA_REV").replace("FECHAREVISION", "FECHA_REV")

   if codigo_fichero:
       codigo_detectado = codigo_fichero
   else:
       match_cod = re.search(r'CODIGO[:\s|]*([A-Z]\.[A-Z]{1,4}\.\d{3})', texto_completo.upper())
       codigo_detectado = "NO DETECTADO"
       if match_cod:
           codigo_detectado = match_cod.group(1)
       else:
           match_suelto = re.search(r'([A-Z]\.[A-Z]{1,4}\.\d{3})', texto_comprimido_seguro)
           if match_suelto: 
               codigo_detectado = match_suelto.group(1)

   # === 2. EXTRACCIÓN DE TÍTULO Y NATURALEZA ===
   patron_titulo = r'(?i)t[íi]tulo\s*[:\s|]\s*([\s\S]+?)(?=\s*(?:naturaleza|c[oó]digo|versi[oó]n|fecha|área|proceso|colegio)\s*[:|]|$)'
   match_tit = re.search(patron_titulo, texto_completo)
   titulo_detectado = "NO DETECTADO"
   if match_tit:
       raw_tit = match_tit.group(1).strip()
       raw_tit = re.split(r'(?i)\s*(?:[:|]\s*)?(?:naturaleza|c[oó]digo|versi[oó]n|fecha|área|proceso|colegio)', raw_tit)[0].strip()
       titulo_detectado = re.sub(r'\s*\|\s*', ' ', raw_tit).strip().strip('|').strip()

   patron_naturaleza = r'(?i)naturaleza(?:[^|\n:]*)[:\s|]\s*([\s\S]+?)(?=\s*(?:t[íi]tulo|c[oó]digo|versi[oó]n|fecha)\s*[:|]|$)'
   match_nat = re.search(patron_naturaleza, texto_completo)
   naturaleza_detectada = "NO DETECTADA"
   if match_nat:
       raw_nat = match_nat.group(1).strip()
       raw_nat = re.split(r'(?i)\s*(?:[:|]\s*)?(?:t[íi]tulo|c[oó]digo|versi[oó]n|fecha)', raw_nat)[0].strip()
       raw_nat = re.sub(r'\s*\|\s*', ' ', raw_nat).strip().strip('|').strip()
       if raw_nat and not raw_nat.upper().startswith("TITULO"):
           naturaleza_detectada = raw_nat

   # === 3. EDICIÓN / VERSIÓN ===
   patron_version = r'(?i)(?:versi[oó]n|revisi[oó]n|rev)\s*[:\s|]\s*(\d+)'
   match_ver = re.search(patron_version, texto_completo)
   version_detectada = "NO NOTADA"
   if match_ver:
       version_detectada = match_ver.group(1)
       if len(version_detectada) == 1: 
           version_detectada = f"0{version_detectada}"
   else:
       texto_sin_fechas = re.sub(r'\d{1,2}/\d{1,2}/\d{2,4}', '', texto_comprimido_seguro)
       match_ver_comp = re.search(r'(?:VERSION|REVISION|REV):?(\d+)', texto_sin_fechas)
       if match_ver_comp:
           version_detectada = match_ver_comp.group(1)
           if len(version_detectada) == 1: version_detectada = f"0{version_detectada}"

   # === 4. EXTRACCIÓN DE FECHAS ===
   fechas_encontradas = re.findall(r'\b\d{1,2}/\d{1,2}/\d{2,4}\b', texto_completo)
   fecha_ini_detectada = "NO DETECTADA"
   fecha_rev_detectada = "NO DETECTADA"
  
   if len(fechas_encontradas) >= 1:
       def clave_cronologica(f_str):
           try:
               partes = f_str.split('/')
               dia = int(partes[0])
               mes = int(partes[1])
               anio = int(partes[2])
               if anio < 100: anio += 2000
               return (anio, mes, dia)
           except:
               return (0, 0, 0)
       
       fechas_ordenadas = sorted(list(set(fechas_encontradas)), key=clave_cronologica)
       if len(fechas_ordenadas) >= 1:
           fecha_ini_detectada = fechas_ordenadas[0]
       if len(fechas_ordenadas) >= 2:
           fecha_rev_detectada = fechas_ordenadas[-1]

   return titulo_detectado, codigo_detectado, version_detectada, fecha_ini_detectada, fecha_rev_detectada, naturaleza_detectada

# ------------------------------------------------------------
#  MOTOR HÍBRIDO AVANZADO (CELDAS + CAPAS FLOTANTES)
# ------------------------------------------------------------
def extraer_todo_excel(ruta):
   """
   Estrategia Maestra: Aplana el archivo convirtiéndolo temporalmente a PDF
   para capturar formas y textos flotantes, y además extrae las celdas crudas.
   """
   acumulado_total = []
   ext = os.path.splitext(ruta)[1].lower()
   nombre_base = os.path.splitext(os.path.basename(ruta))[0]
   carpeta_destino = "/tmp"

   # --- CAPA A: CAPTURAR TEXTO FLOTANTE / DIBUJOS (Conversión PDF) ---
   if pypdf is not None:
       try:
           # Convertimos a PDF sin importar si es .xls o .xlsx para renderizar las capas gráficas
           comando_pdf = ['libreoffice', '--headless', '--convert-to', 'pdf', ruta, '--outdir', carpeta_destino]
           subprocess.run(comando_pdf, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, check=True)
           
           ruta_pdf_temporal = os.path.join(carpeta_destino, nombre_base + '.pdf')
           if os.path.exists(ruta_pdf_temporal):
               reader = pypdf.PdfReader(ruta_pdf_temporal)
               # Evaluamos las primeras 3 hojas renderizadas
               for pagina in reader.pages[:3]:
                   txt_pdf = pagina.extract_text()
                   if txt_pdf:
                       acumulado_total.append(txt_pdf)
               os.remove(ruta_pdf_temporal)
       except:
           pass # Si falla el motor PDF por dependencias, continúa al flujo de celdas

   # --- CAPA B: CAPTURAR CELDAS TRADICIONALES (openpyxl) ---
   if openpyxl is not None:
       ruta_xlsx = ruta
       archivo_temporal_xlsx = False
       
       if ext == '.xls':
           try:
               comando_xlsx = ['libreoffice', '--headless', '--convert-to', 'xlsx', ruta, '--outdir', carpeta_destino]
               subprocess.run(comando_xlsx, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, check=True)
               ruta_xlsx = os.path.join(carpeta_destino, nombre_base + '.xlsx')
               archivo_temporal_xlsx = True
           except:
               pass

       if os.path.exists(ruta_xlsx):
           try:
               wb = openpyxl.load_workbook(ruta_xlsx, data_only=True)
               for hoja in wb.worksheets:
                   for fila in hoja.iter_rows(values_only=True):
                       for celda in fila:
                           if celda is not None:
                               txt_celda = str(celda).strip()
                               if txt_celda: acumulado_total.append(txt_celda)
               if archivo_temporal_xlsx: os.remove(ruta_xlsx)
           except:
               if archivo_temporal_xlsx and os.path.exists(ruta_xlsx): os.remove(ruta_xlsx)

   return " | ".join(acumulado_total)

# ------------------------------------------------------------
#  CONTROLADOR DE LA INTERFAZ
# ------------------------------------------------------------
def abrir_y_analizar():
   ruta_archivo = filedialog.askopenfilename(
       title="Cargar Reporte Excel para Auditoría",
       filetypes=[("Hojas de Cálculo", "*.xlsx *.xls")]
   )
   if not ruta_archivo: return

   nombre_fichero = os.path.basename(ruta_archivo)
   texto_crudo = extraer_todo_excel(ruta_archivo)

   titulo, codigo, version, fecha_ini, fecha_rev, naturaleza = analizar_y_aislar_metadatos_excel(texto_crudo, nombre_fichero)
  
   lbl_archivo.config(text=f"Fichero: {nombre_fichero}", fg="#00f5d4")
   txt_visor.delete("1.0", tk.END)
  
   txt_visor.insert(tk.END, f"📊 AGAPE QA - EXTRACTOR INTEGRAL DE EXCEL (v2.5.0-hybrid)\n")
   txt_visor.insert(tk.END, f"===========================================================\n")
   txt_visor.insert(tk.END, f"📁 EXCEL EVALUADO: {nombre_fichero}\n\n")
  
   txt_visor.insert(tk.END, f"📌 [DATOS AISLADOS DE MATRIZ Y CAPAS FLOTANTES]:\n")
   txt_visor.insert(tk.END, f"   ➔ TÍTULO INTERNO:      {titulo}\n")
   txt_visor.insert(tk.END, f"   ➔ NATURALEZA DOC:      {naturaleza}\n")
   txt_visor.insert(tk.END, f"   ➔ CÓDIGO DEL FORMATO:  {codigo}\n")
   txt_visor.insert(tk.END, f"   ➔ EDICIÓN / REVISIÓN:  {version}\n")
   txt_visor.insert(tk.END, f"   ➔ FECHA INICIAL:       {fecha_ini}\n")
   txt_visor.insert(tk.END, f"   ➔ FECHA DE REVISIÓN:   {fecha_rev}\n")
   txt_visor.insert(tk.END, f"===========================================================\n\n")
  
   txt_visor.insert(tk.END, f"[📋 FLUJO TEXTUAL COMBINADO (CELDAS + FORMATOS GRÁFICOS)]:\n")
   txt_visor.insert(tk.END, texto_crudo if texto_crudo.strip() else "(Sin texto extraíble)")

# ------------------------------------------------------------
#  GUI SETUP (CORREGIDO ERROR DE GEOMETRÍA)
# ------------------------------------------------------------
ventana = tk.Tk()
ventana.title("AGAPE QA - Extractor de Hojas de Cálculo (v2.5.0)")
ventana.geometry("950x720")  # <<--- ¡CORREGIDO AQUÍ! Ya no tiene los paréntesis "950(x720)"

lbl_titulo = tk.Label(ventana, text="Fase 2: Auditor Especializado de Excel (.xlsx / .xls)", font=("Arial", 12, "bold"))
lbl_titulo.pack(pady=10)

btn_subir = tk.Button(
   ventana,
   text="📊 Cargar Libro de Excel (.xlsx / .xls)",
   font=("Arial", 11, "bold"),
   command=abrir_y_analizar,
   bg="#9b5de5",
   fg="white",
   padx=20,
   pady=6
)
btn_subir.pack(pady=5)

lbl_archivo = tk.Label(ventana, text="Ningún libro seleccionado", font=("Arial", 10, "italic"), fg="gray")
lbl_archivo.pack(pady=5)

# Mantenemos el fondo oscuro con la letra verde fosforescente clásica de terminal (#39FF14)
txt_visor = scrolledtext.ScrolledText(ventana, wrap=tk.WORD, font=("Courier New", 10), bg="#0f0f1c", fg="#39FF14")
txt_visor.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

ventana.mainloop()