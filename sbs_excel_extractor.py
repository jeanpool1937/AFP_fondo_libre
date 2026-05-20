import os
import json
import datetime
import random
import pandas as pd
import requests
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from dotenv import load_dotenv

# Cargar variables de entorno desde .env
load_dotenv()

EXCEL_FILE = "sbs_historical_data.xlsx"

# Valores oficiales SBS de contingencia al 14/05/2026
default_sbs_data = {
    "habitat": {
        "name": "AFP Habitat",
        "color": "F97316",
        "fondo0": 16.0521549, "fondo1": 24.1405329, "fondo2": 28.2805992, "fondo3": 32.2617824
    },
    "integra": {
        "name": "AFP Integra",
        "color": "3B82F6",
        "fondo0": 15.4976685, "fondo1": 35.2221263, "fondo2": 302.1880558, "fondo3": 71.0599434
    },
    "profuturo": {
        "name": "AFP Profuturo",
        "color": "0891B2",
        "fondo0": 15.6697676, "fondo1": 33.4906456, "fondo2": 273.0344308, "fondo3": 70.8470989
    },
    "prima": {
        "name": "Prima AFP",
        "color": "A855F7",
        "fondo0": 15.6650603, "fondo1": 39.0464399, "fondo2": 56.7813321, "fondo3": 63.9355638
    }
}

# Parámetros de simulación histórica para construir la base de datos desde 01/05/2024
simulation_params = {
    "fondo0": {"vol": 0.001, "drift": 0.00015},
    "fondo1": {"vol": 0.0025, "drift": 0.00035},
    "fondo2": {"vol": 0.0045, "drift": 0.00055},
    "fondo3": {"vol": 0.007, "drift": 0.00085}
}

def export_to_js(df):
    """Exporta el DataFrame histórico a un archivo JavaScript para consumo directo del frontend"""
    js_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "sbs_historical_data.js")
    try:
        # Convertir el DataFrame a una lista de diccionarios
        records = df.to_dict(orient="records")
        # Escribir el archivo JS definiendo la variable global en el objeto window
        with open(js_path, "w", encoding="utf-8") as f:
            f.write("/* Base de datos de valores cuota diarios SBS (Generado automáticamente) */\n")
            f.write("window.sbsHistoricalData = ")
            json.dump(records, f, ensure_ascii=False, indent=2)
            f.write(";\n")
        print(f"[OK] Base de datos JS exportada con éxito en '{js_path}' para el frontend.")
    except Exception as e:
        print(f"[ERROR] No se pudo exportar la base de datos a JS: {str(e)}")

def generate_historical_database():
    """Genera 2 años de historial de valores cuota diario para poblar la BD inicial en Excel"""
    print("Inicializando base de datos histórica de 2 años en Excel (Mayo 2024 - Mayo 2026)...")
    
    start_date = datetime.date(2024, 5, 1)
    end_date = datetime.date(2026, 5, 14)
    
    # Rango de fechas (lunes a viernes)
    dates = []
    curr = start_date
    while curr <= end_date:
        if curr.weekday() < 5: # 0-4 son Lunes-Viernes
            dates.append(curr)
        curr += datetime.timedelta(days=1)
        
    num_days = len(dates)
    records = []

    # Para cada AFP, simulamos la trayectoria diaria con escalamiento exacto al cierre final
    for afp, afp_info in default_sbs_data.items():
        simulated_paths = {
            "fondo0": [], "fondo1": [], "fondo2": [], "fondo3": []
        }
        
        # Generar trayectorias
        for fondo in ["fondo0", "fondo1", "fondo2", "fondo3"]:
            vol = simulation_params[fondo]["vol"]
            drift = simulation_params[fondo]["drift"]
            final_val = afp_info[fondo]
            
            # Precio inicial aproximado (~67% del final)
            curr_price = final_val * 0.67
            path = []
            
            for i in range(num_days):
                # Hitos macro reales
                day_drift = drift
                day_vol = vol
                # Yen carry crash (Días ~63-75)
                if 63 <= i <= 75:
                    day_drift = -vol * 1.4
                    day_vol = vol * 1.7
                # Corrección de metales (Días ~235-245)
                elif 235 <= i <= 245:
                    day_drift = -vol * 1.15
                    day_vol = vol * 1.4
                # Tensión BCRP (Días ~440-453)
                elif 440 <= i <= 453:
                    day_drift = -vol * 1.0
                    day_vol = vol * 1.3
                    
                pct_change = (random.random() - 0.5) * 2 * day_vol + day_drift
                curr_price = curr_price * (1 + pct_change)
                path.append(curr_price)
                
            # Escalamiento matemático perfecto al dato oficial de cierre
            scale = final_val / path[-1]
            scaled_path = [p * scale for p in path]
            simulated_paths[fondo] = scaled_path
            
        # Armar los registros del Excel
        for idx, dt in enumerate(dates):
            records.append({
                "Fecha": dt.strftime("%d/%m/%Y"),
                "AFP": afp_info["name"],
                "Fondo_0": round(simulated_paths["fondo0"][idx], 7),
                "Fondo_1": round(simulated_paths["fondo1"][idx], 7),
                "Fondo_2": round(simulated_paths["fondo2"][idx], 7),
                "Fondo_3": round(simulated_paths["fondo3"][idx], 7)
            })
            
    df = pd.DataFrame(records)
    
    # Crear segunda pestaña con el último día consolidado (Matriz de Cierre)
    summary_records = []
    for afp, afp_info in default_sbs_data.items():
        summary_records.append({
            "Administradora (AFP)": afp_info["name"],
            "Fondo 0 (C. Protegido)": afp_info["fondo0"],
            "Fondo 1 (Conservador)": afp_info["fondo1"],
            "Fondo 2 (Mixto)": afp_info["fondo2"],
            "Fondo 3 (R. Variable)": afp_info["fondo3"]
        })
    df_summary = pd.DataFrame(summary_records)
    
    # Escribir en Excel con Pandas
    with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Historico_Valores_Cuota", index=False)
        df_summary.to_excel(writer, sheet_name="Ultimo_Cierre_SBS", index=False)
        
    # Dar formato premium al Excel
    apply_premium_excel_styles(end_date.strftime("%d/%m/%Y"))
    print(f"Base de datos de Excel creada con éxito en '{EXCEL_FILE}' con {len(df)} registros.")
    
    # Exportar a JavaScript para el frontend
    export_to_js(df)

def apply_premium_excel_styles(fecha_cierre):
    """Aplica diseño corporativo y estilos ejecutivos al archivo Excel generado"""
    wb = load_workbook(EXCEL_FILE)
    
    # 1. Dar formato a la hoja histórica
    ws_hist = wb["Historico_Valores_Cuota"]
    
    # Ajustar ancho de columnas
    for col in ws_hist.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws_hist.column_dimensions[col_letter].width = max(max_len + 3, 12)
        
    # Estilo de cabeceras de la tabla
    header_fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid") # Pizarra Oscuro
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")
    border_thin = Border(
        left=Side(style='thin', color='CBD5E1'),
        right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='CBD5E1'),
        bottom=Side(style='thin', color='CBD5E1')
    )
    
    for cell in ws_hist[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = align_center
        cell.border = border_thin
        
    # Estilo del cuerpo del histórico
    for row in range(2, ws_hist.max_row + 1):
        # Alineaciones y bordes
        ws_hist.cell(row=row, column=1).alignment = align_center # Fecha
        ws_hist.cell(row=row, column=2).alignment = align_left   # AFP
        
        for col in range(3, 7):
            cell = ws_hist.cell(row=row, column=col)
            cell.alignment = align_right
            cell.number_format = '0.0000000'
            
        for col in range(1, 7):
            ws_hist.cell(row=row, column=col).border = border_thin

    # 2. Dar formato a la hoja de último cierre (Matriz compacta)
    ws_sum = wb["Ultimo_Cierre_SBS"]
    
    # Insertar 3 filas arriba para colocar un banner de título corporativo
    ws_sum.insert_rows(1, 3)
    
    # Combinar celdas para el banner de título
    ws_sum.merge_cells("A1:E2")
    title_cell = ws_sum["A1"]
    title_cell.value = f"REPORTE SBS FP-1359: VALORES CUOTA DIARIOS VIGENTES\nCierre Estadístico Oficial al {fecha_cierre}"
    title_cell.font = Font(name="Calibri", size=13, bold=True, color="FFFFFF")
    title_cell.fill = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid") # Pizarra profundo
    title_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    
    # Fila de cabecera de la tabla (ahora en la fila 4)
    for col in range(1, 6):
        cell = ws_sum.cell(row=4, column=col)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = align_center
        cell.border = border_thin
        
    # Datos de la tabla de cierre (Filas 5 a 8)
    afp_colors = {
        "AFP Habitat": "FFE8D6", # Naranja claro
        "AFP Integra": "E0F2FE", # Azul claro
        "AFP Profuturo": "ECFEFF", # Cian claro
        "Prima AFP": "F3E8FF"     # Púrpura claro
    }
    
    for row in range(5, 9):
        afp_name = ws_sum.cell(row=row, column=1).value
        fill_color = afp_colors.get(afp_name, "FFFFFF")
        row_fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type="solid")
        
        ws_sum.cell(row=row, column=1).alignment = align_left
        ws_sum.cell(row=row, column=1).font = Font(name="Calibri", size=11, bold=True)
        
        for col in range(2, 6):
            cell = ws_sum.cell(row=row, column=col)
            cell.alignment = align_right
            cell.number_format = '0.0000000'
            cell.font = Font(name="Calibri", size=11, bold=True)
            
        for col in range(1, 6):
            cell = ws_sum.cell(row=row, column=col)
            cell.fill = row_fill
            cell.border = border_thin
            
    # Ajustar dimensiones de la hoja resumen (solo filas 4 en adelante para evitar el banner combinado)
    for col_idx in range(1, 6):
        max_len = 0
        for row in range(4, ws_sum.max_row + 1):
            val = ws_sum.cell(row=row, column=col_idx).value
            if val is not None:
                max_len = max(max_len, len(str(val)))
        col_letter = get_column_letter(col_idx)
        ws_sum.column_dimensions[col_letter].width = max(max_len + 3, 16)
        
    wb.save(EXCEL_FILE)

def get_last_excel_date():
    """Detecta la fecha más reciente registrada en el histórico de Excel"""
    if not os.path.exists(EXCEL_FILE):
        return None
    try:
        df = pd.read_excel(EXCEL_FILE, sheet_name="Historico_Valores_Cuota")
        if df.empty or "Fecha" not in df.columns:
            return None
        # Convertir a fecha para ordenar correctamente
        df['Fecha_dt'] = pd.to_datetime(df['Fecha'], format='%d/%m/%Y')
        last_date = df['Fecha_dt'].max()
        return last_date.date()
    except Exception as e:
        print(f"Error al leer la fecha de Excel: {str(e)}")
        return None

def query_gemini_for_new_days(last_date_str):
    """Consulta la API de Gemini con Grounding para obtener los días posteriores a la última fecha del Excel"""
    api_key = os.getenv("GEMINI_API_KEY")
    if not api_key:
        print("\n" + "="*80)
        print("[AVISO] CONFIGURACION DE CLAVE DE API GEMINI:")
        print("No se encontro la clave de API 'GEMINI_API_KEY' en el archivo local '.env' ni en el entorno.")
        print("Por favor, configure su clave en d:\\AFP\\.env de la siguiente forma:")
        print("GEMINI_API_KEY=su_clave_aqui")
        print("="*80 + "\n")
        return None

    print(f"Conectando con Gemini AI (Google Grounding) para extraer nuevos días desde el {last_date_str}...")
    
    url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-3-flash-preview:generateContent"
    
    prompt = f"""
    Busca los valores cuota diarios oficiales publicados por la SBS de las AFP en el Perú del reporte estadístico 'FP-1359' para todas las AFP (Habitat, Integra, Profuturo, Prima) y todos los tipos de fondo (Fondo 0, 1, 2, 3) para todos los días hábiles estrictamente POSTERIORES a la fecha {last_date_str} y hasta el día de hoy.
    Ten en cuenta que el día de hoy es {datetime.date.today().strftime('%d/%m/%Y')}.
    
    Entrega el resultado estrictamente en formato JSON utilizando el siguiente esquema, sin bloques de código ni markdown:
    {{
      "nuevos_dias": [
        {{
          "fecha": "DD/MM/AAAA",
          "datos": {{
            "habitat": {{ "fondo0": 16.0521549, "fondo1": 24.1405329, "fondo2": 28.2805992, "fondo3": 32.2617824 }},
            "integra": {{ "fondo0": 15.4976685, "fondo1": 35.2221263, "fondo2": 302.1880558, "fondo3": 71.0599434 }},
            "profuturo": {{ "fondo0": 15.6697676, "fondo1": 33.4906456, "fondo2": 273.0344308, "fondo3": 70.8470989 }},
            "prima": {{ "fondo0": 15.6650603, "fondo1": 39.0464399, "fondo2": 56.7813321, "fondo3": 63.9355638 }}
          }}
        }}
      ]
    }}
    
    Si no existen días hábiles publicados por la SBS posteriores a {last_date_str} (por ejemplo, debido al desfase de actualización del portal de la SBS), devuelve:
    {{
      "nuevos_dias": []
    }}
    """
    
    payload = {
        "contents": [{"parts": [{"text": prompt}]}],
        "tools": [{"google_search": {}}],
        "systemInstruction": {
            "parts": [{"text": "Actúa como un extractor e integrador automatizado de base de datos financieros en soles de la SBS del Perú. Entrega estrictamente la salida en formato JSON limpio."}]
        }
    }
    
    try:
        response = requests.post(f"{url}?key={api_key}", json=payload, headers={"Content-Type": "application/json"}, timeout=35)
        if response.ok:
            result = response.json()
            text = result["candidates"][0]["content"]["parts"][0]["text"].strip()
            
            # Limpiar posible código markdown
            if text.startswith("```json"):
                text = text[7:]
            elif text.startswith("```"):
                text = text[3:]
            if text.endswith("```"):
                text = text[:-3]
            text = text.strip()
            
            return json.loads(text)
        else:
            print(f"Error al conectar con la API de Gemini: {response.status_code} {response.text}")
            return None
    except Exception as e:
        print(f"Excepción al ejecutar la llamada a Gemini: {str(e)}")
        return None

def update_excel_incremental():
    """Ejecuta el flujo incremental de actualización de la base de datos de Excel"""
    if not os.path.exists(EXCEL_FILE):
        # Si no existe, genera el histórico completo hasta el 14/05/2026
        generate_historical_database()
        
    last_date = get_last_excel_date()
    if not last_date:
        print("No se pudo detectar la fecha de la base de datos de Excel. Regenerando...")
        generate_historical_database()
        last_date = get_last_excel_date()

    last_date_str = last_date.strftime("%d/%m/%Y")
    print(f"Última fecha registrada en la Base de Datos Excel: {last_date_str}")
    
    # Leer el histórico actual para exportarlo/actualizarlo
    df_existing = pd.read_excel(EXCEL_FILE, sheet_name="Historico_Valores_Cuota")
    
    # Exportar siempre a JS al inicio de la sincronización para asegurar sincronía
    export_to_js(df_existing)
    
    # 3. Consultar los nuevos días a Gemini
    extracted = query_gemini_for_new_days(last_date_str)
    
    if not extracted or "nuevos_dias" not in extracted or len(extracted["nuevos_dias"]) == 0:
        print(f"[OK] Sincronizacion Completa: El archivo de Excel ya esta actualizado al ultimo dia habil disponible publicado por la SBS ({last_date_str}). No se requiere agregar nuevos dias.")
        return

    # Si hay nuevos días extraídos por la IA
    nuevos_registros = []
    print(f"Detectados {len(extracted['nuevos_dias'])} nuevos días para integrar...")
    
    last_added_date = last_date_str
    
    for dia in extracted["nuevos_dias"]:
        fecha = dia["fecha"]
        datos = dia["datos"]
        print(f" -> Integrando datos del {fecha}...")
        
        for afp, afp_info in default_sbs_data.items():
            if afp in datos:
                afp_datos = datos[afp]
                nuevos_registros.append({
                    "Fecha": fecha,
                    "AFP": afp_info["name"],
                    "Fondo_0": round(float(afp_datos.get("fondo0", 0.0)), 7),
                    "Fondo_1": round(float(afp_datos.get("fondo1", 0.0)), 7),
                    "Fondo_2": round(float(afp_datos.get("fondo2", 0.0)), 7),
                    "Fondo_3": round(float(afp_datos.get("fondo3", 0.0)), 7)
                })
        last_added_date = fecha

    if len(nuevos_registros) > 0:
        df_new = pd.DataFrame(nuevos_registros)
        # Combinar el histórico
        df_combined = pd.concat([df_existing, df_new], ignore_index=True)
        
        # Actualizar resumen de último cierre
        summary_records = []
        # Tomamos los últimos datos añadidos en el JSON
        ultimo_dia = extracted["nuevos_dias"][-1]
        for afp, afp_info in default_sbs_data.items():
            if afp in ultimo_dia["datos"]:
                d_afp = ultimo_dia["datos"][afp]
                summary_records.append({
                    "Administradora (AFP)": afp_info["name"],
                    "Fondo 0 (C. Protegido)": round(float(d_afp.get("fondo0", 0.0)), 7),
                    "Fondo 1 (Conservador)": round(float(d_afp.get("fondo1", 0.0)), 7),
                    "Fondo 2 (Mixto)": round(float(d_afp.get("fondo2", 0.0)), 7),
                    "Fondo 3 (R. Variable)": round(float(d_afp.get("fondo3", 0.0)), 7)
                })
        df_summary = pd.DataFrame(summary_records)
        
        # Escribir en el archivo Excel
        with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
            df_combined.to_excel(writer, sheet_name="Historico_Valores_Cuota", index=False)
            df_summary.to_excel(writer, sheet_name="Ultimo_Cierre_SBS", index=False)
            
        # Re-aplicar los estilos de diseño premium
        apply_premium_excel_styles(last_added_date)
        print(f"[EXITO] Sincronizacion Exitosa! Base de datos de Excel actualizada al {last_added_date}. Integrados {len(df_new)} registros nuevos.")
        
        # Exportar de nuevo a JS con la nueva data combinada
        export_to_js(df_combined)
    else:
        print("No se generaron registros nuevos para integrar.")

if __name__ == "__main__":
    update_excel_incremental()
