#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
=============================================================================
SIM-CBC - Sincronizador Automático de Google Sheets a sigeca_db.json
=============================================================================
Este script se ejecuta en GitHub Actions para:
1. Conectarse de forma privada y segura a Google Sheets mediante Service Account.
2. Leer las hojas/pestañas de datos.
3. Actualizar sigeca_db.json preservando la integridad del sistema.
4. Finalizar con código 0 para que GitHub Actions haga commit & push si hay cambios.
"""

import os
import sys
import json
import logging
from datetime import datetime

# Configuración de logs
logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")

try:
    from google.oauth2 import service_account
    from googleapiclient.discovery import build
except ImportError:
    logging.error("Faltan dependencias de Google API. Ejecute: pip install google-api-python-client google-auth")
    sys.exit(1)

DB_PATH = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "sigeca_db.json")

def get_google_sheets_service():
    """Obtiene el cliente autenticado de Google Sheets usando las credenciales del secret."""
    sa_key_raw = os.getenv("GCP_SA_KEY")
    if not sa_key_raw:
        logging.error("No se encontró el secreto 'GCP_SA_KEY' en las variables de entorno.")
        sys.exit(1)

    try:
        # Permite formato JSON directo o ruta a archivo
        if sa_key_raw.strip().startswith("{"):
            key_dict = json.loads(sa_key_raw)
            credentials = service_account.Credentials.from_service_account_info(
                key_dict,
                scopes=["https://www.googleapis.com/auth/spreadsheets.readonly"]
            )
        elif os.path.exists(sa_key_raw):
            credentials = service_account.Credentials.from_service_account_file(
                sa_key_raw,
                scopes=["https://www.googleapis.com/auth/spreadsheets.readonly"]
            )
        else:
            # En caso esté codificado en base64
            import base64
            decoded = base64.b64decode(sa_key_raw).decode("utf-8")
            key_dict = json.loads(decoded)
            credentials = service_account.Credentials.from_service_account_info(
                key_dict,
                scopes=["https://www.googleapis.com/auth/spreadsheets.readonly"]
            )

        service = build("sheets", "v4", credentials=credentials)
        return service
    except Exception as e:
        logging.error(f"Error al inicializar credenciales de Google Service Account: {e}")
        sys.exit(1)

def fetch_sheet_values(service, spreadsheet_id, sheet_name):
    """Descarga todas las filas y columnas de una pestaña específica."""
    try:
        result = service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=sheet_name
        ).execute()
        return result.get("values", [])
    except Exception as e:
        logging.warning(f"No se pudo leer la pestaña '{sheet_name}': {e}")
        return []

def sheet_rows_to_dicts(values):
    """Convierte filas con cabecera en lista de diccionarios."""
    if not values or len(values) < 2:
        return []
    headers = [str(h).strip() for h in values[0]]
    rows = []
    for r in values[1:]:
        row_dict = {}
        for idx, h in enumerate(headers):
            val = r[idx] if idx < len(r) else ""
            row_dict[h] = val
        rows.append(row_dict)
    return rows

def sync_data():
    spreadsheet_id = os.getenv("SPREADSHEET_ID")
    if not spreadsheet_id:
        logging.error("No se encontró la variable 'SPREADSHEET_ID'.")
        sys.exit(1)

    # 1. Cargar base de datos actual
    if not os.path.exists(DB_PATH):
        logging.error(f"No se encontró el archivo de base de datos en: {DB_PATH}")
        sys.exit(1)

    with open(DB_PATH, "r", encoding="utf-8") as f:
        db_data = json.load(f)

    # 2. Conectar a Google Sheets API
    service = get_google_sheets_service()

    # 3. Obtener metadatos de las hojas disponibles en la hoja de cálculo
    sheet_metadata = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    sheets_info = sheet_metadata.get("sheets", [])
    sheet_names = [s["properties"]["title"] for s in sheets_info]
    logging.info(f"Pestañas detectadas en la hoja de cálculo: {sheet_names}")

    updated_sections = []

    # 4. Sincronización inteligente por pestaña
    for name in sheet_names:
        name_upper = name.upper().strip()
        raw_values = fetch_sheet_values(service, spreadsheet_id, name)
        if not raw_values:
            continue

        rows_dict = sheet_rows_to_dicts(raw_values)
        if not rows_dict:
            continue

        # A) Pestaña de Evidencias
        if "EVIDENCIA" in name_upper:
            logging.info(f"Actualizando sección 'evidencias' desde '{name}' ({len(rows_dict)} registros)...")
            db_data["evidencias"] = rows_dict
            updated_sections.append("evidencias")

        # B) Pestaña de Medios de Verificación
        elif "MEDIO" in name_upper or "MV" in name_upper:
            logging.info(f"Actualizando sección 'medios_verificacion' desde '{name}' ({len(rows_dict)} registros)...")
            db_data["medios_verificacion"] = rows_dict
            updated_sections.append("medios_verificacion")

        # C) Pestaña de Programas Académicos
        elif "PROGRAMA" in name_upper:
            logging.info(f"Actualizando sección 'programas' desde '{name}' ({len(rows_dict)} registros)...")
            db_data["programas"] = rows_dict
            updated_sections.append("programas")

        # D) Estadísticas: Docentes
        elif "DOCENTE" in name_upper and "RENACYT" not in name_upper:
            if "estadisticas_institucionales" not in db_data:
                db_data["estadisticas_institucionales"] = {}
            logging.info(f"Actualizando 'estadisticas_institucionales.docentes' desde '{name}'...")
            db_data["estadisticas_institucionales"]["docentes"] = rows_dict
            updated_sections.append("docentes")

        # E) Estadísticas: Investigadores RENACYT
        elif "RENACYT" in name_upper:
            if "estadisticas_institucionales" not in db_data:
                db_data["estadisticas_institucionales"] = {}
            logging.info(f"Actualizando 'estadisticas_institucionales.docentes_renacyt' desde '{name}'...")
            if "docentes_renacyt" not in db_data["estadisticas_institucionales"] or not db_data["estadisticas_institucionales"]["docentes_renacyt"]:
                db_data["estadisticas_institucionales"]["docentes_renacyt"] = [{"rows": rows_dict}]
            else:
                db_data["estadisticas_institucionales"]["docentes_renacyt"][0]["rows"] = rows_dict
            updated_sections.append("docentes_renacyt")

        # F) Estadísticas: Deserción
        elif "DESERCION" in name_upper or "DESERCIÓN" in name_upper:
            if "estadisticas_institucionales" not in db_data:
                db_data["estadisticas_institucionales"] = {}
            logging.info(f"Actualizando 'estadisticas_institucionales.desercion_estudiantil' desde '{name}'...")
            db_data["estadisticas_institucionales"]["desercion_estudiantil"] = rows_dict
            updated_sections.append("desercion_estudiantil")

    # Si hubo una pestaña maestra tipo DB_DUMP o JSON
    if "DB_JSON" in [s.upper() for s in sheet_names]:
        raw_json_vals = fetch_sheet_values(service, spreadsheet_id, "DB_JSON")
        if raw_json_vals and len(raw_json_vals) > 0:
            try:
                full_json_str = "".join([row[0] for row in raw_json_vals if len(row) > 0])
                custom_db = json.loads(full_json_str)
                db_data.update(custom_db)
                updated_sections.append("DB_JSON_FULL")
                logging.info("Sincronizado dump completo desde pestaña 'DB_JSON'.")
            except Exception as ex:
                logging.warning(f"No se pudo procesar la pestaña DB_JSON como dump JSON: {ex}")

    # 5. Guardar la base de datos actualizada
    db_data["ultima_sincronizacion"] = datetime.utcnow().isoformat() + "Z"
    
    with open(DB_PATH, "w", encoding="utf-8") as f:
        json.dump(db_data, f, ensure_ascii=False, indent=2)

    logging.info(f"Sincronización completada exitosamente. Secciones procesadas: {updated_sections}")

if __name__ == "__main__":
    sync_data()
