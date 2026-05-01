import sqlite3
import mysql.connector
from mysql.connector import Error
from datetime import datetime

DB_CONFIG_MYSQL = {
    "host": "localhost",
    "user": "root",
    "password": "siab1234",
    "database": "siab"
}

def limpiar_valor(valor, nombre_columna):
    # --- 1. TRANSFORMAR FECHAS (El cambio clave) ---
    # Identificamos columnas que contienen fechas según tu .schema
    columnas_fecha = [
        'fecha_inicio', 'fecha_fin', 'fecha_carga', 
        'firma_bombero_fecha', 'firma_supervisor_fecha', 
        'fecha_anulacion', 'fecha_creacion', 'fecha_modificacion'
    ]
    
    if nombre_columna in columnas_fecha and valor:
        try:
            # Si el valor viene como "DD/MM/YYYY" desde SQLite
            # Lo convertimos a objeto datetime y luego a "YYYY-MM-DD"
            fecha_obj = datetime.strptime(str(valor).strip(), '%d/%m/%Y')
            return fecha_obj.strftime('%Y-%m-%d')
        except ValueError:
            # Si la fecha ya estaba en formato correcto o tiene otro error, 
            # intentamos dejarla pasar o devolver None si está mal
            return valor

    # --- 2. TRADUCIR BOOLEANOS ---
    if nombre_columna in ['autoriza', 'es_encargado', 'anulada']:
        if isinstance(valor, str):
            v = valor.strip().upper()
            if v == 'SI': return 1
            if v == 'NO': return 0
    
    # --- 3. LIMPIEZA DE CELULARES ---
    if nombre_columna == 'nro_cel':
        if isinstance(valor, float) or (isinstance(valor, str) and valor.endswith('.0')):
            try: return str(int(float(valor)))
            except: return valor
            
    # Si es cadena vacía o solo espacios, mandamos NULL
    if isinstance(valor, str) and not valor.strip():
        return None
        
    return valor

def migrar():
    sqlite_conn = None
    mysql_conn = None
    try:
        sqlite_conn = sqlite3.connect('siab.db')
        sqlite_cursor = sqlite_conn.cursor()

        mysql_conn = mysql.connector.connect(**DB_CONFIG_MYSQL)
        mysql_cursor = mysql_conn.cursor()

        mysql_cursor.execute("SET FOREIGN_KEY_CHECKS = 0;")
        print("Restricciones desactivadas...")

        tablas = ['conceptos', 'usuarios', 'legajos', 'actividades', 'actividades_historial', 'notificaciones']

        for tabla in tablas:
            try:
                sqlite_cursor.execute(f"SELECT name FROM sqlite_master WHERE type='table' AND name='{tabla}'")
                if not sqlite_cursor.fetchone():
                    continue

                print(f"Migrando tabla: {tabla}...")
                sqlite_cursor.execute(f"SELECT * FROM {tabla}")
                filas = sqlite_cursor.fetchall()
                columnas = [desc[0] for desc in sqlite_cursor.description]

                if not filas:
                    print(f"  -> '{tabla}' está vacía.")
                    continue

                nombres_cols = ", ".join([f"`{c}`" for c in columnas])
                placeholders = ", ".join(["%s"] * len(columnas))
                query_insert = f"INSERT INTO `{tabla}` ({nombres_cols}) VALUES ({placeholders})"

                datos_para_mysql = []
                for fila in filas:
                    # Aquí aplicamos la nueva lógica de limpieza de fechas
                    fila_limpia = tuple(limpiar_valor(fila[i], columnas[i]) for i in range(len(columnas)))
                    datos_para_mysql.append(fila_limpia)

                mysql_cursor.execute(f"DELETE FROM `{tabla}`")
                mysql_cursor.executemany(query_insert, datos_para_mysql)
                mysql_conn.commit()
                print(f"  -> OK: {len(filas)} registros movidos y formateados.")
            
            except Exception as e:
                print(f"  -> ERROR en tabla '{tabla}': {e}")

        mysql_cursor.execute("SET FOREIGN_KEY_CHECKS = 1;")
        print("\n=== MIGRACIÓN FINALIZADA CON ÉXITO ===")

    except Exception as e:
        print(f"Error crítico: {e}")
    finally:
        if sqlite_conn: sqlite_conn.close()
        if mysql_conn: mysql_conn.close()

if __name__ == "__main__":
    migrar()