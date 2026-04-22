import sqlite3
import mysql.connector
from mysql.connector import Error

DB_CONFIG_MYSQL = {
    "host": "localhost",
    "user": "root",
    "password": "siab1234",
    "database": "siab"
}

def limpiar_valor(valor, nombre_columna):
    # Traducir "SI/NO" a "1/0" para columnas numéricas como 'autoriza' o 'es_encargado'
    if nombre_columna in ['autoriza', 'es_encargado']:
        if isinstance(valor, str):
            v = valor.strip().upper()
            if v == 'SI': return 1
            if v == 'NO': return 0
    
    # Quitar el .0 de los celulares que vienen de Excel
    if nombre_columna == 'nro_cel':
        if isinstance(valor, float) or (isinstance(valor, str) and valor.endswith('.0')):
            try: return str(int(float(valor)))
            except: return valor
            
    # Si el valor es una cadena vacía, mandamos NULL a MySQL
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

        # Tablas a migrar (quité eventos porque ya vimos que no existe en tu .db)
        tablas = ['conceptos', 'usuarios', 'legajos', 'actividades', 'actividades_historial', 'notificaciones']

        for tabla in tablas:
            try:
                # Verificar si existe en SQLite
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

                # Preparar el INSERT
                nombres_cols = ", ".join([f"`{c}`" for c in columnas])
                placeholders = ", ".join(["%s"] * len(columnas))
                query_insert = f"INSERT INTO `{tabla}` ({nombres_cols}) VALUES ({placeholders})"

                # Limpiar y procesar datos fila por fila
                datos_para_mysql = []
                for fila in filas:
                    fila_limpia = tuple(limpiar_valor(fila[i], columnas[i]) for i in range(len(columnas)))
                    datos_para_mysql.append(fila_limpia)

                mysql_cursor.execute(f"DELETE FROM `{tabla}`")
                mysql_cursor.executemany(query_insert, datos_para_mysql)
                mysql_conn.commit()
                print(f"  -> OK: {len(filas)} registros movidos.")
            
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