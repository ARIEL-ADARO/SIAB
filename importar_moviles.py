import pandas as pd
import mysql.connector
import os

def importar_datos():
    # 1. Configuración de conexión (ASEGURATE DE QUE SEAN TUS DATOS)
    config = {
        'user': 'root',
        'password': 'siab1234', 
        'host': 'localhost',
        'database': 'SIAB' 
    }

    archivo_nombre = 'LISTADO_DE_MOVILES.xlsx'

    if not os.path.exists(archivo_nombre):
        print(f"ERROR: No se encuentra el archivo '{archivo_nombre}' en {os.getcwd()}")
        return

    try:
        conn = mysql.connector.connect(**config)
        cursor = conn.cursor()

        # 2. Leer EXCEL directamente
        print(f"Leyendo Excel: {archivo_nombre}...")
        df = pd.read_excel(archivo_nombre)
        df.columns = df.columns.str.strip() 

        print(f"Iniciando carga de {len(df)} filas...")

        for index, row in df.iterrows():
            nro_raw = str(row['N° MOVIL']).strip()
            
            # Saltamos si es S/N o está vacío
            if nro_raw.upper() == 'S/N' or pd.isna(row['N° MOVIL']):
                continue

            try:
                # A. Insertar en unidades_fisicas
                sql_u = """INSERT INTO unidades_fisicas 
                           (marca, anio_fabricacion, capacidad_valor, estado_mecanico) 
                           VALUES (%s, %s, %s, %s)"""
                cursor.execute(sql_u, (
                    row['Marca/Modelo'], 
                    row['Año'], 
                    row['Capacidad'], 
                    'OPERATIVO'
                ))
                id_unidad = cursor.lastrowid

                # B. Insertar en moviles (manejando el espacio en 'MÓVILES ')
                col_tipo = 'MÓVILES' if 'MÓVILES' in df.columns else 'MÓVILES '
                
                sql_m = """INSERT INTO moviles 
                           (nro_movil, id_unidad_actual, tipo_uso, estado_operativo) 
                           VALUES (%s, %s, %s, %s)"""
                cursor.execute(sql_m, (
                    nro_raw, 
                    id_unidad, 
                    row[col_tipo], 
                    'EN SERVICIO'
                ))
                
                print(f"OK -> Móvil {nro_raw}: {row['Marca/Modelo']}")

            except Exception as e_row:
                print(f"Error en fila {index}: {e_row}")

        conn.commit()
        print("\n>>> CARGA DESDE EXCEL FINALIZADA EXITOSAMENTE <<<")

    except Exception as e:
        print(f"Error: {e}")
    finally:
        if 'conn' in locals() and conn.is_connected():
            cursor.close()
            conn.close()

if __name__ == "__main__":
    importar_datos()