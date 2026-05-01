import sys
import os

# Esto le dice a Python que mire también en la carpeta de arriba (C:\SIAB)
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from database import get_db 

def obtener_estado_unidades():
    db = get_db()
    if db is None:
        return []
        
    try:
        cursor = db.cursor(dictionary=True)
        # Agregamos nombre_homenaje y tipo para que coincida con tu nueva tabla
        query = """
            SELECT 
                nro_unidad, 
                nombre_homenaje, 
                tipo, 
                modelo, 
                fecha_vtv, 
                estado 
            FROM moviles 
            WHERE estado != 'BAJA'
        """
        cursor.execute(query)
        unidades = cursor.fetchall()
        return unidades
    finally:
        db.close()