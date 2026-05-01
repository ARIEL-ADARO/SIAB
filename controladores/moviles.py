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
        # Seleccionamos TODO (*) incluyendo ID y los nuevos campos técnicos
        # Agregamos ORDER BY para que el historial se vea ordenado por número
        query = """
            SELECT * FROM moviles 
            WHERE estado != 'BAJA' 
            ORDER BY nro_unidad ASC
        """
        cursor.execute(query)
        unidades = cursor.fetchall()
        return unidades
    finally:
        db.close()