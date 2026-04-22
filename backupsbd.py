import os
import datetime

def hacer_backup():
    # Configuración
    user = "root"
    password = "siab1234"
    db_name = "SIAB"
    
    # Ruta específica donde quieres guardarlo
    folder = r"C:\SIAB\backups"
    
    # Crear la carpeta si no existe
    if not os.path.exists(folder):
        os.makedirs(folder)
    
    # Nombre del archivo con fecha y hora
    fecha = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M")
    filename = f"backup_siab_{fecha}.sql"
    filepath = os.path.join(folder, filename)
    
    # Comando de MySQL
    comando = f"mysqldump -u {user} -p{password} {db_name} > {filepath}"
    
    os.system(comando)
    print(f"Respaldo creado con éxito en: {filepath}")

hacer_backup()