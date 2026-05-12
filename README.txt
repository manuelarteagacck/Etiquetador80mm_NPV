Etiquetador80mm_NPV
---------------------
Aplicación de escritorio Python 3.11+ para imprimir etiquetas térmicas 80mm
usando python-escpos y conexión SQL Server.

Instrucciones:
1. Instalar dependencias: pip install -r requirements.txt
2. Configurar conexión en config.json
3. Ejecutar: python app.py
4. (Opcional) Empaquetar: pyinstaller -F --add-data "back3.jpg;." -n Etiquetador80mm app.py

Si aparece "ModuleNotFoundError: No module named 'win32print'", falta pywin32 en
el Python que esta ejecutando la app. Ejecuta:

python -m pip install pywin32

O reinstala todas las dependencias:

python -m pip install -r requirements.txt
