Etiquetador80mm_NPV
---------------------
Aplicación de escritorio Python 3.11+ para imprimir etiquetas térmicas 80mm
usando python-escpos y conexión SQL Server.

Instrucciones:
1. Instalar dependencias: pip install -r requirements.txt
2. Configurar conexión en config.json. Si la contraseña se captura como
   "password" en texto plano, la app la migra automáticamente a
   "password_encrypted" con Fernet y crea config.key junto al config.
3. Ejecutar: python app.py
4. (Opcional) Empaquetar la interfaz. Es obligatorio incluir los datos de
   python-escpos (en particular escpos/capabilities.json):

   pyinstaller -F -w --collect-data escpos --add-data "back5.png;." --add-data "loguito.png;." --add-data "ticket_printer.ico;." --icon ticket_printer.ico -n Etiquetador80mm app.py

Importante: conserva config.key junto con config.json. Sin esa llave no se puede
descifrar password_encrypted.

Si aparece "ModuleNotFoundError: No module named 'win32print'", falta pywin32 en
el Python que esta ejecutando la app. Ejecuta:

python -m pip install pywin32

O reinstala todas las dependencias:

python -m pip install -r requirements.txt

Instalación final
-----------------
El instalador autocontenido se entrega como:

D:\Etiquetador80mm\Instalador_Etiquetador80mm.exe

La instalación conserva config.json, config.key y paramconf.json existentes,
actualiza la aplicación y el video, crea un acceso principal en el menú Inicio
común y un grupo parcial de Inicio, aplica el anclaje de barra de tareas para
todos los usuarios de Windows 10 y registra la tarea "NPV Backup Precios Venta"
diariamente a las 21:00 como SYSTEM. El usuario que instala recibe además un
intento de anclaje inmediato; los demás perfiles lo reciben al iniciar sesión.

Parámetros de línea de comandos del instalador:

  /passive o --passive   Instala mostrando solamente el progreso, sin preguntas.
  /pasive o --pasive     Alias compatible con la grafía solicitada.
  /silent o --silent     Instala completamente en segundo plano.
  /uninstall             Desinstala y conserva los archivos de configuración.
