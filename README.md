# Formulario de Entrada de Datos en Excel

Este proyecto es una aplicación de escritorio desarrollada en Python utilizando `tkinter` para crear una interfaz gráfica y `openpyxl` para manejar archivos Excel. Permite a los usuarios ingresar datos en un formulario y guardar esos datos en un archivo Excel. También ofrece la funcionalidad de abrir el archivo Excel desde la aplicación y sumar los valores en la columna de "Transacción".

## Requisitos

- Python 3.x
- `openpyxl`
- `Pillow`
- `tkinter` (generalmente incluido con la instalación estándar de Python)

## Instalación

1. **Clona el repositorio**:
    ```bash
    git clone https://github.com/tu_usuario/tu_repositorio.git
    cd tu_repositorio
    ```

2. **Instala las dependencias**:
    Ejecuta el siguiente comando para instalar `openpyxl`:
    ```bash
    pip install openpyxl
    ```
    Ejecuta el siguiente comando para instalar `Pillow`:
    ```bash
    pip install Pillow
    ```
## Uso

1. **Ejecuta la aplicación**:
    Asegúrate de estar en el directorio del proyecto y ejecuta el archivo principal con Python:
    ```bash
    python main.py
    ```

2. **Rellena el formulario**:
    Completa todos los campos del formulario y haz clic en "Guardar" para almacenar los datos en el archivo Excel `datos.xlsx`.

3. **Abrir el archivo Excel**:
    Si deseas abrir el archivo Excel directamente desde la aplicación, haz clic en "Abrir Archivo". Si el archivo aún no ha sido creado, recibirás un mensaje de advertencia.

## Descripción de Funcionalidades

- **Formulario de Entrada de Datos**:
  - Campos: Nombre, Edad, Email, Teléfono, Dirección, Transacción.
  - Validación de entradas para asegurar que todos los campos estén llenos y que los datos sean del tipo adecuado.
  - Verificación básica del formato de email.
  - Limpieza de campos del formulario después de guardar los datos.

- **Suma de Transacciones**:
  - Los valores en la columna de "Transacción" se suman y el total se guarda en la columna 8 del archivo Excel.

- **Compatibilidad Multisistema**:
  - La aplicación es compatible con Windows, macOS y Linux. Se utilizan diferentes métodos para abrir el archivo Excel según el sistema operativo.

## Notas

- El archivo Excel se guarda en el mismo directorio donde se ejecuta el script. Si el archivo no existe, se crea automáticamente con las cabeceras correspondientes.
- La ventana de la aplicación es fija en tamaño y no puede ser redimensionada por el usuario.

## Contribuciones

Las contribuciones son bienvenidas. Si encuentras algún error o tienes alguna sugerencia de mejora, por favor, abre un issue o envía una solicitud de extracción (pull request).

## Licencia

Este proyecto está bajo la Licencia MIT. Consulta el archivo `LICENSE` para más detalles.