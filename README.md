# Generador de Solicitudes de Baja

Este proyecto automatiza la generación de solicitudes de baja para activos, utilizando Python para procesar datos desde un archivo Excel y generar documentos formateados según una plantilla predefinida.
Tener en cuenta que este proyecto está diseñado para funcionar con una plantilla específica y puede requerir ajustes según las necesidades de tu organización.
## Descripción

Este proyecto utiliza Python y la biblioteca `openpyxl` para procesar datos de un archivo Excel y generar solicitudes de baja para activos. La plantilla base para las solicitudes se encuentra en el archivo `F&A-P-01 ANEXO N° 02 SOLICITUD DE BAJA.xlsx`.
## Funcionalidades

- Procesamiento de datos desde un archivo Excel
- Generación de solicitudes de baja para cada activo
- Inserción de imágenes de los activos


## Características

- Generación automática de solicitudes de baja a partir de datos en Excel
- Procesamiento de múltiples solicitudes en un solo archivo
- Soporte para inserción de imágenes de los activos
- Manejo de celdas combinadas y formatos especiales
- Generación de documentos con formato estandarizado

## Requisitos

```
openpyxl==3.1.2
Pillow==11.1.0
```

## Estructura del Proyecto

```
├── app.py                 # Script principal
├── Datos.xlsx             # Archivo con datos de los activos
├── F&A-P-01 ANEXO N° 02 SOLICITUD DE BAJA.xlsx  # Plantilla base
├── imagenes/              # Directorio de imágenes de activos
├── requirements.txt       # Dependencias del proyecto
├── solicitudes_generadas/ # Carpeta donde se guardan las solicitudes
└── venv/                  # Entorno virtual de Python
```

## Uso

1. Asegúrese de tener Python instalado y un entorno virtual activado
2. Instale las dependencias:
   ```
   pip install -r requirements.txt
   ```
3. Prepare el archivo `Datos.xlsx` con la información de los activos
4. Ejecute el script:
   ```
   python app.py
   ```
5. Las solicitudes generadas se guardarán en la carpeta `solicitudes_generadas`

## Formato del Archivo de Datos

El archivo `Datos.xlsx` debe contener las siguientes columnas:

- Código
- Código de Barras
- Descripción
- Marca
- Modelo
- Serie
- Valor Residual
- Valor Mercado
- Ubicación
- Área
- Estado de Conservación
- Estado de Operación
- Especialista
- Firma
- Motivo
- Otro Motivo
- Lugar y Fecha
- Dirección Zonal
- Ruta de Imagen

## Notas

- Las imágenes deben estar en el directorio `imagenes/`
- El programa maneja automáticamente celdas combinadas en la plantilla
- Se genera un único archivo Excel con múltiples hojas, una por cada solicitud