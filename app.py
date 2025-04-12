import openpyxl
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.drawing.image import Image
from PIL import Image as PILImage
import io
import os

def escribir_en_celda(hoja, coordenada, valor):
    """Función para escribir en una celda, manejando el caso de celdas combinadas.
    Si la celda es parte de un rango combinado, escribe en la celda principal del rango."""
    celda = hoja[coordenada]
    if isinstance(celda, MergedCell):
        # Buscar la celda principal del rango combinado
        for rango in hoja.merged_cells.ranges:
            if coordenada in rango:
                # Obtener la celda principal (esquina superior izquierda del rango)
                celda_principal = hoja.cell(row=rango.min_row, column=rango.min_col)
                celda_principal.value = valor
                return
    else:
        # Si no es una celda combinada, escribir directamente
        celda.value = valor

def insertar_imagen(hoja, ruta_imagen, celda_anclaje, ancho=None, alto=None):
    """Función para insertar una imagen en una hoja de Excel.
    
    Args:
        hoja: Hoja de Excel donde se insertará la imagen
        ruta_imagen: Ruta al archivo de imagen
        celda_anclaje: 
        Celda donde se anclará la imagen (ej: 'A1')
        ancho: Ancho deseado de la imagen en píxeles (opcional)
        alto: Alto deseado de la imagen en píxeles (opcional)
    """
    try:
        # Verificar si la imagen existe
        if not os.path.exists(ruta_imagen):
            print(f"La imagen {ruta_imagen} no existe.")
            return
        
        # Abrir y redimensionar la imagen si es necesario
        if ancho is not None or alto is not None:
            img_pil = PILImage.open(ruta_imagen)
            if ancho is not None and alto is not None:
                img_pil = img_pil.resize((ancho, alto))
            elif ancho is not None:
                ratio = ancho / img_pil.width
                alto_nuevo = int(img_pil.height * ratio)
                img_pil = img_pil.resize((ancho, alto_nuevo))
            elif alto is not None:
                ratio = alto / img_pil.height
                ancho_nuevo = int(img_pil.width * ratio)
                img_pil = img_pil.resize((ancho_nuevo, alto))
            
            # Guardar temporalmente la imagen redimensionada
            buffer = io.BytesIO()
            img_pil.save(buffer, format=img_pil.format or 'PNG')
            buffer.seek(0)
            
            # Crear objeto Image de openpyxl desde el buffer
            img = Image(buffer)
        else:
            # Usar la imagen original sin redimensionar
            img = Image(ruta_imagen)
        
        # Obtener la celda de anclaje
        celda = hoja[celda_anclaje]
        
        # Insertar la imagen anclada a la celda
        hoja.add_image(img, celda_anclaje)
        return True
    except Exception as e:
        print(f"Error al insertar la imagen: {e}")
        return False

def llenar_solicitudes(datos_excel, plantilla_excel, carpeta_salida):
    # Cargar el archivo con los datos
    wb_datos = load_workbook(datos_excel)
    hoja_datos = wb_datos["Datos"]

    # Crear la carpeta de salida si no existe
    if not os.path.exists(carpeta_salida):
        os.makedirs(carpeta_salida)
    
    # Crear un único archivo Excel para todas las solicitudes
    # Cargamos directamente la plantilla como base para mantener todos los estilos y formatos
    wb_solicitudes = load_workbook(plantilla_excel)
    # Guardamos la hoja original para usarla como plantilla
    hoja_original = wb_solicitudes.active
    
    # Iterar sobre cada fila de la hoja "Datos"
    for fila in range(2, hoja_datos.max_row + 1):  # Empezamos en la fila 2 para omitir encabezados
        # Extraer el código para nombrar la hoja
        codigo = hoja_datos[f"A{fila}"].value
        
        # Copiar la hoja original para cada activo (esto mantiene todos los formatos, estilos, etc.)
        hoja_solicitud = wb_solicitudes.copy_worksheet(hoja_original)
        hoja_solicitud.title = f"Solicitud_{codigo}"
        
        # Extraer los datos de la fila actual
        codigo = hoja_datos[f"A{fila}"].value
        cod_barra = hoja_datos[f"B{fila}"].value
        descripcion = hoja_datos[f"C{fila}"].value
        marca = hoja_datos[f"D{fila}"].value
        modelo = hoja_datos[f"E{fila}"].value
        serie = hoja_datos[f"F{fila}"].value
        valor_residual = hoja_datos[f"G{fila}"].value
        valor_mercado = hoja_datos[f"H{fila}"].value
        ubicacion = hoja_datos[f"I{fila}"].value
        area = hoja_datos[f"J{fila}"].value
        estado_conservacion = hoja_datos[f"K{fila}"].value
        estado_operacion = hoja_datos[f"L{fila}"].value
        especialista = hoja_datos[f"M{fila}"].value
        firma = hoja_datos[f"N{fila}"].value
        motivo = hoja_datos[f"O{fila}"].value
        otro_motivo = hoja_datos[f"P{fila}"].value
        lugar_fecha = hoja_datos[f"Q{fila}"].value
        direccion_zonal = hoja_datos[f"R{fila}"].value
        # Extraer la ruta de la imagen (nueva columna S)
        ruta_imagen = hoja_datos[f"S{fila}"].value

        # Llenar la plantilla con los datos extraídos usando la función para manejar celdas combinadas
        # Coordenadas corregidas según la estructura visual de la plantilla original
        escribir_en_celda(hoja_solicitud, "B8", f"{codigo}")  # Número de solicitud (N°: 2018-DN-00)
        escribir_en_celda(hoja_solicitud, "C17", f"{codigo}") 
        escribir_en_celda(hoja_solicitud, "H8", descripcion)  # Descripción del activo
        escribir_en_celda(hoja_solicitud, "B10", marca)  # Marca
        escribir_en_celda(hoja_solicitud, "F10", modelo)  # Modelo
        escribir_en_celda(hoja_solicitud, "K10", serie)  # Serie
        escribir_en_celda(hoja_solicitud, "D8", cod_barra)  # Código de barras (corregido a D8 según solicitud)
        escribir_en_celda(hoja_solicitud, "D12", valor_residual)  # Valor residual
        escribir_en_celda(hoja_solicitud, "L12", valor_mercado)  # Valor mercado
        escribir_en_celda(hoja_solicitud, "A1", ubicacion)  # Ubicación física
        escribir_en_celda(hoja_solicitud, "G17", area)  # Área
        escribir_en_celda(hoja_solicitud, "L27", estado_conservacion)  # Estado de conservación
        escribir_en_celda(hoja_solicitud, "B19", estado_operacion)  # Estado de operación  y Opinión técnica
        escribir_en_celda(hoja_solicitud, "D30", especialista)  # Especialista
        escribir_en_celda(hoja_solicitud, "M30", firma)  # Firma
        escribir_en_celda(hoja_solicitud, "C36", motivo)  # Motivo de baja
        if motivo == "Otros" and otro_motivo:
            escribir_en_celda(hoja_solicitud, "l37", otro_motivo)  # Otro motivo de baja
        escribir_en_celda(hoja_solicitud, "D40", lugar_fecha)  # Lugar y fecha
        escribir_en_celda(hoja_solicitud, "B44", direccion_zonal)  # Dirección zonal
        
        # Insertar la imagen si existe la ruta
        if ruta_imagen:
            # Verificar si la ruta es absoluta o relativa
            if not os.path.isabs(ruta_imagen):
                # Si es relativa, construir la ruta absoluta basada en el directorio del script
                ruta_imagen_abs = os.path.join(os.path.dirname(datos_excel), ruta_imagen)
            else:
                ruta_imagen_abs = ruta_imagen
                
            # Insertar la imagen en la celda F22 (ajustar según la ubicación del cuadro de texto en la plantilla)
            # El ancho y alto se pueden ajustar según sea necesario
            insertar_imagen(hoja_solicitud, ruta_imagen_abs, "K18", ancho=200, alto=150)

    # Eliminar la hoja original después de crear todas las copias
    wb_solicitudes.remove(hoja_original)
    
    # Guardar el archivo único con todas las solicitudes
    nombre_archivo = os.path.join(carpeta_salida, "Solicitudes_Baja.xlsx")
    # Verificar si el archivo existe y si está abierto
    try:
        if os.path.exists(nombre_archivo):
            # Intentar eliminar el archivo existente
            try:
                os.remove(nombre_archivo)
            except PermissionError:
                print(f"No se puede sobrescribir {nombre_archivo}. El archivo podría estar abierto.")
                return
        wb_solicitudes.save(nombre_archivo)
        print(f"Proceso completado. Se ha generado el archivo {nombre_archivo} con {len(wb_solicitudes.sheetnames)} hojas.")
    except Exception as e:
        print(f"Error al guardar {nombre_archivo}: {e}")

# Ejecutar el script
if __name__ == "__main__":
    # Obtener la ruta del directorio actual donde se encuentra el script
    directorio_actual = os.path.dirname(os.path.abspath(__file__))
    
    # Rutas de los archivos con rutas absolutas
    datos_excel = os.path.join(directorio_actual, "Datos.xlsx")  # Archivo con los datos
    plantilla_excel = os.path.join(directorio_actual, "F&A-P-01 ANEXO N° 02 SOLICITUD DE BAJA.xlsx")  # Plantilla base
    carpeta_salida = os.path.join(directorio_actual, "solicitudes_generadas")  # Carpeta donde se guardarán los archivos
    
    # Llamar a la función principal
    llenar_solicitudes(datos_excel, plantilla_excel, carpeta_salida)