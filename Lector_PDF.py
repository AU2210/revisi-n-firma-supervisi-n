"""
Módulo de Procesamiento y Blindaje de Documentos PDF
----------------------------------------------------
Este script automatiza la conversión de documentos Word a PDF, 
la inserción precisa de firmas digitales y marcas de agua (VBO), 
y aplica medidas de seguridad avanzadas (aplanamiento de capas y encriptación AES-256)
para garantizar la integridad de los contratos administrativos.
"""

import os
import comtypes.client
import fitz  # PyMuPDF
import base64

def convertir_word_a_pdf(ruta_word: str, ruta_pdf: str):
    """
    Abre Microsoft Word en modo fantasma (invisible), 
    convierte el documento a PDF para no perder el formato, y lo cierra.
    """
    # Necesitamos rutas absolutas para que Word no se pierda
    ruta_w = os.path.abspath(ruta_word)
    ruta_p = os.path.abspath(ruta_pdf)
    
    # Invocamos a Word de forma invisible
    word = comtypes.client.CreateObject('Word.Application')
    word.Visible = False
    
    try:
        doc = word.Documents.Open(ruta_w)
        # 17 es el código interno de Microsoft para guardar como PDF
        doc.SaveAs(ruta_p, FileFormat=17) 
        doc.Close()
    finally:
        # Siempre cerramos Word para no dejar procesos colgados en el sistema
        word.Quit()
        
    return ruta_p

def pdf_a_imagenes(ruta_pdf: str):
    """
    Lee el PDF y convierte cada página en una imagen de alta resolución.
    Ideal para visualización de documentos en interfaces frontend o dispositivos móviles.
    """
    doc = fitz.open(ruta_pdf)
    imagenes = []
    
    for i in range(len(doc)):
        pagina = doc.load_page(i)
        # Zoom 2x para mantener la nitidez en resoluciones altas
        matriz = fitz.Matrix(2, 2) 
        pix = pagina.get_pixmap(matrix=matriz)
        
        # Convertimos la imagen a texto (Base64) para transmisión segura por red
        img_base64 = base64.b64encode(pix.tobytes("png")).decode("utf-8")
        imagenes.append(img_base64)
        
    doc.close()
    return imagenes

def estampar_firma_pdf(ruta_pdf_temp: str, pagina_num: int, pct_x: float, pct_y: float, pct_w: float, pct_h: float, ruta_firma: str, ruta_vbo: str, ruta_salida: str):
    """
    Inserta la firma en coordenadas específicas, aplana la página para evitar 
    extracción de la imagen, aplica marcas VBO y asegura el archivo final.
    """
    doc = fitz.open(ruta_pdf_temp)
    
    # --- 1. ESTAMPAR LA FIRMA PRINCIPAL ---
    pagina_firma = doc.load_page(pagina_num)
    ancho = pagina_firma.rect.width
    alto = pagina_firma.rect.height
    
    # Coordenadas relativas mapeadas desde la UI
    x = ancho * pct_x
    y = alto * pct_y
    w = ancho * pct_w
    h = alto * pct_h
    
    rect_firma = fitz.Rect(x, y, x + w, y + h)
    pagina_firma.insert_image(rect_firma, filename=ruta_firma)
    
    # --- 2. APLANAMIENTO Y APLICACIÓN DE VBO ---
    for i in range(len(doc)):
        page = doc[i]
        
        if i == pagina_num:
            # APLANAR LA PÁGINA DE LA FIRMA (Medida de Seguridad Anti-Extracción)
            pix = page.get_pixmap(matrix=fitz.Matrix(3, 3))
            rect = page.rect
            doc.delete_page(i)
            
            doc.insert_page(i, width=rect.width, height=rect.height)
            nueva_pagina = doc[i]
            nueva_pagina.insert_image(rect, pixmap=pix)
            
        else:
            # INSERTAR VBO EN LAS DEMÁS PÁGINAS (Abajo a la derecha)
            if os.path.exists(ruta_vbo):
                w_vbo, h_vbo = 35, 35 
                margen_x, margen_y = 60, 50 
                
                rect_vbo = fitz.Rect(
                    page.rect.width - w_vbo - margen_x,
                    page.rect.height - h_vbo - margen_y,
                    page.rect.width - margen_x,
                    page.rect.height - margen_y
                )
                page.insert_image(rect_vbo, filename=ruta_vbo)

    # --- 3. BLINDAJE CRIPTOGRÁFICO Y GESTIÓN DE PERMISOS ---
    permisos = fitz.PDF_PERM_ACCESSIBILITY | fitz.PDF_PERM_PRINT 
    
    # Práctica segura: Obtener contraseña de variable de entorno, con un fallback de ejemplo
    contrasena_propietario = os.getenv("PDF_OWNER_PASSWORD", "Configurar_Password_Seguro")
    
    doc.save(
        ruta_salida, 
        encryption=fitz.PDF_ENCRYPT_AES_256,
        owner_pw=contrasena_propietario,     
        user_pw="",            
        permissions=permisos
    )
    doc.close()