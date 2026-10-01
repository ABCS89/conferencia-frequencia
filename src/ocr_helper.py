# -*- coding: utf-8 -*-
"""
ocr_helper.py - Utilitário de OCR nativo do Windows para PDFs em imagem/curvas.

Utiliza pypdfium2 para renderização das páginas e winocr (Windows.Media.Ocr)
para reconhecimento óptico de caracteres 100% local, offline e em pt-BR.
"""
import re
from pathlib import Path

try:
    import pypdfium2 as pdfium
    import winocr
    OCR_DISPONIVEL = True
except ImportError:
    OCR_DISPONIVEL = False


def ocr_disponivel():
    """Retorna True se os módulos de OCR estiverem instalados."""
    return OCR_DISPONIVEL


def normalizar_linha_ocr(linha):
    """
    Corrige pequenos artefatos comuns de OCR:
    - Matrícula com traço lido como underline ou espaço (ex: '24.976_9' -> '24.976-9', '28.491 2' -> '28.491-2').
    - Aspas ou apóstrofos espúrios no início da linha.
    """
    linha = linha.strip()
    # Normaliza formato de matrícula NN.NNN-N
    linha = re.sub(r"^(\d{2})[\.,](\d{3})[-_ ](\d\b)", r"\1.\2-\3", linha)
    return linha


def extrair_linhas_pagina_ocr(page, scale=3.0):
    """
    Renderiza uma página do pdfium e extrai suas linhas de texto via OCR,
    ordenando as palavras horizontalmente por linha e descartando a barra
    lateral de assinaturas digitais do SemPapel.
    """
    img = page.render(scale=scale).to_pil()
    largura, altura = img.size
    
    # Executa OCR nativo do Windows em português do Brasil
    res = winocr.recognize_pil_sync(img, "pt-BR")
    
    words = []
    # 0.95 preserva a coluna de quantidades à direita (~0.92-0.94)
    # e descarta a barra lateral vertical de assinaturas do SemPapel (>= 0.96)
    limite_x = largura * 0.95
    
    for line in res.get("lines", []):
        for w in line.get("words", []):
            rect = w.get("bounding_rect", {})
            x = rect.get("x", 0)
            y = rect.get("y", 0)
            h = rect.get("height", 14)
            text = w.get("text", "")
            if x < limite_x:
                # Se estiver próximo da margem direita (> 0.93), só inclui se contiver dígitos
                # (preserva quantidades como '15,0', '30,0', '884,0' e descarta letras de assinatura)
                if x > largura * 0.93 and not any(c.isdigit() for c in text):
                    continue
                words.append((y, x, h, text))
                
    if not words:
        return []

    # Agrupa palavras por proximidade na coordenada vertical Y
    words.sort(key=lambda w: (w[0], w[1]))
    tolerancia_y = 12 * (scale / 2.0)
    
    linhas_agrupadas = []
    for y, x, h, text in words:
        colocado = False
        for grupo in linhas_agrupadas:
            avg_y = sum(item[0] for item in grupo) / len(grupo)
            if abs(y - avg_y) < tolerancia_y:
                grupo.append((y, x, h, text))
                colocado = True
                break
        if not colocado:
            linhas_agrupadas.append([(y, x, h, text)])
            
    # Ordena linhas de cima para baixo
    linhas_agrupadas.sort(key=lambda g: sum(item[0] for item in g) / len(g))
    
    resultado = []
    for grupo in linhas_agrupadas:
        # Ordena palavras da linha da esquerda para a direita
        grupo.sort(key=lambda item: item[1])
        linha_texto = " ".join(item[3] for item in grupo)
        linha_norm = normalizar_linha_ocr(linha_texto)
        if linha_norm:
            resultado.append(linha_norm)
            
    return resultado


def extrair_texto_pdf_ocr(caminho_pdf, scale=3.0):
    """
    Abre o PDF com pdfium e extrai todas as linhas de texto página a página via OCR.
    Retorna uma lista de strings (uma por página, com quebras de linha \\n).
    """
    if not OCR_DISPONIVEL:
        raise RuntimeError("Módulos de OCR (pypdfium2 e winocr) não estão instalados.")
        
    pdf = pdfium.PdfDocument(str(caminho_pdf))
    paginas_texto = []
    for page in pdf:
        linhas = extrair_linhas_pagina_ocr(page, scale=scale)
        paginas_texto.append("\n".join(linhas))
    return paginas_texto
