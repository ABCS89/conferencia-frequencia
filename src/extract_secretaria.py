# -*- coding: utf-8 -*-
"""
Extrai os registros do relatório "Frequência dos Funcionários" gerado
pela SECRETARIA. Este PDF não vem em formato de tabela real: cada
funcionário aparece em uma linha "MATRICULA NOME OCORRENCIA [DATA QTDE]"
e, se tiver mais de uma ocorrência no mês, as ocorrências extras vêm em
linhas seguintes só com "OCORRENCIA DATA QTDE" (sem repetir matrícula/nome).

Suporta leitura de texto nativo com pdfplumber e fallback transparente
com OCR nativo do Windows (winocr + pypdfium2) para PDFs em imagem/curvas.
"""
import re
import sys
from pathlib import Path
import pdfplumber

# Importa o módulo de OCR auxiliar
try:
    from ocr_helper import (
        ocr_disponivel,
        extrair_linhas_pagina_ocr,
        extrair_texto_pdf_ocr,
        normalizar_linha_ocr
    )
except ImportError:
    from src.ocr_helper import (
        ocr_disponivel,
        extrair_linhas_pagina_ocr,
        extrair_texto_pdf_ocr,
        normalizar_linha_ocr
    )

MATRICULA_RE = re.compile(r"^(\d{2})[\.,](\d{3})[-_ ](\d\b)")
DATE_QTY_TAIL_RE = re.compile(r"(\d{2}/\d{2}/\d{2,4})\s+([\d\.,]+)\s*$")
DATE_ONLY_TAIL_RE = re.compile(r"(\d{2}/\d{2}/\d{2,4})\s*$")


def _normalizar_data_str(data_str, ano_base=2026):
    """Corrige datas com anos incompletos de OCR (ex.: '03/08/202' -> '03/08/2026', '03/08/26' -> '03/08/2026')."""
    if not data_str:
        return None
    partes = data_str.split("/")
    if len(partes) == 3:
        d, m, y = partes
        if len(y) == 2:
            y = f"20{y}"
        elif len(y) == 3 and y.startswith("20"):
            y = f"{y}{str(ano_base)[-1]}"
        return f"{d.zfill(2)}/{m.zfill(2)}/{y}"
    return data_str

# linhas de cabeçalho/rodapé/assinatura a ignorar
IGNORAR_SUBSTRINGS = [
    "PREFEITURA DO MUNICÍPIO",
    "PREFEITURA DO DE PIRACICABA",
    "Frequência dos Funcionários",
    "Referente:",
    "Nro Funcional Nome",
    "Pág.",
    "Página:",
    "Peça do processo",
    "Assinaturas do documento",
    "Código para verificação",
    "Este documento foi assinado",
    "Emitido por",
    "Assinatura ICP",
    "Assinatura do Sistema",
    "válido até",
    "Para verificar a autenticidade",
    "aponte a câmera",
    "gerada automaticamente",
    "materializada por",
    "sempapel.piracicaba",
    "CPF:",
    "NÚCLEO DE APOIO",
    "sistemas.pmp.sp.gov.br",
]

MESES = {
    "janeiro": 1, "fevereiro": 2, "março": 3, "marco": 3, "abril": 4,
    "maio": 5, "junho": 6, "julho": 7, "agosto": 8, "setembro": 9,
    "outubro": 10, "novembro": 11, "dezembro": 12
}


def _parse_cabecalho_linhas(linhas):
    """Auxiliar para extrair código, secretaria e mês/ano de uma lista de linhas."""
    codigo_sec, nome_sec, mes_ano_ref = None, None, None
    for linha in linhas[:25]:
        linha_limpa = linha.strip()

        # 1. Mês/Ano de referência (ex: Referente: Julho/2026 ou Referente: Agosto/2026)
        m_ref = re.search(r"Referente:\s*([A-Za-zçÇ]+)/(\d{4})", linha_limpa)
        if m_ref and not mes_ano_ref:
            mes_str, ano_str = m_ref.groups()
            num_mes = MESES.get(mes_str.lower())
            if num_mes:
                mes_ano_ref = (int(ano_str), num_mes)

        # 2. Código e Nome da Secretaria (ex: 108 SECRETARIA MUNICIPAL DE OBRAS, INFRAESTRUTURA...)
        m_sec = re.match(r"^(\d{3})\s+([A-ZÁÉÍÓÚÀÈÌÒÙÃÕÂÊÎÔÛÇ][A-ZÁÉÍÓÚÀÈÌÒÙÃÕÂÊÎÔÛÇ\s,\.\-/]+)$", linha_limpa)
        if m_sec and not codigo_sec:
            codigo_sec = m_sec.group(1)
            nome_sec = m_sec.group(2).strip()

    return codigo_sec, nome_sec, mes_ano_ref


def extrair_cabecalho(caminho_pdf):
    """Extrai código da secretaria, nome da secretaria e mês/ano de referência do cabeçalho da página 1."""
    codigo_sec, nome_sec, mes_ano_ref = None, None, None
    with pdfplumber.open(caminho_pdf) as pdf:
        if pdf.pages:
            texto = pdf.pages[0].extract_text() or ""
            codigo_sec, nome_sec, mes_ano_ref = _parse_cabecalho_linhas(texto.split("\n"))

    # Fallback via OCR se não encontrou cabeçalho completo no texto nativo
    if (not codigo_sec or not mes_ano_ref) and ocr_disponivel():
        try:
            import pypdfium2 as pdfium
            doc = pdfium.PdfDocument(str(caminho_pdf))
            if len(doc) > 0:
                linhas_p0 = extrair_linhas_pagina_ocr(doc[0])
                c_ocr, n_ocr, m_ocr = _parse_cabecalho_linhas(linhas_p0)
                codigo_sec = codigo_sec or c_ocr
                nome_sec = nome_sec or n_ocr
                mes_ano_ref = mes_ano_ref or m_ocr
        except Exception:
            pass

    return codigo_sec, nome_sec, mes_ano_ref


def _eh_ignoravel(linha):
    linha_limpa = linha.strip()
    if not linha_limpa:
        return True
    if re.match(r"^\d+/\d+$", linha_limpa):  # "1/7", "2/7"...
        return True
    if re.match(r"^\d{2,4}-\d{2}", linha_limpa):  # "00-00", "00-00-000", "10-81"...
        return True
    if re.match(r"^\d{3}\s+[A-ZÁÉÍÓÚÀÈÌÒÙÃÕÂÊÎÔÛÇ]", linha_limpa):  # cabeçalho com código e nome da secretaria
        return True
    if re.match(r"^\d{2,4}[-\d]+\s+[A-ZÁÉÍÓÚÀÈÌÒÙÃÕÂÊÎÔÛÇ]", linha_limpa):  # subdivisão de departamento
        return True
    if "http://" in linha_limpa or "https://" in linha_limpa:
        return True
    for s in IGNORAR_SUBSTRINGS:
        if s.lower() in linha_limpa.lower():
            return True
    return False


def _eh_token_maiusculo(tok):
    """True se o token é 'nome' (todo maiúsculo, ignorando pontuação)."""
    letras = [c for c in tok if c.isalpha()]
    if not letras:
        return True  # token sem letras (ex.: pontuação) não quebra o nome
    return all(c.isupper() for c in letras)


def _separar_nome_ocorrencia(texto):
    """Dado 'ACYR CARDOSO ... Frequência normal' ou 'ADILSON ... Abono',
    retorna (nome, ocorrencia)."""
    tokens = texto.split()
    corte = len(tokens)
    for i, tok in enumerate(tokens):
        if not _eh_token_maiusculo(tok):
            corte = i
            break
    nome = " ".join(tokens[:corte]).strip()
    ocorrencia = " ".join(tokens[corte:]).strip()
    return nome, ocorrencia


def _processar_linhas(linhas):
    """Processa uma lista de linhas de texto e extrai os registros dos funcionários."""
    registros = []
    matricula_atual = None
    nome_atual = None
    nao_reconhecidas = []

    for linha in linhas:
        linha = normalizar_linha_ocr(linha.strip())
        if _eh_ignoravel(linha):
            continue

        m_data_qty = DATE_QTY_TAIL_RE.search(linha)
        data, qtde = (None, None)
        corpo = linha
        if m_data_qty:
            data, qtde_str = m_data_qty.groups()
            corpo = linha[: m_data_qty.start()].strip()
            try:
                qtde = float(qtde_str.replace(".", "").replace(",", "."))
            except ValueError:
                qtde = None
        else:
            m_data_only = DATE_ONLY_TAIL_RE.search(linha)
            if m_data_only:
                data = m_data_only.group(1)
                corpo = linha[: m_data_only.start()].strip()
                qtde = 1.0

        data = _normalizar_data_str(data)

        m_mat = MATRICULA_RE.match(linha)
        if m_mat:
            # Normaliza para o formato padrão XX.XXX-X
            matricula_atual = f"{m_mat.group(1)}.{m_mat.group(2)}-{m_mat.group(3)}"
            pos_fim = m_mat.end()
            if m_data_qty or m_data_only:
                resto_corpo = corpo[pos_fim:].strip()
            else:
                resto_corpo = linha[pos_fim:].strip()
            nome_atual, ocorrencia = _separar_nome_ocorrencia(resto_corpo)
            registros.append({
                "matricula": matricula_atual,
                "nome": nome_atual,
                "ocorrencia": ocorrencia,
                "data": data,
                "qtde_dias": qtde,
            })
        elif (m_data_qty or m_data_only) and corpo and corpo[:1].isalpha() and corpo[:1].isupper():
            # linha de continuação (ocorrência extra do último funcionário)
            if matricula_atual is None:
                nao_reconhecidas.append(linha)
                continue
            ocorrencia = corpo.strip()
            registros.append({
                "matricula": matricula_atual,
                "nome": nome_atual,
                "ocorrencia": ocorrencia,
                "data": data,
                "qtde_dias": qtde,
            })
        else:
            nao_reconhecidas.append(linha)

    return registros, nao_reconhecidas


def extrair(caminho_pdf):
    """Retorna lista de dicts: matricula, nome, ocorrencia, data, qtde_dias."""
    # 1. Tenta extrair diretamente via pdfplumber
    linhas_texto = []
    with pdfplumber.open(caminho_pdf) as pdf:
        for pagina in pdf.pages:
            texto = pagina.extract_text() or ""
            linhas_texto.extend(texto.split("\n"))

    registros, nao_reconhecidas = _processar_linhas(linhas_texto)

    # 2. Se nenhum registro foi encontrado e o OCR estiver disponível, aciona fallback
    if len(registros) == 0 and ocr_disponivel():
        nome_arquivo = Path(caminho_pdf).name
        print(f"  [INFO] O arquivo '{nome_arquivo}' está em formato de imagem/curvas.")
        print(f"         Executando OCR automático nativo do Windows...")
        try:
            paginas_ocr = extrair_texto_pdf_ocr(caminho_pdf)
            linhas_ocr = []
            for p_texto in paginas_ocr:
                linhas_ocr.extend(p_texto.split("\n"))
            registros, nao_reconhecidas = _processar_linhas(linhas_ocr)
            print(f"  [OK] OCR concluído com sucesso: {len(registros)} registro(s) extraído(s)!")
        except Exception as e:
            print(f"  [ERRO] Falha ao executar OCR: {e}")

    return registros, nao_reconhecidas


if __name__ == "__main__":
    import json
    if len(sys.argv) < 2:
        print("Uso: python extract_secretaria.py <caminho_do_pdf>")
        print("Exemplo: python extract_secretaria.py \"../input/secretaria/108 - secretaria.pdf\"")
        sys.exit(1)

    caminho = sys.argv[1]
    cod, nome, ref = extrair_cabecalho(caminho)
    print(f"Cabeçalho: Código={cod}, Nome={nome}, Referência={ref}")
    regs, nao_rec = extrair(caminho)
    print(f"{len(regs)} registros extraídos.")
    if nao_rec:
        print(f"{len(nao_rec)} linha(s) não reconhecida(s) (revisar se parecer dado de verdade):")
        for l in nao_rec[:8]:
            print("   ", repr(l))
    print("\nExemplo de registros extraídos:")
    print(json.dumps(regs[:5], ensure_ascii=False, indent=2))
