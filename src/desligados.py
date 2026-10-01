# -*- coding: utf-8 -*-
"""
desligados.py - Leitor da base de controle de servidores desligados (ODS / Excel).

Permite cruzar a lista de servidores desligados na conferência de frequência,
ignorando divergências de servidores já desligados quando não houver registro no sistema
ou confirmando seu status de desligamento.
"""
import os
import re
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

NAMESPACES = {
    "table": "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
    "text": "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
}


def normalizar_matricula(matricula_str):
    """Extrai os dígitos e preenche com zeros à esquerda até 6 dígitos (padrão XX.XXX-X)."""
    if not matricula_str:
        return ""
    digitos = re.sub(r"\D", "", str(matricula_str))
    if not digitos:
        return ""
    return digitos.zfill(6)


def formatar_matricula(matricula_norm):
    """Converte 6 dígitos para o padrão XX.XXX-X."""
    if len(matricula_norm) == 6:
        return f"{matricula_norm[:2]}.{matricula_norm[2:5]}-{matricula_norm[5]}"
    return matricula_norm


def carregar_desligados_ods(caminho_ods):
    """Lê todas as abas de um arquivo .ods de Controle de Desligamentos
    e retorna dict {matricula_norm: info}."""
    desligados = {}
    if not os.path.exists(caminho_ods):
        return desligados

    with zipfile.ZipFile(caminho_ods) as z:
        with z.open("content.xml") as f:
            tree = ET.parse(f)
            root = tree.getroot()

    for table in root.findall(".//table:table", NAMESPACES):
        aba = table.attrib.get("{urn:oasis:names:tc:opendocument:xmlns:table:1.0}name", "")
        for row in table.findall(".//table:table-row", NAMESPACES):
            cells = []
            for cell in row.findall(".//table:table-cell", NAMESPACES):
                repeated = int(cell.attrib.get("{urn:oasis:names:tc:opendocument:xmlns:table:1.0}number-columns-repeated", 1))
                texts = [t.text or "" for t in cell.findall(".//text:p", NAMESPACES)]
                val = " ".join(texts).strip()
                if repeated > 30:
                    continue
                for _ in range(min(repeated, 10)):
                    cells.append(val)
            while cells and not cells[-1]:
                cells.pop()
            if not cells:
                continue

            # Procura matrícula numérica (5 a 6 dígitos)
            for i, c in enumerate(cells):
                d = re.sub(r"\D", "", c)
                if len(d) in (5, 6) and c.isdigit():
                    mat_norm = d.zfill(6)
                    mat_fmt = formatar_matricula(mat_norm)
                    nome = ""
                    dt_dem = ""
                    for j in range(i + 1, len(cells)):
                        val = cells[j]
                        if not val:
                            continue
                        if re.match(r"^\d{2}/\d{2}/\d{4}$", val):
                            if not dt_dem:
                                dt_dem = val
                        elif not nome and any(ch.isalpha() for ch in val):
                            nome = val

                    desligados[mat_norm] = {
                        "matricula_norm": mat_norm,
                        "matricula_fmt": mat_fmt,
                        "nome": nome,
                        "data_demissao": dt_dem,
                        "aba": aba,
                    }
                    break

    return desligados


def localizar_e_carregar_desligados(diretorio_base):
    """Procura automaticamente por arquivos de desligamento na pasta informada ou na raiz pai."""
    base = Path(diretorio_base)
    pastas_busca = [base]
    if base.name != "input" and (base.parent / "input").exists():
        pastas_busca.append(base.parent / "input")
    elif base.parent.exists() and base.parent != base:
        pastas_busca.append(base.parent)

    for p in pastas_busca:
        candidatos = list(p.glob("*desligamento*.ods")) + \
                     list(p.glob("*desligado*.ods")) + \
                     list(p.glob("*.ods"))

        for arq in candidatos:
            if arq.is_file():
                try:
                    dados = carregar_desligados_ods(str(arq))
                    if dados:
                        return dados, arq.name
                except Exception:
                    continue
    return {}, None
