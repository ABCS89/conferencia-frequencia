# -*- coding: utf-8 -*-
"""
Extrai os registros do relatório "Ocorrência Geral" gerado pelo SISTEMA.
Este PDF já vem em formato de tabela (Divisão | Funcionário | Pessoa |
Data Inicial | Data Final | Qtde Dias | Descrição), então usamos
pdfplumber.extract_tables().
"""
import re
import pdfplumber

MATRICULA_RE = re.compile(r"^\d{2}\.\d{3}-\d$")


def extrair(caminho_pdf):
    """Retorna uma lista de dicts: matricula, nome, data_inicial, data_final,
    qtde_dias, descricao."""
    registros = []
    with pdfplumber.open(caminho_pdf) as pdf:
        for pagina in pdf.pages:
            for tabela in pagina.extract_tables():
                for linha in tabela:
                    linha = [c.strip() if isinstance(c, str) else c for c in linha]
                    # Layout: ['', matricula, nome, data_ini, data_fim, qtde, descricao]
                    # (a 1ª coluna "Divisão" quase sempre vem vazia neste relatório)
                    celulas = [c for c in linha if c not in (None, "")]
                    if len(celulas) < 5:
                        continue
                    matricula_candidata = None
                    resto = celulas
                    for c in celulas:
                        if MATRICULA_RE.match(c):
                            matricula_candidata = c
                            resto = celulas[celulas.index(c) + 1:]
                            break
                    if not matricula_candidata:
                        continue  # linha de cabeçalho/rodapé/total

                    if len(resto) < 4:
                        continue
                    nome, data_ini, data_fim, qtde, *desc = resto
                    if not re.match(r"\d{2}/\d{2}/\d{4}", data_ini):
                        continue
                    descricao = " ".join(" ".join(desc).split()).strip()
                    try:
                        qtde_f = float(qtde.replace(".", "").replace(",", "."))
                    except ValueError:
                        continue

                    registros.append({
                        "matricula": matricula_candidata,
                        "nome": nome.strip(),
                        "data_inicial": data_ini.strip(),
                        "data_final": data_fim.strip(),
                        "qtde_dias": qtde_f,
                        "descricao": descricao,
                    })
    return registros


def extrair_cabecalho_sistema(caminho_pdf):
    """Extrai informações do cabeçalho do relatório do SISTEMA (página 1).
    Exemplo:
        'Referente: 01/09/2026 a 30/09/2026' -> ano_mes = (2026, 9)
    Retorna dict com 'periodo': (data_ini, data_fim) e 'ano_mes': (ano, mes) ou None.
    """
    try:
        with pdfplumber.open(caminho_pdf) as pdf:
            if pdf.pages:
                texto = pdf.pages[0].extract_text() or ""
                m = re.search(r"Referente:\s*(\d{2})/(\d{2})/(\d{4})\s*a\s*(\d{2})/(\d{2})/(\d{4})", texto)
                if m:
                    d_ini, m_ini, a_ini, d_fim, m_fim, a_fim = m.groups()
                    return {
                        "periodo": (f"{d_ini}/{m_ini}/{a_ini}", f"{d_fim}/{m_fim}/{a_fim}"),
                        "ano_mes": (int(a_ini), int(m_ini)),
                    }
    except Exception:
        pass
    return {"periodo": None, "ano_mes": None}


if __name__ == "__main__":
    import sys
    import json
    regs = extrair(sys.argv[1])
    print(f"{len(regs)} registros extraídos.")
    print(json.dumps(regs[:5], ensure_ascii=False, indent=2))
