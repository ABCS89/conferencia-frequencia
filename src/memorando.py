# -*- coding: utf-8 -*-
"""
memorando.py - Leitura e conferência de memorandos de retificação de frequência.

Lê arquivos de memorando (PDFs nativos ou escaneados via OCR), extrai os servidores,
períodos e ocorrências retificadas pela secretaria, e cruza com a lista de retificações
solicitadas para atestar o cumprimento das pendências no próprio relatório Markdown.
"""
import re
import unicodedata
from datetime import date, datetime
from pathlib import Path

import pdfplumber

from ocr_helper import extrair_texto_pdf_ocr, ocr_disponivel
from desligados import normalizar_matricula


def _normalizar_texto(texto):
    """Minúsculas, sem acento, espaços colapsados."""
    if not texto:
        return ""
    nfkd = unicodedata.normalize("NFKD", texto)
    sem_acento = "".join(c for c in nfkd if not unicodedata.combining(c))
    return " ".join(sem_acento.lower().split())


def _parse_data(txt):
    if not txt:
        return None
    try:
        d, m, a = txt.split("/")
        return date(int(a), int(m), int(d))
    except Exception:
        return None


def _extrair_datas_ocorrencia(texto, ano_ref=2026, mes_ref=None):
    """Extrai data_inicio, data_fim e dias a partir de múltiplos formatos de data."""
    if not texto:
        return None, None, None

    # 1. Padrão: 03 e 04/10 ou 03 e 04/10/2026 ou 03 e 04/09
    m_e = re.search(r"\b(\d{2})\s+e\s+(\d{2})/(\d{2})(?:/(\d{4}))?\b", texto)
    if m_e:
        d1, d2, m, a = m_e.group(1), m_e.group(2), m_e.group(3), m_e.group(4)
        ano = int(a) if a else ano_ref
        mes = int(m)
        if mes_ref and mes != mes_ref:
            # Caso a secretaria tenha digitado o mês de expedição por engano ao redigir
            # o memorando (ex: digitou /10 em outubro para frequência de setembro)
            mes = mes_ref
        try:
            dt1 = date(ano, mes, int(d1))
            dt2 = date(ano, mes, int(d2))
            dt_ini, dt_fim = min(dt1, dt2), max(dt1, dt2)
            return dt_ini, dt_fim, (dt_fim - dt_ini).days + 1
        except Exception:
            pass

    # 2. Padrão período com ano completo: 01/09/2026 a 05/09/2026 ou 01/09/2026 ao 05/09/2026
    m_per = re.search(r"\b(\d{2})/(\d{2})/(\d{4})\s*(?:a|ao|até|-)\s*(\d{2})/(\d{2})/(\d{4})\b", texto, re.I)
    if m_per:
        d1, m1, a1, d2, m2, a2 = m_per.groups()
        try:
            dt1 = date(int(a1), int(m1), int(d1))
            dt2 = date(int(a2), int(m2), int(d2))
            dt_ini, dt_fim = min(dt1, dt2), max(dt1, dt2)
            return dt_ini, dt_fim, (dt_fim - dt_ini).days + 1
        except Exception:
            pass

    # 3. Padrão período sem ano: 01/09 a 05/09
    m_per_sem_ano = re.search(r"\b(\d{2})/(\d{2})\s*(?:a|ao|até|-)\s*(\d{2})/(\d{2})\b", texto, re.I)
    if m_per_sem_ano:
        d1, m1, d2, m2 = m_per_sem_ano.groups()
        try:
            mes1 = mes_ref if (mes_ref and int(m1) != mes_ref) else int(m1)
            mes2 = mes_ref if (mes_ref and int(m2) != mes_ref) else int(m2)
            dt1 = date(ano_ref, mes1, int(d1))
            dt2 = date(ano_ref, mes2, int(d2))
            dt_ini, dt_fim = min(dt1, dt2), max(dt1, dt2)
            return dt_ini, dt_fim, (dt_fim - dt_ini).days + 1
        except Exception:
            pass

    # 4. Datas completas DD/MM/AAAA
    datas_completas = re.findall(r"\b(\d{2})/(\d{2})/(\d{4})\b", texto)
    if datas_completas:
        d_objs = []
        for d, m, a in datas_completas:
            try:
                d_objs.append(date(int(a), int(m), int(d)))
            except Exception:
                pass
        if len(d_objs) == 1:
            return d_objs[0], d_objs[0], 1
        elif len(d_objs) >= 2:
            dt_ini, dt_fim = min(d_objs[0], d_objs[1]), max(d_objs[0], d_objs[1])
            return dt_ini, dt_fim, (dt_fim - dt_ini).days + 1

    # 5. Data única sem ano DD/MM
    m_unica = re.search(r"\b(\d{2})/(\d{2})\b", texto)
    if m_unica:
        d, m = m_unica.groups()
        mes = mes_ref if (mes_ref and int(m) != mes_ref) else int(m)
        try:
            dt = date(ano_ref, mes, int(d))
            return dt, dt, 1
        except Exception:
            pass

    return None, None, None


def localizar_memorando(codigo_sec, contexto_dir):
    """Localiza arquivo de memorando de retificação correspondente ao código da secretaria."""
    contexto_dir = Path(contexto_dir)
    if not contexto_dir.exists():
        return None

    padroes = [
        f"{codigo_sec}*retifica*.pdf",
        f"{codigo_sec}*memorando*.pdf",
        f"{codigo_sec}*memo*.pdf",
    ]

    pastas = [contexto_dir]
    if contexto_dir.name != "input" and (contexto_dir.parent / "input").exists():
        pastas.append(contexto_dir.parent / "input")
    elif contexto_dir.parent.exists() and contexto_dir.parent.name == "input":
        pastas.append(contexto_dir.parent)

    for sub in contexto_dir.iterdir():
        if sub.is_dir() and sub.name not in ("secretaria", "sistema"):
            pastas.append(sub)

    for p in pastas:
        if not p.is_dir():
            continue
        for padrao in padroes:
            encontrados = list(p.glob(padrao))
            if encontrados:
                return encontrados[0]

    return None


def extrair_texto_completo_memorando(caminho_pdf):
    """Extrai texto do memorando página a página (usa pdfplumber e recorre ao OCR em 300 DPI em páginas escaneadas)."""
    caminho_pdf = Path(caminho_pdf)
    paginas_texto = []
    
    PALAVRAS_CHAVE_FUNCIONAIS = [
        "retifica", "frequencia", "frequência", "memorando", "funcional",
        "servidor", "servidora", "consta", "adicionar", "alterar", "afastamento"
    ]

    with pdfplumber.open(str(caminho_pdf)) as pdf:
        doc_pdfium = None
        for idx, page in enumerate(pdf.pages):
            t = page.extract_text() or ""
            tem_conteudo_funcional = any(k in t.lower() for k in PALAVRAS_CHAVE_FUNCIONAIS)
            
            # Se a página for imagem/digitalizada sem texto funcional, executa OCR
            if not tem_conteudo_funcional and ocr_disponivel():
                try:
                    if doc_pdfium is None:
                        import pypdfium2 as pdfium
                        doc_pdfium = pdfium.PdfDocument(str(caminho_pdf))
                    from ocr_helper import extrair_linhas_pagina_ocr
                    linhas_ocr = extrair_linhas_pagina_ocr(doc_pdfium[idx], scale=3.0)
                    paginas_texto.append("\n".join(linhas_ocr))
                    continue
                except Exception as e:
                    pass
            paginas_texto.append(t)

    return "\n\n".join(paginas_texto)


def extrair_tabelas_memorando(caminho_pdf, ano_ref=2026, mes_ref=None):
    """Extrai itens estruturados de tabelas em PDFs de memorando (suporta tabelas nativas e com bordas duplas)."""
    itens = []
    if not caminho_pdf:
        return itens
    try:
        with pdfplumber.open(str(caminho_pdf)) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                tabela_nativa_sucesso = False
                for table in tables:
                    if not table or len(table) < 2:
                        continue
                    header = [str(c or "").upper() for c in table[0]]
                    tem_func = any("FUNC" in h or "MATR" in h for h in header)
                    tem_per = any("PER" in h or "DATA" in h for h in header)
                    if not (tem_func or tem_per):
                        continue

                    col_mat = -1
                    col_nome = -1
                    col_oco = -1
                    col_per = -1
                    for idx, h in enumerate(header):
                        if "FUNC" in h or "MATR" in h: col_mat = idx
                        elif "SERV" in h or "NOME" in h: col_nome = idx
                        elif "OCORR" in h or "EVENTO" in h: col_oco = idx
                        elif "PER" in h or "DATA" in h: col_per = idx

                    ult_mat = None
                    ult_nome = None
                    ult_oco = None

                    for row in table[1:]:
                        val_mat = row[col_mat] if col_mat >= 0 and col_mat < len(row) else None
                        val_nome = row[col_nome] if col_nome >= 0 and col_nome < len(row) else None
                        val_oco = row[col_oco] if col_oco >= 0 and col_oco < len(row) else None
                        val_per = row[col_per] if col_per >= 0 and col_per < len(row) else None

                        if val_mat and re.search(r"\d", str(val_mat)):
                            ult_mat = str(val_mat).strip()
                        if val_nome and str(val_nome).strip():
                            ult_nome = " ".join(str(val_nome).split())
                        if val_oco and str(val_oco).strip():
                            ult_oco = " ".join(str(val_oco).split())

                        if not val_per:
                            continue

                        per_str = str(val_per).strip()
                        d_ini, d_fim, dias = _extrair_datas_ocorrencia(per_str, ano_ref=ano_ref, mes_ref=mes_ref)
                        if not d_ini:
                            continue

                        if ult_mat:
                            m_norm = normalizar_matricula(ult_mat)
                            mat_fmt = f"{m_norm[:2]}.{m_norm[2:5]}-{m_norm[5:]}" if len(m_norm) == 6 else ult_mat
                            itens.append({
                                "matricula": mat_fmt,
                                "matricula_norm": m_norm,
                                "nome": ult_nome,
                                "ocorrencia": ult_oco or "Não informada",
                                "data_inicio": d_ini,
                                "data_fim": d_fim,
                                "dias": dias,
                                "texto_original": f"{ult_mat} - {ult_nome} - {ult_oco} - {per_str}",
                            })
                            tabela_nativa_sucesso = True

                # Fallback: tabelas desenhadas com linhas e sem cabeçalhos de coluna rotulados (ex: memo 108)
                if not tabela_nativa_sucesso and len(page.lines) >= 4:
                    h_lines = sorted(list(set(round(l['top'], 0) for l in page.lines if abs(l['top'] - l['bottom']) < 2 and l['width'] > 100)))
                    merged_h = []
                    for y in h_lines:
                        if not merged_h or (y - merged_h[-1]) > 5:
                            merged_h.append(y)
                        else:
                            merged_h[-1] = y
                    
                    if len(merged_h) >= 2:
                        for i in range(len(merged_h) - 1):
                            y0, y1 = merged_h[i], merged_h[i+1]
                            words = [w for w in page.extract_words() if y0 <= w['top'] <= y1 and w['x0'] < 550]
                            if not words:
                                continue
                            txt_row = " ".join(w['text'] for w in sorted(words, key=lambda x: (x['top'], x['x0'])))
                            m_mat = re.search(r"\b(\d{2}[\.,]\d{3}[-_ ]\d|\d{3}[\.,]\d{3}|\d{5,6})\b", txt_row)
                            if m_mat:
                                mat_bruta = m_mat.group(1)
                                mat_norm = normalizar_matricula(mat_bruta)
                                mat_fmt = f"{mat_norm[:2]}.{mat_norm[2:5]}-{mat_norm[5:]}" if len(mat_norm) == 6 else mat_bruta
                                
                                col_esq = [w['text'] for w in sorted([w for w in words if w['x0'] < 200], key=lambda x: (x['top'], x['x0']))]
                                col_dir = [w['text'] for w in sorted([w for w in words if w['x0'] >= 200], key=lambda x: (x['top'], x['x0']))]
                                txt_esq = " ".join(col_esq)
                                txt_dir = " ".join(col_dir)
                                
                                nome = None
                                m_nome = re.search(r"^([A-ZÁÉÍÓÚÀÈÌÒÙÃÕÂÊÎÔÛÇ\s]{4,45}?)(?:\s+N[ºo\.]*|\s+Matr[íi]cula|\s*[-–]|$)", txt_esq, re.I)
                                if m_nome:
                                    cand = m_nome.group(1).strip()
                                    if not any(k in cand.upper() for k in ["NOME", "FUNCIONAL", "MATRICULA"]):
                                        nome = cand.upper()
                                
                                d_ini, d_fim, dias = _extrair_datas_ocorrencia(txt_dir, ano_ref=ano_ref, mes_ref=mes_ref)
                                if not d_ini:
                                    d_ini, d_fim, dias = _extrair_datas_ocorrencia(txt_row, ano_ref=ano_ref, mes_ref=mes_ref)
                                
                                oco = txt_dir
                                oco_limpa = re.sub(r"\b\d{2}/\d{2}(?:/\d{4})?\b|\b\d{2}\s+e\s+\d{2}/\d{2}\b", "", oco)
                                oco_limpa = re.sub(r"^[-–\s]+|[-–\s]+$", "", oco_limpa).strip()
                                
                                itens.append({
                                    "matricula": mat_fmt,
                                    "matricula_norm": mat_norm,
                                    "nome": nome,
                                    "data_inicio": d_ini,
                                    "data_fim": d_fim,
                                    "dias": dias,
                                    "ocorrencia": oco_limpa or "Não informada",
                                    "texto_original": txt_row
                                })
    except Exception:
        pass
    return itens


def extrair_itens_memorando(texto, caminho_pdf=None):
    """
    Analisa o texto do memorando e extrai os itens retificados.
    Suporta múltiplos formatos:
    - Formato estruturado / tabela (com extração visual direta de tabelas e células mescladas)
    - Formato narrativo (parágrafos com 'Onde consta-se ... adicionar/alterar para ...')
    - Formato em tópicos (linhas com matrícula, nome, datas e evento)
    """
    MESES_MAP = {
        "janeiro": 1, "fevereiro": 2, "março": 3, "marco": 3, "abril": 4,
        "maio": 5, "junho": 6, "julho": 7, "agosto": 8, "setembro": 9,
        "outubro": 10, "novembro": 11, "dezembro": 12
    }
    ano_ref = 2026
    mes_ref = None
    m_ref = re.search(r"(janeiro|fevereiro|março|marco|abril|maio|junho|julho|agosto|setembro|outubro|novembro|dezembro)[/\s]+(\d{4})", texto, re.I)
    if m_ref:
        mes_ref = MESES_MAP.get(m_ref.group(1).lower())
        ano_ref = int(m_ref.group(2))
    elif caminho_pdf:
        m_pasta = re.search(r"(\d{4})-(\d{2})", str(caminho_pdf))
        if m_pasta:
            ano_ref = int(m_pasta.group(1))
            mes_ref = int(m_pasta.group(2))

    # 1. Identificação do cabeçalho do memorando (número do memo e processo)
    memo_num = None
    m_num = re.search(r"Memorando\s+([A-Za-z0-9_/\.-]+\s+n[ºo\.]*\s*\d+/\d+)", texto, re.IGNORECASE)
    if not m_num:
        m_num = re.search(r"(?:Memo|Memorando)[/:\s]+([A-Za-z0-9_/\.-]+(?:\s+n[ºo\.]*\s*\d+/\d+)?|\d+/\d+|\d+\.\d+)", texto, re.IGNORECASE)
    if m_num:
        memo_num = m_num.group(1).strip()

    proc_num = None
    m_proc = re.search(r"(?:processo(?:/documento)?|proc\.?)\s*(?:n[ºo\.]*\s*)?([A-Za-z0-9_/\.-]+(?:\s+\d+/\d+)?)", texto, re.IGNORECASE)
    if m_proc:
        proc_num = m_proc.group(1).strip()

    # Se o documento possuir tabela estruturada com dados, utiliza diretamente
    if caminho_pdf:
        itens_tab = extrair_tabelas_memorando(caminho_pdf, ano_ref=ano_ref, mes_ref=mes_ref)
        if itens_tab:
            return {
                "numero_memorando": memo_num,
                "processo": proc_num,
                "itens": itens_tab,
            }

    itens = []
    chaves_existentes = set()

    # Descarta rodapés e páginas de assinatura digital do SemPapel/SolarBPM
    texto_util = re.split(r"Assinaturas\s+do\s+documento|Pe[ç\ufffd\?a-zA-Z]?a\s+do\s+processo/documento", texto, flags=re.IGNORECASE)[0]

    # Limpa linhas de expediente (data de expedição do memorando não deve virar data de ocorrência)
    linhas = texto_util.split("\n")
    linhas_uteis = []
    for l in linhas:
        l_s = l.strip()
        if not l_s:
            continue
        # Ignora linhas de expedição de memorando (ex: "Data: 06/10/2026")
        if re.match(r"^Data:\s*\d{2}/\d{2}/\d{4}", l_s, re.I):
            continue
        if re.match(r"^Piracicaba,\s*\d{1,2}\s+de\s+[a-zç]+\s+de\s+\d{4}", l_s, re.I):
            continue
        # Ignora marcas verticais e rodapés do SemPapel
        if re.search(r"odanissa|otnemucoD|sotnemucoD|ogid[ó\xef\ufffd]?c|ossecorp|\.g[á\xef\ufffd]?P|epapmes", l_s, re.I):
            continue
        linhas_uteis.append(l_s)
    texto_filtrado = "\n".join(linhas_uteis)

    # Divide o texto em blocos por servidor (quebra por parágrafo, marcadores ou nova matrícula)
    blocos = re.split(r"\n\s*\n|(?=[\u2022\uf0b7]\s*)|(?=Onde\s+consta)|(?=[A-Z\s]{4,}\s*[-–]\s*\d{2}[\.,]\d{3})", texto_filtrado)
    
    # Dicionário de ocorrências conhecidas para busca flexível
    OCORRENCIAS_PADRAO = [
        "tratamento de saude",
        "auxilio doenca",
        "ferias regulamentares",
        "ferias premio",
        "abono",
        "abono eleitoral",
        "falta",
        "faltas",
        "faltas clt",
        "faltas efetivos",
        "minutos perdidos",
        "acidente de trabalho",
        "doenca em pessoa da familia",
        "licenca maternidade",
        "gestante",
        "doacao de sangue",
        "gala",
        "nojo",
        "frequencia normal",
        "aguardando pericia sempem",
        "aguardando retorno/pericia auxilio doenc",
    ]

    for bloco in blocos:
        bloco_limpo = " ".join(bloco.split())
        if not bloco_limpo:
            continue

        # Procura matrícula (desconsidera números de processo que antecedem)
        bloco_sem_proc = re.sub(r"processo\s+(?:202\d/)?\d+", "", bloco_limpo, flags=re.I)
        m_mat = re.search(r"\b(\d{2}[\.,]\d{3}[-_ ]\d|\d{3}[\.,]\d{3}|\d{5,6})\b", bloco_sem_proc)
        if not m_mat:
            continue

        mat_bruta = m_mat.group(1)
        mat_norm = normalizar_matricula(mat_bruta)

        # Formata para padrão XX.XXX-X se tiver 6 dígitos
        mat_formatada = f"{mat_norm[:2]}.{mat_norm[2:5]}-{mat_norm[5:]}" if len(mat_norm) == 6 else mat_bruta

        # Procura datas da ocorrência
        d_ini, d_fim, dias = _extrair_datas_ocorrencia(bloco_limpo, ano_ref, mes_ref=mes_ref)
        if not d_ini:
            continue

        # Procura nome do servidor
        nome = None
        padroes_nome = [
            r"sr[a]?\.\s*([A-ZÁÉÍÓÚÀÈÌÒÙÃÕÂÊÎÔÛÇ\s]{4,45}?),\s*(?:no\s+funcional|matr[íi]cula|n[ºo\.]*)",
            r"(?:servidor[a]?|ocupante[^,]+,)\s*(?:sr[a]?\.)?\s*([A-ZÁÉÍÓÚÀÈÌÒÙÃÕÂÊÎÔÛÇ\s]{4,45}?),\s*(?:no\s+funcional|matr[íi]cula|n[ºo\.]*)",
            r"([A-ZÁÉÍÓÚÀÈÌÒÙÃÕÂÊÎÔÛÇ\s]{4,45}?),\s*(?:no\s+funcional|matr[íi]cula|n[ºo\.]*)",
            r"\b\d{2}[\.,]\d{3}[-_ ]\d\s*[-–]\s*([A-ZÁÉÍÓÚÀÈÌÒÙÃÕÂÊÎÔÛÇ\s]{4,45}?)(?:[-–]|\s+\d{2}/\d{2}|\b)",
            r"([A-ZÁÉÍÓÚÀÈÌÒÙÃÕÂÊÎÔÛÇ\s]{4,45}?)\s*[-–]\s*\b\d{2}[\.,]\d{3}[-_ ]\d",
        ]
        for p_nome in padroes_nome:
            m_n = re.search(p_nome, bloco_limpo, re.IGNORECASE)
            if m_n:
                cand = " ".join(m_n.group(1).split()).strip()
                if len(cand) >= 4 and not any(k in cand.upper() for k in ["CARGO", "ASSESSOR", "ONDE", "CONSTA", "PROCESSO", "SERVI"]):
                    nome = cand.upper()
                    break

        # Ocorrência nova informada (De -> Para)
        nova_ocorrencia = None
        m_acao = re.search(r"(?:adicionar|alterar para|retificar para|lan[çc]ar|substituir por|trocar por|considerando:?\s*[-–]?)\s*([A-ZÁÉÍÓÚÀÈÌÒÙÃÕÂÊÎÔÛÇ\s/]+?)(?:\s+de\s+\d+|\s+do\s+dia|\s+no\s+per[íi]odo|\s+referente|\s*;|\s*\.)", bloco_limpo, re.IGNORECASE)
        if m_acao:
            cand_oco = " ".join(m_acao.group(1).split()).strip()
            cand_oco = re.sub(r"\b\d{2}/\d{2}(?:/\d{4})?\b", "", cand_oco).strip()
            if cand_oco:
                nova_ocorrencia = cand_oco
        else:
            # Fallback: procura se alguma ocorrência padrão foi mencionada
            norm_b = _normalizar_texto(bloco_limpo)
            for oc in OCORRENCIAS_PADRAO:
                if oc in norm_b:
                    nova_ocorrencia = oc.title()
                    break

        chave = (mat_norm, d_ini, d_fim)
        if chave not in chaves_existentes:
            chaves_existentes.add(chave)
            itens.append({
                "matricula": mat_formatada,
                "matricula_norm": mat_norm,
                "nome": nome,
                "data_inicio": d_ini,
                "data_fim": d_fim,
                "dias": dias,
                "ocorrencia": nova_ocorrencia.title() if nova_ocorrencia else "Ocorrência Retificada",
                "texto_original": bloco_limpo,
            })

    return {
        "numero_memorando": memo_num,
        "processo": proc_num,
        "itens": itens,
    }


def conferir_memorando_com_solicitacoes(dados_memorando, retificacoes_solicitadas):
    """
    Cruza os itens informados no memorando com as retificações que haviam sido solicitadas.
    Classifica em:
    - atendidas: o memorando retificou conforme solicitado.
    - divergencias_memo: servidor citado no memorando, mas com datas ou tipo inconsistentes.
    - pendentes_restantes: itens solicitados que não foram mencionados no memorando.
    - avulsas: itens no memorando que não estavam na lista de pendências.
    """
    itens_memo = dados_memorando.get("itens", [])
    
    # Mapeia solicitações por matrícula normalizada
    solicitacoes_por_mat = {}
    for r in retificacoes_solicitadas:
        m_norm = normalizar_matricula(r["matricula"])
        solicitacoes_por_mat.setdefault(m_norm, []).append(r)

    atendidas = []
    divergencias_memo = []
    avulsas = []
    solicitacoes_atendidas_ids = set()

    for item in itens_memo:
        m_norm = item["matricula_norm"]
        solics = solicitacoes_por_mat.get(m_norm, [])
        
        if not solics:
            avulsas.append(item)
            continue        # Procura correspondência de período / servidor
        encontrou = False
        for idx, sol in enumerate(solics):
            sol_id = (m_norm, sol["data_inicio"], sol["data_fim"])
            if sol_id in solicitacoes_atendidas_ids:
                continue

            # Verifica sobreposição de datas
            datas_coincidem = False
            if item.get("data_inicio") and item.get("data_fim") and sol.get("data_inicio") and sol.get("data_fim"):
                if (item["data_inicio"] <= sol["data_fim"]) and (item["data_fim"] >= sol["data_inicio"]):
                    datas_coincidem = True
                # Tolerância: mesmo dia inicial ou único apontamento do servidor no mês
                elif (item["data_inicio"].day == sol["data_inicio"].day) or (len(solics) == 1):
                    datas_coincidem = True

            # Compara ocorrência informada com o tipo do sistema / secretaria
            oc_memo_norm = _normalizar_texto(item["ocorrencia"])
            tipo_sis_norm = _normalizar_texto(sol.get("tipo_sistema", ""))
            tipo_sec_norm = _normalizar_texto(sol.get("tipo_secretaria", ""))

            # Tolerância de equivalência de tipos
            compativel = (
                (oc_memo_norm == tipo_sis_norm) or
                ("saude" in oc_memo_norm and "saude" in tipo_sis_norm) or
                ("medico" in oc_memo_norm and "saude" in tipo_sis_norm) or
                ("ferias" in oc_memo_norm and "ferias" in tipo_sis_norm) or
                ("falta" in oc_memo_norm and "falta" in tipo_sis_norm) or
                ("abono" in oc_memo_norm and "abono" in tipo_sis_norm) or
                ("nojo" in oc_memo_norm and "nojo" in tipo_sis_norm) or
                (oc_memo_norm in tipo_sis_norm or tipo_sis_norm in oc_memo_norm) or
                (oc_memo_norm == tipo_sec_norm)
            )

            datas_exatas = (item["data_inicio"] == sol["data_inicio"] and item["data_fim"] == sol["data_fim"])
            abrange_periodo = (item.get("data_inicio") and item.get("data_fim") and item["data_inicio"] <= sol["data_inicio"] and item["data_fim"] >= sol["data_fim"])

            # O memorando de retificação se sobrepõe ao enviado inicialmente e justifica a ocorrência
            if compativel or datas_coincidem or len(solics) == 1:
                dt_memo_str = (
                    f"{item['data_inicio'].strftime('%d/%m/%Y')} a {item['data_fim'].strftime('%d/%m/%Y')}"
                    if item.get("data_inicio") and item["data_inicio"] != item["data_fim"]
                    else (item["data_inicio"].strftime("%d/%m/%Y") if item.get("data_inicio") else "")
                )

                if datas_exatas and compativel:
                    status_desc = "Atendido conforme solicitado"
                elif abrange_periodo and compativel:
                    status_desc = f"Regularizado por memorando de retificação (Período no memo: {dt_memo_str})"
                elif datas_coincidem and compativel:
                    status_desc = f"Regularizado por memorando de retificação (Período retificado no memo: {dt_memo_str})"
                else:
                    detalhes = []
                    if dt_memo_str:
                        detalhes.append(f"Período no memo: {dt_memo_str}")
                    if not compativel:
                        detalhes.append(f"Ocorrência retificada: {item['ocorrencia']}")
                    det_txt = f" ({'; '.join(detalhes)})" if detalhes else ""
                    status_desc = f"Regularizado por memorando de retificação{det_txt}"

                atendidas.append({
                    "matricula": sol["matricula"],
                    "nome": sol["nome"] or item["nome"],
                    "data_inicio": sol["data_inicio"],
                    "data_fim": sol["data_fim"],
                    "dias": sol["dias"],
                    "tipo_retificado": item["ocorrencia"],
                    "data_memo_inicio": item.get("data_inicio"),
                    "data_memo_fim": item.get("data_fim"),
                    "status": status_desc,
                })
                solicitacoes_atendidas_ids.add(sol_id)
                encontrou = True
                break

        if not encontrou:
            avulsas.append(item)

    # Pendências que não foram mencionadas no memorando
    pendentes_restantes = []
    for r in retificacoes_solicitadas:
        m_norm = normalizar_matricula(r["matricula"])
        sol_id = (m_norm, r["data_inicio"], r["data_fim"])
        if sol_id not in solicitacoes_atendidas_ids:
            pendentes_restantes.append(r)

    return {
        "atendidas": atendidas,
        "divergencias_memo": divergencias_memo,
        "pendentes_restantes": pendentes_restantes,
        "avulsas": avulsas,
    }


def gerar_secao_memorando_markdown(resultado_conferencia, dados_memo, nome_arquivo_memo):
    """Gera o bloco Markdown para ser anexado ao final de retificacoes_<secretaria>.md."""
    linhas = []
    linhas.append("---")
    linhas.append("")
    linhas.append("## 📑 Conferência do Memorando de Retificação Recebido")
    
    info_cabecalho = []
    info_cabecalho.append(f"- **Arquivo analisado**: `{nome_arquivo_memo}`")
    if dados_memo.get("numero_memorando"):
        info_cabecalho.append(f"- **Documento**: {dados_memo['numero_memorando']}")
    if dados_memo.get("processo"):
        info_cabecalho.append(f"- **Processo**: {dados_memo['processo']}")
    info_cabecalho.append(f"- **Data da conferência**: {datetime.now().strftime('%d/%m/%Y %H:%M')}")
    
    linhas.extend(info_cabecalho)
    linhas.append("")

    atendidas = resultado_conferencia.get("atendidas", [])
    divergencias = resultado_conferencia.get("divergencias_memo", [])
    pendentes = resultado_conferencia.get("pendentes_restantes", [])
    avulsas = resultado_conferencia.get("avulsas", [])

    # 1. Atendidas
    linhas.append("### ✅ Ocorrências Regularizadas pelo Memorando")
    if atendidas:
        for a in atendidas:
            dt_txt = f"{a['data_inicio'].strftime('%d/%m/%Y')} a {a['data_fim'].strftime('%d/%m/%Y')}" if a['data_inicio'] != a['data_fim'] else a['data_inicio'].strftime('%d/%m/%Y')
            linhas.append(f"- **{a['matricula']} - {a['nome']}** ({dt_txt} - {a['dias']} dia{'s' if a['dias'] != 1 else ''}): Retificado para **{a['tipo_retificado']}** ({a['status']}).")
    else:
        linhas.append("*(Nenhuma ocorrência regularizada identificada no documento)*")
    linhas.append("")

    # 2. Inconsistências / Divergências no Memorando
    if divergencias:
        linhas.append("### ⚠️ Divergências / Inconsistências Identificadas no Memorando")
        linhas.append("As seguintes informações enviadas no memorando apresentam divergência em relação ao sistema ou à solicitação:")
        for d in divergencias:
            linhas.append(f"- **{d['matricula']} - {d['nome']}**:")
            linhas.append(f"  - *Solicitado*: {d['solicitado']}")
            linhas.append(f"  - *Informado no Memorando*: {d['informado_memo']}")
            linhas.append(f"  - *Inconsistência*: {d['motivo']}")
        linhas.append("")

    # 3. Pendências Restantes
    linhas.append("### ❌ Pendências Restantes (Não Mencionadas no Memorando)")
    # Filtra apenas divergências reais que não foram atendidas (ignora sem registro e ocorrências a inserir no sistema)
    pend_reais = [
        p for p in pendentes
        if p.get("tipo_secretaria") != "SEM REGISTRO"
        and not (("minuto" in str(p.get("tipo_secretaria")).lower() or "falta" in str(p.get("tipo_secretaria")).lower()) and p.get("tipo_sistema") == "sem registro em sistema")
    ]
    if pend_reais:
        linhas.append("Os seguintes apontamentos continuam pendentes de regularização:")
        for p in pend_reais:
            dt_txt = f"{p['data_inicio'].strftime('%d/%m/%Y')} a {p['data_fim'].strftime('%d/%m/%Y')}" if p['data_inicio'] != p['data_fim'] else p['data_inicio'].strftime('%d/%m/%Y')
            dias_val = int(p["dias"]) if isinstance(p.get("dias"), (int, float)) and int(p["dias"]) == p["dias"] else p.get("dias", 1)
            linhas.append(f"- {p['matricula']} - {p['nome']} - {p['tipo_secretaria']} --> {p['tipo_sistema']} - {dt_txt} ({dias_val} dia{'s' if dias_val != 1 else ''})")
    else:
        linhas.append("*(Todas as retificações solicitadas à secretaria foram atendidas pelo memorando!)*")
    linhas.append("")

    # 4. Retificações Avulsas / Extras informadas no memorando
    if avulsas:
        linhas.append("### ℹ️ Informações Adicionais / Outros Servidores no Memorando")
        for av in avulsas:
            dt_txt = f"{av['data_inicio'].strftime('%d/%m/%Y')} a {av['data_fim'].strftime('%d/%m/%Y')}" if av['data_inicio'] != av['data_fim'] else av['data_inicio'].strftime('%d/%m/%Y')
            nome_str = f" - {av['nome']}" if av.get("nome") else ""
            linhas.append(f"- {av['matricula']}{nome_str} - {dt_txt} - {av['dias']} dia{'s' if av['dias'] != 1 else ''}: {av['ocorrencia']}")
        linhas.append("")

    return "\n".join(linhas)
