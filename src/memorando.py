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


def extrair_tabelas_memorando(caminho_pdf):
    """Extrai itens estruturados de tabelas em PDFs de memorando (suporta células mescladas verticalmente)."""
    itens = []
    if not caminho_pdf:
        return itens
    try:
        with pdfplumber.open(str(caminho_pdf)) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
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
                        datas = re.findall(r"(\d{2}/\d{2}/\d{4})", per_str)
                        if not datas:
                            continue
                        d_ini = _parse_data(datas[0])
                        d_fim = _parse_data(datas[1]) if len(datas) > 1 else d_ini

                        if ult_mat:
                            m_norm = normalizar_matricula(ult_mat)
                            mat_fmt = f"{m_norm[:2]}.{m_norm[2:5]}-{m_norm[5:]}" if len(m_norm) == 6 else ult_mat
                            dias = (d_fim - d_ini).days + 1 if d_ini and d_fim else 1
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
    itens = []

    # Extrai primeiro por tabelas se o caminho do PDF estiver disponível
    if caminho_pdf:
        itens_tab = extrair_tabelas_memorando(caminho_pdf)
        if itens_tab:
            itens.extend(itens_tab)

    chaves_existentes = {(it["matricula_norm"], it["data_inicio"], it["data_fim"]) for it in itens}

    # 1. Identificação do cabeçalho do memorando (número do memo e processo)
    memo_num = None
    m_num = re.search(r"Memorando\s+([A-Za-z0-9_/\.-]+\s+n[ºo\.]*\s*\d+/\d+)", texto, re.IGNORECASE)
    if not m_num:
        m_num = re.search(r"(?:Memo|Memorando)\s*[:\s]*([A-Za-z0-9_/\.-]+\d+/\d+|\d+/\d+|\d+\.\d+)", texto, re.IGNORECASE)
    if m_num:
        memo_num = m_num.group(1).strip()

    proc_num = None
    m_proc = re.search(r"Processo\s+(?:n[ºo\.]*\s*)?([A-Za-z0-9_/\.-]+)", texto, re.IGNORECASE)
    if m_proc:
        proc_num = m_proc.group(1).strip()

    # Descarta rodapés e páginas de assinatura digital do SemPapel/SolarBPM
    texto_util = re.split(r"Assinaturas\s+do\s+documento|Peça\s+do\s+processo/documento", texto, flags=re.IGNORECASE)[0]

    # Divide o texto em parágrafos e linhas
    blocos = re.split(r"\n\s*\n|(?=Onde\s+consta)|(?=[A-Z\s]{4,}\s*[-–]\s*\d{2}[\.,]\d{3})", texto_util)
    
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

        # Procura matrícula
        # Padrões: 913.286, 91.328-6, 18.290-7, 182907
        m_mat = re.search(r"\b(\d{2}[\.,]\d{3}[-_ ]\d|\d{3}[\.,]\d{3}|\d{5,6})\b", bloco_limpo)
        if not m_mat:
            continue

        mat_bruta = m_mat.group(1)
        mat_norm = normalizar_matricula(mat_bruta)

        # Formata para padrão XX.XXX-X se tiver 6 dígitos
        mat_formatada = f"{mat_norm[:2]}.{mat_norm[2:5]}-{mat_norm[5:]}" if len(mat_norm) == 6 else mat_bruta

        # Procura datas (ex.: 07/08/2026 ao 21/08/2026, de 07/08/2026 a 21/08/2026, 07/08/2026)
        datas = re.findall(r"(\d{2}/\d{2}/\d{4})", bloco_limpo)
        if not datas:
            continue

        d_ini = _parse_data(datas[0])
        d_fim = _parse_data(datas[1]) if len(datas) > 1 else d_ini

        # Quantidade de dias
        dias = None
        m_dias = re.search(r"\b(\d+)\s*dias?\b", bloco_limpo, re.IGNORECASE)
        if m_dias:
            dias = int(m_dias.group(1))
        elif d_ini and d_fim:
            dias = (d_fim - d_ini).days + 1

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
        m_acao = re.search(r"(?:adicionar|alterar para|retificar para|lan[çc]ar|substituir por|trocar por)\s+([A-ZÁÉÍÓÚÀÈÌÒÙÃÕÂÊÎÔÛÇ\s/]+?)(?:\s+de\s+\d+|\s+do\s+dia|\s+no\s+per[íi]odo|\s+referente|\s*;|\s*\.)", bloco_limpo, re.IGNORECASE)
        if m_acao:
            nova_ocorrencia = " ".join(m_acao.group(1).split()).strip()
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
            continue

        # Procura correspondência de período
        encontrou = False
        for idx, sol in enumerate(solics):
            sol_id = (m_norm, sol["data_inicio"], sol["data_fim"])
            
            # Verifica sobreposição de datas
            datas_coincidem = False
            if item["data_inicio"] and item["data_fim"]:
                if (item["data_inicio"] <= sol["data_fim"]) and (item["data_fim"] >= sol["data_inicio"]):
                    datas_coincidem = True

            if datas_coincidem:
                # Compara ocorrência informada com o tipo do sistema
                oc_memo_norm = _normalizar_texto(item["ocorrencia"])
                tipo_sis_norm = _normalizar_texto(sol.get("tipo_sistema", ""))
                
                # Tolerância de equivalência de tipos
                compativel = (oc_memo_norm == tipo_sis_norm) or \
                             ("saude" in oc_memo_norm and "saude" in tipo_sis_norm) or \
                             ("ferias" in oc_memo_norm and "ferias" in tipo_sis_norm) or \
                             ("falta" in oc_memo_norm and "falta" in tipo_sis_norm) or \
                             ("abono" in oc_memo_norm and "abono" in tipo_sis_norm)

                datas_exatas = (item["data_inicio"] == sol["data_inicio"] and item["data_fim"] == sol["data_fim"])
                abrange_periodo = (item["data_inicio"] <= sol["data_inicio"] and item["data_fim"] >= sol["data_fim"])

                if compativel and (datas_exatas or abrange_periodo):
                    status_desc = "Atendido conforme solicitado" if datas_exatas else f"Atendido (Atestado do memo abrange o período: {item['data_inicio'].strftime('%d/%m/%Y')} a {item['data_fim'].strftime('%d/%m/%Y')})"
                    atendidas.append({
                        "matricula": sol["matricula"],
                        "nome": sol["nome"] or item["nome"],
                        "data_inicio": sol["data_inicio"],
                        "data_fim": sol["data_fim"],
                        "dias": sol["dias"],
                        "tipo_retificado": item["ocorrencia"],
                        "status": status_desc,
                    })
                    solicitacoes_atendidas_ids.add(sol_id)
                    encontrou = True
                    break
                else:
                    motivos = []
                    if not (datas_exatas or abrange_periodo):
                        motivos.append(f"Datas informadas no memo ({item['data_inicio'].strftime('%d/%m/%Y')} a {item['data_fim'].strftime('%d/%m/%Y')}) divergem da solicitação ({sol['data_inicio'].strftime('%d/%m/%Y')} a {sol['data_fim'].strftime('%d/%m/%Y')})")
                    if not compativel:
                        motivos.append(f"Ocorrência informada no memo ({item['ocorrencia']}) difere do sistema ({sol.get('tipo_sistema')})")

                    divergencias_memo.append({
                        "matricula": sol["matricula"],
                        "nome": sol["nome"] or item["nome"],
                        "solicitado": f"{sol['data_inicio'].strftime('%d/%m/%Y')} a {sol['data_fim'].strftime('%d/%m/%Y')} ({sol['tipo_secretaria']} --> {sol['tipo_sistema']})",
                        "informado_memo": f"{item['data_inicio'].strftime('%d/%m/%Y')} a {item['data_fim'].strftime('%d/%m/%Y')} ({item['ocorrencia']})",
                        "motivo": "; ".join(motivos),
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
            linhas.append(f"- {p['matricula']} - {p['nome']} - {dt_txt} - {p['dias']} dia{'s' if p['dias'] != 1 else ''} - {p['tipo_secretaria']} --> {p['tipo_sistema']}")
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
