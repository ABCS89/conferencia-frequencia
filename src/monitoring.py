# -*- coding: utf-8 -*-
"""
monitoring.py - Monitoramento de Atrasos (Minutos Perdidos) e Faltas Acumuladas

Gera relatórios dedicados por secretaria para acompanhamento e cobrança
de justificativas para servidores com atrasos acumulados no mês e faltas.
"""
import os
import unicodedata
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

COR_CABECALHO = "1F4E78"
COR_ATRASO = "FFF2CC"      # amarelo suave para atrasos/minutos perdidos
COR_FALTA = "FCE4D6"       # laranja suave para faltas acumuladas
COR_REGULARIZADO = "E2EFDA" # verde suave para ocorrências regularizadas por memorando


def aplicar_retificacoes_memorando(regs_sec, atendidas, nome_memorando=""):
    """
    Substitui/baixa as ocorrências da secretaria que foram regularizadas pelo memorando.
    Ex: as 14 faltas do José Leandro viram Tratamento de Saúde, saindo das faltas pendentes.
    Retorna uma nova lista de registros atualizados e a lista de regularizações aplicadas.
    """
    if not atendidas:
        return [dict(r) for r in regs_sec], []

    import re
    from datetime import datetime

    def _norm_mat(m):
        return re.sub(r"\D", "", str(m or ""))

    atendidas_por_mat = {}
    for at in atendidas:
        m_norm = _norm_mat(at.get("matricula"))
        atendidas_por_mat.setdefault(m_norm, []).append(at)

    regs_atualizados = []
    regularizacoes_aplicadas = []

    for r in regs_sec:
        r_copia = dict(r)
        m_norm = _norm_mat(r_copia.get("matricula"))
        if m_norm not in atendidas_por_mat:
            regs_atualizados.append(r_copia)
            continue

        data_str = r_copia.get("data")
        data_reg = None
        if data_str:
            try:
                partes = data_str.split("/")
                if len(partes) == 3:
                    data_reg = datetime(int(partes[2]), int(partes[1]), int(partes[0])).date()
            except Exception:
                data_reg = None

        foi_regularizado = False
        for at in atendidas_por_mat[m_norm]:
            coincide = False
            if data_reg and at.get("data_inicio") and at.get("data_fim"):
                if at["data_inicio"] <= data_reg <= at["data_fim"]:
                    coincide = True
            elif not data_reg:
                coincide = True

            if coincide:
                oc_antiga = r_copia.get("ocorrencia") or ""
                nova_oc = at.get("tipo_retificado") or "Regularizado"
                r_copia["ocorrencia_original"] = oc_antiga
                r_copia["ocorrencia"] = nova_oc
                r_copia["regularizado_memorando"] = True
                r_copia["memorando_origem"] = nome_memorando

                regularizacoes_aplicadas.append({
                    "matricula": r_copia["matricula"],
                    "nome": r_copia["nome"],
                    "data": data_str,
                    "ocorrencia_anterior": oc_antiga,
                    "ocorrencia_nova": nova_oc,
                    "memorando": nome_memorando,
                    "status": "Regularizado pelo Memorando",
                })
                foi_regularizado = True
                break

        regs_atualizados.append(r_copia)

    return regs_atualizados, regularizacoes_aplicadas


def _normalizar(texto):
    if not texto:
        return ""
    nfkd = unicodedata.normalize("NFKD", texto)
    sem_acento = "".join(c for c in nfkd if not unicodedata.combining(c))
    return " ".join(sem_acento.lower().split())


def _formatar_minutos(minutos):
    """Converte minutos inteiros em formato 'Xh YYmin' e 'X min'."""
    try:
        m = int(round(float(minutos)))
    except (ValueError, TypeError):
        return "0 min", 0
    horas = m // 60
    resto = m % 60
    if horas > 0 and resto > 0:
        texto_h = f"{horas}h {resto}min"
    elif horas > 0:
        texto_h = f"{horas}h"
    else:
        texto_h = f"{resto}min"
    return texto_h, m


def apurar_dados_monitoramento(regs_sec, regs_sis, regularizados=None):
    """Apura minutos perdidos e faltas acumuladas no mês por servidor."""
    nomes = {}
    atrasos_sec = {}
    faltas_sec = {}

    # 1. Processar dados da Secretaria
    for r in regs_sec:
        mat = r["matricula"]
        nome = " ".join(r["nome"].split())
        nomes[mat] = nome
        norm = _normalizar(r.get("ocorrencia") or "")

        # Atrasos / Minutos Perdidos
        if "minut" in norm or "perd" in norm:
            qtd = r.get("qtde_dias") or 0
            try:
                qtd_min = float(qtd)
            except (ValueError, TypeError):
                qtd_min = 0.0
            data_str = r.get("data") or "Não informada"
            item = atrasos_sec.setdefault(mat, {
                "matricula": mat,
                "nome": nome,
                "datas": [],
                "total_minutos": 0.0,
                "ocorrencias": [],
            })
            if data_str not in item["datas"]:
                item["datas"].append(data_str)
            item["total_minutos"] += qtd_min
            item["ocorrencias"].append(f"{data_str} ({int(round(qtd_min))} min)")

        # Faltas
        elif "falta" in norm:
            qtd = r.get("qtde_dias") or 1.0
            try:
                qtd_dias = float(qtd)
            except (ValueError, TypeError):
                qtd_dias = 1.0
            data_str = r.get("data") or "Não informada"
            tipo_falta = r.get("ocorrencia") or "Falta"
            item = faltas_sec.setdefault(mat, {
                "matricula": mat,
                "nome": nome,
                "datas": [],
                "total_dias": 0.0,
                "tipos": set(),
                "ocorrencias": [],
            })
            item["total_dias"] += qtd_dias
            item["tipos"].add(tipo_falta)
            dias_i = int(round(qtd_dias))
            item["ocorrencias"].append(f"{data_str} ({dias_i} dia{'s' if dias_i != 1 else ''})")

    # 2. Processar registros do Sistema para conferência
    sis_minutos_map = {}
    sis_faltas_map = {}
    for r in regs_sis:
        mat = r["matricula"]
        desc = r.get("descricao") or ""
        norm = _normalizar(desc)
        if "minut" in norm or "perd" in norm:
            sis_minutos_map.setdefault(mat, []).append(desc)
        elif "falta" in norm:
            sis_faltas_map.setdefault(mat, []).append(desc)

    # 3. Consolidar Atrasos (Minutos Perdidos)
    lista_atrasos = []
    for mat, info in atrasos_sec.items():
        min_total = info["total_minutos"]
        tempo_str, min_int = _formatar_minutos(min_total)
        status_sis = "Lançado no sistema" if mat in sis_minutos_map else "Sem registro no sistema"
        lista_atrasos.append({
            "matricula": mat,
            "nome": info["nome"],
            "datas": ", ".join(info["datas"]),
            "minutos": min_int,
            "tempo_formatado": tempo_str,
            "detalhes": "; ".join(info["ocorrencias"]),
            "status_sistema": status_sis,
        })
    lista_atrasos.sort(key=lambda x: x["minutos"], reverse=True)

    # 4. Consolidar Faltas
    lista_faltas = []
    for mat, info in faltas_sec.items():
        dias_total = info["total_dias"]
        status_sis = "Lançado no sistema" if mat in sis_faltas_map else "Sem registro no sistema"
        lista_faltas.append({
            "matricula": mat,
            "nome": info["nome"],
            "dias": int(round(dias_total)),
            "tipos": ", ".join(sorted(info["tipos"])),
            "detalhes": "; ".join(info["ocorrencias"]),
            "status_sistema": status_sis,
        })
    lista_faltas.sort(key=lambda x: x["dias"], reverse=True)

    total_minutos_orgao = sum(a["minutos"] for a in lista_atrasos)
    total_tempo_orgao, _ = _formatar_minutos(total_minutos_orgao)
    total_dias_faltas_orgao = sum(f["dias"] for f in lista_faltas)

    resumo = {
        "total_servidores_atrasos": len(lista_atrasos),
        "total_minutos": total_minutos_orgao,
        "total_tempo_formatado": total_tempo_orgao,
        "total_servidores_faltas": len(lista_faltas),
        "total_dias_faltas": total_dias_faltas_orgao,
        "total_regularizados": len(regularizados or []),
    }

    return {
        "resumo": resumo,
        "atrasos": lista_atrasos,
        "faltas": lista_faltas,
        "regularizados": regularizados or [],
    }


def gerar_markdown_monitoramento(dados, caminho_md, mes_label, nome_orgao=""):
    linhas = []
    titulo = f"# Relatório de Monitoramento — Atrasos e Faltas Acumuladas"
    subtitulo = f"### Órgão / Secretaria: {nome_orgao} ({mes_label})" if nome_orgao else f"### Referência: {mes_label}"
    linhas.append(titulo)
    linhas.append(subtitulo)
    linhas.append("")
    linhas.append("> [!NOTE]")
    linhas.append("> Este relatório consolida os atrasos acumulados (minutos perdidos) e faltas apuradas no mês para fins de monitoramento, auditoria interna e posterior cobrança de justificativas.")
    linhas.append("")

    res = dados["resumo"]
    linhas.append("## Resumo Geral do Órgão")
    linhas.append(f"- **Servidores com atrasos (minutos perdidos)**: {res['total_servidores_atrasos']}")
    linhas.append(f"- **Volume total de atrasos no mês**: {res['total_minutos']} minutos ({res['total_tempo_formatado']})")
    linhas.append(f"- **Servidores com faltas no mês**: {res['total_servidores_faltas']}")
    linhas.append(f"- **Total de faltas acumuladas no órgão**: {res['total_dias_faltas']} dia(s)")
    if res.get("total_regularizados", 0) > 0:
        linhas.append(f"- **Ocorrências baixadas/regularizadas via memorando**: {res['total_regularizados']}")
    linhas.append("")

    # Seção de Regularizados
    if dados.get("regularizados"):
        linhas.append("## ✅ Ocorrências Regularizadas por Memorando")
        linhas.append("Ocorrências que foram baixadas e substituídas após conferência do memorando formal da secretaria:")
        linhas.append("")
        linhas.append("| Matrícula | Nome do Servidor | Ocorrência Anterior | Retificado Para | Data / Período | Memorando | Situação |")
        linhas.append("| :--- | :--- | :---: | :---: | :---: | :---: | :---: |")
        for reg in dados["regularizados"]:
            linhas.append(f"| `{reg['matricula']}` | {reg['nome']} | ~~{reg['ocorrencia_anterior']}~~ | **{reg['ocorrencia_nova']}** | {reg['data']} | {reg['memorando']} | {reg['status']} |")
        linhas.append("")

    # Seção de Atrasos
    linhas.append("## ⏱️ Painel de Atrasos (Minutos Perdidos Acumulados)")
    linhas.append("Servidores ordenados pelo maior volume de minutos perdidos acumulados no mês:")
    linhas.append("")
    if dados["atrasos"]:
        linhas.append("| Matrícula | Nome do Servidor | Minutos Acumulados | Equivalência | Datas / Ocorrências | Situação no Sistema |")
        linhas.append("| :--- | :--- | :---: | :---: | :--- | :---: |")
        for a in dados["atrasos"]:
            linhas.append(f"| `{a['matricula']}` | {a['nome']} | **{a['minutos']} min** | **{a['tempo_formatado']}** | {a['detalhes']} | {a['status_sistema']} |")
    else:
        linhas.append("*Nenhum atraso (minutos perdidos) registrado para esta secretaria no mês.*")
    linhas.append("")

    # Seção de Faltas
    linhas.append("## ❌ Painel de Faltas Acumuladas")
    linhas.append("Servidores ordenados pelo maior volume de faltas acumuladas no mês:")
    linhas.append("")
    if dados["faltas"]:
        linhas.append("| Matrícula | Nome do Servidor | Dias de Falta | Tipo(s) de Falta | Datas / Períodos | Situação no Sistema |")
        linhas.append("| :--- | :--- | :---: | :--- | :--- | :---: |")
        for f in dados["faltas"]:
            linhas.append(f"| `{f['matricula']}` | {f['nome']} | **{f['dias']} dia(s)** | {f['tipos']} | {f['detalhes']} | {f['status_sistema']} |")
    else:
        linhas.append("*Nenhuma falta registrada para esta secretaria no mês.*")
    linhas.append("")

    os.makedirs(os.path.dirname(caminho_md), exist_ok=True)
    with open(caminho_md, "w", encoding="utf-8") as f:
        f.write("\n".join(linhas))


def gerar_excel_monitoramento(dados, caminho_xlsx, mes_label, nome_orgao=""):
    wb = Workbook()

    # 1. Aba Resumo
    ws_res = wb.active
    ws_res.title = "Resumo"
    ws_res.append([f"Painel de Monitoramento — {nome_orgao} ({mes_label})"])
    ws_res["A1"].font = Font(bold=True, size=14)
    ws_res.append([])
    res = dados["resumo"]
    ws_res.append(["Indicador", "Valor"])
    ws_res["A3"].font = Font(bold=True, color="FFFFFF")
    ws_res["A3"].fill = PatternFill("solid", fgColor=COR_CABECALHO)
    ws_res["B3"].font = Font(bold=True, color="FFFFFF")
    ws_res["B3"].fill = PatternFill("solid", fgColor=COR_CABECALHO)

    ws_res.append(["Servidores com atrasos (minutos perdidos)", res["total_servidores_atrasos"]])
    ws_res.append(["Total de minutos perdidos acumulados", f"{res['total_minutos']} min ({res['total_tempo_formatado']})"])
    ws_res.append(["Servidores com faltas registradas", res["total_servidores_faltas"]])
    ws_res.append(["Total de faltas acumuladas (dias)", f"{res['total_dias_faltas']} dia(s)"])
    if res.get("total_regularizados", 0) > 0:
        ws_res.append(["Ocorrências baixadas/regularizadas via memorando", res["total_regularizados"]])

    ws_res.column_dimensions["A"].width = 45
    ws_res.column_dimensions["B"].width = 30

    # 2. Aba Atrasos
    ws_atr = wb.create_sheet("Atrasos (Minutos)")
    cols_atr = ["Matrícula", "Nome", "Minutos Acumulados", "Equivalência", "Datas / Ocorrências", "Situação no Sistema"]
    ws_atr.append(cols_atr)
    for cel in ws_atr[1]:
        cel.font = Font(bold=True, color="FFFFFF")
        cel.fill = PatternFill("solid", fgColor=COR_CABECALHO)
        cel.alignment = Alignment(horizontal="center")

    for a in dados["atrasos"]:
        ws_atr.append([
            a["matricula"], a["nome"], a["minutos"], a["tempo_formatado"],
            a["detalhes"], a["status_sistema"]
        ])
        for cel in ws_atr[ws_atr.max_row]:
            cel.fill = PatternFill("solid", fgColor=COR_ATRASO)

    larg_atr = [14, 36, 20, 18, 45, 25]
    for i, w in enumerate(larg_atr, start=1):
        ws_atr.column_dimensions[get_column_letter(i)].width = w
    ws_atr.freeze_panes = "A2"

    # 3. Aba Faltas
    ws_fal = wb.create_sheet("Faltas Acumuladas")
    cols_fal = ["Matrícula", "Nome", "Dias de Falta", "Tipo(s) de Falta", "Datas / Períodos", "Situação no Sistema"]
    ws_fal.append(cols_fal)
    for cel in ws_fal[1]:
        cel.font = Font(bold=True, color="FFFFFF")
        cel.fill = PatternFill("solid", fgColor=COR_CABECALHO)
        cel.alignment = Alignment(horizontal="center")

    for f in dados["faltas"]:
        ws_fal.append([
            f["matricula"], f["nome"], f["dias"], f["tipos"],
            f["detalhes"], f["status_sistema"]
        ])
        for cel in ws_fal[ws_fal.max_row]:
            cel.fill = PatternFill("solid", fgColor=COR_FALTA)

    larg_fal = [14, 36, 16, 25, 45, 25]
    for i, w in enumerate(larg_fal, start=1):
        ws_fal.column_dimensions[get_column_letter(i)].width = w
    ws_fal.freeze_panes = "A2"

    # 4. Aba Regularizados (se houver)
    if dados.get("regularizados"):
        ws_reg = wb.create_sheet("Regularizados (Memorando)")
        cols_reg = ["Matrícula", "Nome", "Ocorrência Anterior", "Retificado Para", "Data / Período", "Memorando", "Situação no Sistema"]
        ws_reg.append(cols_reg)
        for cel in ws_reg[1]:
            cel.font = Font(bold=True, color="FFFFFF")
            cel.fill = PatternFill("solid", fgColor=COR_CABECALHO)
            cel.alignment = Alignment(horizontal="center")

        for r_item in dados["regularizados"]:
            ws_reg.append([
                r_item["matricula"], r_item["nome"], r_item["ocorrencia_anterior"],
                r_item["ocorrencia_nova"], r_item["data"], r_item["memorando"], r_item["status"]
            ])
            for cel in ws_reg[ws_reg.max_row]:
                cel.fill = PatternFill("solid", fgColor=COR_REGULARIZADO)

        larg_reg = [14, 36, 22, 25, 18, 25, 25]
        for i, w in enumerate(larg_reg, start=1):
            ws_reg.column_dimensions[get_column_letter(i)].width = w
        ws_reg.freeze_panes = "A2"

    os.makedirs(os.path.dirname(caminho_xlsx), exist_ok=True)
    wb.save(caminho_xlsx)
