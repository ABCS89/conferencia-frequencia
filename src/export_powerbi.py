# -*- coding: utf-8 -*-
"""
export_powerbi.py - Consolidador de Dados de Monitoramento para o Power BI

Varre todas as pastas mensais em output/*/monitoramento/*.xlsx e gera
uma base de dados consolidada, padronizada e relacional, pronta para
importação direta no Power BI ou ferramentas de BI.

Saída em:
  - output/powerbi/monitoramento_consolidado.xlsx (com abas: Atrasos, Faltas, Resumo_Secretarias)
  - output/powerbi/atrasos.csv
  - output/powerbi/faltas.csv
  - output/powerbi/resumo_secretarias.csv
"""
import re
from pathlib import Path
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

BASE_DIR = Path(__file__).resolve().parent.parent
OUTPUT_DIR = BASE_DIR / "output"
POWERBI_DIR = OUTPUT_DIR / "powerbi"


def extrair_codigo_e_nome(nome_arquivo):
    """Extrai código e nome da secretaria do nome do arquivo (ex: 'monitoramento_103 - PROCURADORIA GERAL.xlsx')."""
    stem = Path(nome_arquivo).stem
    stem = re.sub(r"^monitoramento_", "", stem, flags=re.IGNORECASE)
    m = re.match(r"^(\d{3})\s*-\s*(.+)$", stem)
    if m:
        return m.group(1).strip(), m.group(2).strip()
    m_num = re.match(r"^(\d{3})\b", stem)
    if m_num:
        return m_num.group(1).strip(), stem
    return "", stem


# Mapa de siglas/nomes curtos para exibição elegante no Power BI
MAPA_SIGLAS_SECRETARIAS = {
    "101": "GABINETE",
    "102": "SEMAD",
    "103": "PGM",
    "104": "SMF",
    "106": "SEPLAG",
    "107": "Educação",     # Ajustado conforme solicitado
    "108": "OBRAS",
    "109": "SMADS",
    "110": "AGRIMA",       # Agricultura e Meio Ambiente
    "112": "SEMAC",
    "113": "SEMIC",
    "114": "Saúde",        # Ajustado conforme solicitado
    "116": "GCMP",        # Guarda Civil
    "119": "SELAM",       # Esportes e Lazer
    "120": "SETUR",
    "121": "EMDHAP",
    "122": "SEGOV",
    "123": "IPASP",
    "124": "SEGTRANS",    # Ajustado conforme solicitado (Trânsito/Transportes)
    "125": "SEMA",
}


def obter_sigla_secretaria(codigo, nome):
    """Retorna a sigla/nome curto amigável da secretaria pelo código ou palavra-chave."""
    cod = str(codigo).strip()
    if cod in MAPA_SIGLAS_SECRETARIAS:
        return MAPA_SIGLAS_SECRETARIAS[cod]

    nome_upper = (nome or "").upper()
    if "AGRICULTURA" in nome_upper or "MEIO AMBIENTE" in nome_upper:
        return "AGRIMA"
    if "EDUCA" in nome_upper:
        return "Educação"
    if "PROCURADORIA" in nome_upper:
        return "PGM"
    if "ASSIST" in nome_upper or "DESENVOLVIMENTO SOCIAL" in nome_upper:
        return "SMADS"
    if "CULTURA" in nome_upper:
        return "SEMAC"
    if "GUARDA" in nome_upper:
        return "GCMP"
    if "ESPORTE" in nome_upper or "LAZER" in nome_upper:
        return "SELAM"
    if "SEGURAN" in nome_upper or "TRANSIT" in nome_upper or "TRANSP" in nome_upper:
        return "SEGTRANS"
    if "ADMINISTRA" in nome_upper:
        return "SEMAD"
    if "FINAN" in nome_upper:
        return "SMF"
    if "SAUDE" in nome_upper or cod == "114":
        return "Saúde"
    if "OBRAS" in nome_upper:
        return "OBRAS"

    return cod if cod else nome


def consolidar_monitoramento():
    """Varre todas as pastas em output/<mes>/monitoramento/*.xlsx e unifica em tabelas."""
    POWERBI_DIR.mkdir(parents=True, exist_ok=True)

    linhas_atrasos = []
    linhas_faltas = []
    linhas_regularizados = []
    linhas_resumo = []

    # Procura pastas mensais em output (ex: output/2026-08/monitoramento)
    # Também suporta se houver a pasta legada output/monitoramento
    pastas_monitoramento = []
    for item in sorted(OUTPUT_DIR.iterdir()):
        if not item.is_dir() or item.name == "powerbi":
            continue
        monit_sub = item / "monitoramento"
        if monit_sub.is_dir():
            pastas_monitoramento.append((item.name, monit_sub))
        elif item.name == "monitoramento":
            pastas_monitoramento.append(("geral", item))

    print(f"[Power BI] Pastas de monitoramento encontradas: {len(pastas_monitoramento)}")

    for mes_id, pasta in pastas_monitoramento:
        arquivos = sorted(pasta.glob("monitoramento_*.xlsx"))
        print(f"  -> Lendo lote '{mes_id}': {len(arquivos)} arquivo(s)")

        for arq in arquivos:
            cod_sec, nome_sec = extrair_codigo_e_nome(arq.name)
            sigla_sec = obter_sigla_secretaria(cod_sec, nome_sec)

            try:
                wb = load_workbook(arq, data_only=True)
            except Exception as e:
                print(f"     [AVISO] Falha ao abrir {arq.name}: {e}")
                continue

            # 1. Ler Aba Atrasos (Minutos) primeiro para apurar quem tem >= 300 min
            nome_aba_atrasos = None
            for s in wb.sheetnames:
                if "atraso" in s.lower() or "minuto" in s.lower():
                    nome_aba_atrasos = s
                    break

            atrasos_desta_sec = []
            if nome_aba_atrasos:
                ws_atr = wb[nome_aba_atrasos]
                for r in range(2, ws_atr.max_row + 1):
                    mat = ws_atr.cell(row=r, column=1).value
                    if not mat:
                        continue
                    nome = ws_atr.cell(row=r, column=2).value or ""
                    min_val = ws_atr.cell(row=r, column=3).value or 0
                    equiv = ws_atr.cell(row=r, column=4).value or ""
                    det = ws_atr.cell(row=r, column=5).value or ""
                    status = ws_atr.cell(row=r, column=6).value or ""

                    try: min_int = int(min_val)
                    except: min_int = 0

                    acima_300 = "Sim" if min_int >= 300 else "Não"
                    status_questionar = "Questionar (>= 300 min)" if min_int >= 300 else "Tolerado (< 300 min)"

                    item_atr = {
                        "AnoMes": mes_id,
                        "MesReferencia": mes_id,
                        "CodigoSecretaria": cod_sec,
                        "Secretaria": sigla_sec,
                        "NomeCompletoSecretaria": nome_sec,
                        "OrgaoCompleto": f"{cod_sec} - {sigla_sec}" if cod_sec else sigla_sec,
                        "Matricula": str(mat).strip(),
                        "NomeServidor": str(nome).strip(),
                        "MinutosAcumulados": min_int,
                        "HorasEquivalentes": round(min_int / 60.0, 2),
                        "Acima300Min": acima_300,
                        "QuestionarSecretaria": status_questionar,
                        "EquivalenciaTexto": str(equiv).strip(),
                        "OcorrenciasDatas": str(det).strip(),
                        "SituacaoSistema": str(status).strip(),
                    }
                    linhas_atrasos.append(item_atr)
                    atrasos_desta_sec.append(item_atr)

            # Contadores de servidores >= 300 min para o resumo
            servs_mais_300 = [a for a in atrasos_desta_sec if a["MinutosAcumulados"] >= 300]
            qtd_serv_mais_300 = len(servs_mais_300)
            tot_minutos_mais_300 = sum(a["MinutosAcumulados"] for a in servs_mais_300)

            # 2. Ler Aba Faltas (Dias)
            nome_aba_faltas = None
            for s in wb.sheetnames:
                if "falta" in s.lower():
                    nome_aba_faltas = s
                    break

            faltas_desta_sec = []
            if nome_aba_faltas:
                ws_flt = wb[nome_aba_faltas]
                for r in range(2, ws_flt.max_row + 1):
                    mat = ws_flt.cell(row=r, column=1).value
                    if not mat:
                        continue
                    nome = ws_flt.cell(row=r, column=2).value or ""
                    dias_val = ws_flt.cell(row=r, column=3).value or 0
                    tipos = ws_flt.cell(row=r, column=4).value or ""
                    det = ws_flt.cell(row=r, column=5).value or ""
                    status = ws_flt.cell(row=r, column=6).value or ""

                    try: dias_int = int(dias_val)
                    except: dias_int = 0

                    acima_4_faltas = "Sim" if dias_int >= 4 else "Não"
                    status_questionar_falta = "Questionar (>= 4 faltas)" if dias_int >= 4 else "Tolerado (< 4 faltas)"

                    item_flt = {
                        "AnoMes": mes_id,
                        "MesReferencia": mes_id,
                        "CodigoSecretaria": cod_sec,
                        "Secretaria": sigla_sec,
                        "NomeCompletoSecretaria": nome_sec,
                        "OrgaoCompleto": f"{cod_sec} - {sigla_sec}" if cod_sec else sigla_sec,
                        "Matricula": str(mat).strip(),
                        "NomeServidor": str(nome).strip(),
                        "DiasFalta": dias_int,
                        "Acima4Faltas": acima_4_faltas,
                        "QuestionarFalta": status_questionar_falta,
                        "TiposFalta": str(tipos).strip(),
                        "OcorrenciasDatas": str(det).strip(),
                        "SituacaoSistema": str(status).strip(),
                    }
                    linhas_faltas.append(item_flt)
                    faltas_desta_sec.append(item_flt)

            servs_mais_4_faltas = [f for f in faltas_desta_sec if f["DiasFalta"] >= 4]
            qtd_serv_mais_4_faltas = len(servs_mais_4_faltas)
            tot_dias_mais_4_faltas = sum(f["DiasFalta"] for f in servs_mais_4_faltas)

            # 3. Ler Aba Regularizados (Memorando)
            nome_aba_reg = None
            for s in wb.sheetnames:
                if "regularizado" in s.lower():
                    nome_aba_reg = s
                    break

            regularizados_desta_sec = []
            if nome_aba_reg:
                ws_reg = wb[nome_aba_reg]
                for r in range(2, ws_reg.max_row + 1):
                    mat = ws_reg.cell(row=r, column=1).value
                    if not mat:
                        continue
                    nome = ws_reg.cell(row=r, column=2).value or ""
                    oc_ant = ws_reg.cell(row=r, column=3).value or ""
                    oc_nova = ws_reg.cell(row=r, column=4).value or ""
                    dt_per = ws_reg.cell(row=r, column=5).value or ""
                    memo = ws_reg.cell(row=r, column=6).value or ""
                    status = ws_reg.cell(row=r, column=7).value or ""

                    item_reg = {
                        "AnoMes": mes_id,
                        "MesReferencia": mes_id,
                        "CodigoSecretaria": cod_sec,
                        "Secretaria": sigla_sec,
                        "NomeCompletoSecretaria": nome_sec,
                        "OrgaoCompleto": f"{cod_sec} - {sigla_sec}" if cod_sec else sigla_sec,
                        "Matricula": str(mat).strip(),
                        "NomeServidor": str(nome).strip(),
                        "OcorrenciaAnterior": str(oc_ant).strip(),
                        "RetificadoPara": str(oc_nova).strip(),
                        "DataPeriodo": str(dt_per).strip(),
                        "MemorandoOrigem": str(memo).strip(),
                        "SituacaoSistema": str(status).strip(),
                    }
                    linhas_regularizados.append(item_reg)
                    regularizados_desta_sec.append(item_reg)

            qtd_regularizados = len(regularizados_desta_sec)

            # 4. Ler Aba Resumo
            mes_label = mes_id
            qtd_serv_atraso = len(atrasos_desta_sec)
            tot_minutos = sum(a["MinutosAcumulados"] for a in atrasos_desta_sec)
            qtd_serv_falta = len(faltas_desta_sec)
            tot_dias_falta = sum(f["DiasFalta"] for f in faltas_desta_sec)

            if "Resumo" in wb.sheetnames:
                ws_res = wb["Resumo"]
                titulo_res = ws_res.cell(row=1, column=1).value or ""
                m_label = re.search(r"\((.+?)\)", str(titulo_res))
                if m_label:
                    mes_label = m_label.group(1).strip()
                    for a in atrasos_desta_sec:
                        a["MesReferencia"] = mes_label
                    for f in faltas_desta_sec:
                        f["MesReferencia"] = mes_label
                    for reg_it in regularizados_desta_sec:
                        reg_it["MesReferencia"] = mes_label

                for row in range(3, ws_res.max_row + 1):
                    ind = str(ws_res.cell(row=row, column=1).value or "").strip().lower()
                    val = ws_res.cell(row=row, column=2).value
                    if "servidores com atrasos" in ind:
                        try: qtd_serv_atraso = int(val)
                        except: pass
                    elif "total de minutos perdidos" in ind:
                        m_val = re.search(r"(\d+)\s*min", str(val))
                        if m_val:
                            tot_minutos = int(m_val.group(1))
                        else:
                            try: tot_minutos = int(val)
                            except: pass
                    elif "servidores com faltas" in ind:
                        try: qtd_serv_falta = int(val)
                        except: pass
                    elif "total de faltas acumuladas" in ind:
                        m_dias = re.search(r"(\d+)", str(val))
                        if m_dias:
                            tot_dias_falta = int(m_dias.group(1))
                    elif "regularizada" in ind:
                        try: qtd_regularizados = int(val)
                        except: pass

            linhas_resumo.append({
                "AnoMes": mes_id,
                "MesReferencia": mes_label,
                "CodigoSecretaria": cod_sec,
                "Secretaria": sigla_sec,
                "NomeCompletoSecretaria": nome_sec,
                "OrgaoCompleto": f"{cod_sec} - {sigla_sec}" if cod_sec else sigla_sec,
                "QtdServidoresAtraso": qtd_serv_atraso,
                "TotalMinutosPerdidos": tot_minutos,
                "TotalHorasPerdidas": round(tot_minutos / 60.0, 2),
                "QtdServidoresMais300Min": qtd_serv_mais_300,
                "TotalMinutosMais300Min": tot_minutos_mais_300,
                "TotalHorasMais300Min": round(tot_minutos_mais_300 / 60.0, 2),
                "QtdServidoresFalta": qtd_serv_falta,
                "TotalDiasFalta": tot_dias_falta,
                "QtdServidoresMais4Faltas": qtd_serv_mais_4_faltas,
                "TotalDiasMais4Faltas": tot_dias_mais_4_faltas,
                "QtdRegularizadosMemorando": qtd_regularizados,
            })

    # Gravar Excel Consolidado
    caminho_xlsx = POWERBI_DIR / "monitoramento_consolidado.xlsx"
    wb_out = Workbook()

    cor_header = "1F4E78"
    font_header = Font(bold=True, color="FFFFFF")
    fill_header = PatternFill("solid", fgColor=cor_header)

    # Aba 1: Resumo_Secretarias
    ws1 = wb_out.active
    ws1.title = "Resumo_Secretarias"
    if linhas_resumo:
        headers1 = list(linhas_resumo[0].keys())
        ws1.append(headers1)
        for cel in ws1[1]:
            cel.font = font_header
            cel.fill = fill_header
            cel.alignment = Alignment(horizontal="center")
        for item in linhas_resumo:
            ws1.append(list(item.values()))
        for col_idx in range(1, len(headers1) + 1):
            ws1.column_dimensions[get_column_letter(col_idx)].width = 22

    # Aba 2: Fato_Atrasos
    ws2 = wb_out.create_sheet("Fato_Atrasos")
    if linhas_atrasos:
        headers2 = list(linhas_atrasos[0].keys())
        ws2.append(headers2)
        for cel in ws2[1]:
            cel.font = font_header
            cel.fill = fill_header
            cel.alignment = Alignment(horizontal="center")
        for item in linhas_atrasos:
            ws2.append(list(item.values()))
        for col_idx in range(1, len(headers2) + 1):
            ws2.column_dimensions[get_column_letter(col_idx)].width = 22

    # Aba 3: Fato_Faltas
    ws3 = wb_out.create_sheet("Fato_Faltas")
    if linhas_faltas:
        headers3 = list(linhas_faltas[0].keys())
        ws3.append(headers3)
        for cel in ws3[1]:
            cel.font = font_header
            cel.fill = fill_header
            cel.alignment = Alignment(horizontal="center")
        for item in linhas_faltas:
            ws3.append(list(item.values()))
        for col_idx in range(1, len(headers3) + 1):
            ws3.column_dimensions[get_column_letter(col_idx)].width = 22

    # Aba 4: Fato_Regularizados
    ws4 = wb_out.create_sheet("Fato_Regularizados")
    if linhas_regularizados:
        headers4 = list(linhas_regularizados[0].keys())
        ws4.append(headers4)
        for cel in ws4[1]:
            cel.font = font_header
            cel.fill = fill_header
            cel.alignment = Alignment(horizontal="center")
        for item in linhas_regularizados:
            ws4.append(list(item.values()))
        for col_idx in range(1, len(headers4) + 1):
            ws4.column_dimensions[get_column_letter(col_idx)].width = 22

    wb_out.save(caminho_xlsx)
    print(f"\n[OK] Base consolidada salva em: {caminho_xlsx.relative_to(BASE_DIR)}")
    print(f"     -> Resumo Secretarias: {len(linhas_resumo)} linhas")
    print(f"     -> Fato Atrasos:       {len(linhas_atrasos)} registros de servidores")
    print(f"     -> Fato Faltas:        {len(linhas_faltas)} registros de servidores")
    print(f"     -> Fato Regularizados: {len(linhas_regularizados)} registros regularizados")

    # Gravar CSVs individuais (caso o usuário queira importar CSV)
    import csv
    def salvar_csv(linhas, nome_arquivo):
        if not linhas:
            return
        p_csv = POWERBI_DIR / nome_arquivo
        with open(p_csv, "w", encoding="utf-8-sig", newline="") as f:
            writer = csv.DictWriter(f, fieldnames=linhas[0].keys(), delimiter=";")
            writer.writeheader()
            writer.writerows(linhas)
        print(f"[OK] CSV gerado: {p_csv.relative_to(BASE_DIR)}")

    salvar_csv(linhas_resumo, "resumo_secretarias.csv")
    salvar_csv(linhas_atrasos, "fato_atrasos.csv")
    salvar_csv(linhas_faltas, "fato_faltas.csv")
    salvar_csv(linhas_regularizados, "fato_regularizados.csv")

    return {
        "resumo": len(linhas_resumo),
        "atrasos": len(linhas_atrasos),
        "faltas": len(linhas_faltas),
        "regularizados": len(linhas_regularizados),
        "arquivo_excel": caminho_xlsx,
    }


if __name__ == "__main__":
    consolidar_monitoramento()
