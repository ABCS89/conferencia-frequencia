# -*- coding: utf-8 -*-
"""
main.py - Conferência de Frequência

Uso:
    python src/main.py
        -> Detecta e processa automaticamente todos os pares em input/secretaria e input/sistema

    python src/main.py <pdf_secretaria> <pdf_sistema> [--mes AAAA-MM]
        -> Processa um par específico

Gera em output/:
  - conferencia_<CODIGO - NOME DA SECRETARIA>.xlsx  (aba "Retificações" + aba "Resumo")
  - retificacoes_<CODIGO - NOME DA SECRETARIA>.md   (se houver retificação)
  - retificacoes_<CODIGO - NOME DA SECRETARIA>.txt  (se houver retificação)
"""
import sys
import os
import re
from pathlib import Path
from datetime import datetime

# Garante saída compatível com consoles Windows (cp1252 / UTF-8)
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

# Garante que imports locais funcionem
sys.path.insert(0, str(Path(__file__).resolve().parent))

from extract_secretaria import extrair as extrair_secretaria, extrair_cabecalho
from extract_sistema import extrair as extrair_sistema, extrair_cabecalho_sistema
from compare import comparar_por_dia, mes_referencia, mes_referencia_registros
from desligados import localizar_e_carregar_desligados, normalizar_matricula
from monitoring import (
    apurar_dados_monitoramento,
    gerar_markdown_monitoramento,
    gerar_excel_monitoramento,
)
from memorando import (
    localizar_memorando,
    extrair_texto_completo_memorando,
    extrair_itens_memorando,
    conferir_memorando_com_solicitacoes,
    gerar_secao_memorando_markdown,
)
from export_powerbi import consolidar_monitoramento

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# Diretórios base garantidos independentemente de onde o script é chamado
BASE_DIR = Path(__file__).resolve().parent.parent
INPUT_DIR = BASE_DIR / "input"
OUTPUT_DIR = BASE_DIR / "output"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

COR_CABECALHO = "1F4E78"
COR_SEM_REGISTRO_SECRETARIA = "FFF2CC"   # secretaria não tinha nada -> precisa lançar
COR_SEM_REGISTRO_SISTEMA = "D9E1F2"      # sistema não tinha nada -> conferir/lançar lá
COR_TIPO_DIFERENTE = "F8CBAD"            # os dois têm algo, mas tipos diferentes
COR_SOBREPOSICAO = "E1D5E7"              # lilás: sobreposição de eventos na mesma data
COR_REGULARIZADO = "E2EFDA"              # verde suave: regularizado por memorando

MESES_PT = ["", "janeiro", "fevereiro", "março", "abril", "maio", "junho",
            "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"]


def formatar_mes_ano(ano_mes):
    """Formata tupla (ano, mes) como 'MM/AAAA (Mês/AAAA)'."""
    if not ano_mes or len(ano_mes) != 2:
        return "não identificado"
    ano, mes = ano_mes
    if 1 <= mes <= 12:
        nome_mes = MESES_PT[mes].capitalize()
        return f"{str(mes).zfill(2)}/{ano} ({nome_mes}/{ano})"
    return f"{str(mes).zfill(2)}/{ano}"


def validar_competencia_documentos(ano_mes_esperado, mes_cabecalho_sec, mes_moda_sec,
                                  mes_cabecalho_sis, mes_moda_sis,
                                  nome_sec_pdf="", nome_sis_pdf=""):
    """Verifica se os arquivos da Secretaria e do Sistema coincidem entre si
    e com o mês esperado do lote (pasta). Retorna lista de avisos (vazia se estiver tudo correto).
    """
    avisos = []
    mes_sec = mes_cabecalho_sec or mes_moda_sec
    mes_sis = mes_cabecalho_sis or mes_moda_sis

    # 1. Secretaria vs Pasta do lote
    if ano_mes_esperado and mes_sec and mes_sec != ano_mes_esperado:
        avisos.append(
            f"Arquivo da Secretaria ('{nome_sec_pdf}') aponta para {formatar_mes_ano(mes_sec)}, "
            f"divergente da pasta do lote ({formatar_mes_ano(ano_mes_esperado)})."
        )

    # 2. Sistema vs Pasta do lote
    if ano_mes_esperado and mes_sis and mes_sis != ano_mes_esperado:
        avisos.append(
            f"Arquivo do Sistema ('{nome_sis_pdf}') aponta para {formatar_mes_ano(mes_sis)}, "
            f"divergente da pasta do lote ({formatar_mes_ano(ano_mes_esperado)})."
        )

    # 3. Secretaria vs Sistema entre si
    if mes_sec and mes_sis and mes_sec != mes_sis:
        avisos.append(
            f"Divergência entre documentos: Secretaria refere-se a {formatar_mes_ano(mes_sec)} "
            f"enquanto Sistema refere-se a {formatar_mes_ano(mes_sis)}."
        )

    # 4. Inconsistência interna (cabeçalho vs lançamentos)
    if mes_cabecalho_sec and mes_moda_sec and mes_cabecalho_sec != mes_moda_sec:
        avisos.append(
            f"Arquivo da Secretaria ('{nome_sec_pdf}') traz cabeçalho {formatar_mes_ano(mes_cabecalho_sec)}, "
            f"mas a maioria dos registros lançados é de {formatar_mes_ano(mes_moda_sec)}."
        )

    if mes_cabecalho_sis and mes_moda_sis and mes_cabecalho_sis != mes_moda_sis:
        avisos.append(
            f"Arquivo do Sistema ('{nome_sis_pdf}') traz cabeçalho {formatar_mes_ano(mes_cabecalho_sis)}, "
            f"mas a maioria dos registros lançados é de {formatar_mes_ano(mes_moda_sis)}."
        )

    return avisos


def sanitizar_nome_arquivo(texto):
    """Remove caracteres proibidos no sistema de arquivos do Windows."""
    for ch in r'\/:*?"<>|':
        texto = texto.replace(ch, "-")
    return " ".join(texto.split()).strip(" .-")


def _cor_da_linha(r):
    if r.get("regularizado_memorando"):
        return COR_REGULARIZADO
    if r.get("sobreposicao"):
        return COR_SOBREPOSICAO
    if r["tipo_secretaria"] == "SEM REGISTRO":
        return COR_SEM_REGISTRO_SECRETARIA
    if r["tipo_sistema"] == "sem registro em sistema":
        return COR_SEM_REGISTRO_SISTEMA
    return COR_TIPO_DIFERENTE


def _texto_data(r):
    if r["data_inicio"] == r["data_fim"]:
        return r["data_inicio"].strftime("%d/%m/%Y")
    return f'{r["data_inicio"].strftime("%d/%m/%Y")} a {r["data_fim"].strftime("%d/%m/%Y")}'


def _linha_texto(r):
    if r.get("instrucao_direta"):
        return (f'{r["matricula"]} - {r["nome"]} - {_texto_data(r)} - {r["dias"]} '
                f'dia{"s" if r["dias"] != 1 else ""} - {r["instrucao_direta"]}')
    obs = f' ({r["observacao"]})' if r.get("observacao") else ""
    return (f'{r["matricula"]} - {r["nome"]} - {_texto_data(r)} - {r["dias"]} '
            f'dia{"s" if r["dias"] != 1 else ""} - {r["tipo_secretaria"]} --> {r["tipo_sistema"]}{obs}')


def _eh_lancamento_somente_sec(r):
    """Retorna True se for registro de minutos perdidos ou falta informado pela secretaria sem lançamento no sistema."""
    tipo_sec_norm = r["tipo_secretaria"].lower()
    eh_alvo = ("minuto" in tipo_sec_norm or "falta" in tipo_sec_norm)
    return eh_alvo and r["tipo_sistema"] == "sem registro em sistema"


_eh_minutos_somente_sec = _eh_lancamento_somente_sec


def _eh_divergencia_secretaria(r):
    """Retorna True se for uma divergência que a secretaria precisa retificar via memorando.
    Registros de 'SEM REGISTRO' na secretaria ou de minutos perdidos/faltas que a secretaria já informou
    não são retificações a cargo da secretaria (são tratados internamente/inseridos no sistema).
    """
    if r["tipo_secretaria"] == "SEM REGISTRO":
        return False
    if _eh_lancamento_somente_sec(r):
        return False
    return True


def gerar_excel(retificacoes, sobreposicoes, caminho_saida, total_sec, total_sis, mes_label, nome_orgao="", itens_regularizados=None, avisos_mes=None):
    wb = Workbook()

    ws = wb.active
    ws.title = "Retificações"
    colunas = ["Matrícula", "Nome", "Data Início", "Data Fim", "Dias",
               "Tipo Atual (a corrigir)", "Tipo Correto", "Observação / Conflito"]
    ws.append(colunas)
    for cel in ws[1]:
        cel.font = Font(bold=True, color="FFFFFF")
        cel.fill = PatternFill("solid", fgColor=COR_CABECALHO)
        cel.alignment = Alignment(horizontal="center")

    chaves_reg = {(normalizar_matricula(a["matricula"]), a["data_inicio"], a["data_fim"]) for a in (itens_regularizados or [])}

    ret_divergencias = [r for r in retificacoes if _eh_divergencia_secretaria(r)]
    ret_sem_registro = [r for r in retificacoes if not _eh_divergencia_secretaria(r)]
    ret_pendentes = [r for r in ret_divergencias if (normalizar_matricula(r["matricula"]), r["data_inicio"], r["data_fim"]) not in chaves_reg]

    # Ordena colocando divergências primeiro, depois servidores sem registro na secretaria / a inserir no sistema
    ret_ordenadas = ret_divergencias + ret_sem_registro

    for r in ret_ordenadas:
        ch_r = (normalizar_matricula(r["matricula"]), r["data_inicio"], r["data_fim"])
        foi_reg = ch_r in chaves_reg
        r_copia = dict(r)
        if foi_reg:
            r_copia["regularizado_memorando"] = True

        obs_texto = r_copia.get("observacao", "")
        if r_copia["tipo_secretaria"] == "SEM REGISTRO" and not obs_texto:
            obs_texto = "Sem registro na secretaria (verificar possível desligamento/situação funcional)"
        elif _eh_lancamento_somente_sec(r_copia) and not obs_texto:
            tipo_label = "Falta informada" if "falta" in r_copia["tipo_secretaria"].lower() else "Minutos perdidos informados"
            obs_texto = f"{tipo_label} pela secretaria (inserir no sistema)"

        if foi_reg:
            obs_col = f"Regularizado por memorando de retificação | {obs_texto}" if obs_texto else "Regularizado por memorando de retificação"
        elif r_copia.get("instrucao_direta") and obs_texto:
            obs_col = f"{r_copia['instrucao_direta'].capitalize()} | {obs_texto}"
        elif r_copia.get("instrucao_direta"):
            obs_col = r_copia["instrucao_direta"].capitalize()
        else:
            obs_col = obs_texto

        ws.append([
            r_copia["matricula"], r_copia["nome"],
            r_copia["data_inicio"].strftime("%d/%m/%Y"), r_copia["data_fim"].strftime("%d/%m/%Y"),
            r_copia["dias"], r_copia["tipo_secretaria"], r_copia["tipo_sistema"],
            obs_col,
        ])
        cor = _cor_da_linha(r_copia)
        for cel in ws[ws.max_row]:
            cel.fill = PatternFill("solid", fgColor=cor)

    larguras = [12, 34, 14, 14, 8, 30, 36, 60]
    for i, w in enumerate(larguras, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"

    if sobreposicoes:
        ws_sob = wb.create_sheet("Sobreposições")
        cols_sob = ["Matrícula", "Nome", "Origem", "Data Início", "Data Fim", "Dias",
                    "Eventos Sobrepostos", "Observação"]
        ws_sob.append(cols_sob)
        for cel in ws_sob[1]:
            cel.font = Font(bold=True, color="FFFFFF")
            cel.fill = PatternFill("solid", fgColor=COR_CABECALHO)
            cel.alignment = Alignment(horizontal="center")

        for s in sobreposicoes:
            ws_sob.append([
                s["matricula"], s["nome"], s["origem"],
                s["data_inicio"].strftime("%d/%m/%Y"), s["data_fim"].strftime("%d/%m/%Y"),
                s["dias"], s["eventos"], s["observacao"],
            ])
            for cel in ws_sob[ws_sob.max_row]:
                cel.fill = PatternFill("solid", fgColor=COR_SOBREPOSICAO)

        larg_sob = [12, 34, 12, 14, 14, 8, 38, 60]
        for i, w in enumerate(larg_sob, start=1):
            ws_sob.column_dimensions[get_column_letter(i)].width = w
        ws_sob.freeze_panes = "A2"

    ws2 = wb.create_sheet("Resumo")
    titulo_resumo = f"Conferência de Frequência — {nome_orgao} ({mes_label})" if nome_orgao else f"Conferência de Frequência ({mes_label})"
    ws2.append([titulo_resumo, ""])
    ws2["A1"].font = Font(bold=True, size=14)
    ws2.append(["Gerado em", datetime.now().strftime("%d/%m/%Y %H:%M")])
    ws2.append(["Órgão / Secretaria", nome_orgao or "Não identificado"])
    ws2.append(["Mês de referência", mes_label])
    if avisos_mes:
        for av in avisos_mes:
            ws2.append(["⚠️ ALERTA DE COMPETÊNCIA/PASTA", av])
            cel_a = ws2.cell(row=ws2.max_row, column=1)
            cel_b = ws2.cell(row=ws2.max_row, column=2)
            cel_a.font = Font(bold=True, color="9C0006")
            cel_a.fill = PatternFill("solid", fgColor="FFC7CE")
            cel_b.font = Font(bold=True, color="9C0006")
            cel_b.fill = PatternFill("solid", fgColor="FFC7CE")
        ws2.append([])
    ws2.append(["Registros extraídos (secretaria)", total_sec])
    ws2.append(["Registros extraídos (sistema)", total_sis])
    ws2.append(["Total de retificações solicitadas à secretaria", len(ret_divergencias)])
    if itens_regularizados:
        ws2.append(["Regularizadas por memorando recebido", len(itens_regularizados)])
        ws2.append(["Pendentes de retificação pela secretaria", len(ret_pendentes)])
    ws2.append(["Servidores sem registro na frequência / a inserir no sistema", len(ret_sem_registro)])
    ws2.append(["Sobreposições de eventos detectadas", len(sobreposicoes)])
    ws2.append([])
    ws2.append(["Legenda de cores", ""])
    ws2["A11" if not itens_regularizados else "A13"].font = Font(bold=True)
    if itens_regularizados:
        ws2.append(["Verde suave", "ocorrência regularizada pelo memorando de retificação"])
    ws2.append(["Amarelo", "secretaria não tinha esse servidor na folha (verificar desligamento)"])
    ws2.append(["Azul", "sistema não tem nada nesse dia (conferir se é pra lançar lá)"])
    ws2.append(["Laranja", "os dois têm algo lançado, mas de tipo diferente"])
    ws2.append(["Lilás", "sobreposição de eventos na mesma data (conflito interno a verificar)"])
    ws2.column_dimensions["A"].width = 40
    ws2.column_dimensions["B"].width = 55

    os.makedirs(os.path.dirname(caminho_saida), exist_ok=True)
    wb.save(caminho_saida)


def gerar_txt(retificacoes, sobreposicoes, caminho_saida, itens_regularizados=None, avisos_mes=None):
    linhas = []
    if avisos_mes:
        linhas.append("=== ⚠️ ALERTA: COMPETÊNCIA DIVERGENTE OU ARQUIVO EM PASTA INCORRETA ===")
        for av in avisos_mes:
            linhas.append(f"  * {av}")
        linhas.append("")

    if sobreposicoes:
        linhas.append("=== ATENÇÃO: SOBREPOSIÇÃO DE EVENTOS NA MESMA DATA ===")
        for s in sobreposicoes:
            dias_txt = (f"{s['data_inicio'].strftime('%d/%m/%Y')} a {s['data_fim'].strftime('%d/%m/%Y')}"
                        if s["data_inicio"] != s["data_fim"] else s["data_inicio"].strftime("%d/%m/%Y"))
            linhas.append(f"- {s['matricula']} - {s['nome']} - {dias_txt} ({s['dias']} dia{'s' if s['dias'] != 1 else ''}) - {s['origem']}: {s['eventos']}")
        linhas.append("")

    chaves_reg = {(normalizar_matricula(a["matricula"]), a["data_inicio"], a["data_fim"]) for a in (itens_regularizados or [])}
    ret_divergencias = [r for r in retificacoes if _eh_divergencia_secretaria(r) and (normalizar_matricula(r["matricula"]), r["data_inicio"], r["data_fim"]) not in chaves_reg]
    ret_sem_registro = [r for r in retificacoes if not _eh_divergencia_secretaria(r) and (normalizar_matricula(r["matricula"]), r["data_inicio"], r["data_fim"]) not in chaves_reg]

    if ret_divergencias:
        linhas.append("=== RETIFICAÇÕES A REALIZAR ===")
        linhas.append("Favor enviar memorando de retificação das seguintes frequências:")
        linhas.append("")
        for r in ret_divergencias:
            linhas.append(f"- {_linha_texto(r)}")
            if r.get("nota_responsabilidade"):
                linhas.append(f"  [Responsabilidade: {r['nota_responsabilidade']}]")
        linhas.append("")

    if ret_sem_registro:
        linhas.append("=== OCORRÊNCIAS SEM REGISTRO NA FREQUÊNCIA DA SECRETARIA (A VERIFICAR / INSERIR NO SISTEMA) ===")
        for r in ret_sem_registro:
            linhas.append(f"- {_linha_texto(r)}")
            if r.get("nota_responsabilidade"):
                linhas.append(f"  [Responsabilidade: {r['nota_responsabilidade']}]")
        linhas.append("")

    os.makedirs(os.path.dirname(caminho_saida), exist_ok=True)
    with open(caminho_saida, "w", encoding="utf-8") as f:
        f.write("\n".join(linhas) + "\n")


def gerar_markdown(retificacoes, sobreposicoes, caminho_saida, mes_label, total_sec, total_sis, nome_orgao="", secao_memorando=None, itens_regularizados=None, avisos_mes=None):
    linhas = []
    titulo = f"# Retificações de Frequência — {nome_orgao} ({mes_label})" if nome_orgao else f"# Retificações de Frequência ({mes_label})"
    linhas.append(titulo)
    linhas.append("")
    if avisos_mes:
        linhas.append("> [!CAUTION]")
        linhas.append("> ### ⚠️ ALERTA: ARQUIVO COM COMPETÊNCIA DIVERGENTE OU PASTA INCORRETA")
        linhas.append("> Foram identificadas divergências de mês de competência nos documentos analisados:")
        for av in avisos_mes:
            linhas.append(f"> - {av}")
        linhas.append(">")
        linhas.append("> *Atenção: Os lançamentos foram filtrados para o mês do lote. Se os documentos forem de meses distintos ou estiverem na pasta incorreta, os resultados da conferência podem conter falsas divergências ou dados faltantes.*")
        linhas.append("")
    linhas.append(f"- Gerado em: {datetime.now().strftime('%d/%m/%Y %H:%M')}")
    if nome_orgao:
        linhas.append(f"- Órgão / Secretaria: {nome_orgao}")
    linhas.append(f"- Mês de referência: {mes_label}")
    linhas.append(f"- Registros extraídos (secretaria): {total_sec}")
    linhas.append(f"- Registros extraídos (sistema): {total_sis}")

    chaves_reg = {(normalizar_matricula(a["matricula"]), a["data_inicio"], a["data_fim"]) for a in (itens_regularizados or [])}
    ret_divergencias_total = [r for r in retificacoes if _eh_divergencia_secretaria(r)]
    ret_divergencias = [r for r in ret_divergencias_total if (normalizar_matricula(r["matricula"]), r["data_inicio"], r["data_fim"]) not in chaves_reg]

    ret_sem_registro_total = [r for r in retificacoes if not _eh_divergencia_secretaria(r)]
    ret_sem_registro = [r for r in ret_sem_registro_total if (normalizar_matricula(r["matricula"]), r["data_inicio"], r["data_fim"]) not in chaves_reg]

    if itens_regularizados:
        linhas.append(f"- **Retificações solicitadas inicialmente: {len(ret_divergencias_total)}**")
        linhas.append(f"- **Regularizadas pelo memorando recebido: {len(itens_regularizados)}**")
        linhas.append(f"- **Retificações pendentes a realizar: {len(ret_divergencias)}**")
    else:
        linhas.append(f"- **Total de retificações: {len(ret_divergencias)}**")

    if ret_sem_registro:
        tem_inserir = any(_eh_lancamento_somente_sec(r) for r in ret_sem_registro)
        rotulo_sem_reg = "Servidores sem registro na frequência / a inserir no sistema" if tem_inserir else "Servidores sem registro na frequência (a verificar)"
        linhas.append(f"- **{rotulo_sem_reg}: {len(ret_sem_registro)}**")
    if sobreposicoes:
        linhas.append(f"- **⚠️ Sobreposições de eventos detectadas: {len(sobreposicoes)}**")
    linhas.append("")

    if sobreposicoes:
        linhas.append("## ⚠️ Sobreposição de Eventos na Mesma Data")
        linhas.append("Foram identificados servidores com múltiplos registros para a mesma data (conflito no sistema ou secretaria):")
        linhas.append("")
        for s in sobreposicoes:
            dias_txt = (f"{s['data_inicio'].strftime('%d/%m/%Y')} a {s['data_fim'].strftime('%d/%m/%Y')}"
                        if s["data_inicio"] != s["data_fim"] else s["data_inicio"].strftime("%d/%m/%Y"))
            linhas.append(f"- **{s['matricula']} - {s['nome']}**: {dias_txt} ({s['dias']} dia{'s' if s['dias'] != 1 else ''}) — *{s['origem']}*: **{s['eventos']}**")
        linhas.append("")
        linhas.append("> [!IMPORTANT]")
        linhas.append("> **Regra de Responsabilidade em Conflitos Falta vs. Afastamento Médico**:")
        linhas.append("> - **Faltas relatadas pela secretaria**: A secretaria deve retificar as frequências substituindo as faltas pelo afastamento médico.")
        linhas.append("> - **Faltas registradas no sistema**: Caso constem faltas no sistema em datas cobertas por atestado médico, cabe ao RH/sistema retirá-las.")
        linhas.append("")

    if ret_divergencias:
        linhas.append("## Retificações a Realizar")
        linhas.append("Favor enviar memorando de retificação das seguintes frequências:")
        linhas.append("")
        for r in ret_divergencias:
            linhas.append(f"- {_linha_texto(r)}")
            if r.get("nota_responsabilidade"):
                linhas.append(f"  > **Responsabilidade**: {r['nota_responsabilidade']}")
        linhas.append("")
    elif secao_memorando:
        linhas.append("## Retificações a Realizar")
        linhas.append("*(Todas as retificações a cargo da secretaria foram regularizadas pelo memorando!)*")
        linhas.append("")

    if ret_sem_registro:
        linhas.append("## Ocorrências sem Registro na Frequência da Secretaria")
        tem_sem_freq = any(r["tipo_secretaria"] == "SEM REGISTRO" for r in ret_sem_registro)
        tem_inserir = any(_eh_lancamento_somente_sec(r) for r in ret_sem_registro)

        if tem_sem_freq and tem_inserir:
            linhas.append("Os seguintes servidores possuem lançamentos no sistema que **não constam** no relatório da secretaria (verificar se o servidor já está desligado/exonerado ou se houve omissão na lista) ou ocorrências informadas pela secretaria (faltas / minutos perdidos) que **devem ser inseridas no sistema**:")
        elif tem_inserir:
            linhas.append("Os seguintes servidores possuem ocorrências informadas pela secretaria (faltas / minutos perdidos) sem lançamento no sistema (ocorrências a serem inseridas no sistema):")
        else:
            linhas.append("Os seguintes servidores possuem lançamentos no sistema, mas **não constam** no relatório de frequência entregue pela secretaria (verificar se o servidor já está desligado/exonerado ou se houve omissão na lista):")
        linhas.append("")
        for r in ret_sem_registro:
            linhas.append(f"- {_linha_texto(r)}")
            if r.get("nota_responsabilidade"):
                linhas.append(f"  > **Responsabilidade**: {r['nota_responsabilidade']}")
        linhas.append("")

    if secao_memorando:
        linhas.append(secao_memorando)
        linhas.append("")

    os.makedirs(os.path.dirname(caminho_saida), exist_ok=True)
    with open(caminho_saida, "w", encoding="utf-8") as f:
        f.write("\n".join(linhas).rstrip() + "\n")


def processar_par(caminho_secretaria, caminho_sistema, ano_mes=None, desligados=None, pasta_saida_mes=None):
    """Processa um par de relatórios (secretaria + sistema) e gera os arquivos em output/<mes>/."""
    caminho_secretaria = Path(caminho_secretaria)
    caminho_sistema = Path(caminho_sistema)

    if desligados is None:
        desligados, _ = localizar_e_carregar_desligados(INPUT_DIR)

    # Identificar mês esperado da pasta (se aplicável)
    ano_mes_pasta = ano_mes
    p_nome = pasta_saida_mes or caminho_secretaria.parent.name
    if ano_mes_pasta is None and p_nome and re.match(r"^\d{4}-\d{2}$", p_nome):
        a_str, m_str = p_nome.split("-")
        ano_mes_pasta = (int(a_str), int(m_str))

    print("\n" + "=" * 65)
    print(f"PROCESSANDO CONFERÊNCIA:")
    print(f"  Secretaria: {caminho_secretaria.name}")
    print(f"  Sistema:    {caminho_sistema.name}")
    if ano_mes_pasta:
        print(f"  Lote/Pasta: {formatar_mes_ano(ano_mes_pasta)}")
    print("=" * 65)

    # 1. Extrair cabeçalho da secretaria (código, nome e mês/ano)
    codigo_sec, nome_sec, mes_cabecalho = extrair_cabecalho(str(caminho_secretaria))
    if ano_mes is None:
        ano_mes = ano_mes_pasta or mes_cabecalho

    if codigo_sec and nome_sec:
        nome_orgao = f"{codigo_sec} - {nome_sec}"
    elif nome_sec:
        nome_orgao = nome_sec
    else:
        m_arq = re.match(r"^(\d{3})\b", caminho_secretaria.stem)
        nome_orgao = m_arq.group(1) if m_arq else caminho_secretaria.stem

    print(f"Identificação do Órgão: {nome_orgao}")

    # 2. Extrair registros da secretaria
    print("Lendo registros da secretaria...")
    regs_sec, nao_reconhecidas = extrair_secretaria(str(caminho_secretaria))
    print(f"  -> {len(regs_sec)} registro(s) extraído(s)")
    if nao_reconhecidas:
        print(f"  -> {len(nao_reconhecidas)} linha(s) ignorada(s)/ruído")

    # 3. Extrair registros do sistema e cabeçalho
    print("Lendo registros do sistema...")
    regs_sis = extrair_sistema(str(caminho_sistema))
    print(f"  -> {len(regs_sis)} registro(s) extraído(s)")
    info_sis = extrair_cabecalho_sistema(str(caminho_sistema))
    mes_cabecalho_sis = info_sis.get("ano_mes")

    # 3.1. Validar compatibilidade de meses entre documentos e pasta
    mes_moda_sec = mes_referencia(regs_sec)
    mes_moda_sis = mes_referencia_registros(regs_sis, "data_inicial")

    avisos_mes = validar_competencia_documentos(
        ano_mes_esperado=ano_mes_pasta,
        mes_cabecalho_sec=mes_cabecalho,
        mes_moda_sec=mes_moda_sec,
        mes_cabecalho_sis=mes_cabecalho_sis,
        mes_moda_sis=mes_moda_sis,
        nome_sec_pdf=caminho_secretaria.name,
        nome_sis_pdf=caminho_sistema.name,
    )
    if avisos_mes:
        print("\n" + "!" * 65)
        print("⚠️  [ALERTA: DIVERGÊNCIA DE COMPETÊNCIA / ARQUIVO NA PASTA ERRADA]")
        for av in avisos_mes:
            print(f"  • {av}")
        print("  -> Atenção: Os resultados da conferência podem conter distorções!")
        print("!" * 65 + "\n")

    # 4. Comparar dia a dia
    retificacoes, sobreposicoes, (ano_ref, mes_ref) = comparar_por_dia(
        regs_sec, regs_sis, ano_mes=ano_mes, desligados=desligados
    )
    mes_label = f"{MESES_PT[mes_ref]}/{ano_ref}" if (ano_ref and mes_ref) else "mês não identificado"
    print(f"Mês de referência apurado: {mes_label}")
    print(f"Total de retificações encontradas: {len(retificacoes)}")
    if sobreposicoes:
        print(f"  -> [ATENÇÃO] {len(sobreposicoes)} caso(s) de sobreposição de eventos na mesma data detectado(s)!")

    # 5. Localizar e conferir memorando de retificação (se houver)
    cod_busca = codigo_sec
    if not cod_busca:
        m_arq = re.match(r"^(\d{3})\b", caminho_secretaria.stem)
        if m_arq:
            cod_busca = m_arq.group(1)

    itens_atendidos_memo = []
    nome_memo_str = ""
    secao_memo = None
    if cod_busca:
        # Busca prioritariamente na pasta onde o relatório da secretaria está
        arq_memo = localizar_memorando(cod_busca, caminho_secretaria.parent)
        if arq_memo:
            nome_memo_str = arq_memo.name
            print(f"\n[MEMORANDO] Localizado memorando de retificação: {arq_memo.name}")
            try:
                texto_memo = extrair_texto_completo_memorando(arq_memo)
                dados_memo = extrair_itens_memorando(texto_memo, caminho_pdf=arq_memo)
                res_conf = conferir_memorando_com_solicitacoes(dados_memo, retificacoes)
                secao_memo = gerar_secao_memorando_markdown(res_conf, dados_memo, arq_memo.name)
                itens_atendidos_memo = res_conf.get("atendidas", [])
                print(f"  -> Itens lidos no memorando: {len(dados_memo['itens'])}")
                print(f"  -> Atendidas/Regularizadas: {len(res_conf['atendidas'])}")
                if res_conf['divergencias_memo']:
                    print(f"  -> Inconsistências no documento: {len(res_conf['divergencias_memo'])}")
                qtd_pend = len([p for p in res_conf['pendentes_restantes'] if _eh_divergencia_secretaria(p)])
                print(f"  -> Pendências restantes: {qtd_pend}")
            except Exception as err:
                print(f"  -> [AVISO] Falha ao processar memorando {arq_memo.name}: {err}")

    # 6. Definir pastas de saída mensais
    if pasta_saida_mes:
        mes_pasta_nome = pasta_saida_mes
    elif ano_ref and mes_ref:
        mes_pasta_nome = f"{ano_ref}-{str(mes_ref).zfill(2)}"
    else:
        p_name = caminho_secretaria.parent.name
        mes_pasta_nome = p_name if re.match(r"^\d{4}-\d{2}$", p_name) else "geral"

    dir_conferencia = OUTPUT_DIR / mes_pasta_nome / "conferencia"
    dir_retificacoes = OUTPUT_DIR / mes_pasta_nome / "retificacoes"
    dir_monitoramento = OUTPUT_DIR / mes_pasta_nome / "monitoramento"

    for _d in (dir_conferencia, dir_retificacoes, dir_monitoramento):
        _d.mkdir(parents=True, exist_ok=True)

    nome_saida = sanitizar_nome_arquivo(nome_orgao)
    caminho_xlsx = dir_conferencia / f"conferencia_{nome_saida}.xlsx"
    gerar_excel(retificacoes, sobreposicoes, str(caminho_xlsx), len(regs_sec), len(regs_sis), mes_label, nome_orgao=nome_orgao, itens_regularizados=itens_atendidos_memo, avisos_mes=avisos_mes)
    print(f"[OK] Excel salvo em: {caminho_xlsx.relative_to(BASE_DIR)}")

    caminho_md = dir_retificacoes / f"retificacoes_{nome_saida}.md"
    caminho_txt = dir_retificacoes / f"retificacoes_{nome_saida}.txt"

    if retificacoes or sobreposicoes or secao_memo or avisos_mes:
        gerar_markdown(retificacoes, sobreposicoes, str(caminho_md), mes_label, len(regs_sec), len(regs_sis), nome_orgao=nome_orgao, secao_memorando=secao_memo, itens_regularizados=itens_atendidos_memo, avisos_mes=avisos_mes)
        print(f"[OK] Markdown salvo em: {caminho_md.relative_to(BASE_DIR)}")

        gerar_txt(retificacoes, sobreposicoes, str(caminho_txt), itens_regularizados=itens_atendidos_memo, avisos_mes=avisos_mes)
        print(f"[OK] Texto puro salvo em: {caminho_txt.relative_to(BASE_DIR)}")
    else:
        if caminho_md.exists():
            caminho_md.unlink()
        if caminho_txt.exists():
            caminho_txt.unlink()
        print("[INFO] Nenhuma retificação necessária (100% de conformidade!).")

    # 7. Gerar relatório de monitoramento de atrasos e faltas acumuladas
    # Baixa e substitui as ocorrências que foram regularizadas pelo memorando
    from monitoring import aplicar_retificacoes_memorando
    regs_sec_monit, regularizados = aplicar_retificacoes_memorando(
        regs_sec, itens_atendidos_memo, nome_memorando=nome_memo_str
    )
    dados_monit = apurar_dados_monitoramento(regs_sec_monit, regs_sis, regularizados=regularizados)
    if dados_monit["atrasos"] or dados_monit["faltas"] or dados_monit.get("regularizados"):
        caminho_monit_md = dir_monitoramento / f"monitoramento_{nome_saida}.md"
        gerar_markdown_monitoramento(dados_monit, str(caminho_monit_md), mes_label, nome_orgao=nome_orgao, avisos_mes=avisos_mes)
        print(f"[OK] Monitoramento Markdown: {caminho_monit_md.relative_to(BASE_DIR)}")

        caminho_monit_xlsx = dir_monitoramento / f"monitoramento_{nome_saida}.xlsx"
        gerar_excel_monitoramento(dados_monit, str(caminho_monit_xlsx), mes_label, nome_orgao=nome_orgao, avisos_mes=avisos_mes)
        print(f"[OK] Monitoramento Excel:    {caminho_monit_xlsx.relative_to(BASE_DIR)}")

    return {
        "orgao": nome_orgao,
        "mes": mes_label,
        "pasta_mes": mes_pasta_nome,
        "registros_sec": len(regs_sec),
        "registros_sis": len(regs_sis),
        "retificacoes": len(retificacoes),
        "atrasos_minutos": dados_monit["resumo"]["total_minutos"],
        "faltas_acumuladas": dados_monit["resumo"]["total_dias_faltas"],
        "avisos_mes": avisos_mes,
    }


def descobrir_pares_na_pasta(pasta):
    """Encontra pares de secretaria e sistema dentro de uma pasta de mês."""
    pasta = Path(pasta)
    pares = []

    # 1. Se houver subpastas secretaria/ e sistema/
    sec_sub = pasta / "secretaria"
    sis_sub = pasta / "sistema"
    if sec_sub.is_dir() and sis_sub.is_dir():
        for f_sec in sorted(sec_sub.glob("*.pdf")):
            m = re.match(r"^(.+?)\s*-\s*secretaria\.pdf$", f_sec.name, re.IGNORECASE)
            if m:
                prefixo = m.group(1).strip()
                f_sis = sis_sub / f"{prefixo} - sistema.pdf"
                if f_sis.exists():
                    pares.append((f_sec, f_sis))
                    continue
            m_num = re.match(r"^(\d+)", f_sec.stem)
            if m_num:
                cod = m_num.group(1)
                sis_cands = list(sis_sub.glob(f"{cod}*.pdf"))
                if sis_cands:
                    pares.append((f_sec, sis_cands[0]))
                    continue
        if pares:
            return pares

    # 2. Arquivos diretamente na pasta do mês (sem subpastas)
    todos_pdfs = list(pasta.glob("*.pdf"))
    sec_files = [f for f in todos_pdfs if re.search(r"secretaria", f.name, re.IGNORECASE)]
    sis_files = [f for f in todos_pdfs if re.search(r"sistema", f.name, re.IGNORECASE)]

    for f_sec in sorted(sec_files):
        m = re.match(r"^(.+?)\s*-\s*secretaria\.pdf$", f_sec.name, re.IGNORECASE)
        if m:
            prefixo = m.group(1).strip()
            cand = [f for f in sis_files if re.match(rf"^{re.escape(prefixo)}\s*-\s*sistema\.pdf$", f.name, re.IGNORECASE)]
            if cand:
                pares.append((f_sec, cand[0]))
                continue

        m_num = re.match(r"^(\d+)", f_sec.stem)
        if m_num:
            cod = m_num.group(1)
            cand = [f for f in sis_files if re.match(rf"^{cod}\b", f.stem)]
            if cand:
                pares.append((f_sec, cand[0]))
                continue

    return pares


def descobrir_lotes(input_dir, mes_filtro=None):
    """Descobre lotes mensais em input/."""
    input_dir = Path(input_dir)
    lotes = []

    # Procura subpastas de meses na raiz de input
    for item in sorted(input_dir.iterdir()):
        if not item.is_dir():
            continue
        if item.name.lower() in ("secretaria", "sistema", "desligados", ".git", "__pycache__"):
            continue
        if mes_filtro and item.name != mes_filtro:
            continue

        pares = descobrir_pares_na_pasta(item)
        if pares:
            lotes.append({
                "mes_id": item.name,
                "pasta": item,
                "pares": pares,
            })

    # Compatibilidade com modo legado se não encontrou pastas mensais
    if not lotes and not mes_filtro:
        sec_dir = input_dir / "secretaria"
        sis_dir = input_dir / "sistema"
        if sec_dir.is_dir() and sis_dir.is_dir():
            pares_legado = descobrir_pares_na_pasta(input_dir)
            if pares_legado:
                lotes.append({
                    "mes_id": None,
                    "pasta": input_dir,
                    "pares": pares_legado,
                })

    return lotes


def selecionar_lotes_interativo(lotes_disponiveis):
    """Permite ao usuário escolher qual mês processar, sugerindo o mais recente como padrão."""
    if not lotes_disponiveis:
        return []

    # Ordena com o mais recente primeiro (ex: '2026-09', depois '2026-08')
    lotes_ordenados = sorted(lotes_disponiveis, key=lambda l: l["mes_id"] or "", reverse=True)

    # Se só houver 1 lote, processa ele diretamente
    if len(lotes_ordenados) == 1:
        print(f"\n[INFO] Mês identificado em input/: {lotes_ordenados[0]['mes_id'] or 'raiz'}")
        return lotes_ordenados

    # Se não for terminal interativo (ex: execução em background), assume o último mês cadastrado
    if not sys.stdin.isatty():
        print(f"\n[INFO] Modo automático/não interativo. Selecionando o último mês cadastrado: {lotes_ordenados[0]['mes_id']}")
        return [lotes_ordenados[0]]

    print("\n" + "=" * 65)
    print("  CONFERÊNCIA DE FREQUÊNCIA — SELEÇÃO DE MÊS")
    print("=" * 65)
    print("Meses identificados em input/:")
    for idx, lote in enumerate(lotes_ordenados, start=1):
        extra = " (Último mês cadastrado - PADRÃO)" if idx == 1 else ""
        print(f"  [{idx}] {lote['mes_id'] or 'raiz'} ({len(lote['pares'])} secretarias){extra}")
    print("  [T] Processar TODOS os meses")
    print("  [S] Sair")
    print("-" * 65)

    while True:
        try:
            escolha = input("Escolha o mês a processar [1]: ").strip().upper()
        except (EOFError, KeyboardInterrupt):
            print("\nOperação cancelada pelo usuário.")
            sys.exit(0)

        if escolha in ("", "1"):
            return [lotes_ordenados[0]]
        elif escolha == "T":
            return lotes_disponiveis
        elif escolha == "S":
            print("Operação encerrada.")
            sys.exit(0)
        elif escolha.isdigit():
            idx_num = int(escolha)
            if 1 <= idx_num <= len(lotes_ordenados):
                return [lotes_ordenados[idx_num - 1]]
        print("Opção inválida. Digite o número correspondente, 'T' para todos ou 'S' para sair.")


def main():
    args_raw = sys.argv[1:]
    ano_mes = None
    mes_filtro = None
    modo_todos = "--todos" in args_raw
    modo_ultimo = "--ultimo" in args_raw

    args = []
    i = 0
    while i < len(args_raw):
        a = args_raw[i]
        if a == "--mes" and i + 1 < len(args_raw):
            mes_filtro = args_raw[i + 1]
            if "-" in mes_filtro:
                ano_str, mes_str = mes_filtro.split("-", 1)
                if ano_str.isdigit() and mes_str.isdigit():
                    ano_mes = (int(ano_str), int(mes_str))
            i += 2
        elif a.startswith("--"):
            i += 1
        else:
            args.append(a)
            i += 1

    desligados, arq_desligados = localizar_e_carregar_desligados(INPUT_DIR)

    if len(args) == 2:
        caminho_secretaria, caminho_sistema = args
        processar_par(caminho_secretaria, caminho_sistema, ano_mes=ano_mes, desligados=desligados, pasta_saida_mes=mes_filtro)
    elif len(args) == 0:
        lotes = descobrir_lotes(INPUT_DIR, mes_filtro=mes_filtro)
        if not lotes:
            print("\n[AVISO] Nenhum par de relatórios encontrado em:")
            print(f"  - {INPUT_DIR}")
            print("\nEstrutura esperada por mês:")
            print("  input/<AAAA-MM>/<COD> - secretaria.pdf")
            print("  input/<AAAA-MM>/<COD> - sistema.pdf")
            print("  input/<AAAA-MM>/<COD> - retificação.pdf (opcional)")
            print("  input/Controle Desligamentos 2026.ods (na raiz de input)")
            print("\nUso manual:")
            print("  python src/main.py <pdf_secretaria> <pdf_sistema> [--mes AAAA-MM]\n")
            sys.exit(1)

        # Se não especificou filtro nem modo todos, seleciona via menu ou último cadastrado
        if not mes_filtro and not modo_todos:
            if modo_ultimo:
                lotes_ordenados = sorted(lotes, key=lambda l: l["mes_id"] or "", reverse=True)
                lotes_para_rodar = [lotes_ordenados[0]]
            else:
                lotes_para_rodar = selecionar_lotes_interativo(lotes)
        else:
            lotes_para_rodar = lotes

        print("\n" + "#" * 65)
        print(f"  CONFERÊNCIA DE FREQUÊNCIA — PROCESSAMENTO")
        print(f"  Mês(es) selecionado(s): {', '.join(l['mes_id'] or 'raiz' for l in lotes_para_rodar)}")
        if desligados:
            print(f"  Controle de Desligamentos: {len(desligados)} registros ativos ({arq_desligados})")
        print("#" * 65)

        todos_resultados = []
        for lote in lotes_para_rodar:
            mes_id = lote["mes_id"]
            pares = lote["pares"]
            print(f"\n>>> INICIANDO LOTE: {mes_id or 'raiz'} ({len(pares)} pares) <<<")

            ano_mes_lote = ano_mes
            if ano_mes_lote is None and mes_id and re.match(r"^\d{4}-\d{2}$", mes_id):
                a_str, m_str = mes_id.split("-")
                ano_mes_lote = (int(a_str), int(m_str))

            for sec, sis in pares:
                res = processar_par(sec, sis, ano_mes=ano_mes_lote, desligados=desligados, pasta_saida_mes=mes_id)
                todos_resultados.append(res)

        print("\n" + "#" * 65)
        print("RESUMO FINAL DO PROCESSAMENTO:")
        print("#" * 65)
        for r in todos_resultados:
            detalhe_extras = []
            if r.get("atrasos_minutos"):
                detalhe_extras.append(f"Atrasos: {r['atrasos_minutos']} min")
            if r.get("faltas_acumuladas"):
                detalhe_extras.append(f"Faltas: {r['faltas_acumuladas']} dia(s)")
            extras_str = f" | {' | '.join(detalhe_extras)}" if detalhe_extras else ""
            pasta_info = f" [{r['pasta_mes']}]" if r.get("pasta_mes") else ""
            aviso_flag = " ⚠️ [MÊS DIVERGENTE]" if r.get("avisos_mes") else ""
            print(f"  *{pasta_info} {r['orgao']} ({r['mes']}): {r['retificacoes']} retificação(ões){extras_str}{aviso_flag}")
            if r.get("avisos_mes"):
                for av in r["avisos_mes"]:
                    print(f"     -> ⚠️ {av}")
        print("#" * 65 + "\n")

        # Atualiza automaticamente a base consolidada para o Power BI
        try:
            print("[Power BI] Atualizando base de dados consolidada para o dashboard...")
            consolidar_monitoramento()
        except Exception as e:
            print(f"[Power BI] Aviso ao consolidar dados para o Power BI: {e}")
    else:
        print("Uso: python src/main.py <pdf_secretaria> <pdf_sistema> [--mes AAAA-MM]")
        print("  Ou rode apenas 'python src/main.py' para selecionar e processar os lotes de input/")
        sys.exit(1)


if __name__ == "__main__":
    main()
