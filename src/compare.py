# -*- coding: utf-8 -*-
"""
Compara os registros extraídos da SECRETARIA com os do SISTEMA dia a dia,
dentro do mês de referência, e monta a lista de "retificações": para
cada matrícula e cada dia em que o tipo registrado pela secretaria não
bate com o tipo registrado pelo sistema, gera uma linha no formato

    matricula - nome - data(s) - X dia(s) - tipo_secretaria --> tipo_sistema

(dias consecutivos com o mesmo par "de/para" são agrupados num intervalo,
igual ao que se faz na conferência manual).

Comparar dia a dia — em vez de somar quantidades totais por tipo — já
resolve sozinho o problema de eventos que também ocupam dias de outro
mês: nos dias que efetivamente caem dentro do mês de referência, o tipo
bate normalmente, mesmo que a duração total do evento (fora do mês) não
bata com o que está na secretaria.
"""
import unicodedata
import re
from datetime import date, timedelta
from collections import Counter

try:
    from desligados import normalizar_matricula
except ImportError:
    def normalizar_matricula(m):
        return re.sub(r"\D", "", str(m)).zfill(6) if m else ""

# Tipos que representam "nada a registrar" e não entram na comparação
TIPOS_IGNORADOS = {"frequencia normal"}

# Aliases: grafias diferentes para o mesmo tipo de ocorrência
# (chave e valor já devem estar normalizados por _normalizar)
ALIASES = {
    "falta": "falta",
    "faltas": "falta",
    "faltas efetivos": "falta",
    "faltas clt": "falta",
    "falta efetivos": "falta",
    "falta clt": "falta",
    "minutos perdidos": "minutos perdidos",
    "minutos perdidos clt": "minutos perdidos",
    "minutos perdidos clt-medico plantonista": "minutos perdidos",
    "minutos perd efetivos": "minutos perdidos",
    "minutos perdidos efetivos": "minutos perdidos",
}

ROTULO_SEM_REGISTRO_SECRETARIA = "SEM REGISTRO"
ROTULO_SEM_REGISTRO_SISTEMA = "sem registro em sistema"
ROTULO_FREQUENCIA_NORMAL = "Frequência normal"


def _normalizar(texto):
    """minúsculas, sem acento, espaços colapsados — para comparar tipos
    escritos com capitalização/formatação diferente nos dois relatórios."""
    if not texto:
        return ""
    nfkd = unicodedata.normalize("NFKD", texto)
    sem_acento = "".join(c for c in nfkd if not unicodedata.combining(c))
    return " ".join(sem_acento.lower().split())


def _tipo_canonico(texto):
    if not texto:
        return ""
    norm = _normalizar(texto)
    if not norm:
        return ""

    # Minutos perdidos / atrasos: mesmo evento independentemente do rótulo
    # (ex.: 'minutos perdidos', 'minutos perdidos clt', 'minutos perd efetivos')
    if "minut" in norm and ("perd" in norm or "atras" in norm):
        return "minutos perdidos"

    # Faltas: mesmo evento independentemente do rótulo
    # (ex.: 'falta', 'faltas', 'faltas clt', 'faltas efetivos')
    if "falta" in norm:
        return "falta"

    return ALIASES.get(norm, norm)


def _parse_data(txt):
    if not txt:
        return None
    try:
        d, m, a = txt.split("/")
        return date(int(a), int(m), int(d))
    except (ValueError, AttributeError):
        return None


def mes_referencia(regs_secretaria):
    """Descobre o mês/ano de referência do relatório a partir da moda das
    datas dos registros da secretaria. Retorna (ano, mes) ou None."""
    contagem = Counter()
    for r in regs_secretaria:
        d = _parse_data(r.get("data"))
        if d:
            contagem[(d.year, d.month)] += 1
    if not contagem:
        return None
    return contagem.most_common(1)[0][0]


def _limites_mes(ano, mes):
    primeiro = date(ano, mes, 1)
    ultimo_excl = (date(ano, mes + 1, 1) if mes < 12 else date(ano + 1, 1, 1))
    ultimo_incl = ultimo_excl - timedelta(days=1)
    return primeiro, ultimo_incl


def _expandir_dias(registros, campo_tipo, campo_data_ini, campo_data_fim,
                    campo_qtde, primeiro_dia, ultimo_dia_incl):
    """Expande cada ocorrência em entradas por dia, recortadas ao mês.
    Retorna (mapa, nomes):
      mapa: {(matricula, date): [rotulo_original_1, rotulo_original_2, ...]}
      nomes: {matricula: nome}
    """
    mapa = {}
    nomes = {}
    for r in registros:
        matricula = r["matricula"]
        nomes.setdefault(matricula, " ".join(r["nome"].split()))

        tipo_raw = r[campo_tipo]
        tipo_canon = _tipo_canonico(tipo_raw)
        if tipo_canon in TIPOS_IGNORADOS:
            continue

        d_ini = _parse_data(r.get(campo_data_ini))
        if d_ini is None:
            continue

        if campo_data_fim:
            d_fim = _parse_data(r.get(campo_data_fim)) or d_ini
        elif "minut" in tipo_canon:
            # Minutos perdidos é um evento pontual no dia informado;
            # o campo qtde representa minutos acumulados no mês, NÃO dias de calendário!
            d_fim = d_ini
        else:
            # a secretaria só informa a data de início + quantidade de
            # dias corridos a partir dela
            qtd = r.get(campo_qtde) or 1
            try:
                qtd_i = max(int(round(qtd)), 1)
            except (TypeError, ValueError):
                qtd_i = 1
            d_fim = d_ini + timedelta(days=qtd_i - 1)

        inicio = max(d_ini, primeiro_dia)
        fim = min(d_fim, ultimo_dia_incl)
        d = inicio
        while d <= fim:
            lista = mapa.setdefault((matricula, d), [])
            if tipo_raw not in lista:
                lista.append(tipo_raw)
            d += timedelta(days=1)

    return mapa, nomes


def detectar_sobreposicoes(mapa_dias, nomes, primeiro_dia, ultimo_dia_incl, origem="Sistema"):
    """Identifica quando há mais de um tipo canônico diferente registrado no mesmo dia
    para o mesmo servidor. Agrupa períodos contínuos com a mesma sobreposição.
    Retorna lista de dicts:
      matricula, nome, origem, data_inicio, data_fim, dias, eventos, observacao
    """
    sobreposicoes = []
    for matricula in sorted(nomes):
        nome = nomes[matricula]
        pendente = None

        def fechar():
            nonlocal pendente
            if pendente is None:
                return
            dias = (pendente["fim"] - pendente["inicio"]).days + 1
            tipos_str = " + ".join(pendente["tipos"])
            sobreposicoes.append({
                "matricula": matricula,
                "nome": nome,
                "origem": origem,
                "data_inicio": pendente["inicio"],
                "data_fim": pendente["fim"],
                "dias": dias,
                "eventos": tipos_str,
                "observacao": f"Sobreposição no {origem.lower()}: {tipos_str}",
            })
            pendente = None

        d = primeiro_dia
        while d <= ultimo_dia_incl:
            lista = mapa_dias.get((matricula, d), [])
            tipos_unicos = []
            tipos_canons_vistos = set()
            for t in lista:
                canon = _tipo_canonico(t)
                if canon not in tipos_canons_vistos:
                    tipos_canons_vistos.add(canon)
                    tipos_unicos.append(t)

            if len(tipos_unicos) > 1:
                tipos_tuple = tuple(sorted(tipos_unicos))
                if pendente and pendente["tipos_tuple"] == tipos_tuple:
                    pendente["fim"] = d
                else:
                    fechar()
                    pendente = {
                        "tipos": list(tipos_tuple),
                        "tipos_tuple": tipos_tuple,
                        "inicio": d,
                        "fim": d,
                    }
            else:
                fechar()
            d += timedelta(days=1)
        fechar()

    sobreposicoes.sort(key=lambda x: (x["matricula"], x["data_inicio"]))
    return sobreposicoes


def comparar_por_dia(regs_secretaria, regs_sistema, ano_mes=None, desligados=None):
    """Retorna (retificacoes, sobreposicoes, (ano_ref, mes_ref)).

    Cada item de `retificacoes`:
      matricula, nome, data_inicio, data_fim, dias,
      tipo_secretaria (rótulo ou "SEM REGISTRO"),
      tipo_sistema (rótulo ou "sem registro em sistema"),
      observacao (str com detalhes, ex: sobreposição),
      sobreposicao (bool indicando conflito de eventos na mesma data)

    Cada item de `sobreposicoes`:
      matricula, nome, origem, data_inicio, data_fim, dias, eventos, observacao
    """
    ano_ref, mes_ref = ano_mes or mes_referencia(regs_secretaria) or (None, None)
    if not ano_ref:
        return [], [], (None, None)

    primeiro_dia, ultimo_dia_incl = _limites_mes(ano_ref, mes_ref)

    mapa_sec, nomes_sec = _expandir_dias(
        regs_secretaria, "ocorrencia", "data", None, "qtde_dias",
        primeiro_dia, ultimo_dia_incl)
    mapa_sis, nomes_sis = _expandir_dias(
        regs_sistema, "descricao", "data_inicial", "data_final", "qtde_dias",
        primeiro_dia, ultimo_dia_incl)

    nomes = {**nomes_sis, **nomes_sec}  # nomes_sec prevalece (fonte com todo mundo)

    # Identifica sobreposições em ambos os lados
    sobreposicoes_sis = detectar_sobreposicoes(mapa_sis, nomes, primeiro_dia, ultimo_dia_incl, origem="Sistema")
    sobreposicoes_sec = detectar_sobreposicoes(mapa_sec, nomes, primeiro_dia, ultimo_dia_incl, origem="Secretaria")
    sobreposicoes = sobreposicoes_sis + sobreposicoes_sec

    # Mapeia servidores com minutos perdidos no mês em cada lado
    mats_minutos_sec = set()
    for (mat, d), tipos in mapa_sec.items():
        if any(_tipo_canonico(t) == "minutos perdidos" for t in tipos):
            mats_minutos_sec.add(mat)

    mats_minutos_sis = set()
    for (mat, d), tipos in mapa_sis.items():
        if any(_tipo_canonico(t) == "minutos perdidos" for t in tipos):
            mats_minutos_sis.add(mat)

    # Servidores com minutos perdidos em ambos os relatórios no mês
    mats_minutos_em_ambos = mats_minutos_sec & mats_minutos_sis

    retificacoes = []
    for matricula in sorted(nomes):
        nome = nomes[matricula]
        pendente = None  # dict: par, inicio, fim

        def fechar():
            nonlocal pendente
            if pendente is None:
                return
            rotulo_sec, rotulo_sis, obs, sob = pendente["par"]
            dias = (pendente["fim"] - pendente["inicio"]).days + 1
            if rotulo_sec:
                tipo_sec = rotulo_sec
            elif matricula in nomes_sec:
                tipo_sec = ROTULO_FREQUENCIA_NORMAL
            else:
                tipo_sec = ROTULO_SEM_REGISTRO_SECRETARIA

            mat_norm = normalizar_matricula(matricula)
            info_desl = desligados.get(mat_norm) if desligados else None

            # 1. Regra solicitada: Quando a ocorrência NÃO estiver no arquivo do sistema
            # (ou seja, rotulo_sis é vazio / sem registro em sistema) e o servidor for desligado,
            # IGNORA O DESLIGADO!
            if not rotulo_sis and info_desl:
                pendente = None
                return

            # 2. Quando o servidor tiver lançamentos no sistema, mas NÃO na secretaria (SEM REGISTRO):
            # Se for confirmado desligado, sinaliza a confirmação do desligamento
            if not rotulo_sec and matricula not in nomes_sec and info_desl:
                dt_str = f" em {info_desl['data_demissao']}" if info_desl.get("data_demissao") else ""
                obs = f"Servidor desligado{dt_str} (confirmado em Controle de Desligamentos)"

            retificacoes.append({
                "matricula": matricula,
                "nome": nome,
                "data_inicio": pendente["inicio"],
                "data_fim": pendente["fim"],
                "dias": dias,
                "tipo_secretaria": tipo_sec,
                "tipo_sistema": rotulo_sis or ROTULO_SEM_REGISTRO_SISTEMA,
                "observacao": obs,
                "sobreposicao": sob,
                "desligado": bool(info_desl),
                "data_demissao": info_desl.get("data_demissao") if info_desl else None,
            })
            pendente = None

        d = primeiro_dia
        while d <= ultimo_dia_incl:
            itens_sec = mapa_sec.get((matricula, d), [])
            itens_sis = mapa_sis.get((matricula, d), [])

            # Desduplica por chave canônica preservando rótulo
            tipos_sec_unicos = []
            canons_sec = set()
            for t in itens_sec:
                c = _tipo_canonico(t)
                if c not in canons_sec:
                    canons_sec.add(c)
                    tipos_sec_unicos.append(t)

            tipos_sis_unicos = []
            canons_sis = set()
            for t in itens_sis:
                c = _tipo_canonico(t)
                if c not in canons_sis:
                    canons_sis.add(c)
                    tipos_sis_unicos.append(t)

            divergente = (canons_sec != canons_sis)

            # Ajuste de Minutos Perdidos em dias diferentes:
            # Minutos perdidos é apurado no acumulado do mês. Se o servidor possui minutos
            # perdidos registrados em AMBOS os relatórios no mês, o lançamento em datas
            # diferentes com frequência normal no dia oposto NÃO deve ser considerado divergência.
            if divergente and matricula in mats_minutos_em_ambos:
                if (canons_sec == {"minutos perdidos"} and not canons_sis) or \
                   (canons_sis == {"minutos perdidos"} and not canons_sec):
                    divergente = False

            # Caso especial: Aguardando retorno/perícia auxílio doença quando presente em AMBOS os arquivos
            # Deve ser solicitada retificação para atualização da ocorrência junto ao SEMPEM
            ambos_aguardando_retorno = False
            if not divergente and canons_sec:
                for c in canons_sec:
                    if "aguardando retorno" in c or ("pericia" in c and "auxilio" in c):
                        ambos_aguardando_retorno = True
                        break

            if divergente or ambos_aguardando_retorno:
                tem_sob_sec = len(canons_sec) > 1
                tem_sob_sis = len(canons_sis) > 1
                tem_sob = tem_sob_sec or tem_sob_sis

                tipos_sec_unicos.sort()
                tipos_sis_unicos.sort()

                val_sec = " + ".join(tipos_sec_unicos) if tipos_sec_unicos else None
                val_sis = " + ".join(tipos_sis_unicos) if tipos_sis_unicos else None

                obs_partes = []
                if tem_sob_sis:
                    val_sis = f"{val_sis} [SOBREPOSIÇÃO]"
                    obs_partes.append(f"Sobreposição no sistema: {' + '.join(tipos_sis_unicos)}")
                if tem_sob_sec:
                    val_sec = f"{val_sec} [SOBREPOSIÇÃO]"
                    obs_partes.append(f"Sobreposição na secretaria: {' + '.join(tipos_sec_unicos)}")

                # 1. Quando na secretaria for Aguardando perícia sempem e no sistema estiver sem registro
                if val_sec and "sempem" in _normalizar(val_sec) and not val_sis:
                    obs_partes.append("solicitar ao sempem")

                # Quando na secretaria for Minutos perdidos ou Falta e no sistema estiver sem registro
                if val_sec and _tipo_canonico(val_sec) in ("minutos perdidos", "falta") and not val_sis:
                    if "inserir no sistema" not in obs_partes:
                        obs_partes.append("inserir no sistema")

                # 2. Quando constar no Sistema como Aguardando perícia / retorno
                # (esteja em ambos os arquivos ou com outra ocorrência na secretaria)
                sis_eh_aguardando_pericia = val_sis and any(
                    "aguardando" in c and ("pericia" in c or "retorno" in c or "sempem" in c)
                    for c in canons_sis
                )
                if ambos_aguardando_retorno or sis_eh_aguardando_pericia:
                    if "solicitar ao sempem atualização da ocorrencia" not in obs_partes:
                        obs_partes.append("solicitar ao sempem atualização da ocorrencia")

                obs = "; ".join(obs_partes)
                par = (val_sec, val_sis, obs, tem_sob)

                if pendente and pendente["par"] == par:
                    pendente["fim"] = d
                else:
                    fechar()
                    pendente = {"par": par, "inicio": d, "fim": d}
            else:
                fechar()
            d += timedelta(days=1)
        fechar()

    retificacoes = _consolidar_conflitos_saude_faltas(
        retificacoes, regs_secretaria, regs_sistema, primeiro_dia, ultimo_dia_incl
    )
    retificacoes.sort(key=lambda x: (x["matricula"], x["data_inicio"]))
    return retificacoes, sobreposicoes, (ano_ref, mes_ref)


TIPOS_SAUDE = {
    "tratamento de saude",
    "aguardando retorno/pericia auxilio doenc",
    "auxilio doenca",
    "acidente de trabalho",
}

TIPOS_FALTA = {
    "falta",
    "faltas efetivos",
    "faltas clt",
}


def _consolidar_conflitos_saude_faltas(retificacoes, regs_sec, regs_sis, primeiro_dia, ultimo_dia_incl):
    """Consolida períodos contínuos de afastamento por saúde (ex.: Tratamento de Saúde)
    que coincidem com faltas, unificando retificações fragmentadas e expressando
    claramente a ação e a responsabilidade de cada parte (Secretaria vs Sistema/RH)."""
    saude_sis = {}
    for r in regs_sis:
        desc = r.get("descricao") or ""
        if _tipo_canonico(desc) in TIPOS_SAUDE:
            mat = r["matricula"]
            d_ini = _parse_data(r.get("data_inicial"))
            if not d_ini:
                continue
            d_fim = _parse_data(r.get("data_final")) or d_ini
            ini_rec = max(d_ini, primeiro_dia)
            fim_rec = min(d_fim, ultimo_dia_incl)
            if ini_rec <= fim_rec:
                saude_sis.setdefault(mat, []).append((ini_rec, fim_rec, desc))

    saude_sec = {}
    for r in regs_sec:
        ocorr = r.get("ocorrencia") or ""
        if _tipo_canonico(ocorr) in TIPOS_SAUDE:
            mat = r["matricula"]
            d_ini = _parse_data(r.get("data"))
            if not d_ini:
                continue
            qtd = r.get("qtde_dias") or 1
            try:
                qtd_i = max(int(round(qtd)), 1)
            except (TypeError, ValueError):
                qtd_i = 1
            d_fim = d_ini + timedelta(days=qtd_i - 1)
            ini_rec = max(d_ini, primeiro_dia)
            fim_rec = min(d_fim, ultimo_dia_incl)
            if ini_rec <= fim_rec:
                saude_sec.setdefault(mat, []).append((ini_rec, fim_rec, ocorr))

    novas_ret = []
    por_mat = {}
    for r in retificacoes:
        por_mat.setdefault(r["matricula"], []).append(r)

    for mat, lista in por_mat.items():
        # Caso 1: Tratamento de Saúde no Sistema e Faltas na Secretaria
        if mat in saude_sis:
            for ini_s, fim_s, desc_s in saude_sis[mat]:
                no_int = [r for r in lista if ini_s <= r["data_inicio"] and r["data_fim"] <= fim_s]
                if no_int and any(_tipo_canonico(r["tipo_secretaria"]) in TIPOS_FALTA for r in no_int):
                    tem_falta_sis = any(r.get("sobreposicao") for r in no_int)
                    ret_cons = {
                        "matricula": mat,
                        "nome": no_int[0]["nome"],
                        "data_inicio": ini_s,
                        "data_fim": fim_s,
                        "dias": (fim_s - ini_s).days + 1,
                        "tipo_secretaria": "Falta",
                        "tipo_sistema": desc_s,
                        "instrucao_direta": f"as faltas que coincidem com {desc_s.lower()} devem ser retificadas.",
                        "responsabilidade": "Secretaria",
                        "nota_responsabilidade": (
                            f"Como as faltas constam no relatório da secretaria, cabe à secretaria retificar. "
                            f"Adicionalmente, as faltas lançadas no sistema no mesmo período são de responsabilidade do RH/sistema retirar."
                            if tem_falta_sis else
                            f"Como as faltas constam no relatório da secretaria, cabe à secretaria retificar."
                        ),
                        "observacao": (
                            f"As faltas relatadas pela secretaria coincidem com {desc_s.lower()} registrado no sistema "
                            f"e devem ser retificadas pela secretaria. No sistema também constam faltas que devem ser retiradas pelo RH."
                            if tem_falta_sis else
                            f"As faltas relatadas pela secretaria coincidem com {desc_s.lower()} registrado no sistema "
                            f"e devem ser retificadas pela secretaria."
                        ),
                        "sobreposicao": tem_falta_sis,
                    }
                    lista = [r for r in lista if r not in no_int]
                    lista.append(ret_cons)

        # Caso 2: Tratamento de Saúde na Secretaria e Faltas no Sistema
        if mat in saude_sec:
            for ini_s, fim_s, desc_s in saude_sec[mat]:
                no_int = [r for r in lista if ini_s <= r["data_inicio"] and r["data_fim"] <= fim_s]
                if no_int and any(_tipo_canonico(r["tipo_sistema"]) in TIPOS_FALTA for r in no_int):
                    ret_cons = {
                        "matricula": mat,
                        "nome": no_int[0]["nome"],
                        "data_inicio": ini_s,
                        "data_fim": fim_s,
                        "dias": (fim_s - ini_s).days + 1,
                        "tipo_secretaria": desc_s,
                        "tipo_sistema": "Faltas Efetivos",
                        "instrucao_direta": f"as faltas registradas no sistema que coincidem com {desc_s.lower()} devem ser retiradas no sistema pelo RH.",
                        "responsabilidade": "Sistema (RH)",
                        "nota_responsabilidade": (
                            f"No relatório da secretaria consta {desc_s.lower()}; "
                            f"as faltas registradas no sistema são de responsabilidade do RH/sistema retirar."
                        ),
                        "observacao": f"A secretaria relatou {desc_s.lower()}; as faltas no sistema são de responsabilidade do RH/sistema retirar.",
                        "sobreposicao": True,
                    }
                    lista = [r for r in lista if r not in no_int]
                    lista.append(ret_cons)

        novas_ret.extend(lista)

    return novas_ret
