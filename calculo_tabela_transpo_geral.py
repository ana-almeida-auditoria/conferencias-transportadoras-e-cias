# =============================================================================
# SCRIPT UNIFICADO: calculo_tabela_cia_v2.py
# Busca AWBs por período/fornecedor, localiza o trecho (com match de zona) e
# calcula o valor de tabela (tipos 1, 4, 5, 8, 11, 14, 23, 32 + percentual_frete)
# =============================================================================

import math
import os
import re
from collections import defaultdict
from pathlib import Path

import pandas as pd
import mysql.connector
from dotenv import load_dotenv
from mysql.connector import Error

# =============================================================================
# CONFIGURAÇÕES
# =============================================================================
DATA_INICIAL = "2026-08-20"   # período de busca (data_entrada)
DATA_FINAL   = "2026-09-20"
FORNECEDOR   = ["89", "169715", "91", "47719", "31432", "12414", "588319", "98292"]            # id do fornecedor (notas.fornecedor / fornecedores.id_local)

ARQUIVO_SAIDA = "C:\\Users\\ana.almeida\\Downloads\\resultado_tabela_cia.xlsx"

TAM_BLOCO = 2000   # máximo de ids por IN (...) em cada consulta


# =============================================================================
# CONEXÃO COM O BANCO
# =============================================================================
load_dotenv(Path(__file__).parent / ".env")
print("HOST:", os.getenv("DB_HOST"))

DB_CONFIG = {
    "host": os.getenv("DB_HOST"),
    "port": int(os.getenv("DB_PORT", 3306)),
    "user": os.getenv("DB_USER"),
    "password": os.getenv("DB_PASSWORD"),
    "database": os.getenv("DB_NAME"),
}

def conectar_mysql():
    try:
        conn = mysql.connector.connect(**DB_CONFIG)
        if conn.is_connected():
            return conn
    except Error as e:
        print(f"Erro ao conectar: {e}")
        return None

# =============================================================================
# MÓDULO DE CÁLCULO DE TABELA (tipos 1, 4, 5, 6, 8, 11, 14, 23, 32 + TDE(24) + percentual_frete)
# =============================================================================
TIPOS_VALIDOS = {1, 4, 5, 6, 8, 11, 14, 22, 23, 25, 32}  # calculados no fluxo normal (calcular_tipos)

TIPOS_ADVALOREM           = {32}        # base = valor_transp; MAX((valor_transp * excedente/100) + franquia, minimo)
TIPOS_ADVALOREM_DIRETO    = {4, 5}      # base = valor_transp direto (sem ratio)
TIPOS_FRACAO              = {23}        # CEIL(peso / fracao) * excedente (sem franquia/minimo)
TIPOS_FRACAO_PROPORCIONAL = {14}        # CEIL(peso / fracao) * excedente (igual tipo 23, sem franquia/minimo)
TIPOS_NORMAIS             = {1, 6, 8, 11, 22, 25}  # (base - franquia) * excedente + minimo
# Tipo 6 (Despacho) e mutuamente exclusivo com tipo 1: se o trecho tiver tipo 6, tipo 1 e ignorado.

TIPO_TDE = 24  # calculado FORA do fluxo normal: depende de fornecedores.tad e usa nf_total como base

# --- Desconto por cliente (relatorios.cliente_trecho) ---
FORNECEDORES_CLIENTE_TRECHO = {"169715"}  # so consulta a cliente_trecho para estes fornecedores (JEM)
TIPOS_FRETE_PESO = {8}      # com desconto: excedente*(1-desc_peso) + minimo*(1-desc_minimo), somados
TIPOS_ADV        = {4, 32}  # adv_exc_fixo (> 0) substitui o excedente da faixa
TIPO_GRIS        = 5        # gris_exc_fixo (> 0) substitui o excedente da faixa


def _localizar_faixa(base: float, faixas: list[dict], ratio: float | None = None) -> dict | None:
    """
    Localiza a faixa onde: inicio <= base <= fim
    Se fim for NULL/0, não há limite superior (aceita qualquer base >= inicio).

    Tipo 32 (ad valorem por valor_transp): localiza faixa por valor_transp (inicio <= valor_transp <= fim).
    Tipos 1, 4, 5, 8, 11: localizam a faixa pela propria base (peso ou valor_transp).
    """
    comparador = ratio if ratio is not None else base

    for faixa in faixas:
        inicio = float(faixa.get("inicio", 0) or 0)
        fim    = faixa.get("fim")
        fim    = float(fim) if fim else 0

        if fim == 0:
            if comparador >= inicio:
                return faixa
        else:
            if inicio <= comparador <= fim:
                return faixa
    return None


def _calcular_faixa(base: float, faixa: dict, tipo: int, ratio: float | None = None, desc: dict | None = None) -> float:
    """
    Aplica o calculo de uma faixa individual conforme o tipo:

    Tipos normais (1, 6, 11): base = peso
      resultado = MAX((base - franquia) * excedente, minimo)  -- usa MAX, nao soma

    Tipo 8: base = peso
      resultado = MAX((base - franquia) * excedente, minimo)  -- igual tipos normais

    Tipo 4 (ad valorem direto, por valor_transp):
      resultado = minimo + (valor_transp - franquia, se > 0) * (excedente / 100)
      faixa localizada diretamente por valor_transp (nao por ratio)

    Tipo 5 (GRIS por valor_transp):
      resultado = MAX(valor_transp * (excedente / 100), minimo)
      sem franquia; se calculado < minimo cobra minimo; sem somar minimo ao calculado

    Tipo 32 (Advalorem por valor_transp):
      resultado = MAX((valor_transp * (excedente / 100)) + franquia, minimo)
      franquia = taxa fixa por minuta; localiza faixa por valor_transp (inicio <= valor_transp <= fim)

    Tipo fracao (23):
      resultado = CEIL(peso / fracao) * excedente_por_fracao
      onde fracao = faixa.inicio  (nao usa franquia nem minimo)

    Tipo 14 (pedagio por fracao com CEIL):
      resultado = CEIL(peso / fracao) * excedente
      onde fracao = faixa.inicio  (igual tipo 23: sem franquia, sem minimo, com arredondamento para cima)

    desc (linha da relatorios.cliente_trecho, opcional):
      Frete peso (TIPOS_FRETE_PESO): sem desconto = MAX(excedente, minimo)
                                     com desconto = excedente * (1 - desconto_frete_peso/100)
                                                  + minimo    * (1 - desconto_frete_minimo/100)
      Advalorem (TIPOS_ADV)        : adv_exc_fixo  > 0 substitui o excedente
      GRIS (TIPO_GRIS)             : gris_exc_fixo > 0 substitui o excedente
    """
    minimo    = float(faixa.get("minimo",    0) or 0)
    franquia  = float(faixa.get("franquia",  0) or 0)
    excedente = float(faixa.get("excedente", 0) or 0)

    fator_exc   = 1.0
    soma_partes = False
    if desc:
        if tipo in TIPOS_FRETE_PESO:
            soma_partes = True
            fator_exc = 1 - float(desc.get("desconto_frete_peso")   or 0) / 100
            minimo   *= 1 - float(desc.get("desconto_frete_minimo") or 0) / 100
        elif tipo in TIPOS_ADV and float(desc.get("adv_exc_fixo") or 0) > 0:
            excedente = float(desc["adv_exc_fixo"])
        elif tipo == TIPO_GRIS and float(desc.get("gris_exc_fixo") or 0) > 0:
            excedente = float(desc["gris_exc_fixo"])

    # --- Tipos 23 e 14: fracao com CEIL (sem franquia/minimo) ---
    if tipo in TIPOS_FRACAO or tipo in TIPOS_FRACAO_PROPORCIONAL:
        fracao = float(faixa.get("inicio", 1) or 1)
        return math.ceil(base / fracao) * excedente

    # --- Tipo 32: advalorem por valor_transp — MAX((valor_transp * excedente/100) + franquia, minimo) ---
    if tipo in TIPOS_ADVALOREM:
        calculado = (base * (excedente / 100)) + franquia
        return max(calculado, minimo) if minimo > 0 else calculado

    # --- Tipo 5: GRIS por valor_transp — MAX(calculado, minimo), sem franquia, sem somar minimo ---
    if tipo == 5:
        calculado = base * (excedente / 100) if excedente > 0 else 0
        return max(calculado, minimo) if minimo > 0 else calculado

    # --- Tipo 4: ad valorem direto (base = valor_transp, faixa localizada por valor_transp) ---
    if tipo == 4:
        base_exc = (base - franquia) if franquia > 0 else base
        Y = (base_exc * (excedente / 100)) if excedente > 0 else 0
        Z = minimo if minimo > 0 else 0
        return Y + Z

    # --- Tipos normais (1, 6, 8, 11): base = peso ---
    # Tipo 6 (Despacho) segue a mesma logica dos tipos normais.
    # Tipo 8: usa MAX(calculado, minimo) em vez de somar o minimo.
    # Com franquia : base_exc = peso - franquia
    # Sem franquia : base_exc = peso
    if franquia > 0 and base <= franquia:
        return minimo

    base_exc  = (base - franquia) if franquia  > 0 else base
    calculado = (base_exc * excedente * fator_exc) if excedente > 0 else 0

    if tipo == 8:
        if soma_partes:
            return calculado + minimo
        return max(calculado, minimo) if minimo > 0 else calculado

    # Tipos 1, 6, 11: soma minimo
    Z = minimo if minimo > 0 else 0
    return calculado + Z


def calcular_tipos(peso: float, valor_transp: float, faixas_por_tipo: dict[int, list[dict]], desc: dict | None = None) -> dict:
    """
    Recebe o peso e valor_transp do AWB e um dict com as faixas agrupadas por tipo.
    Retorna o resultado de cada tipo e o valor total do trecho antes do aero minimo.
    """
    resultados = {}

    # Tipo 6 (Despacho) e mutuamente exclusivo com tipo 1: se o trecho tiver tipo 6, tipo 1 e ignorado.
    tipos_a_ignorar = set()
    if 6 in faixas_por_tipo:
        tipos_a_ignorar.add(1)

    for tipo, faixas in faixas_por_tipo.items():
        if tipo not in TIPOS_VALIDOS:
            continue
        if tipo in tipos_a_ignorar:
            continue

        # Define a base de localizacao/calculo de cada tipo
        if tipo in TIPOS_ADVALOREM:
            base  = valor_transp
            ratio = None
        elif tipo in TIPOS_ADVALOREM_DIRETO:
            base  = valor_transp
            ratio = None
        else:
            # 1, 8, 11, 14, 23 -> calculados por peso
            base  = peso
            ratio = None

        if tipo in TIPOS_FRACAO or tipo in TIPOS_FRACAO_PROPORCIONAL:
            # Tipos 14 e 23: o campo "inicio" representa o tamanho da fracao, nao um limite
            # inferior de peso. Por isso NAO passam pelo _localizar_faixa.
            faixa_match = faixas[0] if faixas else None
        else:
            faixa_match = _localizar_faixa(base, faixas, ratio=ratio)

        if faixa_match is None:
            resultados[tipo] = 0.0
            continue

        resultados[tipo] = _calcular_faixa(base, faixa_match, tipo, ratio=ratio, desc=desc)

    # Separa tipos de fracao (23 e 14) dos demais para aplicar aero minimo corretamente.
    # Tipos 23 e 14 ficam fora do aero minimo e sao somados por ultimo.
    valor_sem_fracao = sum(v for t, v in resultados.items() if t not in (23, 14))
    valor_fracao      = resultados.get(23, 0.0) + resultados.get(14, 0.0)

    return {
        "por_tipo":                resultados,
        "total_antes_aero_minimo": sum(resultados.values()),
        "valor_sem_fracao":        valor_sem_fracao,
        "valor_fracao":            valor_fracao,
    }


def calcular_tde(nf_total: float, faixas_tipo_24: list[dict], tem_tde: bool) -> float:
    """
    Calcula o valor do TDE (tipo 24), somado POR FORA do fluxo normal de tipos.

    Gate: so calcula se tem_tde=True (fornecedores.tad = 1 para o notas.destinatario).
    Se tad = 0 ou nao ha faixa de tipo 24 cadastrada para o trecho, retorna 0.

    Base de localizacao/calculo: nf_total (valor do AWB), NAO peso nem valor_transp.
    Formula: mesma logica dos tipos normais (minimo + franquia + excedente),
    incluindo a regra "se nf_total <= franquia: resultado = minimo".
    """
    if not tem_tde or not faixas_tipo_24:
        return 0.0

    faixa = _localizar_faixa(nf_total, faixas_tipo_24)
    if faixa is None:
        return 0.0

    return _calcular_faixa(nf_total, faixa, TIPO_TDE)


def aplicar_aero_minimo(valor_sem_fracao: float, valor_fracao: float, aero_minimo: float) -> float:
    """
    Aplica a regra do aero minimo do trecho:
    - Aero minimo compara apenas com os tipos SEM o tipo 23
    - Tipo 23 (fracao) sempre e somado por ultimo
    valor_com_minimo = MAX(valor_sem_fracao, aero_minimo) + valor_fracao
    """
    aero_minimo = float(aero_minimo or 0)
    base        = max(valor_sem_fracao, aero_minimo) if aero_minimo > 0 else valor_sem_fracao
    return base + valor_fracao


def aplicar_percentual_frete(valor_com_minimo: float, percentual_frete: float) -> float:
    """
    Aplica o percentual_frete do tabela_trecho sobre o valor ja calculado
    (apos a soma de todos os tipos e a aplicacao do aero minimo).

    valor_final = valor_com_minimo * (1 + percentual_frete / 100)

    Se percentual_frete for 0/nulo, o valor permanece inalterado.
    """
    percentual_frete = float(percentual_frete or 0)
    return valor_com_minimo * (1 + percentual_frete / 100)


def calcular_valor_trecho(
    peso: float,
    valor_transp: float,
    aero_minimo: float,
    faixas_por_tipo: dict[int, list[dict]],
    percentual_frete: float = 0,
    nf_total: float = 0,
    tem_tde: bool = False,
    desc: dict | None = None,
) -> dict:
    """
    Funcao principal do modulo de calculo.
    Orquestra o calculo completo de um trecho para um AWB.

    Ordem de aplicacao:
    1. Calcula cada tipo do fluxo normal (1, 4, 5, 8, 11, 14, 23, 32)
    2. Separa fracao (23) dos demais
    3. valor_com_minimo = MAX(soma_sem_fracao, aero_minimo) + fracao
    4. valor_com_tde = valor_com_minimo + TDE (tipo 24, somado por fora, so se tem_tde=True)
    5. valor_final = valor_com_tde * (1 + percentual_frete / 100)
    """
    resultado = calcular_tipos(peso, valor_transp, faixas_por_tipo, desc=desc)

    valor_com_minimo = aplicar_aero_minimo(
        resultado["valor_sem_fracao"],
        resultado["valor_fracao"],
        aero_minimo,
    )

    tde = calcular_tde(nf_total, faixas_por_tipo.get(TIPO_TDE, []), tem_tde)
    valor_com_tde = valor_com_minimo + tde

    valor_final = aplicar_percentual_frete(valor_com_tde, percentual_frete)

    resultado["tde"]              = tde
    resultado["valor_com_minimo"] = valor_com_minimo
    resultado["valor_com_tde"]    = valor_com_tde
    resultado["valor_final"]      = valor_final
    return resultado


# =============================================================================
# CLIENTE_TRECHO: seleção do desconto (candidatas cacheadas por rota/serviço)
# =============================================================================
def candidatas_cliente_trecho(regras: list[dict], cidade_origem: int, cidade_destino: int, id_servico: int) -> list[dict]:
    """
    Regras da relatorios.cliente_trecho compatíveis com origem/destino/serviço do AWB,
    ordenadas da melhor para a pior. id_origem/id_destino = 0 -> qualquer cidade. Prioridade:
      1 = origem exata  x destino exato
      2 = origem exata  x destino 0
      3 = origem 0      x destino exato
      4 = origem 0      x destino 0
    Desempate: servico exato antes de servico 0 (e, persistindo o empate, ordem de leitura).
    """
    achadas = []
    for r in regras:
        orig, dest, serv = int(r["id_origem"] or 0), int(r["id_destino"] or 0), int(r["servico"] or 0)
        if orig not in (0, cidade_origem) or dest not in (0, cidade_destino) or serv not in (0, id_servico):
            continue
        prioridade = {(True, True): 1, (True, False): 2, (False, True): 3, (False, False): 4}[(orig != 0, dest != 0)]
        achadas.append(((prioridade, 0 if serv != 0 else 1), r))
    achadas.sort(key=lambda x: x[0])  # sort estável
    return [r for _, r in achadas]


def selecionar_cliente_trecho(candidatas: list[dict], peso: float, data_entrada) -> dict | None:
    """
    Primeira candidata (já ordenada por prioridade) que atende aos filtros:
    peso_inicio <= peso <= peso_fim e data_entrada <= vigencia (NULL = sem limite).
    """
    data_ref = data_entrada.date() if hasattr(data_entrada, "date") else data_entrada
    for r in candidatas:
        p_ini, p_fim = r.get("peso_inicio"), r.get("peso_fim")
        if (p_ini is not None and peso < float(p_ini)) or (p_fim is not None and peso > float(p_fim)):
            continue
        vig = r.get("vigencia")
        if vig is not None and data_ref is not None:
            vig = vig.date() if hasattr(vig, "date") else vig
            if data_ref > vig:
                continue
        return r
    return None


# =============================================================================
# TRECHO: match origem/destino feito em Python (sem JOIN explosivo no banco)
# =============================================================================
def _int(v):
    try:
        return int(v)
    except (TypeError, ValueError):
        return None


def consultar_em_blocos(cursor, sql: str, ids) -> list[dict]:
    """Executa `sql` (com o marcador {ph}) em blocos de TAM_BLOCO ids e concatena as linhas."""
    ids, linhas = list(ids), []
    for i in range(0, len(ids), TAM_BLOCO):
        bloco = ids[i:i + TAM_BLOCO]
        cursor.execute(sql.format(ph=",".join(["%s"] * len(bloco))), bloco)
        linhas.extend(cursor.fetchall())
    return linhas


def melhor_trecho(trechos: list[dict], cidade_origem, cidade_destino, rotas_uf: dict, zonas: dict):
    """
    Prioridade de match origem/destino (mesma regra da query antiga):
      1 = origem rota x destino rota          4 = origem rota x destino zona
      2 = origem UF   x destino UF            5 = origem UF   x destino zona
      3 = origem rota x destino UF (ou UF x rota)   6 = demais (origem 0 ou destino 0)
    Origem so bate por rota ou UF (ou 0). Destino bate por rota, UF, zona (ou 0).
    Desempate determinístico: menor id_trecho.
    """
    rota_o, rota_d = cidade_origem in rotas_uf, cidade_destino in rotas_uf
    ufs_o,  ufs_d  = rotas_uf.get(cidade_origem, ()), rotas_uf.get(cidade_destino, ())
    zonas_d        = zonas.get(cidade_destino, ())
    melhor = None
    for tt in trechos:
        o, d = tt["origem"], tt["destino"]
        if o is None or d is None:
            continue
        o_rota, o_uf = rota_o and o == cidade_origem,  o in ufs_o
        d_rota, d_uf, d_zona = rota_d and d == cidade_destino, d in ufs_d, d in zonas_d
        if not (o == 0 or o_rota or o_uf):
            continue
        if not (d == 0 or d_rota or d_uf or d_zona):
            continue
        if o_rota and d_rota:
            p = 1
        elif o_uf and d_uf:
            p = 2
        elif (o_rota and d_uf) or (o_uf and d_rota):
            p = 3
        elif o_rota and d_zona:
            p = 4
        elif o_uf and d_zona:
            p = 5
        else:
            p = 6
        chave = (p, tt["id_trecho"])
        if melhor is None or chave < melhor[0]:
            melhor = (chave, tt)
    return melhor


# =============================================================================
# EXPORTAÇÃO XLSX (padrão skyu-formatacao-xlsx)
# =============================================================================
FORMATOS = {
    "text":      None,
    "int":       None,
    "inteiro":   "#,##0_ ;[Red]-#,##0 ",
    "decimal":   "#,##0.00_ ;[Red]-#,##0.00 ",
    "moeda":     "R$ #,##0.00_ ;[Red]-R$ #,##0.00 ",
    "pct":       "0.0%_ ;[Red]-0.0% ",
    "data":      "dd/mm/yyyy",
    "data_hora": "dd/mm/yyyy hh:mm",
}

CENTER  = {"left": "left", "mid": "center", "right": "right"}
DEFAULT = {"format": "text", "size": 8.0, "center": "left"}


def _montar_abas(df_resultado: pd.DataFrame) -> dict:
    return {
        "Resultado": {
            "df": df_resultado, "sort_col": "cod_awb",
            "cols": """
                awb                         size:14    center:mid
                cod_awb                     format:int      size:11    center:mid
                fornecedor                  size:28    center:left
                valor_cobrado               format:decimal  size:15    center:mid
                id_tabela                   format:int      size:11    center:mid
                peso                        format:decimal  size:11    center:mid
                id_trecho                   format:int      size:11    center:mid
                valor_tabela                format:decimal  size:14    center:mid
                diferença                   format:decimal  size:13    center:mid
                data_emissão                format:data     size:14    center:mid
                origem                      size:10    center:mid
                destino                     size:10    center:mid
                serviço_awb                 size:14    center:mid
                status                      size:12    center:mid
            """,
        },
    }


_RE_TOKEN = re.compile(r"(format|size|center)\s*:\s*(\S+)", re.IGNORECASE)


def _parse_cols(spec, nome_aba=""):
    cfg = {}
    for n_linha, linha in enumerate((spec or "").splitlines(), 1):
        linha = linha.split("#")[0].strip()
        if not linha:
            continue
        tokens = list(_RE_TOKEN.finditer(linha))
        col = (linha[:tokens[0].start()] if tokens else linha).strip()
        if not col:
            raise ValueError(f"[{nome_aba}] linha {n_linha}: falta o nome da coluna")
        d = dict(DEFAULT)
        for t in tokens:
            chave, valor = t.group(1).lower(), t.group(2).lower()
            d[chave] = float(valor) if chave == "size" else valor
        if d["format"] not in FORMATOS:
            raise ValueError(f"[{nome_aba}] '{col}': format '{d['format']}' inválido. Use: {', '.join(FORMATOS)}")
        if d["center"] not in CENTER:
            raise ValueError(f"[{nome_aba}] '{col}': center '{d['center']}' inválido. Use: {', '.join(CENTER)}")
        cfg[col] = d
    return cfg


def formatar_aba(ws, wb, df, cols_cfg=None):
    cols_cfg = cols_cfg or {}
    _fmt_cache = {}

    def _get_fmt(align, num_format):
        key = (align, num_format)
        if key not in _fmt_cache:
            props = {"align": align, "valign": "vcenter"}
            if num_format:
                props["num_format"] = num_format
            _fmt_cache[key] = wb.add_format(props)
        return _fmt_cache[key]

    def snake_to_titulo(col):
        return col.replace("_", " ").upper()

    fmt_header = wb.add_format({
        "bold": True, "font_color": "#FFFFFF", "bg_color": "#261957",
        "align": "center", "valign": "vcenter", "text_wrap": True,
    })

    cols, n_rows = list(df.columns), len(df)

    ws.set_row(0, 30)
    for c_idx, col in enumerate(cols):
        ws.write(0, c_idx, snake_to_titulo(col), fmt_header)

    for c_idx, col in enumerate(cols):
        cfg = cols_cfg.get(col, DEFAULT)
        fmt = _get_fmt(CENTER[cfg["center"]], FORMATOS[cfg["format"]])
        ws.set_column(c_idx, c_idx, cfg["size"], fmt)
        if n_rows and pd.api.types.is_datetime64_any_dtype(df[col]):
            valores = [None if pd.isna(v) else v.to_pydatetime() for v in df[col]]
            ws.write_column(1, c_idx, valores, fmt)

    ws.autofilter(0, 0, n_rows, len(cols) - 1)


# =============================================================================
# EXECUÇÃO PRINCIPAL
# =============================================================================
def main():
    conn = conectar_mysql()
    if conn is None:
        raise SystemExit("❌ Encerrando: não foi possível conectar ao banco.")

    cursor = conn.cursor(dictionary=True)

    # -------------------------------------------------------------------------
    # QUERY 1: notas do período/fornecedor (sem tabela_trecho — o match é feito em Python)
    # -------------------------------------------------------------------------
    ph_forn = ",".join(["%s"] * len(FORNECEDOR))
    SQL_NOTAS = f"""
    SELECT
        n.id_nota,
        n.awb,
        n.data_entrada,
        n.peso,
        n.nf_total,
        n.valor_transp,
        f.id_local         AS id_fornecedor,
        f.fantasia         AS fornecedor_nome,
        dbv.tabela         AS id_tabela,
        f_dest.tad         AS tad,
        a_orig.cidade      AS cidade_origem,
        a_dest.cidade      AS cidade_destino,
        eq.id_equipamento  AS id_servico
    FROM notas n
    INNER JOIN fornecedores f ON f.id_local = n.fornecedor
    INNER JOIN personalizados.vigencia_tabela_awb dbv
        ON dbv.id_fornecedor = f.id_local
        AND dbv.data_inicial <= n.data_entrada
        AND (dbv.data_final >= n.data_entrada OR dbv.data_final IS NULL)
    LEFT JOIN fornecedores f_dest ON f_dest.id_local = n.destinatario
    INNER JOIN aero a_orig ON a_orig.id_aero = n.origem
    INNER JOIN aero a_dest ON a_dest.id_aero = n.destino
    INNER JOIN equipamento eq ON eq.servico_cia = n.servico
    WHERE n.data_entrada BETWEEN %s AND %s
        AND n.fornecedor IN ({ph_forn})
    ORDER BY n.id_nota
    """
    cursor.execute(SQL_NOTAS, (DATA_INICIAL, DATA_FINAL, *FORNECEDOR))
    rows_notas = cursor.fetchall()
    print(f" - notas (brutas) -> [{len(rows_notas):,} linhas]".replace(",", "."))

    if not rows_notas:
        cursor.close()
        conn.close()
        raise SystemExit("⚠️  Nenhum AWB encontrado para o período/fornecedor informado.")

    # -------------------------------------------------------------------------
    # QUERY 2-4: trechos das tabelas vigentes, rotas (UF) e zonas — tabelas pequenas
    # -------------------------------------------------------------------------
    ids_tabela  = {r["id_tabela"] for r in rows_notas}
    cidades_ref = {_int(r["cidade_origem"]) for r in rows_notas} | {_int(r["cidade_destino"]) for r in rows_notas}
    cidades_ref.discard(None)
    cidades_dst = {_int(r["cidade_destino"]) for r in rows_notas}
    cidades_dst.discard(None)

    rows_tt = consultar_em_blocos(cursor, """
    SELECT
        tt.id_trecho,
        tt.id_tabela,
        tt.servico,
        tt.origem,
        tt.destino,
        tt.aero_minimo,
        tt.percentual_frete
    FROM tabela_trecho tt
    WHERE tt.status = 1
        AND tt.id_tabela IN ({ph})
    """, ids_tabela)

    trechos_por: dict[tuple, list[dict]] = defaultdict(list)
    for t in rows_tt:
        t["origem"], t["destino"] = _int(t["origem"]), _int(t["destino"])
        trechos_por[(str(t["id_tabela"]), _int(t["servico"]))].append(t)

    rotas_uf: dict[int, set] = defaultdict(set)
    for r in consultar_em_blocos(cursor, "SELECT r.id_rota, r.uf_ibge FROM rotas r WHERE r.id_rota IN ({ph})", cidades_ref):
        rotas_uf[int(r["id_rota"])].add(_int(r["uf_ibge"]))

    zonas: dict[int, set] = defaultdict(set)
    for z in consultar_em_blocos(cursor, "SELECT czg.id_zona, czg.id_rota FROM cidade_zona_grupo czg WHERE czg.id_rota IN ({ph})", cidades_dst):
        zonas[int(z["id_rota"])].add(_int(z["id_zona"]))

    print(f" - tabela_trecho -> [{len(rows_tt):,} trechos | {len(rotas_uf):,} rotas | {len(zonas):,} cidades com zona]".replace(",", "."))

    # Melhor trecho por AWB (menor prioridade; empate -> menor id_trecho).
    # Resultado só depende de (tabela, serviço, origem, destino) -> cache.
    cache_trecho: dict[tuple, tuple | None] = {}
    mapa_notas: dict = {}
    for row in rows_notas:
        tab  = str(row["id_tabela"])
        serv = _int(row["id_servico"])
        co, cd = _int(row["cidade_origem"]), _int(row["cidade_destino"])
        k = (tab, serv, co, cd)
        if k not in cache_trecho:
            cache_trecho[k] = melhor_trecho(trechos_por.get((tab, serv), ()), co, cd, rotas_uf, zonas)
        achado = cache_trecho[k]
        if achado is None:
            continue
        chave, tt = achado
        atual = mapa_notas.get(row["id_nota"])
        if atual is None or chave < atual["_chave"]:
            mapa_notas[row["id_nota"]] = {
                **row,
                "id_trecho":        tt["id_trecho"],
                "aero_minimo":      tt["aero_minimo"],
                "percentual_frete": tt["percentual_frete"],
                "_chave":           chave,
            }

    print(f" - notas_trecho -> [{len(mapa_notas):,} AWBs com trecho encontrado]".replace(",", "."))

    if not mapa_notas:
        cursor.close()
        conn.close()
        raise SystemExit("⚠️  Nenhum AWB com trecho encontrado para o período/fornecedor informado.")

    # -------------------------------------------------------------------------
    # QUERY: faixas de cada trecho encontrado (só os tipos que o cálculo usa)
    # -------------------------------------------------------------------------
    ids_trecho = list({r["id_trecho"] for r in mapa_notas.values()})
    tipos_sql  = ",".join(str(t) for t in sorted(TIPOS_VALIDOS | {TIPO_TDE}))
    rows_faixas = consultar_em_blocos(cursor, f"""
    SELECT
        tf.id_trecho,
        tf.id_tabela,
        tf.tipo,
        tf.indice,
        tf.inicio,
        tf.fim,
        tf.minimo,
        tf.franquia,
        tf.excedente
    FROM tabela_faixas tf
    WHERE tf.id_trecho IN ({{ph}})
        AND tf.deleted_at IS NULL
        AND tf.tipo IN ({tipos_sql})
    ORDER BY tf.id_trecho, tf.tipo, tf.indice
    """, ids_trecho)

    # Agrupa faixas: {(id_trecho, id_tabela): {tipo: [faixas]}}
    faixas_agrupadas: dict[tuple, dict[int, list]] = {}
    for f in rows_faixas:
        chave = (f["id_trecho"], str(f["id_tabela"]))
        faixas_agrupadas.setdefault(chave, {}).setdefault(int(f["tipo"]), []).append(f)

    print(f" - faixas       -> [{len(rows_faixas):,} registros carregados]".replace(",", "."))

    # -------------------------------------------------------------------------
    # QUERY: descontos por cliente (relatorios.cliente_trecho)
    # -------------------------------------------------------------------------
    ids_cliente = list({r["id_fornecedor"] for r in mapa_notas.values() if str(r["id_fornecedor"]) in FORNECEDORES_CLIENTE_TRECHO})
    rows_cliente = consultar_em_blocos(cursor, """
    SELECT
        ct.id_trecho_cliente,
        ct.id_cliente,
        ct.id_origem,
        ct.id_destino,
        ct.servico,
        ct.peso_inicio,
        ct.peso_fim,
        ct.vigencia,
        ct.desconto_frete_peso,
        ct.desconto_frete_minimo,
        ct.adv_exc_fixo,
        ct.gris_exc_fixo
    FROM relatorios.cliente_trecho ct
    WHERE ct.id_cliente IN ({ph})
        AND ct.status = 1
    """, ids_cliente)

    regras_cliente: dict[int, list[dict]] = {}
    for r in rows_cliente:
        regras_cliente.setdefault(int(r["id_cliente"]), []).append(r)

    print(f" - cliente_trecho -> [{len(rows_cliente):,} registros carregados]".replace(",", "."))

    # -------------------------------------------------------------------------
    # CÁLCULO DO VALOR DE TABELA POR AWB
    # -------------------------------------------------------------------------
    resultados = []
    cache_cand: dict[tuple, list[dict]] = {}

    for id_nota, dados in mapa_notas.items():
        nf_total         = float(dados["nf_total"]         or 0)
        peso             = float(dados["peso"]             or 0)
        valor_transp     = float(dados["valor_transp"]     or 0)
        aero_minimo      = float(dados["aero_minimo"]      or 0)
        percentual_frete = float(dados["percentual_frete"] or 0)
        id_trecho        = dados["id_trecho"]
        id_tabela        = dados["id_tabela"]

        faixas_por_tipo = faixas_agrupadas.get((id_trecho, str(id_tabela)), {})
        tem_tde = int(dados.get("tad") or 0) == 1

        id_forn = int(dados["id_fornecedor"])
        kc = (id_forn, int(dados["cidade_origem"] or 0), int(dados["cidade_destino"] or 0), int(dados["id_servico"] or 0))
        if kc not in cache_cand:
            cache_cand[kc] = candidatas_cliente_trecho(regras_cliente.get(id_forn, []), kc[1], kc[2], kc[3])
        desc = selecionar_cliente_trecho(cache_cand[kc], peso, dados["data_entrada"])

        calc = calcular_valor_trecho(
            peso,
            valor_transp,
            aero_minimo,
            faixas_por_tipo,
            percentual_frete=percentual_frete,
            nf_total=nf_total,
            tem_tde=tem_tde,
            desc=desc,
        )
        # calc_cheio = calcular_valor_trecho(   # reativar junto com "valor_tabela_sem_desconto"
        #     peso, valor_transp, aero_minimo, faixas_por_tipo,
        #     percentual_frete=percentual_frete, nf_total=nf_total, tem_tde=tem_tde,
        # ) if desc else calc

        resultados.append({
            "awb":                       dados["awb"],
            "cod_awb":                   id_nota,
            "fornecedor":                dados["fornecedor_nome"],
            "valor_cobrado":             nf_total,
            "id_tabela":                 id_tabela,
            "peso":                      peso,
            "id_trecho":                 id_trecho,
            # "id_trecho_cliente":         desc["id_trecho_cliente"] if desc else None,
            # "valor_tabela_sem_desconto": round(calc_cheio["valor_final"], 2),
            "valor_tabela":              round(calc["valor_final"], 2),
            "diferença":                 round(calc["valor_final"] - nf_total, 2),
            # "tde":                       round(calc["tde"], 2),
            # "percentual_frete":          percentual_frete,
        })

    df_resultado = pd.DataFrame(resultados)
    print(f" - df_resultado -> [{df_resultado.shape[0]:,} linhas x {df_resultado.shape[1]} colunas]".replace(",", "."))

    # -------------------------------------------------------------------------
    # QUERY: dados complementares de db_awb (cod_awb é PK INT)
    # -------------------------------------------------------------------------
    rows_dawb = consultar_em_blocos(cursor, """
    SELECT
        da.cod_awb,
        da.emissao_awb,
        da.origem,
        da.destino,
        da.servico_awb,
        da.status_awb
    FROM personalizados.db_awb da
    WHERE da.cod_awb IN ({ph})
    AND da.status_awb = "ATIVO"
    """, df_resultado["cod_awb"].dropna().tolist())

    df_dawb = pd.DataFrame(rows_dawb, columns=["cod_awb", "emissao_awb", "origem", "destino", "servico_awb", "status_awb"])
    df_dawb["emissao_awb"] = pd.to_datetime(df_dawb["emissao_awb"], errors="coerce")
    print(f" - df_dawb      -> [{df_dawb.shape[0]:,} linhas x {df_dawb.shape[1]} colunas]".replace(",", "."))

    # LEFT: mantém AWBs sem registro em db_awb (campos vazios); chave temporária, cod_awb segue com o tipo original
    df_dawb["_chave"]      = pd.to_numeric(df_dawb["cod_awb"], errors="coerce").astype("Int64")
    df_resultado["_chave"] = pd.to_numeric(df_resultado["cod_awb"], errors="coerce").astype("Int64")
    df_resultado = (
        df_resultado
        .merge(df_dawb.drop(columns="cod_awb"), on="_chave", how="left")
        .drop(columns="_chave")
        .rename(columns={"emissao_awb": "data_emissão", "servico_awb": "serviço_awb", "status_awb": "status"})
    )
    print(f" - df_resultado -> [{df_resultado.shape[0]:,} linhas x {df_resultado.shape[1]} colunas]".replace(",", "."))

    cursor.close()
    conn.close()

    # -------------------------------------------------------------------------
    # EXPORTAÇÃO (xlsxwriter, dois passos)
    # -------------------------------------------------------------------------
    ABAS = _montar_abas(df_resultado)

    with pd.ExcelWriter(
        ARQUIVO_SAIDA,
        engine="xlsxwriter",
        datetime_format="dd/mm/yyyy",
        date_format="dd/mm/yyyy",
    ) as writer:

        # PASSO 1 — escrever
        for nome_aba, cfg in ABAS.items():
            _df = cfg["df"]
            if cfg.get("sort_col"):
                _df = _df.sort_values(cfg["sort_col"], na_position="last").reset_index(drop=True)
            _df.to_excel(writer, sheet_name=nome_aba, index=False)
            ABAS[nome_aba]["_df_sorted"] = _df

        # PASSO 2 — formatar
        wb = writer.book
        for nome_aba, cfg in ABAS.items():
            _df       = cfg["_df_sorted"]
            _cols_cfg = _parse_cols(cfg.get("cols", ""), nome_aba)
            _fantasma = [c for c in _cols_cfg if c not in _df.columns]
            if _fantasma:
                print(f"⚠ [{nome_aba}] coluna declarada e inexistente no df: {', '.join(_fantasma)}")
            formatar_aba(writer.sheets[nome_aba], wb, _df, cols_cfg=_cols_cfg)

    print("Exportado:")
    for nome_aba, cfg in ABAS.items():
        _df = cfg["_df_sorted"]
        print(f" - {'df_resultado':<27} -> [{_df.shape[0]:,} linhas x {_df.shape[1]} colunas]".replace(",", "."))
    print(f"\n✔ {ARQUIVO_SAIDA}")


if __name__ == "__main__":
    main()