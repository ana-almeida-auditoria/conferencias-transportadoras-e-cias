# =============================================================================
# SCRIPT UNIFICADO: calculo_tabela_cia_v2.py
# Busca AWBs por período/fornecedor, localiza o trecho (com match de zona) e
# calcula o valor de tabela (tipos 1, 4, 5, 8, 11, 14, 23, 32 + percentual_frete)
# =============================================================================

import math
import pandas as pd
import os
from dotenv import load_dotenv
from pathlib import Path
import mysql.connector
from mysql.connector import Error
from sqlalchemy import create_engine
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

# =============================================================================
# CONFIGURAÇÕES
# =============================================================================
DATA_INICIAL = "2026-08-17"   # período de busca (data_entrada)
DATA_FINAL   = "2026-09-23"
FORNECEDOR   = ["31432"]            # id do fornecedor (notas.fornecedor / fornecedores.id_local) 

ARQUIVO_SAIDA = "C:\\Users\\ana.almeida\\Downloads\\resultado_tabela_cia.xlsx"


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

# exemplo com sqlalchemy
engine = create_engine(
    f"mysql+mysqlconnector://{DB_CONFIG['user']}:{DB_CONFIG['password']}"
    f"@{DB_CONFIG['host']}:{DB_CONFIG['port']}/{DB_CONFIG['database']}"
)

# =============================================================================
# MÓDULO DE CÁLCULO DE TABELA (tipos 1, 4, 5, 6, 8, 11, 14, 23, 32 + TDE(24) + percentual_frete)
# =============================================================================
TIPOS_VALIDOS = {1, 4, 5, 6, 8, 11, 14, 23, 32}  # calculados no fluxo normal (calcular_tipos)

TIPOS_ADVALOREM           = {32}        # base = valor_transp; MAX((valor_transp * excedente/100) + franquia, minimo)
TIPOS_ADVALOREM_DIRETO    = {4, 5}      # base = valor_transp direto (sem ratio)
TIPOS_FRACAO              = {23}        # CEIL(peso / fracao) * excedente (sem franquia/minimo)
TIPOS_FRACAO_PROPORCIONAL = {14}        # CEIL(peso / fracao) * excedente (igual tipo 23, sem franquia/minimo)
TIPOS_NORMAIS             = {1, 6, 8, 11}  # (base - franquia) * excedente + minimo
# Tipo 6 (Despacho) e mutuamente exclusivo com tipo 1: se o trecho tiver tipo 6, tipo 1 e ignorado.

TIPO_TDE = 24  # calculado FORA do fluxo normal: depende de fornecedores.tad e usa nf_total como base


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


def _calcular_faixa(base: float, faixa: dict, tipo: int, ratio: float | None = None) -> float:
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
    """
    minimo    = float(faixa.get("minimo",    0) or 0)
    franquia  = float(faixa.get("franquia",  0) or 0)
    excedente = float(faixa.get("excedente", 0) or 0)

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
    calculado = (base_exc * excedente) if excedente > 0 else 0

    if tipo == 8:
        return max(calculado, minimo) if minimo > 0 else calculado

    # Tipos 1, 6, 11: soma minimo
    Z = minimo if minimo > 0 else 0
    return calculado + Z


def calcular_tipos(peso: float, valor_transp: float, faixas_por_tipo: dict[int, list[dict]]) -> dict:
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

        resultados[tipo] = _calcular_faixa(base, faixa_match, tipo, ratio=ratio)

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
    resultado = calcular_tipos(peso, valor_transp, faixas_por_tipo)

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
# EXECUÇÃO PRINCIPAL
# =============================================================================
def main():
    conn = conectar_mysql()
    if conn is None:
        raise SystemExit("❌ Encerrando: não foi possível conectar ao banco.")

    cursor = conn.cursor(dictionary=True)

    # -------------------------------------------------------------------------
    # QUERY PRINCIPAL: notas do período/fornecedor + fornecedor + trecho (com prioridade)
    # -------------------------------------------------------------------------
    # Prioridade de match origem/destino:
    #   1 = origem rota    x destino rota
    #   2 = origem UF      x destino UF
    #   3 = origem rota    x destino UF   (ou origem UF x destino rota)
    #   4 = origem rota    x destino zona
    #   5 = origem UF      x destino zona
    #
    # Origem so bate por rota ou UF. Destino pode bater por rota, UF ou zona
    # (tabela_trecho.destino = cidade_zona_grupo.id_zona -> cidade_zona_grupo.id_rota).
    SQL_NOTAS_TRECHO = """
    SELECT
        n.id_nota,
        n.awb,
        n.data_entrada,
        n.peso,
        n.nf_total,
        n.valor_transp,
        f.id_local        AS id_fornecedor,
        f.fantasia         AS fornecedor_nome,
        dbv.tabela AS id_tabela,
        f_dest.tad         AS tad,
        tt.id_trecho,
        tt.aero_minimo,
        tt.percentual_frete,
        CASE
            WHEN r_orig.id_rota      IS NOT NULL AND r_dest.id_rota      IS NOT NULL THEN 1
            WHEN r_orig_uf.id_rota   IS NOT NULL AND r_dest_uf.id_rota   IS NOT NULL THEN 2
            WHEN r_orig.id_rota      IS NOT NULL AND r_dest_uf.id_rota   IS NOT NULL THEN 3
            WHEN r_orig_uf.id_rota   IS NOT NULL AND r_dest.id_rota      IS NOT NULL THEN 3
            WHEN r_orig.id_rota      IS NOT NULL AND r_dest_zona.id_rota IS NOT NULL THEN 4
            WHEN r_orig_uf.id_rota   IS NOT NULL AND r_dest_zona.id_rota IS NOT NULL THEN 5
            ELSE 6
        END AS prioridade_trecho
    FROM notas n
    INNER JOIN fornecedores f
        ON f.id_local = n.fornecedor
    LEFT JOIN personalizados.vigencia_tabela_awb dbv
        ON dbv.id_fornecedor = f.id_local
        AND dbv.data_inicial <= n.data_entrada
    LEFT JOIN fornecedores f_dest
        ON f_dest.id_local = n.destinatario
    INNER JOIN aero a_orig
        ON a_orig.id_aero = n.origem
    INNER JOIN aero a_dest
        ON a_dest.id_aero = n.destino
    INNER JOIN equipamento eq
        ON eq.servico_cia = n.servico
    INNER JOIN tabela_trecho tt
        ON tt.servico    = eq.id_equipamento
        AND tt.id_tabela = dbv.tabela
        AND tt.status    = 1
        AND (
            tt.origem = 0
            OR EXISTS (SELECT 1 FROM rotas r WHERE r.id_rota = tt.origem AND r.id_rota = a_orig.cidade)
            OR EXISTS (SELECT 1 FROM rotas r WHERE r.uf_ibge = tt.origem AND r.id_rota = a_orig.cidade)
        )
        AND (
            tt.destino = 0
            OR EXISTS (SELECT 1 FROM rotas r WHERE r.id_rota = tt.destino AND r.id_rota = a_dest.cidade)
            OR EXISTS (SELECT 1 FROM rotas r WHERE r.uf_ibge = tt.destino AND r.id_rota = a_dest.cidade)
            OR EXISTS (SELECT 1 FROM cidade_zona_grupo czg WHERE czg.id_zona = tt.destino AND czg.id_rota = a_dest.cidade)
        )
    LEFT JOIN rotas r_orig
        ON r_orig.id_rota      = tt.origem  AND r_orig.id_rota      = a_orig.cidade
    LEFT JOIN rotas r_orig_uf
        ON r_orig_uf.uf_ibge   = tt.origem  AND r_orig_uf.id_rota   = a_orig.cidade
    LEFT JOIN rotas r_dest
        ON r_dest.id_rota      = tt.destino AND r_dest.id_rota      = a_dest.cidade
    LEFT JOIN rotas r_dest_uf
        ON r_dest_uf.uf_ibge   = tt.destino AND r_dest_uf.id_rota   = a_dest.cidade
    LEFT JOIN cidade_zona_grupo r_dest_zona
        ON r_dest_zona.id_zona = tt.destino AND r_dest_zona.id_rota = a_dest.cidade
    WHERE
        n.data_entrada BETWEEN %s AND %s
        AND n.fornecedor = %s
    ORDER BY
        n.id_nota,
        prioridade_trecho ASC
    """

    placeholders = ",".join(["%s"] * len(FORNECEDOR))
    sql = SQL_NOTAS_TRECHO.replace("n.fornecedor = %s", f"n.fornecedor IN ({placeholders})")
    cursor.execute(sql, (DATA_INICIAL, DATA_FINAL, *FORNECEDOR))
    rows_notas = cursor.fetchall()

    # Mantém apenas o melhor trecho por AWB (menor prioridade numérica = melhor match)
    mapa_notas = {}
    for row in rows_notas:
        id_nota = row["id_nota"]
        if id_nota not in mapa_notas:
            mapa_notas[id_nota] = row

    print(f" - notas_trecho -> [{len(mapa_notas):,} AWBs com trecho encontrado]".replace(",", "."))

    if not mapa_notas:
        cursor.close()
        conn.close()
        raise SystemExit("⚠️  Nenhum AWB encontrado para o período/fornecedor informado.")

    # -------------------------------------------------------------------------
    # QUERY: faixas de cada trecho encontrado (filtrado por id_tabela)
    # -------------------------------------------------------------------------
    ids_trecho = list({r["id_trecho"] for r in mapa_notas.values()})

    fmt        = ",".join(["%s"] * len(ids_trecho))
    SQL_FAIXAS = f"""
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
    WHERE tf.id_trecho IN ({fmt})
    AND tf.deleted_at IS NULL
    ORDER BY tf.id_trecho, tf.tipo, tf.indice
    """
    cursor.execute(SQL_FAIXAS, ids_trecho)
    rows_faixas = cursor.fetchall()

    # Agrupa faixas: {(id_trecho, id_tabela): {tipo: [faixas]}}
    faixas_agrupadas: dict[tuple, dict[int, list]] = {}
    for f in rows_faixas:
        chave = (f["id_trecho"], str(f["id_tabela"]))
        tipo  = int(f["tipo"])
        faixas_agrupadas.setdefault(chave, {}).setdefault(tipo, []).append(f)

    print(f" - faixas       -> [{len(rows_faixas):,} registros carregados]".replace(",", "."))

    # -------------------------------------------------------------------------
    # CÁLCULO DO VALOR DE TABELA POR AWB
    # -------------------------------------------------------------------------
    resultados = []

    for id_nota, dados in mapa_notas.items():
        nf_total         = float(dados["nf_total"]         or 0)
        peso             = float(dados["peso"]             or 0)
        valor_transp     = float(dados["valor_transp"]     or 0)
        aero_minimo      = float(dados["aero_minimo"]      or 0)
        percentual_frete = float(dados["percentual_frete"] or 0)
        id_trecho        = dados["id_trecho"]
        id_tabela        = dados["id_tabela"]

        chave           = (id_trecho, str(id_tabela))
        faixas_por_tipo = faixas_agrupadas.get(chave, {})

        tem_tde = int(dados.get("tad") or 0) == 1

        calc = calcular_valor_trecho(
            peso,
            valor_transp,
            aero_minimo,
            faixas_por_tipo,
            percentual_frete=percentual_frete,
            nf_total=nf_total,
            tem_tde=tem_tde,
        )

        resultados.append({
            "AWB":              dados["awb"],
            "ID NOTA":          id_nota,
            "FORNECEDOR":       dados["fornecedor_nome"],
            "VALOR":            nf_total,
            "ID TABELA":        id_tabela,
            "PESO":             peso,
            "ID TRECHO":        id_trecho,
            "PERCENTUAL FRETE": percentual_frete,
            "TDE":              round(calc["tde"], 2),
            "VALOR TABELA":     round(calc["valor_final"], 2),
            "DIFERENÇA":        round(calc["valor_final"] - nf_total, 2),
            # --- DEBUG (temporario, remover depois de validar) ---
            # "DEBUG_TIPOS_ENCONTRADOS": ",".join(str(t) for t in sorted(faixas_por_tipo.keys())),
            # "DEBUG_TIPO_1":  round(calc["por_tipo"].get(1,  0), 2),
            # "DEBUG_TIPO_4":  round(calc["por_tipo"].get(4,  0), 2),
            # "DEBUG_TIPO_5":  round(calc["por_tipo"].get(5,  0), 2),
            # "DEBUG_TIPO_6":  round(calc["por_tipo"].get(6,  0), 2),
            # "DEBUG_TIPO_8":  round(calc["por_tipo"].get(8,  0), 2),
            # "DEBUG_TIPO_11": round(calc["por_tipo"].get(11, 0), 2),
            # "DEBUG_TIPO_14": round(calc["por_tipo"].get(14, 0), 2),
            # "DEBUG_TIPO_23": round(calc["por_tipo"].get(23, 0), 2),
            # "DEBUG_TIPO_32": round(calc["por_tipo"].get(32, 0), 2), 
        })

    df_resultado = pd.DataFrame(resultados)
    print(f" - df_resultado -> [{df_resultado.shape[0]:,} linhas x {df_resultado.shape[1]} colunas]".replace(",", "."))

    # -------------------------------------------------------------------------
    # QUERY: dados complementares de db_awb (mesmo padrão do main antigo)
    # -------------------------------------------------------------------------
    ids_awb = df_resultado["AWB"].dropna().tolist()

    if ids_awb:
        fmt_awb  = ",".join(["%s"] * len(ids_awb))
        SQL_DAWB = f"""
        SELECT
            da.cod_awb,
            da.emissao_awb,
            da.responsavel_transferencia,
            da.origem,
            da.destino,
            da.servico_awb,
            da.status_awb
        FROM personalizados.db_awb da
        WHERE da.cod_awb IN ({fmt_awb})
        """
        cursor.execute(SQL_DAWB, ids_awb)
        rows_dawb = cursor.fetchall()
    else:
        rows_dawb = []

    df_dawb = pd.DataFrame(rows_dawb)
    print(f" - df_dawb      -> [{df_dawb.shape[0]:,} linhas x {df_dawb.shape[1]} colunas]".replace(",", "."))

    # Merge: LEFT para manter AWBs sem registro em db_awb (aparecem com campos vazios)
    if not df_dawb.empty:
        df_dawb = df_dawb.rename(columns={"cod_awb": "AWB"})
        df_dawb["AWB"]      = df_dawb["AWB"].astype(str).str.strip()
        df_resultado["AWB"] = df_resultado["AWB"].astype(str).str.strip()
        df_resultado = df_resultado.merge(df_dawb, on="AWB", how="left")
        df_resultado = df_resultado.rename(columns={
            "emissao_awb":               "DATA EMISSÃO",
            "responsavel_transferencia": "RESPONSÁVEL TRANSFERÊNCIA",
            "origem":                    "ORIGEM",
            "destino":                    "DESTINO",
            "servico_awb":               "SERVIÇO AWB",
            "status_awb":                "STATUS",
        })

    print(f" - df_resultado -> [{df_resultado.shape[0]:,} linhas x {df_resultado.shape[1]} colunas]".replace(",", "."))

    # -------------------------------------------------------------------------
    # EXPORTAÇÃO COM FORMATAÇÃO (mesmo padrão visual do main antigo)
    # -------------------------------------------------------------------------
    df_resultado.to_excel(ARQUIVO_SAIDA, index=False)

    wb = load_workbook(ARQUIVO_SAIDA)
    ws = wb.active

    fill_header = PatternFill(start_color="1F3864", end_color="1F3864", fill_type="solid")
    font_header = Font(color="FFFFFF", bold=True)
    font_body   = Font(color="000000")
    alinhamento = Alignment(horizontal="center", vertical="center")
    borda_lado  = Side(style="thin", color="595959")
    borda       = Border(left=borda_lado, right=borda_lado, top=borda_lado, bottom=borda_lado)

    for cell in ws[1]:
        cell.fill      = fill_header
        cell.font      = font_header
        cell.alignment = alinhamento
        cell.border    = borda

    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.font      = font_body
            cell.alignment = alinhamento
            cell.border    = borda

    for col in ws.columns:
        max_len = max((len(str(c.value)) if c.value is not None else 0) for c in col)
        ws.column_dimensions[col[0].column_letter].width = max_len + 4

    wb.save(ARQUIVO_SAIDA)
    print(f"\n Arquivo gerado: {ARQUIVO_SAIDA}")

    # -------------------------------------------------------------------------
    # ENCERRAMENTO
    # -------------------------------------------------------------------------
    cursor.close()
    conn.close()


if __name__ == "__main__":
    main()