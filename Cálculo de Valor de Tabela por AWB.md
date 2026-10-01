---
schema_version: 2
tipo: auditoria
categoria: auditoria-operacional
subcategoria: calculo-tabela-transpo
status: nao-verificado
resumo: Recalcula o valor de tabela de cada AWB de um fornecedor no período (trecho, faixas, aero mínimo, TDE, percentual de frete e desconto por cliente) e compara com o valor cobrado (nf_total), exportando o resultado em Excel.
revisado_em: 2026-10-01
setores_destino:
  - auditoria schemas:
  - personalizados
  - relatorios execucao:
  - script
---
# Cálculo de Valor de Tabela por AWB
### `calculo_tabela_transpo_geral.py`

## 1. Objetivo

Recalcular, AWB a AWB, o valor que deveria ser cobrado segundo a tabela de frete vigente do fornecedor e comparar com o valor efetivamente lançado no AWB (`notas.nf_total`). A diferença indica cobrança acima ou abaixo da tabela.

O cálculo segue, nesta ordem:

1. Localiza o **trecho** da tabela vigente que corresponde à origem/destino do AWB.
2. Calcula cada **tipo de cobrança** pelas faixas do trecho (`tabela_faixas`), aplicando o **desconto do cliente** (`cliente_trecho`) quando houver.
3. Aplica o **aero mínimo** do trecho.
4. Soma o **TDE** (tipo 24), se o destinatário tiver TDE.
5. Aplica o **percentual de frete** do trecho.
6. Compara com `nf_total` e exporta.

---

## 2. Parâmetros de execução

|Parâmetro|Exemplo|Uso|
|---|---|---|
|`DATA_INICIAL`|`2026-09-01`|Início do período, sobre `notas.data_entrada`|
|`DATA_FINAL`|`2026-09-25`|Fim do período, sobre `notas.data_entrada` (inclusivo)|
|`FORNECEDOR`|`["169715"]`|Lista de ids de fornecedor (`notas.fornecedor` = `fornecedores.id_local`)|
|`ARQUIVO_SAIDA`|`...\Downloads\resultado_tabela_cia.xlsx`|Caminho do Excel gerado|

Constantes de regra de negócio (bloco do módulo de cálculo):

|Constante|Valor|Significado|
|---|---|---|
|`TIPOS_VALIDOS`|`{1, 4, 5, 6, 8, 11, 14, 22, 23, 25, 32}`|Tipos calculados no fluxo normal|
|`TIPO_TDE`|`24`|TDE, calculado fora do fluxo normal|
|`FORNECEDORES_CLIENTE_TRECHO`|`{"169715"}`|Fornecedores que consultam a `cliente_trecho` (hoje só JEM)|
|`TIPOS_FRETE_PESO`|`{8}`|Tipos que recebem desconto de frete peso/mínimo|
|`TIPOS_ADV`|`{4, 32}`|Tipos de advalorem que podem ter o excedente substituído|
|`TIPO_GRIS`|`5`|Tipo de GRIS que pode ter o excedente substituído|

Conexão: credenciais lidas do `.env` na mesma pasta do script (`DB_HOST`, `DB_PORT`, `DB_USER`, `DB_PASSWORD`, `DB_NAME`). As tabelas sem schema qualificado (`notas`, `fornecedores`, `aero`, `equipamento`, `tabela_trecho`, `tabela_faixas`, `rotas`, `cidade_zona_grupo`) são lidas do banco definido em `DB_NAME`.

---

## 3. Fontes de dados

|Tabela|Papel|
|---|---|
|`notas`|AWBs: peso, `nf_total`, `valor_transp`, origem, destino, serviço, fornecedor, destinatário|
|`fornecedores` (f)|Nome do fornecedor do AWB|
|`fornecedores` (f_dest)|Flag `tad` do destinatário (habilita TDE)|
|`personalizados.vigencia_tabela_awb`|Tabela vigente do fornecedor na data de entrada|
|`aero`|Converte origem/destino do AWB em cidade (`aero.cidade` = `rotas.id_rota`)|
|`equipamento`|Converte `notas.servico` (`servico_cia`) em `id_equipamento`|
|`tabela_trecho`|Trechos da tabela: origem, destino, serviço, aero mínimo, percentual de frete|
|`rotas` / `cidade_zona_grupo`|Match de origem/destino por cidade, UF ou zona|
|`tabela_faixas`|Faixas de cada tipo de cobrança do trecho|
|`relatorios.cliente_trecho`|Descontos por cliente e trecho (somente fornecedores em `FORNECEDORES_CLIENTE_TRECHO`)|
|`personalizados.db_awb`|Dados complementares do AWB para o relatório|

---

## 4. Filtros

### 4.1 AWBs (`SQL_NOTAS_TRECHO`)

- `notas.data_entrada BETWEEN DATA_INICIAL AND DATA_FINAL`.
- `notas.fornecedor IN (FORNECEDOR)`.
- Tabela vigente: `vigencia_tabela_awb.data_inicial <= data_entrada` e (`data_final >= data_entrada` ou `data_final IS NULL`).
- Trecho ativo: `tabela_trecho.status = 1`, mesmo `id_tabela` da vigência e mesmo serviço (`tabela_trecho.servico = equipamento.id_equipamento`, com `equipamento.servico_cia = notas.servico`).
- Origem do trecho compatível: `tabela_trecho.origem = 0`, ou igual à cidade de origem, ou igual à UF (`rotas.uf_ibge`) da cidade de origem.
- Destino do trecho compatível: `tabela_trecho.destino = 0`, ou igual à cidade, ou à UF, ou a uma zona (`cidade_zona_grupo.id_zona`) que contenha a cidade de destino.

> AWBs sem vigência, sem equipamento, sem aeroporto cadastrado ou sem trecho compatível **não aparecem** no resultado (os JOINs são obrigatórios).

### 4.2 Faixas (`SQL_FAIXAS`)

- Somente os trechos encontrados para os AWBs.
- `tabela_faixas.deleted_at IS NULL`.
- Ordenadas por `id_trecho`, `tipo`, `indice`.
- Agrupadas por `(id_trecho, id_tabela)` → `{tipo: [faixas]}`.

### 4.3 Descontos por cliente (`SQL_CLIENTE_TRECHO`)

- Só é executada se algum AWB carregado for de fornecedor presente em `FORNECEDORES_CLIENTE_TRECHO`. Para os demais fornecedores, a consulta não roda e o cálculo é feito sem desconto.
- `cliente_trecho.id_cliente` = id do fornecedor do AWB.
- `cliente_trecho.status = 1`.

---

## 5. Seleção do trecho da tabela

Um AWB pode casar com mais de um trecho. É mantido **apenas o de menor `prioridade_trecho`**:

|Prioridade|Origem|Destino|
|---|---|---|
|1|Cidade (rota)|Cidade (rota)|
|2|UF|UF|
|3|Cidade|UF — ou UF × Cidade|
|4|Cidade|Zona|
|5|UF|Zona|
|6|Demais casos (ex.: origem/destino `0` = qualquer)||

Origem só casa por cidade ou UF. Destino casa por cidade, UF ou zona.

---

## 6. Seleção da linha da `cliente_trecho` (desconto)

Aplicada por AWB, apenas sobre as linhas do próprio fornecedor (`id_cliente`).

**Filtros da linha:**

- `id_origem` = cidade de origem do AWB **ou** `0` (qualquer origem).
- `id_destino` = cidade de destino do AWB **ou** `0` (qualquer destino).
- `servico` = `id_equipamento` do AWB **ou** `0` (vale para todos os serviços da tabela que tenham o trecho).
- `peso_inicio <= peso <= peso_fim` (limites nulos são ignorados).
- `data_entrada <= vigencia` — `vigencia` é a **data final** de aplicação do desconto; nula = sem limite.

**Prioridade (menor vence):**

| Prioridade | `id_origem`  | `id_destino` |
| ---------- | ------------ | ------------ |
| 1          | Cidade exata | Cidade exata |
| 2          | Cidade exata | `0`          |
| 3          | `0`          | Cidade exata |
| 4          | `0`          | `0`          |

Desempate dentro da mesma prioridade: `servico` específico antes de `servico = 0`.

Se nenhuma linha passar nos filtros, o AWB é calculado sem desconto.

---

## 7. Regras de cálculo por tipo (`tabela_faixas`)

### 7.1 Localização da faixa

- Regra geral: primeira faixa (na ordem de `indice`) com `inicio <= base <= fim`. Se `fim` for nulo ou `0`, não há limite superior.
- Tipos **14 e 23**: não localizam faixa — usam sempre a primeira faixa do tipo, pois `inicio` é o tamanho da fração.
- Se nenhuma faixa for localizada, o tipo vale `0`.
- Tipos fora de `TIPOS_VALIDOS` (exceto o 24) são ignorados.

### 7.2 Base de cada tipo

|Base|Tipos|
|---|---|
|`valor_transp`|4, 5, 32|
|`peso`|1, 6, 8, 11, 14, 22, 23, 25|
|`nf_total`|24 (TDE)|

### 7.3 Fórmulas sem desconto

|Tipo|Descrição|Fórmula|
|---|---|---|
|1, 6, 11, 22, 25|Normais (peso)|Se `peso <= franquia`: `mínimo`. Senão: `(peso − franquia) × excedente + mínimo`|
|8|Frete peso por faixa|Se `peso <= franquia`: `mínimo`. Senão: `MAX((peso − franquia) × excedente, mínimo)`|
|4|Advalorem direto|`(valor_transp − franquia) × excedente/100 + mínimo`|
|5|GRIS|`MAX(valor_transp × excedente/100, mínimo)`|
|32|Advalorem|`MAX(valor_transp × excedente/100 + franquia, mínimo)`|
|14|Pedágio por fração|`CEIL(peso / inicio) × excedente` (sem franquia, sem mínimo)|
|23|Fração|`CEIL(peso / inicio) × excedente` (sem franquia, sem mínimo)|
|24|TDE|Mesma lógica dos tipos normais, com base `nf_total` — ver seção 8|

Regras em todas as fórmulas:

- Franquia só é descontada quando `> 0`.
- Mínimo só é considerado quando `> 0`.
- **Tipo 6 (Despacho) exclui o tipo 1**: se o trecho tiver faixa de tipo 6, o tipo 1 é ignorado.

### 7.4 Ajustes da `cliente_trecho` (quando há linha selecionada)

O desconto é aplicado **em cada tipo**, sobre o valor cheio da `tabela_faixas`, **antes** do aero mínimo, do TDE e do percentual de frete.

|Tipo|Coluna(s)|Regra|
|---|---|---|
|8 (frete peso)|`desconto_frete_peso`, `desconto_frete_minimo`|Calculado **por partes e somado**: `(peso − franquia) × excedente × (1 − desconto_frete_peso/100)` **+** `mínimo × (1 − desconto_frete_minimo/100)`. Se `peso <= franquia`: só o mínimo com desconto|
|4, 32 (advalorem)|`adv_exc_fixo`|Se `> 0`, **substitui** o excedente (%) da faixa. Franquia e mínimo da faixa permanecem|
|5 (GRIS)|`gris_exc_fixo`|Se `> 0`, **substitui** o excedente (%) da faixa. Mínimo da faixa permanece. Se `0`, mantém o GRIS cheio|
|14, 23, demais|—|Sem alteração|

Colunas da `cliente_trecho` **ignoradas por enquanto**: `desconto_frete_coleta`, `desconto_frete_entrega`, `desconto_advalorem`, `desconto_gris`, `acrescimo_*`, `minimo_*`, `fixo_gris`, `desconto_fixo_peso_excedente`, `zona_*`, `estado_*`.

---

## 8. Composição do valor final

```
soma_sem_fracao  = soma de todos os tipos, exceto 14 e 23
valor_fracao     = tipo 14 + tipo 23
valor_com_minimo = MAX(soma_sem_fracao, aero_minimo) + valor_fracao   (MAX só se aero_minimo > 0)
valor_com_tde    = valor_com_minimo + TDE
valor_final      = valor_com_tde × (1 + percentual_frete / 100)
```

- **Aero mínimo** (`tabela_trecho.aero_minimo`): compara apenas com os tipos sem fração; tipos 14 e 23 são somados depois.
- **TDE (tipo 24)**: só é calculado se `fornecedores.tad = 1` para o **destinatário** do AWB e se o trecho tiver faixa de tipo 24. Base = `nf_total`. Não recebe desconto da `cliente_trecho`.
- **Percentual de frete** (`tabela_trecho.percentual_frete`): aplicado por último; nulo/0 não altera o valor.

---

## 9. Exemplo validado — JEM (169715), tabela 2325

Dados do AWB: peso 30,00 · valor transportado 1.961,16 · origem 3550308 · destino 1302603 · trecho 3742879.

`cliente_trecho` selecionada: **12700** (`id_origem = 0`, `id_destino = 1302603`), pois não há linha com origem 3550308 e destino 1302603 juntos.

|Componente|Tabela cheia|Com desconto|
|---|---|---|
|Advalorem|1.961,16 × 0,20% = 3,92|1.961,16 × 0,15% (`adv_exc_fixo`) = 2,94|
|GRIS|1.961,16 × 0,23% = 4,51|`gris_exc_fixo = 0` → 4,51|
|Pedágio|CEIL(30 / 100) × 5,02 = 5,02|5,02|
|Frete peso — excedente|(30 − 10) × 56,44 = 1.128,80|1.128,80 − 66,35% = 379,84|
|Frete peso — mínimo|167,96|167,96 − 40,53% = 99,89|
|Frete peso — resultado|MAX(1.128,80; 167,96) = 1.128,80|379,84 + 99,89 = 479,73|

---

## 10. Saída (Excel)

Uma linha por AWB com trecho encontrado.

|Coluna|Origem|
|---|---|
|AWB|`notas.awb`|
|ID NOTA|`notas.id_nota`|
|FORNECEDOR|`fornecedores.fantasia`|
|VALOR|`notas.nf_total` (valor cobrado)|
|ID TABELA|`vigencia_tabela_awb.tabela`|
|PESO|`notas.peso`|
|ID TRECHO|`tabela_trecho.id_trecho` selecionado|
|ID TRECHO CLIENTE|`cliente_trecho.id_trecho_cliente` aplicado (vazio se sem desconto)|
|VALOR TABELA SEM DESCONTO|Valor final calculado sem a `cliente_trecho`|
|VALOR TABELA|Valor final calculado (com desconto, se houver)|
|DIFERENÇA|`VALOR TABELA − VALOR` (positivo = cobrado abaixo da tabela)|
|TDE|Valor do tipo 24|
|PERCENTUAL FRETE|`tabela_trecho.percentual_frete`|
|DATA EMISSÃO, ORIGEM, DESTINO, SERVIÇO AWB, STATUS|`personalizados.db_awb` (LEFT JOIN por `cod_awb = ID NOTA`)|

Valores arredondados em 2 casas. Formatação via openpyxl: cabeçalho `#1F3864` com fonte branca em negrito, todas as células centralizadas com borda fina e largura de coluna automática.

---

## 11. Pendências conhecidas

- **Desempate de trechos**: AWBs com dois trechos na mesma prioridade ficam com o primeiro retornado pelo banco (sem critério definido).
- **`cliente_trecho.servico`**: comparado com `equipamento.id_equipamento`; a confirmar se o correto é `notas.servico`.
- **Escopo do desconto**: válido apenas para o tipo 8 (frete peso), tipos 4/32 (advalorem) e 5 (GRIS), e apenas para o fornecedor 169715.

