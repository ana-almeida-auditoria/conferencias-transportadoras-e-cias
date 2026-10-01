## O que esse relatório aponta?

Mostra se o valor cobrado em cada **AWB** (o documento de transporte emitido pelo prestador de transporte) está de acordo com a **tabela de preços combinada** com esse prestador.

O relatório trata de **AWB**. Não é minuta, coleta, cotação nem fatura.

Para cada AWB, o relatório refaz a conta do frete pela tabela e coloca lado a lado o valor cobrado e o valor que deveria ter sido cobrado. Assim fica fácil identificar cobranças acima ou abaixo do combinado.

## Quais filtros são utilizados?

- **Período:** AWBs que deram entrada entre duas datas escolhidas.
- **Prestador:** um ou mais prestadores de transporte escolhidos.
- **Tabela vigente:** só vale a tabela de preços que estava valendo na data de entrada do AWB.
- **Trecho cadastrado:** o percurso do AWB (cidade de origem até cidade de destino, no serviço contratado) precisa existir na tabela. AWBs sem percurso cadastrado não entram no relatório.

## Quais regras foram aplicadas?

1. **Escolha do percurso:** quando mais de uma linha da tabela serve, vale a mais específica, nesta ordem:
    - cidade com cidade;
    - estado com estado;
    - cidade com estado;
    - por último, regiões (zonas).
2. **Cálculo das taxas:** cada taxa é calculada conforme a tabela:
    - **Frete peso:** pelo peso, com valor mínimo.
    - **Advalorem e GRIS:** percentual sobre o valor da mercadoria, com valor mínimo.
    - **Pedágio:** cobrado a cada fração de peso, arredondando para cima.
3. **Mínimo do trecho:** se a soma das taxas ficar abaixo do mínimo do percurso, cobra-se o mínimo. O pedágio é somado depois.
4. **TDE (taxa de dificuldade de entrega):** é somada só quando o destinatário está marcado como local de difícil entrega.
5. **Percentual de frete:** acréscimo percentual do percurso, aplicado por último sobre o total.
6. **Desconto negociado (hoje só para o prestador JEM):** se houver desconto cadastrado para aquele percurso, ele é aplicado antes do mínimo, da TDE e do percentual de frete:
    - **Frete peso:** um percentual de desconto sobre a parte do peso e outro sobre o valor mínimo, somando as duas partes.
    - **Advalorem e GRIS:** o percentual da tabela é trocado pelo percentual negociado, quando existir.
    - **Pedágio:** não muda.
    - **Escolha do desconto:** primeiro procura um desconto para a origem e o destino exatos; se não achar, aceita "qualquer origem" ou "qualquer destino"; por último, o desconto que vale para todos os percursos.
    - **Validade:** o desconto só vale até a data de vigência cadastrada.

## Qual é o resultado final?

Uma planilha Excel com **uma linha por AWB**, contendo:

- AWB, prestador, peso e o valor cobrado;
- a tabela e o percurso usados no cálculo, e o desconto aplicado, quando houver;
- o valor calculado pela tabela, **sem** e **com** desconto;
- a **diferença** entre o valor da tabela e o valor cobrado:
    - **positiva:** foi cobrado **menos** que a tabela;
    - **negativa:** foi cobrado **mais** que a tabela;
- dados complementares do AWB (emissão, origem, destino, serviço e status).