# Especificação — Geração do Relatório de Medição ("A GERADORA")

> **Escopo deste documento:** apenas a **lógica de geração do relatório**. A partir das duas planilhas enviadas na página de upload, o backend deve gerar **1 relatório PDF por cont_id**, no layout do modelo `A Geradora ES.xlsx`. O PDF gerado é o insumo que o fluxo N8N existente anexa junto com XML/PDF/BOLETO.
>
> **Status:** rascunho v1 — estruturar junto antes de implementar. Itens marcados com 🔶 são decisões pendentes.

---

## 1. Objetivo

Transformar as duas planilhas abaixo em um **relatório de medição em PDF por cliente (cont_id)**, idêntico em conteúdo e layout ao modelo `A Geradora ES.xlsx`, para que o fluxo N8N `Baixar_PDF_XML_BL_test` anexe esse PDF junto aos documentos do Omie (NFSe/XML, PDF e Boleto) de cada cliente.

A chave que liga tudo é o **`cont_id`** (CONT ID).

---

## 2. Entradas

### 2.1 Planilha "Banco de Dados" (`Banco de Dados Sillion ...xlsx`)

Base por equipamento. A aba que interessa para o relatório é **`Detalhamentos PDF`**.

- **Cabeçalho de período** (linhas 1–2): `DATA INICIAL` (F2), `DATA FINAL` (G2), `QTD. DIAS` (H2).
- **Tabela de detalhe** (a partir da linha 4 = cabeçalho; dados na linha 5+):

  | Col | Campo | Uso |
  |-----|-------|-----|
  | A | CONT ID | **chave** / agrupamento |
  | B | CLIENTE ID | identificação |
  | C | CLIENTE NOME | identificação / título |
  | D | NOME | linha de detalhe |
  | E | PLACA | linha de detalhe |
  | F | OPERAÇÃO DO MÊS | linha de detalhe |
  | G | DATA DA OPERAÇÃO | linha de detalhe |
  | H | DIAS / QTD. | linha de detalhe |
  | I | MENSALIDADE | valor |
  | J | PROPORCIONAL MENSAL | valor (entra no total) |
  | K | VALOR OPERAÇÃO* | valor (entra no total) |
  | L | OBSERVAÇÕES | livre |

  \* No modelo `A Geradora ES` a coluna K se chama **"VALOR OPERAÇÃO"**; na aba `Detalhamentos PDF` o título vem como **"OPERAÇÃO"**. Tratar como o mesmo campo (valor de operação avulsa). 🔶 *confirmar nome final da coluna.*

- ⚠️ **A aba vem agrupada com linhas de subtotal.** Para cada cont_id existe uma linha extra `"<cont_id> Total"` e, ao final, uma linha `"Total Geral"`. **Essas linhas de subtotal devem ser ignoradas na leitura dos detalhes** (são resultado de agrupamento do Excel, não são equipamentos).

### 2.2 Planilha "TOT_sillion" (`TOT Sillion ...xlsx`)

Soma por cliente do que será faturado. Aba principal **`Tot`**, com `CONT_ID` como identificador:

| Campo | Uso no relatório |
|-------|------------------|
| CONT_ID | **chave** de cruzamento |
| CLIENTE_ID, NOME | identificação |
| TOTAL | valor de referência do cliente |
| DESCONTO, ST NF, IR, IMP. FED., ISS | impostos/ajustes |
| LIQUIDO | valor líquido |

Uso: **valor de conferência (reconciliação)** do total do relatório por cont_id. Outras abas da TOT (Clientes, RPS, NF, Boletas, Boletos) **não** entram na geração do relatório — pertencem ao fluxo de NF/boleto do N8N.

---

## 3. Saída

- **1 arquivo PDF por cont_id**, layout do modelo `A Geradora ES.xlsx`.
- Gerado apenas para cont_ids **que tenham linhas de detalhe** na aba `Detalhamentos PDF`. 🔶 *confirmar: gerar também para clientes que estão na TOT mas não têm detalhe? (ver §6)*
- **Nomenclatura do arquivo** (proposta, para o N8N casar por cont_id):
  `{cont_id}_{CLIENTE_NOME_sanitizado}_medicao_{AAAA-MM}.pdf`
  ex.: `526_A_GERADORA_MINAS_GERAIS_medicao_2026-05.pdf`
  🔶 *confirmar padrão de nome esperado pelo N8N.*

---

## 4. Layout do PDF (modelo "A Geradora")

Estrutura de cada página/relatório, reproduzindo `A Geradora ES.xlsx`:

```
┌───────────────────────────────────────────────────────────────┐
│  PLANILHA DE MEDIÇÃO                                            │
│  PERÍODO   DATA INICIAL: 01/05/2026   DATA FINAL: 31/05/2026    │
│            QTD. DIAS: 31                                        │
│                                                                 │
│  CLIENTE: {CLIENTE NOME}   (CONT ID: {x} | CLIENTE ID: {y})     │
│                                          TOTAL (R$): {total}    │
├───────────────────────────────────────────────────────────────┤
│ NOME | PLACA | OPERAÇÃO DO MÊS | DATA | DIAS | MENSALIDADE |    │
│       PROPORCIONAL MENSAL | VALOR OPERAÇÃO | OBSERVAÇÕES         │
│ ...uma linha por equipamento (placa)...                         │
├───────────────────────────────────────────────────────────────┤
│                                   TOTAL GERAL (R$): {total}     │
└───────────────────────────────────────────────────────────────┘
```

Observações do modelo:
- O bloco de cabeçalho de período é o **mesmo** para todos os clientes (vem do cabeçalho da aba).
- As colunas A/B/C (CONT ID, CLIENTE ID, CLIENTE NOME) podem aparecer como **identificação do cliente no topo** (não repetir em cada linha) — no modelo elas existem mas o nome se repete. 🔶 *definir: repetir colunas A–C por linha (como no Excel) ou só no cabeçalho do bloco.*

---

## 5. Lógica de geração (passo a passo)

1. **Ler período** do cabeçalho da aba `Detalhamentos PDF` (F2/G2/H2).
2. **Ler detalhes**: linhas a partir da 5, **descartando** linhas de subtotal (`"… Total"`, `"Total Geral"`) e linhas em branco.
3. **Agrupar por `cont_id`** (coluna A).
4. **Ler TOT** (aba `Tot`) em um dicionário por `CONT_ID` → {NOME, TOTAL, LIQUIDO}.
5. Para cada cont_id com detalhes:
   1. Montar cabeçalho (período + identificação do cliente).
   2. Listar as linhas de detalhe ordenadas (por PLACA → DATA). 🔶 *confirmar ordenação.*
   3. Calcular **TOTAL GERAL** (ver §6).
   4. Conferir contra `TOT.TOTAL` do cont_id (validação — ver §6).
   5. Renderizar o PDF e salvar com a nomenclatura de §3.
6. Disponibilizar os PDFs para o fluxo N8N (ver §7).

---

## 6. Regras de cálculo e reconciliação

- **TOTAL GERAL (R$)** do relatório = **Σ PROPORCIONAL MENSAL (col J) + Σ VALOR OPERAÇÃO (col K)** das linhas de detalhe do cont_id.
  - No modelo `A Geradora ES`: 2 placas × 245 (proporcional) = **490**, sem valor de operação → TOTAL 490. ✔️
- A linha de subtotal `"<cont_id> Total"` na aba já traz esse total agregado — pode ser usada como **conferência cruzada** contra o valor calculado.
- **PROPORCIONAL MENSAL**: quando `DIAS = QTD. DIAS do período` (mês cheio), proporcional = mensalidade. Caso contrário é proporcional aos dias. 🔶 *confirmar se o backend recalcula ou só lê o valor já calculado na planilha.* Recomendação inicial: **ler o valor pronto** (a planilha já vem calculada) e não recalcular.
- **Reconciliação com TOT**: comparar `TOTAL GERAL` do relatório com `TOT.TOTAL` do mesmo cont_id.
  - ⚠️ Observado nos dados de teste: vários cont_ids da aba `Detalhamentos PDF` **não foram encontrados na aba `Tot`** (chaves não bateram). Isso precisa ser entendido — pode ser diferença de período, UN diferente, ou cont_id formatado diferente. 🔶 **decisão pendente importante.**
- Tratar valores `#N/A`, vazios e texto em colunas numéricas como **0** no somatório.

---

## 7. Ponto de integração com o N8N (handoff)

> Detalhamento fora do escopo deste MD, registrado só para alinhar a interface.

- O fluxo `Baixar_PDF_XML_BL_test` hoje: Webhook → busca no Omie por cliente (NFSe/XML via `Obter_Docs`, Boleto via `Obter_Boleto`/`PesquisarLancamentos_por_Cliente`) → agrupa → `Agrupa_Zip` → `Upload file` (Google Drive) → `Create folder` / `Share file` → e-mail (`Send a message`).
- **Onde o relatório entra:** o PDF de medição de cada cont_id deve ser **adicionado ao mesmo agrupamento por cliente** (entrar no `Agrupa_Zip` / pasta do Drive / anexo do e-mail), casando pelo `cont_id`.
- **Interface mínima necessária:** o backend precisa expor os PDFs de forma que o N8N os recupere por `cont_id` (ex.: pasta/URL por cont_id, ou retorno no payload do webhook). 🔶 *definir o mecanismo: arquivo em disco/Drive, base64 no payload, ou endpoint de download.*

---

## 8. Casos de borda

- **cont_id com detalhe mas sem registro na TOT** → gerar relatório mesmo assim? (visto nos dados de teste). 🔶
- **cont_id na TOT sem nenhuma linha de detalhe** → não gera relatório (nada a detalhar) — confirmar.
- **Linhas de subtotal e "Total Geral"** → sempre descartar.
- **Cliente com nome divergente** entre `Detalhamentos PDF` e `Tot` (ex.: "A GERADORA - MINAS GERAIS") → definir qual fonte manda no título. Sugestão: usar o nome da aba de detalhe.
- **Valores negativos / ISS** → não entram no relatório de medição (são da TOT/NF).
- **Caracteres especiais no nome** → sanitizar para o nome do arquivo.

---

## 9. Decisões pendentes (resumo dos 🔶)

1. Nome final da coluna K ("VALOR OPERAÇÃO" vs "OPERAÇÃO").
2. Gerar relatório para cont_id sem correspondência na TOT?
3. Repetir colunas A–C por linha ou só no cabeçalho do bloco.
4. Ordenação das linhas de detalhe.
5. Backend recalcula proporcional ou lê valor pronto da planilha.
6. Entender por que cont_ids do detalhe não batem com a TOT.
7. Mecanismo de handoff dos PDFs para o N8N (disco/Drive/base64/endpoint) e padrão de nome do arquivo.

---

## 10. Próximos passos

1. Fechar as decisões da §9.
2. Validar o layout do PDF contra um exemplo real (gerar 1 PDF de teste para o cont_id 526 e comparar com `A Geradora ES.xlsx`).
3. Especificar a §7 (integração N8N) em um MD próprio, se necessário.
