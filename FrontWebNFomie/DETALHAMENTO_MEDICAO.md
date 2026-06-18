# Detalhamento da Medição — geração de relatórios + envio em lote

Nova funcionalidade do front Streamlit: a partir de **um arquivo já tratado**, o
front gera **1 relatório PDF por cont_id** (layout "PLANILHA DE MEDIÇÃO" / modelo
"A Geradora") e envia **tudo em um único lote** ao N8N, cada relatório separado e
codificado em **Base64**.

Todo o tratamento acontece **no front** — o N8N recebe os PDFs prontos.

## Seletor (cliente específico)

A seção de upload fica atrás de um **seletor** (`st.toggle`) no formulário:
**"Anexar Detalhamento da Medição (cliente específico)"**.

- **Desativado (padrão):** o app segue como hoje — só a extração de NFSe/XML/Boleto
  pelo botão Download. O painel de upload não aparece.
- **Ativado:** o painel de upload aparece e a string identificadora
  **`enviar_email`** é enviada no campo `opcao` em **ambas** as requisições
  (Download e lote). Quando desativado, `opcao` vai vazio (`""`).

---

## Arquivos

| Arquivo | Papel |
|---------|-------|
| `relatorio_medicao.py` | Leitura da planilha, agrupamento por cont_id, geração dos PDFs e montagem do payload do lote. Não depende do Streamlit. |
| `app.py` | Seção "Detalhamento da Medição": upload, preview e botão de envio do lote. |

---

## Entrada (arquivo enviado)

- `.xlsx` contendo a aba **`Detalhamentos PDF`** (se não existir aba com esse nome, usa a primeira).
- **Período**: lido automaticamente do cabeçalho — rótulos `DATA INICIAL`, `DATA FINAL`, `QTD. DIAS`, com o valor na célula imediatamente abaixo.
- **Cabeçalho da tabela**: linha cujo primeiro campo é `CONT ID` (detectado automaticamente; fallback = linha 5).
- **Dados**: linhas após o cabeçalho. **Linhas de subtotal** (`"576 Total"`, `"Total Geral"`, qualquer célula com "Total") são **ignoradas**.

Colunas esperadas (ordem do modelo):
`CONT ID | CLIENTE ID | CLIENTE NOME | NOME | PLACA | OPERAÇÃO DO MÊS | DATA DA OPERAÇÃO | DIAS / QTD. | MENSALIDADE | PROPORCIONAL MENSAL | VALOR OPERAÇÃO | OBSERVAÇÕES`

---

## Regra de cálculo

- **TOTAL GERAL** (por cont_id) = Σ `PROPORCIONAL MENSAL` + Σ `VALOR OPERAÇÃO`.
- Valores em branco / texto / `#N/A` contam como **0**.
- `PROPORCIONAL MENSAL` é lido pronto da planilha (não é recalculado).

---

## Saída — nome do arquivo

`{cont_id}_{CLIENTE_NOME_sanitizado}_medicao_{AAAA-MM}.pdf`

Ex.: `576_PORTO_DO_ACU_OPERACOES_S_A_G_PRUMO_medicao_2026-05.pdf`

---

## Contrato do lote (payload do webhook)

`POST` JSON único contendo **todos** os relatórios, cada um separado no array `lote`:

```json
{
  "email": "willian.silva@sillion.com.br",
  "empresa": "Sillion",
  "tipo": "detalhamento_medicao",
  "opcao": "enviar_email",
  "periodo": {
    "data_inicial": "2026-05-01",
    "data_final": "2026-05-31",
    "qtd_dias": 31
  },
  "qtd_relatorios": 10,
  "lote": [
    {
      "cont_id": "576",
      "cliente_id": "23804",
      "cliente_nome": "PORTO DO ACU OPERACOES S.A. - G PRUMO",
      "filename": "576_PORTO_DO_ACU_..._medicao_2026-05.pdf",
      "mime_type": "application/pdf",
      "total": 5001.39,
      "qtd_veiculos": 34,
      "file_base64": "JVBERi0xLjQK..."
    }
  ]
}
```

| Campo (envelope) | Tipo | Descrição |
|------------------|------|-----------|
| `email` | string | E-mail corporativo (validado @sillion.com.br) |
| `empresa` | string | `Sitrack` ou `Sillion` |
| `tipo` | string | Sempre `detalhamento_medicao` (para o Switch do N8N rotear) |
| `opcao` | string | `enviar_email` quando o seletor está ativo; `""` quando inativo |
| `periodo` | object | `data_inicial`, `data_final` (ISO `YYYY-MM-DD`), `qtd_dias` |
| `qtd_relatorios` | int | Quantidade de itens no lote |
| `lote` | array | Um item por cont_id (ver abaixo) |

| Campo (item do lote) | Tipo | Descrição |
|----------------------|------|-----------|
| `cont_id` | string | Chave do cliente (casa com a busca no Omie) |
| `cliente_id` | string | ID do cliente |
| `cliente_nome` | string | Nome do cliente |
| `filename` | string | Nome do PDF |
| `mime_type` | string | `application/pdf` |
| `total` | number | Total geral do relatório |
| `qtd_veiculos` | int | Nº de linhas/veículos |
| `file_base64` | string | PDF em Base64 |

### Consumir no N8N

No nó **Webhook** (POST), iterar `{{ $json.lote }}` (ex.: nó **Split Out** / **Item Lists**) e, para cada item, reconstruir o binário:

```javascript
const it = $json;  // item do lote
const buffer = Buffer.from(it.file_base64, 'base64');
return [{
  json: { cont_id: it.cont_id, cliente_nome: it.cliente_nome, total: it.total },
  binary: { data: { data: buffer, mimeType: it.mime_type, fileName: it.filename } }
}];
```

O PDF de cada `cont_id` deve ser **anexado ao mesmo agrupamento** dos documentos
baixados do Omie (NFSe/XML, PDF, Boleto), casando por `cont_id`.

---

## Configuração

- Reusa `N8N_WEBHOOK_URL` por padrão. Para um webhook dedicado ao lote, defina
  em `.streamlit/secrets.toml`:

  ```toml
  N8N_WEBHOOK_URL      = "https://.../webhook/<id-extracao>"
  N8N_WEBHOOK_LOTE_URL = "https://.../webhook/<id-lote-medicao>"   # opcional
  ```

- Dependências adicionais (já em `requirements.txt`): `openpyxl`, `reportlab`.

---

## Decisões em aberto (ajustáveis)

1. Webhook dedicado para o lote (`N8N_WEBHOOK_LOTE_URL`) ou reusar o principal.
2. Mostrar/baixar os PDFs no próprio front antes de enviar (preview individual).
3. Incluir bloco de impostos (ISS/LÍQUIDO) no PDF — hoje só TOTAL GERAL.
4. Limite de tamanho do lote (muitos PDFs em Base64 num único POST) — avaliar
   envio em partes se o nº de clientes crescer muito.
