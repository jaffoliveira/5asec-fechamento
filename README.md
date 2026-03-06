# 🧺 5ASEC — Fechamento de Caixa

Aplicação web para automatizar o fechamento diário das 4 lojas.

## Funcionalidades

| Funcionalidade | Como funciona |
|---|---|
| **Parseia PDF de fechamento** | Arraste os PDFs exportados do sistema Windows |
| **Importa Excel da Rede** | Preenche crédito e débito automaticamente |
| **Entrada manual** | Formulários para depósito (Itaú) e sangrias detalhadas |
| **Dashboard gerencial** | KPIs, gráficos de barras e pizza, tabela comparativa |
| **Tendência histórica** | Evolução 30 dias ao carregar o template de fechamento |
| **Exporta Excel** | Preenche o template "Fechamento de Caixa - 2026.xlsx" |
| **Exporta CSV** | Backup ou integração com outras ferramentas |

---

## Fluxo diário

```
1. Abra o app no navegador
2. Selecione a data do fechamento
3. Arraste os PDFs das 4 lojas → clique "Processar PDFs"
4. Para cada loja, importe o Excel da Rede (crédito/débito)
5. Informe depósito/PIX manualmente (Itaú) para cada loja
6. Preencha as sangrias detalhadas (água, café, banco, etc.)
7. Veja o dashboard → aba "Dashboard"
8. Exporte o template preenchido → aba "Exportar"
```

---

## Como hospedar (Streamlit Cloud — gratuito)

### Passo 1 — Crie um repositório no GitHub

1. Acesse [github.com](https://github.com) e faça login (crie conta se necessário)
2. Clique em **New repository**
3. Nome sugerido: `5asec-fechamento`
4. Deixe como **Private** (não aparece para ninguém)
5. Clique **Create repository**

### Passo 2 — Suba os arquivos

Opção mais simples (sem precisar instalar nada):
1. Na página do repositório clique **Add file → Upload files**
2. Arraste os arquivos `app.py` e `requirements.txt`
3. Clique **Commit changes**

### Passo 3 — Deploy no Streamlit Cloud

1. Acesse [share.streamlit.io](https://share.streamlit.io)
2. Faça login com sua conta GitHub
3. Clique **New app**
4. Selecione seu repositório `5asec-fechamento`
5. Branch: `main` | Main file path: `app.py`
6. Clique **Deploy!**
7. Aguarde ~2 minutos → você receberá um link permanente

**O link funciona em qualquer dispositivo (computador, celular, tablet).**

---

## Como executar localmente (Windows)

```bash
# Instale Python 3.10+ se ainda não tiver
# Abra o Prompt de Comando na pasta do projeto

pip install -r requirements.txt
streamlit run app.py
```

O navegador abrirá automaticamente em `http://localhost:8501`

---

## Como exportar o PDF no sistema Windows (5ASEC)

No sistema Windows, na tela de **Leitura X**:
1. Clique em **[F11] Imprimir**
2. Na janela de impressão, selecione **"Microsoft Print to PDF"** (ou "Salvar como PDF")
3. Salve com o nome da loja (ex: `pompeia_04032026.pdf`)
4. Arraste para o app

---

## Identificação automática das lojas

O app identifica a loja pelo texto do PDF (endereço/telefone no cabeçalho):

| Loja | Palavras-chave detectadas |
|---|---|
| West Side / Pompéia | `POMPEIA, 1700` ou `3672 1466` |
| West Zone / Sonda | `CARLOS VICARI` ou `SONDA POMPEIA` |
| West Place / Girassol | `GIRASSOL` *(adicione endereço no app.py)* |
| West Station / Paulistânia | `PAULISTANIA, 91` ou `3675 4094` |

Se a loja não for identificada automaticamente, o app pedirá seleção manual.

---

## Importação do Excel da Rede

1. Acesse o portal da Rede e exporte o relatório do dia em Excel
2. Na barra lateral do app, faça upload em **"Excel da Rede"**
3. Selecione a loja correspondente
4. Clique **"Importar Rede"** — crédito e débito são preenchidos automaticamente

> Se o app não detectar os valores automaticamente (formato diferente), ele mostrará uma prévia do arquivo para você identificar as colunas. Nesse caso, entre em contato para ajustar o parser ao formato específico da sua conta Rede.

---

## Personalização

Para adicionar o endereço da loja **West Place / Girassol** na detecção automática:

1. Abra `app.py`
2. Encontre `STORE_KEYWORDS`
3. Adicione o endereço/telefone da loja na lista `"PLACE"`

```python
STORE_KEYWORDS = {
    ...
    "PLACE": ["GIRASSOL", "RUA FULANO DE TAL, 123", "3XXX-XXXX"],
    ...
}
```
