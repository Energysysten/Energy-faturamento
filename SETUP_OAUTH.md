# Configuração do Gmail API — Passo a Passo

Este guia configura o acesso da automação à caixa energysystenfaturamento@gmail.com.
Faça isso **uma única vez**.

---

## Pré-requisitos

- Acesso ao e-mail energysystenfaturamento@gmail.com (para autorizar no navegador)
- Python 3.8+ instalado no computador

---

## Passo 1 — Instalar as dependências Python

Abra o Terminal e execute:

```bash
pip3 install google-auth-oauthlib google-auth-httplib2 google-api-python-client pdfplumber openpyxl
```

---

## Passo 2 — Criar o projeto no Google Cloud Console

1. Acesse: https://console.cloud.google.com
2. Faça login com a conta Google que gerencia o projeto
   (pode ser sua conta pessoal — não precisa ser a conta de faturamento)
3. No canto superior esquerdo, clique em **"Selecionar projeto"** → **"Novo projeto"**
4. Nome: `Faturamento Energy Systen` → clique em **Criar**

---

## Passo 3 — Ativar a Gmail API

1. Com o projeto selecionado, vá em: **APIs e Serviços → Biblioteca**
2. Pesquise `Gmail API`
3. Clique em **Gmail API** → **Ativar**

---

## Passo 4 — Configurar a Tela de Consentimento OAuth

1. Vá em: **APIs e Serviços → Tela de consentimento OAuth**
2. Tipo de usuário: **Externo** → Criar
3. Preencha:
   - Nome do app: `Automação Faturamento`
   - E-mail de suporte: (seu e-mail)
   - Informações de contato: (seu e-mail)
4. Clique em **Salvar e continuar** (nas telas seguintes pode deixar em branco)
5. Na tela **Usuários de teste**: clique em **Adicionar usuários**
   → adicione: `energysystenfaturamento@gmail.com`
6. Clique em **Salvar e continuar**

---

## Passo 5 — Criar as Credenciais OAuth

1. Vá em: **APIs e Serviços → Credenciais**
2. Clique em **+ Criar credenciais → ID do cliente OAuth**
3. Tipo de aplicativo: **App para computador**
4. Nome: `Script Faturamento`
5. Clique em **Criar**
6. Na janela que abrir, clique em **Baixar JSON**
7. **Renomeie o arquivo baixado para `credentials.json`**
8. Mova o arquivo para a pasta do projeto:
   ```
   /Users/leonardocarmo/Documents/Claude/Projects/Faturamento/credentials.json
   ```

---

## Passo 6 — Primeira execução (autorização)

No Terminal, navegue até a pasta do projeto e execute:

```bash
cd "/Users/leonardocarmo/Documents/Claude/Projects/Faturamento"
python3 processar_medicoes.py
```

- Uma janela do navegador abrirá automaticamente
- Faça login com **energysystenfaturamento@gmail.com**
- Clique em **Continuar** (pode aparecer aviso de app não verificado — é normal)
- Autorize as permissões de leitura/modificação do Gmail
- O arquivo `token.json` será salvo automaticamente
- O script processará qualquer PDF já recebido

**Nas próximas execuções**, o script roda sem abrir o navegador — o token é renovado automaticamente.

---

## Resultado esperado

Após a execução você verá no Terminal algo como:

```
[2026-05-16 08:00:01] ============================================================
[2026-05-16 08:00:01] Iniciando processamento de folhas de medição
[2026-05-16 08:00:02] 2 email(s) com PDF encontrado(s).
[2026-05-16 08:00:03] Processando: 'Folha de Medição - Março/26' de fornecedor@email.com
[2026-05-16 08:00:04]   Anexo salvo: 1011557405 - PORANGATU ENERGY MARÇO26.pdf
[2026-05-16 08:00:05]   Planilha atualizada: folha 1011557405 | R$ 204.834,99
[2026-05-16 08:00:06] Concluído. 1 folha(s) processada(s).
```

- PDFs salvos em: `Anexos/`
- Planilha atualizada: `Controle_Medicoes.xlsx`
- Emails marcados com label `Medicao-Processada` no Gmail

---

## Dúvidas

Se tiver algum erro, peça ajuda ao Claude e cole a mensagem de erro — ele consegue diagnosticar.
