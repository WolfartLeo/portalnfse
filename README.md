# PortalNFSe - UI (MVP)

Interface web simples (Streamlit) em cima do robô `bot_nfse.py`, com **login**, seleção de **competência** e **clientes**, execução e **download** do resultado.

## O que este MVP entrega
- Login/senha do **aplicativo**
- Cadastro de clientes (em Excel)
- Processamento por competência e por cliente
- Download da competência em ZIP (PDF/XML + LOG)

## Como rodar (Windows)
```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
streamlit run app.py
```

## Segurança / Governança
- `users.json`, `config.local.json` e a planilha real **não vão para o Git** (estão no `.gitignore`).
- Na primeira execução é criado um admin padrão:
  - usuário: `admin`
  - senha: `admin123`  (troque imediatamente)

## URL pública sem pagar hosting
Padrão “raiz”: rodar no PC/servidor do escritório (Windows) e expor por túnel (Cloudflare Tunnel).


## Nota (Streamlit + threads)
Este MVP evita `missing ScriptRunContext` porque o worker thread **não chama st.*** e comunica progresso via fila.


## Fix WinError 193
Este pacote força o uso do **Selenium Manager** (desabilita webdriver_manager) para evitar driver baixado errado/corrompido no Windows.


## Regra: nota cancelada (Portal)
O robô detecta nota cancelada pela coluna **Situação** (ícone `tb-cancelada.svg` ou tooltip `NFS-e cancelada`). Mesmo assim baixa XML/PDF, mas no LOG final zera todos os valores monetários.


## Login: bcrypt
O login usa hash seguro (passlib/bcrypt). Se aparecer `MissingBackendError`, instale com a venv ativa:

- `pip install bcrypt`

Se seu Python for 3.13 e der erro de build, use Python 3.12 para criar a venv.


## Fix Auditoria: PermissionError com '~$'
No Windows, se o Excel estiver aberto, ele cria arquivos temporários `~$*.xlsx` que ficam bloqueados. A auditoria ignora esses arquivos e segue.


## Login (PBKDF2)
O login usa `passlib` com `pbkdf2_sha256` (sem bcrypt). Se existir `users.json` antigo com hash bcrypt, apague/renomeie o arquivo para recriar o ADMIN.


## UI/UX
Inclui skin (CSS) e tema Streamlit em `.streamlit/config.toml` para deixar o app mais apresentável.
