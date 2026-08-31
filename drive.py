"""
Envio dos documentos gerados (Excel, Word, DARF) para uma pasta no Google
Drive -- usa LOGIN OAUTH (sua propria conta Google), nao uma conta de
servico. Trocamos de abordagem porque a organizacao tinha uma politica de
seguranca bloqueando a criacao de chaves de conta de servico, e o proprio
Google recomenda OAuth para esse tipo de uso (um script rodando no
computador de alguem, nao um servidor).

Configuracao necessaria, NENHUMA delas fica no codigo nem e enviada pra
lugar nenhum alem do seu computador:

    credentials/oauth_client.json
        -- baixado do Cloud Console: Credenciais > Criar credenciais >
           ID do cliente OAuth > tipo "Aplicativo para computador"

    drive_config.json
        -- {"pasta_raiz_id": "ID_DA_PASTA_RAIZ_NO_DRIVE_COMPARTILHADO"}

Na primeira vez que alguem gerar um processo com o Drive configurado, abre
uma janela do navegador pedindo login com a conta Google e autorizacao de
acesso ao Drive. Depois disso, um arquivo credentials/token.json guarda
essa autorizacao -- nao pede login de novo nas proximas vezes.

Se credentials/oauth_client.json ou drive_config.json nao existirem, o
envio ao Drive e simplesmente PULADO (retorna None) -- o app continua
funcionando 100% normal sem essa parte, o Drive e so um extra por cima do
que ja funciona.
"""

from __future__ import annotations

import io
import json
from datetime import date
from pathlib import Path

PASTA_APP = Path(__file__).parent
PASTA_CREDENCIAIS = PASTA_APP / "credentials"
CAMINHO_CLIENT_SECRET = PASTA_CREDENCIAIS / "oauth_client.json"
CAMINHO_TOKEN = PASTA_CREDENCIAIS / "token.json"
CAMINHO_CONFIG = PASTA_APP / "drive_config.json"

SCOPES = ["https://www.googleapis.com/auth/drive"]

MIME_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
MIME_DOCX = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
MIME_PDF = "application/pdf"


def esta_configurado() -> bool:
    """True se o client OAuth e o arquivo de config existirem. (O token.json
    ainda nao precisa existir -- ele e criado sozinho no primeiro uso.)"""
    return CAMINHO_CLIENT_SECRET.exists() and CAMINHO_CONFIG.exists()


def _carregar_config() -> dict:
    with open(CAMINHO_CONFIG, encoding="utf-8") as f:
        return json.load(f)


def _obter_credenciais():
    # Imports feitos aqui dentro (nao no topo do arquivo) para que o app
    # inteiro nao quebre se essas bibliotecas nao estiverem instaladas em
    # algum ambiente onde o Drive nunca vai ser usado.
    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow

    creds = None
    if CAMINHO_TOKEN.exists():
        creds = Credentials.from_authorized_user_file(str(CAMINHO_TOKEN), SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            # Abre o navegador pedindo login -- so acontece na primeira vez
            # (ou se o token for revogado/expirar sem poder renovar sozinho).
            flow = InstalledAppFlow.from_client_secrets_file(
                str(CAMINHO_CLIENT_SECRET), SCOPES,
            )
            creds = flow.run_local_server(port=0)

        PASTA_CREDENCIAIS.mkdir(parents=True, exist_ok=True)
        with open(CAMINHO_TOKEN, "w", encoding="utf-8") as f:
            f.write(creds.to_json())

    return creds


def _obter_servico():
    from googleapiclient.discovery import build

    creds = _obter_credenciais()
    return build("drive", "v3", credentials=creds, cache_discovery=False)


def _encontrar_ou_criar_pasta(servico, nome: str, pasta_pai_id: str) -> str:
    """Reaproveita a pasta se ja existir (evita duplicar "2026-08" toda vez
    que alguem calcula um processo no mesmo mes)."""
    nome_escapado = nome.replace("'", "\\'")
    query = (
        f"name = '{nome_escapado}' and mimeType = 'application/vnd.google-apps.folder' "
        f"and '{pasta_pai_id}' in parents and trashed = false"
    )
    resultado = servico.files().list(
        q=query,
        fields="files(id, name)",
        supportsAllDrives=True,
        includeItemsFromAllDrives=True,
        corpora="allDrives",
    ).execute()
    encontrados = resultado.get("files", [])
    if encontrados:
        return encontrados[0]["id"]

    metadata = {
        "name": nome,
        "mimeType": "application/vnd.google-apps.folder",
        "parents": [pasta_pai_id],
    }
    pasta = servico.files().create(
        body=metadata, fields="id", supportsAllDrives=True,
    ).execute()
    return pasta["id"]


def _upload_arquivo(servico, nome: str, conteudo: bytes, mime_type: str, pasta_id: str) -> str:
    from googleapiclient.http import MediaIoBaseUpload

    media = MediaIoBaseUpload(io.BytesIO(conteudo), mimetype=mime_type, resumable=False)
    metadata = {"name": nome, "parents": [pasta_id]}
    arquivo = servico.files().create(
        body=metadata,
        media_body=media,
        fields="id, webViewLink",
        supportsAllDrives=True,
    ).execute()
    return arquivo.get("webViewLink", "")


def enviar_processo_para_drive(
    nome_cliente: str,
    xlsx_bytes: bytes,
    docx_bytes: bytes,
    darf_bytes: bytes,
) -> dict[str, str] | None:
    """
    Sobe os 3 documentos pra uma subpasta organizada por ano-mes e depois
    por cliente+data, dentro da pasta raiz configurada em drive_config.json:

        <pasta raiz> / 2026-08 / DOUGLAS BENICIO DA SILVA - 2026-08-31 /
            FORMULARIO_DSI_REMATE_2026-08-31.xlsx
            Demonstrativo_Calculo_2026-08-31.docx
            DARF_2026-08-31.pdf

    Devolve um dict {"excel": link, "word": link, "darf": link}, ou None
    se o Drive nao estiver configurado (nesse caso o chamador deve seguir
    sem erro -- e um recurso opcional).
    """
    if not esta_configurado():
        return None

    config = _carregar_config()
    pasta_raiz_id = config["pasta_raiz_id"]

    servico = _obter_servico()

    hoje = date.today()
    pasta_mes = _encontrar_ou_criar_pasta(servico, f"{hoje:%Y-%m}", pasta_raiz_id)
    nome_pasta_processo = f"{nome_cliente or 'Sem nome'} - {hoje:%Y-%m-%d}"
    pasta_processo = _encontrar_ou_criar_pasta(servico, nome_pasta_processo, pasta_mes)

    links = {
        "excel": _upload_arquivo(
            servico, f"FORMULARIO_DSI_REMATE_{hoje:%Y-%m-%d}.xlsx",
            xlsx_bytes, MIME_XLSX, pasta_processo,
        ),
        "word": _upload_arquivo(
            servico, f"Demonstrativo_Calculo_{hoje:%Y-%m-%d}.docx",
            docx_bytes, MIME_DOCX, pasta_processo,
        ),
        "darf": _upload_arquivo(
            servico, f"DARF_{hoje:%Y-%m-%d}.pdf",
            darf_bytes, MIME_PDF, pasta_processo,
        ),
    }
    return links