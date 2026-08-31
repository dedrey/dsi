"""
Historico de processos ja calculados, salvo em disco (historico.json, ao
lado do app.py) para poder usar um processo anterior como "espelho"/modelo
ao iniciar um novo -- sem precisar redigitar tudo de novo quando o cliente
ou o tipo de mercadoria se repete.

Formato de cada entrada:
{
    "timestamp": "2026-08-31T14:32:00",
    "resumo": "31/08/2026 14:32 — DOUGLAS BENICIO DA SILVA — NCM 8711.90.00",
    "campos": { <key do widget>: <valor (string, ou "YYYY-MM-DD" p/ datas)> }
}
"""

from __future__ import annotations

import json
from datetime import date, datetime
from pathlib import Path

CAMINHO_HISTORICO = Path(__file__).parent / "historico.json"
LIMITE_ENTRADAS = 30

# Campos de data precisam de tratamento especial (string <-> date) ao salvar
# e ao carregar de volta para os widgets.
CAMPOS_DATA = ("data_desembarque_val", "data_darf_val", "data_chegada_val")


def carregar_historico() -> list[dict]:
    if not CAMINHO_HISTORICO.exists():
        return []
    try:
        with open(CAMINHO_HISTORICO, encoding="utf-8") as f:
            return json.load(f)
    except (json.JSONDecodeError, OSError):
        return []


def _serializar_campos(campos: dict) -> dict:
    """Converte os valores dos widgets (que podem incluir objetos date) para
    algo serializavel em JSON."""
    saida = {}
    for chave, valor in campos.items():
        if chave in CAMPOS_DATA:
            saida[chave] = valor.isoformat() if isinstance(valor, date) else None
        else:
            saida[chave] = valor
    return saida


def desserializar_campos(campos: dict) -> dict:
    """Caminho inverso: strings ISO viram objetos date de novo, prontos
    para serem colocados no session_state dos widgets de data."""
    saida = {}
    for chave, valor in campos.items():
        if chave in CAMPOS_DATA:
            saida[chave] = date.fromisoformat(valor) if valor else None
        else:
            saida[chave] = valor
    return saida


def salvar_no_historico(campos: dict, resumo: str) -> None:
    """Adiciona uma nova entrada e mantem so as LIMITE_ENTRADAS mais
    recentes (evita o arquivo crescer sem controle)."""
    historico = carregar_historico()
    historico.append(
        {
            "timestamp": datetime.now().isoformat(timespec="seconds"),
            "resumo": resumo,
            "campos": _serializar_campos(campos),
        }
    )
    historico = historico[-LIMITE_ENTRADAS:]
    with open(CAMINHO_HISTORICO, "w", encoding="utf-8") as f:
        json.dump(historico, f, ensure_ascii=False, indent=2)