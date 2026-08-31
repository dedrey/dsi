"""
Preenchimento do DARF (Documento de Arrecadacao de Receitas Federais) a
partir do modelo oficial (ModeloDarf.pdf), que tem campos de formulario de
verdade. Gera um PDF com 4 paginas -- uma DARF para cada tributo federal da
DSI (I.I., I.P.I., PIS, COFINS). O ICMS NAO tem DARF -- e pago pela guia
estadual (GRPR/GNRE), ja tratada em outro lugar do app.

Nomes dos campos do formulario original (levantados com pypdf.get_fields()):
    Nome         -> Nome / Telefone
    NI           -> Numero do CPF ou CNPJ
    Receita      -> Codigo da Receita
    Referência   -> Numero de Referencia
    Vencimento   -> Data de Vencimento
    Apuração     -> Periodo de Apuracao
    Principal    -> Valor do Principal
    Multa        -> Valor da Multa
    Juros        -> Valor dos Juros
    Total        -> Valor Total
"""

from __future__ import annotations

import io
from datetime import date

from pypdf import PdfReader, PdfWriter
from pypdf.generic import BooleanObject, NameObject, RectangleObject, TextStringObject

from calculo import CODIGOS_RECEITA_FEDERAL, ResultadoTributos
from preenchimento import DadosProcesso

# Conforme instrucao: este campo e SEMPRE o mesmo, independente do tributo.
NUMERO_REFERENCIA_DSI = "0910600"

# Recorte da pagina original: tira o titulo "Modelo de Documento..." e a
# tabela de instrucoes de preenchimento que vinham no PDF de referencia,
# deixando so a guia em si (o que faz sentido entregar ao cliente).
_CORTE_PAGINA = RectangleObject((28, 418, 587, 738))


def _fmt_moeda(valor: float) -> str:
    s = f"{valor:,.2f}"
    return s.replace(",", "§").replace(".", ",").replace("§", ".")


def _preencher_uma_pagina(template_bytes: bytes, campos: dict, prefixo: str) -> PdfWriter:
    reader = PdfReader(io.BytesIO(template_bytes))
    writer = PdfWriter()
    writer.append(reader)
    for page in writer.pages:
        writer.update_page_form_field_values(page, campos)
        page.cropbox = _CORTE_PAGINA
        page.mediabox = _CORTE_PAGINA
        # Renomeia os campos com um prefixo unico por tributo. Sem isso,
        # quando juntamos varias paginas desse MESMO template num PDF so,
        # campos de mesmo nome (ex.: "Receita") em paginas diferentes sao
        # tratados como o MESMO campo compartilhando um unico valor --
        # so o primeiro que for preenchido "vence" em todas as paginas.
        for annot in page.get("/Annots") or []:
            obj = annot.get_object()
            nome_atual = obj.get("/T")
            if nome_atual:
                obj[NameObject("/T")] = TextStringObject(f"{prefixo}_{nome_atual}")
    # Forca os leitores de PDF a recalcular a aparencia dos campos, senao
    # alguns visualizadores mostram os campos em branco ate o usuario clicar
    # neles.
    writer.set_need_appearances_writer(True)
    return writer


def gerar_darf_pdf(
    template_bytes: bytes,
    dados: DadosProcesso,
    r: ResultadoTributos,
    data_darf: date | None,
    telefone: str = "",
) -> bytes:
    """Gera um PDF com 4 paginas (I.I., I.P.I., PIS, COFINS), cada uma um
    DARF preenchido a partir do template oficial."""
    data_fmt = data_darf.strftime("%d/%m/%Y") if data_darf else ""
    nome_telefone = dados.nome_importador + (f" - {telefone}" if telefone else "")

    tributos = [
        ("I.I.", CODIGOS_RECEITA_FEDERAL["I.I."], r.ii),
        ("I.P.I.", CODIGOS_RECEITA_FEDERAL["I.P.I."], r.ipi),
        ("PIS", CODIGOS_RECEITA_FEDERAL["PIS"], r.pis),
        ("COFINS", CODIGOS_RECEITA_FEDERAL["COFINS"], r.cofins),
    ]

    writer_final = PdfWriter()
    for nome_tributo, codigo_receita, valor in tributos:
        campos = {
            "Nome": nome_telefone,
            "NI": dados.cpf_cnpj,
            "Receita": codigo_receita,
            "Referência": NUMERO_REFERENCIA_DSI,
            "Apuração": data_fmt,
            "Vencimento": data_fmt,
            "Principal": _fmt_moeda(valor),
            "Total": _fmt_moeda(valor),
        }
        prefixo = nome_tributo.replace(".", "")
        writer_pagina = _preencher_uma_pagina(template_bytes, campos, prefixo)
        writer_final.append(writer_pagina)

    writer_final.set_need_appearances_writer(True)
    buf = io.BytesIO()
    writer_final.write(buf)
    return buf.getvalue()