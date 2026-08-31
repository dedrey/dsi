import json
from datetime import date
from pathlib import Path

import streamlit as st
import streamlit.components.v1 as components

from calculo import (
    calcular_tributos,
    CODIGOS_RECEITA_FEDERAL,
    COFINS_ALIQUOTA_PADRAO,
    PIS_ALIQUOTA_PADRAO,
)
from preenchimento import DadosProcesso, preencher_excel, preencher_word
from darf import gerar_darf_pdf
from historico import carregar_historico, desserializar_campos, salvar_no_historico
from drive import enviar_processo_para_drive, esta_configurado as drive_configurado

# Caminho absoluto (baseado na pasta deste arquivo, nao na pasta de onde o
# comando "streamlit run" foi executado) -- evita erro de "arquivo nao
# encontrado" se o app for rodado a partir de outro diretorio.
PASTA_APP = Path(__file__).parent
PASTA_TEMPLATES = PASTA_APP / "templates"

st.set_page_config(page_title="DSI", layout="centered")


# ---------------------------------------------------------------------------
# Estilo (fontes/tabela) + mascara de digitacao (virgula automatica e NCM
# com pontos). O CSS roda via st.markdown normalmente. O JS PRECISA rodar
# via components.html (st.markdown nunca executa <script>, o navegador
# ignora scripts inseridos via innerHTML por seguranca) -- o script mira em
# window.parent porque os campos do formulario ficam fora do iframe do
# componente.
# ---------------------------------------------------------------------------
st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=Space+Grotesk:wght@500;600;700&family=IBM+Plex+Mono:wght@400;500;600&display=swap');

    html, body, [class*="css"] {
        font-family: 'Inter', -apple-system, sans-serif;
    }
    h1, h2, h3 {
        font-family: 'Space Grotesk', 'Inter', sans-serif !important;
        letter-spacing: -0.01em;
    }

    [data-testid="stForm"] {
        background: #FFFFFF;
        border: 1px solid #D7DCE3;
        border-radius: 12px;
        padding: 1.75rem 1.75rem 1.25rem 1.75rem;
    }
    /* Borda sempre visivel nos campos (antes so aparecia ao focar/clicar) */
    [data-testid="stTextInput"] input,
    [data-testid="stDateInput"] input,
    [data-testid="stSelectbox"] div[data-baseweb="select"] > div {
        border: 1.5px solid #0F4C5C !important;
        border-radius: 8px !important;
    }
    [data-testid="stTextInput"] input {
        font-family: 'IBM Plex Mono', monospace;
        text-align: right;
    }
    [data-testid="stTextInput"] input:focus,
    [data-testid="stDateInput"] input:focus {
        border: 1.5px solid #9C6B2E !important;
        box-shadow: none !important;
    }
    /* Esconde a dica "Press Enter to submit form" que o Streamlit mostra
       embaixo de cada campo de texto dentro de um form. */
    [data-testid="InputInstructions"] {
        display: none !important;
    }
    [data-testid="stFormSubmitButton"] button {
        background: #0F4C5C;
        color: #FFFFFF;
        border: none;
        border-radius: 8px;
        font-weight: 600;
        letter-spacing: 0.02em;
        padding: 0.6rem 0;
    }
    [data-testid="stFormSubmitButton"] button:hover {
        background: #0C3A47;
        color: #FFFFFF;
    }
    [data-testid="stDownloadButton"] button {
        background: #9C6B2E;
        color: #FFFFFF;
        border: none;
        border-radius: 8px;
        font-weight: 600;
    }
    [data-testid="stDownloadButton"] button:hover {
        background: #7A511F;
        color: #FFFFFF;
    }
    [data-testid="stMetricValue"] {
        font-family: 'IBM Plex Mono', monospace;
        color: #9C6B2E;
    }
    [data-testid="stAlert"] {
        border-radius: 10px;
        border-left: 4px solid #0F4C5C;
    }

    .dsi-eyebrow {
        font-family: 'IBM Plex Mono', monospace;
        font-size: 0.78rem;
        letter-spacing: 0.06em;
        text-transform: uppercase;
        color: #5B6472;
        margin: 0.25rem 0 0.75rem 0;
    }
    .dsi-ledger {
        width: 100%;
        border-collapse: collapse;
        margin-top: 0.25rem;
        font-size: 0.92rem;
        background: #FFFFFF;
        border: 1px solid #D7DCE3;
        border-radius: 10px;
        overflow: hidden;
    }
    .dsi-ledger th {
        background: #0F4C5C;
        color: #FFFFFF;
        font-family: 'Inter', sans-serif;
        font-weight: 600;
        text-align: left;
        padding: 0.6rem 0.8rem;
        font-size: 0.75rem;
        text-transform: uppercase;
        letter-spacing: 0.04em;
    }
    .dsi-ledger td {
        padding: 0.55rem 0.8rem;
        border-bottom: 1px solid #E2E5EA;
        color: #1A2233;
        font-family: 'IBM Plex Mono', monospace;
    }
    .dsi-ledger td.label {
        font-family: 'Inter', sans-serif;
        font-weight: 600;
    }
    .dsi-ledger tr:nth-child(even) td {
        background: #F7F9FB;
    }
    .dsi-ledger td.num {
        text-align: right;
    }
    .dsi-ledger tr.total td {
        background: #FBF3E7;
        border-top: 2px solid #9C6B2E;
        border-bottom: none;
        font-weight: 700;
        color: #7A511F;
        font-family: 'Inter', sans-serif;
    }
    .dsi-ledger tr.total td.num {
        font-family: 'IBM Plex Mono', monospace;
        font-weight: 700;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

components.html(
    """
    <script>
    (function () {
        var doc = window.parent.document;

        function setValue(input, value) {
            if (input.value === value) return;
            var setter = Object.getOwnPropertyDescriptor(
                window.parent.HTMLInputElement.prototype, 'value'
            ).set;
            setter.call(input, value);
            input.dispatchEvent(new window.parent.Event('input', { bubbles: true }));
            window.parent.requestAnimationFrame(function () {
                input.setSelectionRange(value.length, value.length);
            });
        }

        function maskMoeda(input) {
            var digits = input.value.replace(/\\D/g, '');
            if (digits === '') { setValue(input, ''); return; }
            digits = digits.replace(/^0+(?=\\d)/, '');
            while (digits.length < 3) digits = '0' + digits;
            var cents = digits.slice(-2);
            var inteiro = digits.slice(0, -2).replace(/\\B(?=(\\d{3})+(?!\\d))/g, '.');
            setValue(input, inteiro + ',' + cents);
        }

        function maskNcm(input) {
            var digits = input.value.replace(/\\D/g, '').slice(0, 8);
            var out = digits;
            if (digits.length > 6) {
                out = digits.slice(0, 4) + '.' + digits.slice(4, 6) + '.' + digits.slice(6);
            } else if (digits.length > 4) {
                out = digits.slice(0, 4) + '.' + digits.slice(4);
            }
            setValue(input, out);
        }

        function maskCpfCnpj(input) {
            var digits = input.value.replace(/\\D/g, '').slice(0, 14);
            var out;
            if (digits.length <= 11) {
                out = digits
                    .replace(/(\\d{3})(\\d)/, '$1.$2')
                    .replace(/(\\d{3})(\\d)/, '$1.$2')
                    .replace(/(\\d{3})(\\d{1,2})$/, '$1-$2');
            } else {
                out = digits
                    .replace(/(\\d{2})(\\d)/, '$1.$2')
                    .replace(/(\\d{3})(\\d)/, '$1.$2')
                    .replace(/(\\d{3})(\\d)/, '$1/$2')
                    .replace(/(\\d{4})(\\d{1,2})$/, '$1-$2');
            }
            setValue(input, out);
        }

        var CAMPOS_MOEDA = [
            'Valor Aduaneiro / Valor da Mercadoria (R$)',
            'Alíquota I.I. (%) — Portal Aduaneiras',
            'Alíquota I.P.I. (%) — Portal Aduaneiras',
            'Alíquota ICMS (%) — ajuste se o produto tiver alíquota diferenciada',
            'Alíquota PIS (%) — Portal Aduaneiras',
            'Alíquota COFINS (%) — Portal Aduaneiras',
            'Valor FOB do item (US$)',
            'Valor do frete (US$)',
            'Valor do seguro (US$)'
        ];
        var CAMPO_NCM = 'NCM (Classificação Fiscal)';
        var CAMPO_CPF_CNPJ = 'CPF/CNPJ';

        function anexarMascaras() {
            CAMPOS_MOEDA.forEach(function (label) {
                doc.querySelectorAll('input[aria-label="' + label + '"]').forEach(function (input) {
                    if (input.dataset.mascarado) return;
                    input.dataset.mascarado = '1';
                    input.addEventListener('input', function () { maskMoeda(input); });
                });
            });
            doc.querySelectorAll('input[aria-label="' + CAMPO_NCM + '"]').forEach(function (input) {
                if (input.dataset.mascarado) return;
                input.dataset.mascarado = '1';
                input.addEventListener('input', function () { maskNcm(input); });
            });
            doc.querySelectorAll('input[aria-label="' + CAMPO_CPF_CNPJ + '"]').forEach(function (input) {
                if (input.dataset.mascarado) return;
                input.dataset.mascarado = '1';
                input.addEventListener('input', function () { maskCpfCnpj(input); });
            });
        }

        anexarMascaras();
        new window.parent.MutationObserver(anexarMascaras).observe(doc.body, {
            childList: true, subtree: true,
        });
    })();
    </script>
    """,
    height=0,
)


def fmt_moeda(valor: float) -> str:
    """Formata no padrao brasileiro: 2.367,22"""
    s = f"{valor:,.2f}"
    return s.replace(",", "§").replace(".", ",").replace("§", ".")


def fmt_pct(valor: float) -> str:
    return f"{valor:.2f}".replace(".", ",")


def parse_numero_br(texto: str) -> float:
    """Converte um numero digitado no padrao brasileiro ("2.367,22") para float."""
    texto = (texto or "").strip()
    if not texto:
        return 0.0
    texto = texto.replace(".", "").replace(",", ".")
    try:
        return float(texto)
    except ValueError:
        return 0.0


@st.cache_data
def carregar_icms_por_uf() -> dict:
    with open("icms_uf.json", encoding="utf-8") as f:
        dados = json.load(f)
    dados.pop("_fonte", None)
    return dados


def codigo_icms(uf: str) -> str:
    return "1210 (GRPR)" if uf == "PR" else "100056 (GNRE)"


def montar_tabela_html(r, ncm: str, uf: str) -> str:
    linhas = [
        ("I.I.", r.valor_mercadoria, r.aliquota_ii, CODIGOS_RECEITA_FEDERAL["I.I."], r.ii),
        ("I.P.I.", r.base_ipi, r.aliquota_ipi, CODIGOS_RECEITA_FEDERAL["I.P.I."], r.ipi),
        ("PIS", r.valor_mercadoria, r.aliquota_pis, CODIGOS_RECEITA_FEDERAL["PIS"], r.pis),
        ("COFINS", r.valor_mercadoria, r.aliquota_cofins, CODIGOS_RECEITA_FEDERAL["COFINS"], r.cofins),
        ("ICMS", r.base_icms, r.aliquota_icms, codigo_icms(uf), r.icms),
    ]
    linhas_html = "".join(
        f"<tr><td class='label'>{nome}</td>"
        f"<td class='num'>{fmt_moeda(base)}</td>"
        f"<td class='num'>{fmt_pct(aliq)}%</td>"
        f"<td>{codigo}</td>"
        f"<td class='num'>{fmt_moeda(imposto)}</td></tr>"
        for nome, base, aliq, codigo, imposto in linhas
    )
    ncm_html = f"<p class='dsi-eyebrow'>NCM (Classificação Fiscal): {ncm}</p>" if ncm else ""
    return f"""
    {ncm_html}
    <table class="dsi-ledger">
        <thead>
            <tr>
                <th>Tributo</th>
                <th style="text-align:right">Base de Cálculo (R$)</th>
                <th style="text-align:right">Alíquota</th>
                <th>Código Receita</th>
                <th style="text-align:right">Imposto a Recolher (R$)</th>
            </tr>
        </thead>
        <tbody>
            {linhas_html}
            <tr class="total">
                <td colspan="4">TOTAL DE TRIBUTOS</td>
                <td class="num">R$ {fmt_moeda(r.total_tributos)}</td>
            </tr>
        </tbody>
    </table>
    """


icms_por_uf = carregar_icms_por_uf()

# ---------------------------------------------------------------------------
# Valores padrao de cada campo, aplicados so na primeira vez (setdefault) --
# nunca por cima de um valor ja existente no session_state. Isso evita o
# aviso do Streamlit de "widget criado com valor padrao E com valor setado
# via Session State API", que acontecia quando um campo tinha `value=...` E
# também recebia um valor carregado do historico ao mesmo tempo.
# ---------------------------------------------------------------------------
_PADROES = {
    "ncm_texto": "",
    "valor_mercadoria_texto": "",
    "aliquota_ii_texto": "",
    "aliquota_ipi_texto": "",
    "uf_selecionada": "PR",
    "aliquota_icms_texto": fmt_pct(icms_por_uf["PR"]),
    "aliquota_pis_texto": fmt_pct(PIS_ALIQUOTA_PADRAO),
    "aliquota_cofins_texto": fmt_pct(COFINS_ALIQUOTA_PADRAO),
    "nome_importador_texto": "",
    "cpf_cnpj_texto": "",
    "telefone_texto": "",
    "endereco_texto": "",
    "rg_texto": "",
    "nacionalidade_texto": "BRASILEIRA",
    "data_desembarque_val": None,
    "pais_procedencia_texto": "PARAGUAI",
    "termo_entrada_texto": "",
    "valor_fob_texto": "",
    "valor_frete_texto": "",
    "valor_seguro_texto": "",
    "taxa_conversao_texto": "",
    "item_qtde_texto": "1",
    "item_unidade_texto": "UNID.",
    "item_descricao_texto": "",
    "local_data_texto": f"FOZ DO IGUAÇU, {date.today():%d/%m/%Y}",
    "data_darf_val": None,
    "transportador_texto": "",
    "identificacao_veiculo_texto": "",
    "data_chegada_val": None,
    "numero_conhecimento_texto": "",
    "qtde_volumes_texto": "",
    "peso_bruto_texto": "",
    "peso_liquido_texto": "",
    "depositario_texto": "",
}
for _chave, _valor_padrao in _PADROES.items():
    st.session_state.setdefault(_chave, _valor_padrao)

st.title("DSI")

# ---------------------------------------------------------------------------
# Historico -- carregar um processo anterior como modelo. Fica FORA do
# form de propósito: precisa rodar (e escrever no session_state de cada
# campo) antes dos widgets do form serem criados mais abaixo no script.
# ---------------------------------------------------------------------------
historico_salvo = carregar_historico()
if historico_salvo:
    opcao_em_branco = "— Novo processo em branco —"
    opcoes_historico = [opcao_em_branco] + [h["resumo"] for h in reversed(historico_salvo)]
    col_hist1, col_hist2 = st.columns([4, 1])
    with col_hist1:
        escolha_historico = st.selectbox(
            "Usar processo anterior como modelo", opcoes_historico, key="escolha_historico",
        )
    with col_hist2:
        st.write("")
        clique_carregar = st.button("Carregar", use_container_width=True)
    if clique_carregar and escolha_historico != opcao_em_branco:
        entrada = next(
            h for h in reversed(historico_salvo) if h["resumo"] == escolha_historico
        )
        for chave, valor in desserializar_campos(entrada["campos"]).items():
            st.session_state[chave] = valor
        st.rerun()


def _ao_mudar_uf():
    nova_uf = st.session_state.get("uf_selecionada")
    if nova_uf:
        st.session_state["aliquota_icms_texto"] = fmt_pct(icms_por_uf[nova_uf])


st.subheader("Cliente")
uf_selecionada = st.selectbox(
    "UF do cliente", sorted(icms_por_uf.keys()),
    key="uf_selecionada",
    on_change=_ao_mudar_uf,
)
st.caption("A alíquota de ICMS abaixo se atualiza sozinha ao trocar a UF.")

with st.form("dados_tributos"):
    col1, col2 = st.columns(2)

    with col1:
        st.subheader("Mercadoria")
        ncm_texto = st.text_input(
            "NCM (Classificação Fiscal)", max_chars=10, key="ncm_texto",
        )
        valor_mercadoria_texto = st.text_input(
            "Valor Aduaneiro / Valor da Mercadoria (R$)", placeholder="0,00",
            key="valor_mercadoria_texto",
        )
        aliquota_ii_texto = st.text_input(
            "Alíquota I.I. (%) — Portal Aduaneiras", placeholder="0,00",
            key="aliquota_ii_texto",
        )
        aliquota_ipi_texto = st.text_input(
            "Alíquota I.P.I. (%) — Portal Aduaneiras", placeholder="0,00",
            key="aliquota_ipi_texto",
        )

    with col2:
        st.subheader("ICMS / PIS / COFINS")
        aliquota_icms_texto = st.text_input(
            "Alíquota ICMS (%) — ajuste se o produto tiver alíquota diferenciada",
            key="aliquota_icms_texto",
        )
        aliquota_pis_texto = st.text_input(
            "Alíquota PIS (%) — Portal Aduaneiras",
            key="aliquota_pis_texto",
        )
        aliquota_cofins_texto = st.text_input(
            "Alíquota COFINS (%) — Portal Aduaneiras",
            key="aliquota_cofins_texto",
        )

    st.divider()
    st.subheader("Dados do processo")
    col3, col4 = st.columns(2)

    with col3:
        st.markdown("**Importador**")
        nome_importador_texto = st.text_input("Nome / Nome Empresarial", key="nome_importador_texto")
        cpf_cnpj_texto = st.text_input("CPF/CNPJ", max_chars=18, key="cpf_cnpj_texto")
        telefone_texto = st.text_input("Telefone (opcional, para o DARF)", key="telefone_texto")
        endereco_texto = st.text_input("Endereço completo", key="endereco_texto")
        rg_texto = st.text_input("RG / Passaporte", key="rg_texto")
        nacionalidade_texto = st.text_input(
            "Nacionalidade", key="nacionalidade_texto",
        )
        data_desembarque_val = st.date_input(
            "Data do desembarque", format="DD/MM/YYYY", key="data_desembarque_val",
        )

    with col4:
        st.markdown("**Carga e item**")
        pais_procedencia_texto = st.text_input(
            "País de procedência", key="pais_procedencia_texto",
        )
        termo_entrada_texto = st.text_input(
            "Termo de entrada (Nº do Termo de Retenção)",
            placeholder="0917500-236227/2026", key="termo_entrada_texto",
        )
        valor_fob_texto = st.text_input(
            "Valor FOB do item (US$)", placeholder="0,00", key="valor_fob_texto",
        )
        valor_frete_texto = st.text_input(
            "Valor do frete (US$)", placeholder="0,00", key="valor_frete_texto",
        )
        valor_seguro_texto = st.text_input(
            "Valor do seguro (US$)", placeholder="0,00", key="valor_seguro_texto",
        )
        taxa_conversao_texto = st.text_input(
            "Taxa de conversão (R$)", placeholder="0,0000", key="taxa_conversao_texto",
        )

    item_qtde_texto = st.text_input("Quantidade do item", key="item_qtde_texto")
    item_unidade_texto = st.text_input("Unidade", key="item_unidade_texto")
    item_descricao_texto = st.text_input("Descrição do item", key="item_descricao_texto")
    local_data_texto = st.text_input(
        "Local e data (declaração)", key="local_data_texto",
    )
    data_darf_val = st.date_input(
        "Data do DARF (período de apuração / vencimento)", format="DD/MM/YYYY",
        key="data_darf_val",
    )

    with st.expander("Outros dados (opcional, raramente usados)"):
        col5, col6 = st.columns(2)
        with col5:
            transportador_texto = st.text_input("Transportador", key="transportador_texto")
            identificacao_veiculo_texto = st.text_input(
                "Identificação do veículo", key="identificacao_veiculo_texto",
            )
            data_chegada_val = st.date_input(
                "Data de chegada", format="DD/MM/YYYY", key="data_chegada_val",
            )
            numero_conhecimento_texto = st.text_input(
                "Nº do conhecimento / etiqueta de bagagem", key="numero_conhecimento_texto",
            )
        with col6:
            qtde_volumes_texto = st.text_input("Qtde volumes", key="qtde_volumes_texto")
            peso_bruto_texto = st.text_input("Peso bruto (kg)", key="peso_bruto_texto")
            peso_liquido_texto = st.text_input("Peso líquido (kg)", key="peso_liquido_texto")
            depositario_texto = st.text_input("Depositário/Armazém", key="depositario_texto")

    calcular = st.form_submit_button("Calcular tributos", use_container_width=True)

# ---------------------------------------------------------------------------
# O calculo e a geracao dos documentos rodam quando o form e enviado, e o
# resultado fica guardado em st.session_state -- assim ele continua na tela
# (e os botoes de download continuam funcionando) mesmo depois do rerun que
# o clique num st.download_button dispara. Sem isso, clicar em "Baixar"
# fazia a tela voltar pro estado inicial.
# ---------------------------------------------------------------------------
if calcular:
    try:
        r = calcular_tributos(
            valor_mercadoria=parse_numero_br(valor_mercadoria_texto),
            aliquota_ii=parse_numero_br(aliquota_ii_texto),
            aliquota_ipi=parse_numero_br(aliquota_ipi_texto),
            aliquota_icms=parse_numero_br(aliquota_icms_texto),
            aliquota_pis=parse_numero_br(aliquota_pis_texto),
            aliquota_cofins=parse_numero_br(aliquota_cofins_texto),
        )
    except ValueError as e:
        st.session_state.pop("resultado", None)
        st.error(str(e))
    else:
        try:
            item_qtde_val = int(parse_numero_br(item_qtde_texto)) if item_qtde_texto.strip() else 1
        except ValueError:
            item_qtde_val = 1

        dados_processo = DadosProcesso(
            nome_importador=nome_importador_texto,
            cpf_cnpj=cpf_cnpj_texto,
            endereco=endereco_texto,
            rg_passaporte=rg_texto,
            nacionalidade=nacionalidade_texto or "BRASILEIRA",
            data_desembarque=data_desembarque_val,
            ncm=ncm_texto.replace(".", ""),
            pais_procedencia=pais_procedencia_texto,
            termo_entrada=termo_entrada_texto,
            valor_frete_usd=parse_numero_br(valor_frete_texto),
            valor_seguro_usd=parse_numero_br(valor_seguro_texto),
            taxa_conversao=parse_numero_br(taxa_conversao_texto),
            item_qtde=item_qtde_val,
            item_unidade=item_unidade_texto or "UNID.",
            item_descricao=item_descricao_texto,
            valor_fob_usd=parse_numero_br(valor_fob_texto),
            local_data=local_data_texto,
            transportador=transportador_texto,
            identificacao_veiculo=identificacao_veiculo_texto,
            data_chegada=data_chegada_val,
            numero_conhecimento=numero_conhecimento_texto,
            qtde_volumes=qtde_volumes_texto,
            peso_bruto_kg=peso_bruto_texto,
            peso_liquido_kg=peso_liquido_texto,
            depositario=depositario_texto,
        )

        caminho_xlsx = PASTA_TEMPLATES / "FORMULARIO_DSI__REMATE.xlsx"
        caminho_docx = PASTA_TEMPLATES / "FORMULARIO_DSI_REMATE.docx"
        caminho_darf = PASTA_TEMPLATES / "ModeloDarf.pdf"
        try:
            with open(caminho_xlsx, "rb") as f:
                xlsx_bytes = preencher_excel(f.read(), dados_processo, r)
            with open(caminho_docx, "rb") as f:
                docx_bytes = preencher_word(f.read(), dados_processo, r)
            with open(caminho_darf, "rb") as f:
                darf_bytes = gerar_darf_pdf(
                    f.read(), dados_processo, r, data_darf_val, telefone=telefone_texto,
                )
        except FileNotFoundError:
            st.session_state.pop("resultado", None)
            st.error(
                "Não encontrei os templates dentro da pasta `templates/`. "
                "Confirme que os TRÊS arquivos abaixo existem, com esses nomes "
                "e extensões EXATOS:\n\n"
                f"- `{caminho_xlsx}`\n"
                f"- `{caminho_docx}`\n"
                f"- `{caminho_darf}`"
            )
        except ValueError as e:
            st.session_state.pop("resultado", None)
            st.error(str(e))
        else:
            links_drive = None
            if drive_configurado():
                try:
                    links_drive = enviar_processo_para_drive(
                        nome_importador_texto, xlsx_bytes, docx_bytes, darf_bytes,
                    )
                except Exception as e:
                    st.warning(f"Documentos gerados normalmente, mas não consegui enviar ao Google Drive: {e}")

            st.session_state["resultado"] = {
                "r": r,
                "ncm": ncm_texto,
                "uf": uf_selecionada,
                "xlsx_bytes": xlsx_bytes,
                "docx_bytes": docx_bytes,
                "darf_bytes": darf_bytes,
                "links_drive": links_drive,
            }

            # Salva este processo no historico para poder ser usado como
            # modelo de um processo futuro.
            campos_atuais = {
                "ncm_texto": ncm_texto,
                "valor_mercadoria_texto": valor_mercadoria_texto,
                "aliquota_ii_texto": aliquota_ii_texto,
                "aliquota_ipi_texto": aliquota_ipi_texto,
                "uf_selecionada": uf_selecionada,
                "aliquota_icms_texto": aliquota_icms_texto,
                "aliquota_pis_texto": aliquota_pis_texto,
                "aliquota_cofins_texto": aliquota_cofins_texto,
                "nome_importador_texto": nome_importador_texto,
                "cpf_cnpj_texto": cpf_cnpj_texto,
                "telefone_texto": telefone_texto,
                "endereco_texto": endereco_texto,
                "rg_texto": rg_texto,
                "nacionalidade_texto": nacionalidade_texto,
                "data_desembarque_val": data_desembarque_val,
                "pais_procedencia_texto": pais_procedencia_texto,
                "termo_entrada_texto": termo_entrada_texto,
                "valor_fob_texto": valor_fob_texto,
                "valor_frete_texto": valor_frete_texto,
                "valor_seguro_texto": valor_seguro_texto,
                "taxa_conversao_texto": taxa_conversao_texto,
                "item_qtde_texto": item_qtde_texto,
                "item_unidade_texto": item_unidade_texto,
                "item_descricao_texto": item_descricao_texto,
                "local_data_texto": local_data_texto,
                "data_darf_val": data_darf_val,
                "transportador_texto": transportador_texto,
                "identificacao_veiculo_texto": identificacao_veiculo_texto,
                "data_chegada_val": data_chegada_val,
                "numero_conhecimento_texto": numero_conhecimento_texto,
                "qtde_volumes_texto": qtde_volumes_texto,
                "peso_bruto_texto": peso_bruto_texto,
                "peso_liquido_texto": peso_liquido_texto,
                "depositario_texto": depositario_texto,
            }
            resumo_historico = (
                f"{date.today():%d/%m/%Y} — "
                f"{nome_importador_texto or 'Sem nome'} — "
                f"NCM {ncm_texto or '—'} — R$ {fmt_moeda(r.valor_mercadoria)}"
            )
            salvar_no_historico(campos_atuais, resumo_historico)

if "resultado" in st.session_state:
    res = st.session_state["resultado"]
    r = res["r"]

    st.subheader("Demonstrativo de Cálculo dos Tributos")
    st.markdown(montar_tabela_html(r, res["ncm"], res["uf"]), unsafe_allow_html=True)
    st.metric("TOTAL DE TRIBUTOS", f"R$ {fmt_moeda(r.total_tributos)}")

    st.subheader("Guia estadual de ICMS")
    if res["uf"] == "PR":
        st.info(
            "Cliente do **Paraná** → GRPR: "
            "https://emitirgrpr.sefa.pr.gov.br/arrecadacao/emitir/guiatela\n\n"
            "- Categoria: **ICMS**\n"
            "- Código de arrecadação: **1210** - Recolhimento Antecipado - Entradas do Exterior\n"
            "- Declaração Simplificada de Importação (DSI): **0910600 SEMPRE**"
        )
    else:
        st.info(
            f"Cliente de **{res['uf']}** (fora do Paraná) → GNRE: "
            "https://www.gnre.pe.gov.br:444/gnre/v/guia/index\n\n"
            "- Código Receita: **100056** - ICMS Importação"
        )

    st.subheader("Documentos oficiais preenchidos")
    col_dl1, col_dl2, col_dl3 = st.columns(3)
    with col_dl1:
        st.download_button(
            "📊 Baixar Excel",
            data=res["xlsx_bytes"],
            file_name=f"FORMULARIO_DSI_REMATE_{date.today():%Y-%m-%d}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    with col_dl2:
        st.download_button(
            "📄 Baixar Word",
            data=res["docx_bytes"],
            file_name=f"Demonstrativo_Calculo_{date.today():%Y-%m-%d}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )
    with col_dl3:
        st.download_button(
            "🧾 Baixar DARF (PDF)",
            data=res["darf_bytes"],
            file_name=f"DARF_{date.today():%Y-%m-%d}.pdf",
            mime="application/pdf",
            use_container_width=True,
        )

    if res.get("links_drive"):
        links = res["links_drive"]
        st.success(
            "☁️ Também enviado ao Google Drive:  "
            f"[Excel]({links['excel']}) · [Word]({links['word']}) · [DARF]({links['darf']})"
        )
    elif not drive_configurado():
        st.caption("☁️ Envio automático ao Google Drive: não configurado (opcional).")

with st.expander("Conferir com o exemplo do seu PASSO A PASSO DSI.txt"):
    st.write("VM = 2.367,22 · Alíquota I.I. = 18% · Alíquota I.P.I. = 35% · ICMS (PR) = 19,5%")
    st.write(
        "Esperado (cálculo manual): I.I=426,10 · I.P.I=977,66 · PIS=49,71 · "
        "COFINS=228,44 · ICMS≈980,84 · Total≈2.662,75"
    )
    st.caption(
        "Diferenças de 1 centavo entre o cálculo manual e o automático são normais "
        "(arredondamento em cada etapa)."
    )