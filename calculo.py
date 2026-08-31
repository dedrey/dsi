"""
Motor de calculo dos tributos da DSI (Declaracao Simplificada de Importacao).

Formulas (conforme especificado):
- I.I     = VM x %II
- I.P.I   = (VM + I.I) x %IPI
- PIS     = VM x %PIS      (padrao 2,10%)
- COFINS  = VM x %COFINS   (padrao 9,65%)
- Base ICMS = (VM + I.I + I.P.I + PIS + COFINS) / (1 - %ICMS_UF)
- ICMS    = Base ICMS x %ICMS_UF
- Total   = I.I + I.P.I + PIS + COFINS + ICMS

Cada valor intermediario e arredondado para 2 casas decimais antes de entrar
no proximo calculo. E assim que o Demonstrativo de Calculo oficial funciona
(cada celula da tabela precisa "fechar" sozinha) e e o mesmo caminho que o
calculo manual em PASSO_A_PASSO_DSI.txt segue. Pequenas diferencas de
centavos em relacao a uma conta feita 100% na calculadora sao esperadas.
"""

from dataclasses import dataclass


PIS_ALIQUOTA_PADRAO = 2.10
COFINS_ALIQUOTA_PADRAO = 9.65

# Codigos de receita (DARF) usados no recolhimento de cada tributo federal,
# conforme PASSO_A_PASSO_DSI.txt
CODIGOS_RECEITA_FEDERAL = {
    "I.I.": "0086",
    "I.P.I.": "1038",
    "PIS": "5602",
    "COFINS": "5629",
}


@dataclass
class ResultadoTributos:
    valor_mercadoria: float
    aliquota_ii: float
    aliquota_ipi: float
    aliquota_pis: float
    aliquota_cofins: float
    aliquota_icms: float

    ii: float
    base_ipi: float
    ipi: float
    pis: float
    cofins: float
    base_icms: float
    icms: float
    total_tributos: float


def _r2(valor: float) -> float:
    """Arredondamento comercial para 2 casas decimais."""
    return round(valor + 1e-9, 2)


def calcular_tributos(
    valor_mercadoria: float,
    aliquota_ii: float,
    aliquota_ipi: float,
    aliquota_icms: float,
    aliquota_pis: float = PIS_ALIQUOTA_PADRAO,
    aliquota_cofins: float = COFINS_ALIQUOTA_PADRAO,
) -> ResultadoTributos:
    """
    Calcula II, IPI, PIS, COFINS e ICMS de uma DSI.

    Todas as aliquotas em formato percentual (ex.: 18 para 18%, nao 0.18).
    Levanta ValueError se os dados de entrada nao fizerem sentido.
    """
    if valor_mercadoria <= 0:
        raise ValueError("Valor da mercadoria deve ser maior que zero.")
    if not (0 <= aliquota_icms < 100):
        raise ValueError("Aliquota de ICMS invalida (deve ser entre 0 e 100).")

    vm = _r2(valor_mercadoria)

    ii = _r2(vm * aliquota_ii / 100)

    base_ipi = _r2(vm + ii)
    ipi = _r2(base_ipi * aliquota_ipi / 100)

    pis = _r2(vm * aliquota_pis / 100)
    cofins = _r2(vm * aliquota_cofins / 100)

    soma_para_icms = _r2(vm + ii + ipi + pis + cofins)
    base_icms = _r2(soma_para_icms / (1 - aliquota_icms / 100))
    icms = _r2(base_icms * aliquota_icms / 100)

    total = _r2(ii + ipi + pis + cofins + icms)

    return ResultadoTributos(
        valor_mercadoria=vm,
        aliquota_ii=aliquota_ii,
        aliquota_ipi=aliquota_ipi,
        aliquota_pis=aliquota_pis,
        aliquota_cofins=aliquota_cofins,
        aliquota_icms=aliquota_icms,
        ii=ii,
        base_ipi=base_ipi,
        ipi=ipi,
        pis=pis,
        cofins=cofins,
        base_icms=base_icms,
        icms=icms,
        total_tributos=total,
    )