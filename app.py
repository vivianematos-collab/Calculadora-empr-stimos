# -*- coding: utf-8 -*-
"""
app.py - Calculadora de Empréstimo (Streamlit)
==============================================
Interface web para a classe `CalculadoraEmprestimo`. Este arquivo precisa
estar na MESMA PASTA de `calculadora_emprestimo.py`.

Instalação (uma única vez):
    pip install -r requirements.txt

Para abrir a aplicação:
    streamlit run app.py
"""

from __future__ import annotations

import tempfile
from datetime import date, datetime
from pathlib import Path

import altair as alt
import pandas as pd
import streamlit as st

from calculadora_emprestimo import CalculadoraEmprestimo

st.set_page_config(page_title="Calculadora de Empréstimo", page_icon="📊", layout="wide")

# Cores das séries (ordem fixa: a cor acompanha a rubrica, não a posição)
COR_AMORTIZACAO = "#2a78d6"
COR_JUROS = "#eb6834"
COR_CORRECAO = "#1baf7a"
# Rótulos dos eixos com ponto como separador de milhar (padrão brasileiro)
EIXO_BR = "replace(format(datum.value, ',.0f'), /,/g, '.')"


# ---------------------------------------------------------------------- #
# Funções auxiliares
# ---------------------------------------------------------------------- #
def fmt_br(valor: float, casas: int = 2) -> str:
    """Número no padrão brasileiro: 1.234.567,89"""
    texto = f"{valor:,.{casas}f}"
    return texto.replace(",", "#").replace(".", ",").replace("#", ".")


def moeda(valor: float) -> str:
    return "R$ " + fmt_br(valor)


def exibir_tabela(dados: pd.DataFrame, altura: int = 460) -> None:
    """Mostra um DataFrame com datas em dd/mm/aaaa e valores no padrão brasileiro."""
    formatos = {}
    for coluna in dados.columns:
        if pd.api.types.is_datetime64_any_dtype(dados[coluna]):
            formatos[coluna] = lambda d: d.strftime("%d/%m/%Y")
        elif pd.api.types.is_float_dtype(dados[coluna]):
            formatos[coluna] = fmt_br
    pd.set_option("styler.render.max_elements", max(262_144, dados.size + 1))
    st.dataframe(dados.style.format(formatos), hide_index=True, height=altura,
                 width="stretch")


def ler_serie(arquivo) -> pd.Series:
    """Série de número-índice: 1ª coluna = data (dia/mês/ano), 2ª = número-índice."""
    if arquivo.name.lower().endswith((".xlsx", ".xlsm", ".xls")):
        dados = pd.read_excel(arquivo)
    else:
        dados = pd.read_csv(arquivo, sep=None, engine="python", decimal=",")
    datas = pd.to_datetime(dados.iloc[:, 0], dayfirst=True)
    valores = pd.to_numeric(dados.iloc[:, 1])
    serie = pd.Series(valores.to_numpy(), index=datas).dropna()
    if serie.empty:
        raise ValueError("O arquivo da série histórica não tem dados.")
    return serie


def excel_em_bytes(calculadora: CalculadoraEmprestimo) -> bytes:
    with tempfile.TemporaryDirectory() as pasta:
        caminho = calculadora.exportar_excel(Path(pasta) / "tabela_emprestimo.xlsx")
        return caminho.read_bytes()


# ---------------------------------------------------------------------- #
# Entrada de dados (barra lateral)
# ---------------------------------------------------------------------- #
st.title("Calculadora de Empréstimo")
st.caption(
    "Tabela de pagamentos, apropriação diária de juros e correção monetária "
    "e resumo mensal para contabilização."
)

with st.sidebar:
    st.header("Dados do contrato")

    valor_contrato = st.number_input(
        "Valor do contrato (R$)", min_value=0.01, value=10_000_000.00,
        step=100_000.00, format="%.2f",
    )
    data_contratacao = st.date_input(
        "Data da contratação", value=date.today(), format="DD/MM/YYYY"
    )

    st.subheader("Correção monetária")
    opcao_indice = st.selectbox("Índice", ["Sem correção", "IPCA", "CDI", "IGPM", "Outro"])
    indice = None
    taxa_indice = None
    correcao_modo = "pagar"
    if opcao_indice != "Sem correção":
        nome_indice = opcao_indice
        if opcao_indice == "Outro":
            nome_indice = st.text_input("Nome do índice", value="INDICE").strip() or "INDICE"
        fonte = st.radio(
            "Variação do índice",
            ["Taxa anual esperada", "Série histórica (arquivo)"],
        )
        if fonte == "Taxa anual esperada":
            indice = nome_indice
            taxa_indice = st.number_input(
                f"Taxa anual esperada do {nome_indice} (%)",
                min_value=-99.0, value=4.50, step=0.25, format="%.4f",
            ) / 100.0
        else:
            arquivo = st.file_uploader(
                "Arquivo Excel ou CSV", type=["xlsx", "xls", "csv"],
                help="1ª coluna: data (dia/mês/ano). 2ª coluna: número-índice.",
            )
            if st.checkbox("Projetar o índice após o fim da série", value=True):
                taxa_indice = st.number_input(
                    "Taxa anual de projeção (%)",
                    min_value=-99.0, value=4.50, step=0.25, format="%.4f",
                ) / 100.0
            if arquivo is None:
                st.info("Envie o arquivo da série histórica para calcular.")
                st.stop()
            try:
                indice = ler_serie(arquivo)
            except Exception as erro:
                st.error(f"Não foi possível ler o arquivo: {erro}")
                st.stop()
        correcao_modo = st.radio(
            "Cobrança da correção",
            ["pagar", "incorporar"],
            format_func=lambda v: "Paga junto com a parcela"
            if v == "pagar" else "Incorporada ao saldo devedor",
        )

    st.subheader("Juros e amortização")
    spread = st.number_input(
        "Spread ao ano (%)", min_value=-99.0, value=5.00, step=0.25, format="%.4f"
    ) / 100.0
    base_dias = st.radio(
        "Capitalização dos juros",
        ["365", "252"],
        format_func=lambda v: "365 - dias corridos" if v == "365"
        else "252 - dias úteis (ANBIMA)",
    )
    sistema = st.radio(
        "Sistema de amortização",
        ["SAC", "PRICE"],
        format_func=lambda v: "SAC - amortização constante" if v == "SAC"
        else "PRICE - parcelas constantes",
    )
    periodicidade = st.selectbox(
        "Periodicidade",
        ["mensal", "bimestral", "trimestral", "quadrimestral", "semestral", "anual"],
        format_func=str.capitalize,
    )
    prazo_total = int(st.number_input(
        "Prazo total (nº de parcelas, inclui carência)", min_value=1, value=60, step=1
    ))
    carencia = int(st.number_input(
        "Carência (períodos sem amortização)",
        min_value=0, max_value=max(prazo_total - 1, 0), value=0, step=1,
    ))
    juros_carencia = "capitalizar"
    if carencia > 0:
        juros_carencia = st.radio(
            "Juros na carência",
            ["capitalizar", "pagar"],
            format_func=lambda v: "Capitalizados no saldo devedor"
            if v == "capitalizar" else "Pagos a cada período",
        )

    st.subheader("Fee inicial")
    fee_tipo = st.radio(
        "Tipo de fee",
        ["percentual", "fixo"],
        format_func=lambda v: "Percentual do contrato" if v == "percentual"
        else "Valor fixo (R$)",
    )
    if fee_tipo == "percentual":
        fee_inicial = st.number_input(
            "Fee (%)", min_value=0.0, value=0.0, step=0.25, format="%.4f"
        ) / 100.0
    else:
        fee_inicial = st.number_input(
            "Fee (R$)", min_value=0.0, value=0.0, step=1_000.0, format="%.2f"
        )
    fee_modo = "deduzir"
    if fee_inicial > 0:
        fee_modo = st.radio(
            "Cobrança do fee",
            ["deduzir", "financiar"],
            format_func=lambda v: "Deduzido do desembolso" if v == "deduzir"
            else "Somado ao saldo devedor",
        )

# ---------------------------------------------------------------------- #
# Cálculo
# ---------------------------------------------------------------------- #
try:
    calculadora = CalculadoraEmprestimo(
        valor_contrato=valor_contrato,
        indice=indice,
        spread=spread,
        data_contratacao=datetime.combine(data_contratacao, datetime.min.time()),
        periodicidade=periodicidade,
        sistema_amortizacao=sistema,
        prazo_total=prazo_total,
        carencia=carencia,
        fee_inicial=fee_inicial,
        base_dias=base_dias,
        taxa_indice_projetada=taxa_indice,
        juros_carencia=juros_carencia,
        fee_tipo=fee_tipo,
        fee_modo=fee_modo,
        correcao_modo=correcao_modo,
    )
    if isinstance(indice, pd.Series):
        calculadora.definir_indice(nome=nome_indice)

    with st.spinner("Calculando..."):
        pagamentos = calculadora.gerar_tabela().reset_index()
        diaria = calculadora.gerar_tabela_diaria()
        mensal = calculadora.gerar_resumo_mensal()
        try:
            custo_efetivo = calculadora.custo_efetivo_anual()
        except ValueError:
            custo_efetivo = None
        arquivo_excel = excel_em_bytes(calculadora)
except (ValueError, ImportError) as erro:
    st.error(f"Não foi possível calcular: {erro}")
    st.stop()

# ---------------------------------------------------------------------- #
# Resultados
# ---------------------------------------------------------------------- #
c1, c2, c3 = st.columns(3)
c4, c5, _ = st.columns(3)
c1.metric("Desembolso líquido", moeda(calculadora.desembolso_liquido))
c2.metric("Total de juros", moeda(pagamentos["Juros"].sum()))
c3.metric("Total de correção", moeda(pagamentos["Correção Monetária"].sum()))
c4.metric("Total pago", moeda(pagamentos["Parcela"].sum()))
c5.metric(
    "Custo efetivo anual",
    "n/d" if custo_efetivo is None else fmt_br(custo_efetivo * 100) + "%",
)

st.download_button(
    "Baixar Excel (4 abas)",
    data=arquivo_excel,
    file_name=f"tabela_emprestimo_{datetime.now():%Y%m%d_%H%M}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary",
)

aba_pag, aba_mensal, aba_diaria, aba_graficos, aba_resumo = st.tabs(
    ["Tabela de Pagamentos", "Contabilização Mensal", "Cálculo Diário", "Gráficos", "Resumo"]
)

with aba_pag:
    st.caption("Uma linha por parcela.")
    exibir_tabela(pagamentos)

with aba_mensal:
    st.caption(
        "Uma linha por mês de competência: juros e correção apropriados no mês "
        "(soma do cálculo diário), pagamentos e saldo no fechamento."
    )
    exibir_tabela(mensal)

with aba_diaria:
    st.caption(
        "Apropriação pro rata die. Em cada linha: Saldo Final = Saldo Inicial "
        "+ Correção + Juros - Parcela."
    )
    exibir_tabela(diaria)

with aba_graficos:
    # Saldo devedor no fechamento de cada mês (uma série: o título a identifica)
    st.subheader("Saldo devedor no fechamento de cada mês")
    saldo_mes = mensal[["Competência", "Saldo Devedor Final"]].copy()
    saldo_mes["Saldo"] = saldo_mes["Saldo Devedor Final"].map(moeda)
    grafico_saldo = (
        alt.Chart(saldo_mes)
        .mark_line(color=COR_AMORTIZACAO, strokeWidth=2)
        .encode(
            x=alt.X("Competência:T", title=None, axis=alt.Axis(format="%m/%Y")),
            y=alt.Y("Saldo Devedor Final:Q", title="R$",
                    axis=alt.Axis(labelExpr=EIXO_BR)),
            tooltip=[
                alt.Tooltip("Competência:T", format="%d/%m/%Y"),
                alt.Tooltip("Saldo:N", title="Saldo devedor"),
            ],
        )
        .properties(height=300)
    )
    st.altair_chart(grafico_saldo, width="stretch")

    # Composição de cada parcela paga
    st.subheader("Composição das parcelas")
    composicao = (
        diaria[diaria["Parcela"] > 0][
            ["Data", "Nº Parcela", "Amortização", "Juros Pagos", "Correção Paga"]
        ]
        .rename(columns={"Juros Pagos": "Juros", "Correção Paga": "Correção monetária"})
        .melt(id_vars=["Data", "Nº Parcela"], var_name="Componente", value_name="Valor")
    )
    if composicao.empty:
        st.info("Não há parcelas pagas neste cenário.")
    else:
        ordem = ["Amortização", "Juros", "Correção monetária"]
        composicao["ordem"] = composicao["Componente"].map(ordem.index)
        composicao["Valor (R$)"] = composicao["Valor"].map(moeda)
        grafico_parcelas = (
            alt.Chart(composicao)
            .mark_bar()
            .encode(
                x=alt.X("Nº Parcela:O", title="Nº da parcela",
                        axis=alt.Axis(labelAngle=0, labelOverlap=True)),
                y=alt.Y("Valor:Q", title="R$", stack="zero",
                        axis=alt.Axis(labelExpr=EIXO_BR)),
                color=alt.Color(
                    "Componente:N",
                    scale=alt.Scale(
                        domain=ordem,
                        range=[COR_AMORTIZACAO, COR_JUROS, COR_CORRECAO],
                    ),
                    legend=alt.Legend(title=None, orient="top"),
                ),
                order=alt.Order("ordem:Q"),
                tooltip=[
                    alt.Tooltip("Nº Parcela:O"),
                    alt.Tooltip("Data:T", format="%d/%m/%Y", title="Vencimento"),
                    alt.Tooltip("Componente:N"),
                    alt.Tooltip("Valor (R$):N", title="Valor"),
                ],
            )
            .properties(height=320)
        )
        st.altair_chart(grafico_parcelas, width="stretch")
    st.caption("Os valores exatos estão nas abas de tabela e no Excel.")

with aba_resumo:
    linhas = []
    for descricao, valor, formato in calculadora.resumo():
        if valor is None:
            continue
        if isinstance(valor, datetime):
            texto = valor.strftime("%d/%m/%Y")
        elif formato == "0.00%" and not isinstance(valor, str):
            texto = fmt_br(valor * 100) + "%"
        elif formato == "#,##0.00":
            texto = moeda(valor)
        else:
            texto = str(valor)
        linhas.append({"Item": descricao, "Valor": texto})
    st.dataframe(pd.DataFrame(linhas), hide_index=True, width="stretch",
                 height=38 * (len(linhas) + 1))
