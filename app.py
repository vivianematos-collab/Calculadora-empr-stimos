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

import io
import tempfile
from datetime import date, datetime
from pathlib import Path

import altair as alt
import pandas as pd
import streamlit as st

from dateutil.relativedelta import relativedelta

from calculadora_emprestimo import CalculadoraEmprestimo, ler_cronograma

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


MESES_POR_PERIODO = {
    "mensal": 1, "bimestral": 2, "trimestral": 3,
    "quadrimestral": 4, "semestral": 6, "anual": 12,
}


def cronograma_padrao(inicio, periodicidade, parcelas, carencia, total) -> pd.DataFrame:
    """Cronograma inicial: datas pela periodicidade e amortização em partes
    iguais após a carência (a última parcela absorve o arredondamento)."""
    meses = MESES_POR_PERIODO[periodicidade]
    datas = [pd.Timestamp(inicio) + relativedelta(months=meses * k)
             for k in range(1, parcelas + 1)]
    carencia = min(max(carencia, 0), parcelas - 1)
    pagantes = parcelas - carencia
    valores = [0.0] * carencia + [round(total / pagantes, 2)] * pagantes
    valores[-1] = round(total - sum(valores[:-1]), 2)
    return pd.DataFrame({"Data": pd.to_datetime(datas), "Amortização": valores})


def cronograma_em_excel(cronograma: pd.DataFrame, rotulo: str) -> bytes:
    """Planilha-modelo do cronograma, para editar no Excel e enviar de volta."""
    from openpyxl import Workbook

    wb = Workbook()
    ws = wb.active
    ws.title = "Cronograma"
    ws.append(["Data de vencimento", rotulo])
    for data, valor in cronograma.itertuples(index=False, name=None):
        ws.append([pd.Timestamp(data).to_pydatetime(), float(valor)])
    for linha in ws.iter_rows(min_row=2):
        linha[0].number_format = "DD/MM/YYYY"
        linha[1].number_format = "#,##0.00"
    ws.column_dimensions["A"].width = 20
    ws.column_dimensions["B"].width = 22
    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()


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
    modo_vencimentos = st.radio(
        "Vencimentos das parcelas",
        ["periodicidade", "cronograma"],
        format_func=lambda v: "Periodicidade fixa" if v == "periodicidade"
        else "Cronograma informado",
        help="No cronograma informado, você define a data de cada parcela e, "
             "se quiser, quanto do principal é amortizado em cada uma.",
    )
    usa_cronograma = modo_vencimentos == "cronograma"
    rotulos_sistema = {
        "SAC": "SAC - amortização constante",
        "PRICE": "PRICE - parcelas constantes",
        "valor": "Informada no cronograma (R$)",
        "percentual": "Informada no cronograma (% do principal)",
    }
    sistema = st.radio(
        "Sistema de amortização",
        ["SAC", "PRICE", "valor", "percentual"] if usa_cronograma else ["SAC", "PRICE"],
        format_func=rotulos_sistema.get,
    )
    amortizacao_informada = sistema in ("valor", "percentual")
    if usa_cronograma:
        st.caption(
            "Periodicidade, prazo e carência abaixo servem para montar o "
            "cronograma inicial (botão **Recriar pela periodicidade**)."
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
        "Carência (períodos sem amortização)", min_value=0, value=0, step=1,
        help="Deve ser menor que o número de parcelas.",
    ))
    juros_carencia = "capitalizar"
    if carencia > 0 or amortizacao_informada:
        juros_carencia = st.radio(
            "Juros na carência",
            ["capitalizar", "pagar"],
            format_func=lambda v: "Capitalizados no saldo devedor"
            if v == "capitalizar" else "Pagos a cada período",
            help="Vale para as parcelas iniciais sem amortização.",
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
# Cronograma de pagamento informado (área principal)
# ---------------------------------------------------------------------- #
cronograma = None
if usa_cronograma:
    # Total a amortizar: saldo devedor inicial (R$) ou 100 (%)
    valor_fee = valor_contrato * fee_inicial if fee_tipo == "percentual" else fee_inicial
    saldo_inicial = valor_contrato + (valor_fee if fee_modo == "financiar" else 0.0)
    total_esperado = 100.0 if sistema == "percentual" else saldo_inicial
    rotulo_amortizacao = "Amortização (%)" if sistema == "percentual" else "Amortização (R$)"

    def novo_cronograma() -> pd.DataFrame:
        return cronograma_padrao(
            data_contratacao, periodicidade, prazo_total, carencia, total_esperado
        )

    estado = st.session_state
    if "cronograma_base" not in estado:
        estado["cronograma_base"] = novo_cronograma()
        estado["cronograma_versao"] = 0
        estado["cronograma_unidade"] = "percentual" if sistema == "percentual" else "valor"

    # Ao trocar entre R$ e %, converte os valores já digitados para a nova unidade.
    unidade_atual = "percentual" if sistema == "percentual" else "valor"
    if amortizacao_informada and estado["cronograma_unidade"] != unidade_atual:
        convertido = estado.get("cronograma_editado", estado["cronograma_base"]).copy()
        soma_anterior = float(convertido["Amortização"].sum())
        if soma_anterior > 0:
            convertido["Amortização"] = (
                convertido["Amortização"] / soma_anterior * total_esperado
            ).round(2)
            ultima = convertido.index[convertido["Amortização"] > 0][-1]
            convertido.loc[ultima, "Amortização"] += round(
                total_esperado - float(convertido["Amortização"].sum()), 2
            )
        estado["cronograma_base"] = convertido
        estado["cronograma_versao"] += 1
        estado["cronograma_unidade"] = unidade_atual

    with st.expander("Cronograma de pagamento", expanded=True):
        st.write(
            "Edite as datas"
            + (" e a amortização de cada parcela" if amortizacao_informada else "")
            + " diretamente na tabela (use a última linha para incluir parcelas e "
            "a caixa de seleção à esquerda para excluir), ou envie um arquivo."
        )
        col_tabela, col_acoes = st.columns([3, 2])

        with col_acoes:
            if st.button("Recriar pela periodicidade"):
                estado["cronograma_base"] = novo_cronograma()
                estado["cronograma_versao"] += 1
            arquivo_cronograma = st.file_uploader(
                "Enviar cronograma (Excel ou CSV)", type=["xlsx", "xls", "csv"],
                help="1ª coluna: data de vencimento (dia/mês/ano). "
                     "2ª coluna (opcional): amortização da parcela.",
            )
            if arquivo_cronograma is not None:
                identificador = (arquivo_cronograma.name, arquivo_cronograma.size)
                if estado.get("cronograma_arquivo") != identificador:
                    try:
                        lido = ler_cronograma(arquivo_cronograma)
                        if "Amortização" not in lido.columns:
                            lido["Amortização"] = 0.0
                        estado["cronograma_base"] = lido
                        estado["cronograma_versao"] += 1
                        estado["cronograma_arquivo"] = identificador
                    except Exception as erro:
                        st.error(f"Não foi possível ler o arquivo: {erro}")
            else:
                estado.pop("cronograma_arquivo", None)

        with col_tabela:
            editado = st.data_editor(
                estado["cronograma_base"],
                key=f"editor_cronograma_{estado['cronograma_versao']}",
                num_rows="dynamic",
                hide_index=True,
                height=320,
                width="stretch",
                column_order=["Data", "Amortização"] if amortizacao_informada else ["Data"],
                column_config={
                    "Data": st.column_config.DateColumn(
                        "Data de vencimento", format="DD/MM/YYYY", required=True
                    ),
                    "Amortização": st.column_config.NumberColumn(
                        rotulo_amortizacao, min_value=0.0, format="%.2f"
                    ),
                },
            )

        editado = editado.dropna(subset=["Data"]).copy()
        editado["Data"] = pd.to_datetime(editado["Data"])
        editado["Amortização"] = pd.to_numeric(editado["Amortização"]).fillna(0.0)
        editado = editado.sort_values("Data").reset_index(drop=True)
        estado["cronograma_editado"] = editado
        soma_divergente = False

        with col_acoes:
            st.download_button(
                "Baixar cronograma atual (modelo Excel)",
                data=cronograma_em_excel(editado[["Data", "Amortização"]], rotulo_amortizacao),
                file_name="cronograma_pagamento.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            st.write(f"**Parcelas:** {len(editado)}")
            if amortizacao_informada:
                soma = float(editado["Amortização"].sum())
                unidade = "%" if sistema == "percentual" else ""
                prefixo = "" if sistema == "percentual" else "R$ "
                st.write(
                    f"**Soma das amortizações:** {prefixo}{fmt_br(soma)}{unidade}  \n"
                    f"**Deve somar:** {prefixo}{fmt_br(total_esperado)}{unidade}"
                )
                soma_divergente = abs(soma - total_esperado) > 0.01
                if soma_divergente:
                    st.warning(
                        f"Diferença de {prefixo}{fmt_br(soma - total_esperado)}{unidade}. "
                        "Ajuste os valores para calcular."
                    )

    if editado.empty:
        st.info("Inclua ao menos uma parcela no cronograma para calcular.")
        st.stop()
    if soma_divergente:
        st.stop()
    cronograma = editado[["Data", "Amortização"]] if amortizacao_informada else editado[["Data"]]

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
        sistema_amortizacao="SAC" if amortizacao_informada else sistema,
        prazo_total=prazo_total,
        carencia=carencia,
        cronograma=cronograma,
        cronograma_tipo="percentual" if sistema == "percentual" else "valor",
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
