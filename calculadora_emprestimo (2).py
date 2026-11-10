# -*- coding: utf-8 -*-
"""
calculadora_emprestimo.py
=========================
Tabela de pagamentos de empréstimo indexado (índice + spread), com SAC ou PRICE,
carência, fee inicial e capitalização em dias corridos (365) ou dias úteis (252).

Como usar (não é preciso saber Python)
--------------------------------------
    python calculadora_emprestimo.py
O programa faz as perguntas na tela, uma a uma, mostra um resumo para
conferência e grava a tabela de pagamentos em um arquivo Excel (.xlsx).
Para rodar o exemplo pronto, sem perguntas:  python calculadora_emprestimo.py --exemplo

No Google Colab
---------------
Em uma célula, rode antes:   !pip install bizdays
Depois cole este código em outra célula e execute. As perguntas aparecem
abaixo da célula, cada uma com uma caixinha de texto para a resposta; enquanto
o programa aguarda, a célula fica marcada como "executando". Ao final, o
Excel é baixado pelo navegador.

Dependências
------------
    pip install pandas numpy python-dateutil openpyxl
    pip install bizdays          # necessário apenas para base '252'
No modo interativo, se `bizdays` ou `openpyxl` não estiverem instalados, o
programa se oferece para instalá-los automaticamente.

Dias úteis (base 252): usa a biblioteca `bizdays` (https://pypi.org/project/bizdays/),
que já traz embutido o calendário de feriados ANBIMA - o padrão do mercado
brasileiro para DU/252. Nenhum feriado é codificado à mão neste script.

Cálculo diário e contabilização mensal
--------------------------------------
Além da tabela por parcela (`gerar_tabela`), a classe calcula a apropriação
diária, pro rata die, de juros e correção monetária (`gerar_tabela_diaria`) e
o resumo por mês de competência (`gerar_resumo_mensal`), para os lançamentos
contábeis mensais. O Excel exportado traz as três visões.

Convenções adotadas (todas parametrizáveis onde indicado)
---------------------------------------------------------
* Taxas em decimal: spread=0.05 significa 5% a.a.
* `prazo_total` é o número total de parcelas e INCLUI a carência.
  Parcelas de amortização = prazo_total - carencia.
* Carência: juros capitalizados no saldo devedor (padrão) ou pagos
  (`juros_carencia='pagar'`).
* Fee inicial: deduzido do desembolso (padrão) ou somado ao saldo devedor
  (`fee_modo='financiar'`); percentual sobre o contrato (padrão) ou valor fixo.
* Em cada período: (1) corrige-se o saldo devedor pelo índice;
  (2) calculam-se os juros do spread sobre o saldo corrigido;
  (3) apura-se a amortização.
* Juros compostos por dia: fator = (1 + spread) ** (dias / base).
* Correção monetária (`correcao_modo`):
  - 'pagar' (padrão): a correção do período é paga junto com a parcela, isto é,
    Parcela = Amortização + Juros + Correção Monetária. O saldo devedor não
    cresce com a correção e a amortização incide sobre o principal
    (no SAC, Principal / nº de parcelas de amortização).
  - 'incorporar': a correção é somada ao saldo devedor e amortizada ao longo
    das parcelas restantes (Parcela = Amortização + Juros, com SAC e PRICE
    recalculados sobre o saldo corrigido).
  Na carência com juros capitalizados nada é pago, então a correção também é
  capitalizada no saldo, em qualquer dos dois modos.
"""

from __future__ import annotations

import importlib
import subprocess
import sys
from datetime import date, datetime
from pathlib import Path
from typing import Callable, Optional, Union

import numpy as np
import pandas as pd
from dateutil.relativedelta import relativedelta

TipoData = Union[datetime, date, str, pd.Timestamp]
TipoSerie = Union[dict, pd.Series]
TipoFuncaoIndice = Callable[[pd.Timestamp, pd.Timestamp], float]


class CalculadoraEmprestimo:
    """Estrutura o fluxo de pagamentos de um empréstimo indexado.

    Parâmetros
    ----------
    valor_contrato : float
        Principal financiado.
    indice : None | str | dict | pandas.Series | callable
        Indexador de correção monetária.
        - None ........ sem correção monetária.
        - str ......... apenas o nome (ex.: 'IPCA'); informe também
                        `taxa_indice_projetada` ou chame `definir_indice()`.
        - dict/Series . série de NÚMERO-ÍNDICE {data: nível}. O fator do período
                        é nível(data_fim) / nível(data_início). Para datas além
                        do último ponto, a série é extrapolada com
                        `taxa_indice_projetada` (se informada).
        - callable .... função f(data_inicio, data_fim) -> fator de correção
                        (ex.: 1.0045 para 0,45% no período).
    spread : float
        Juros ao ano acima do índice, em decimal (0.05 = 5% a.a.).
    data_contratacao : datetime | date | str
        Data de início do contrato (data do desembolso).
    periodicidade : str
        'mensal', 'bimestral', 'trimestral', 'quadrimestral', 'semestral' ou 'anual'.
    sistema_amortizacao : str
        'SAC' ou 'PRICE'.
    prazo_total : int
        Número total de parcelas (inclui a carência).
    carencia : int
        Períodos sem amortização de principal.
    fee_inicial : float
        Fee cobrado na cabeça (percentual em decimal ou valor fixo, ver `fee_tipo`).
    base_dias : str
        '365' (dias corridos) ou '252' (dias úteis, calendário ANBIMA).
    taxa_indice_projetada : float, opcional
        Taxa anual esperada do índice (decimal), usada como projeção.
    juros_carencia : str
        'capitalizar' (padrão) ou 'pagar'.
    fee_tipo : str
        'percentual' (padrão) ou 'fixo'.
    fee_modo : str
        'deduzir' (padrão - reduz o desembolso líquido) ou 'financiar'
        (soma o fee ao saldo devedor).
    correcao_modo : str
        'pagar' (padrão - correção paga junto com a parcela) ou 'incorporar'
        (correção somada ao saldo devedor).
    calendario : objeto, opcional
        Calendário com método `.bizdays(data_ini, data_fim)`. Se omitido e a
        base for '252', carrega-se o calendário ANBIMA da biblioteca `bizdays`.
    """

    COLUNAS = [
        "Data",
        "Saldo Devedor Inicial",
        "Juros",
        "Amortização",
        "Parcela",
        "Correção Monetária",
        "Saldo Devedor Final",
    ]

    _MESES_POR_PERIODO = {
        "mensal": 1,
        "bimestral": 2,
        "trimestral": 3,
        "quadrimestral": 4,
        "semestral": 6,
        "anual": 12,
    }

    # ------------------------------------------------------------------ #
    # Inicialização e validação
    # ------------------------------------------------------------------ #
    def __init__(
        self,
        valor_contrato: float,
        indice: Union[None, str, TipoSerie, TipoFuncaoIndice],
        spread: float,
        data_contratacao: TipoData,
        periodicidade: str,
        sistema_amortizacao: str,
        prazo_total: int,
        carencia: int = 0,
        fee_inicial: float = 0.0,
        base_dias: Union[str, int] = "365",
        *,
        taxa_indice_projetada: Optional[float] = None,
        juros_carencia: str = "capitalizar",
        fee_tipo: str = "percentual",
        fee_modo: str = "deduzir",
        correcao_modo: str = "pagar",
        calendario=None,
    ) -> None:
        self.valor_contrato = float(valor_contrato)
        self.spread = float(spread)
        self.data_contratacao = pd.Timestamp(data_contratacao).normalize()
        self.periodicidade = str(periodicidade).strip().lower()
        self.sistema_amortizacao = str(sistema_amortizacao).strip().upper()
        self.prazo_total = int(prazo_total)
        self.carencia = int(carencia)
        self.fee_inicial = float(fee_inicial)
        self.base_dias = str(base_dias).strip()
        self.juros_carencia = str(juros_carencia).strip().lower()
        self.fee_tipo = str(fee_tipo).strip().lower()
        self.fee_modo = str(fee_modo).strip().lower()
        self.correcao_modo = str(correcao_modo).strip().lower()
        self._calendario = calendario

        self._validar_parametros()

        # Estado do indexador
        self.nome_indice: Optional[str] = None
        self._taxa_indice: Optional[float] = None
        self._serie_indice: Optional[pd.Series] = None
        self._funcao_indice: Optional[TipoFuncaoIndice] = None
        self._tem_indice = indice is not None or taxa_indice_projetada is not None

        if isinstance(indice, str):
            self.definir_indice(nome=indice, taxa_anual=taxa_indice_projetada)
        elif isinstance(indice, (dict, pd.Series)):
            self.definir_indice(serie=indice, taxa_anual=taxa_indice_projetada)
        elif callable(indice):
            self.definir_indice(funcao=indice)
        elif indice is None:
            if taxa_indice_projetada is not None:
                self.definir_indice(taxa_anual=taxa_indice_projetada)
        else:
            raise TypeError(
                "`indice` deve ser None, str, dict, pandas.Series ou uma função."
            )

    def _validar_parametros(self) -> None:
        if self.valor_contrato <= 0:
            raise ValueError("`valor_contrato` deve ser positivo.")
        if self.spread <= -1:
            raise ValueError("`spread` deve ser maior que -100%.")
        if self.periodicidade not in self._MESES_POR_PERIODO:
            raise ValueError(
                f"`periodicidade` inválida: {self.periodicidade!r}. "
                f"Use uma de {sorted(self._MESES_POR_PERIODO)}."
            )
        if self.sistema_amortizacao not in ("SAC", "PRICE"):
            raise ValueError("`sistema_amortizacao` deve ser 'SAC' ou 'PRICE'.")
        if self.prazo_total <= 0:
            raise ValueError("`prazo_total` deve ser um inteiro positivo.")
        if not 0 <= self.carencia < self.prazo_total:
            raise ValueError(
                "`carencia` deve ser >= 0 e menor que `prazo_total` "
                "(é preciso ao menos uma parcela de amortização)."
            )
        if self.base_dias not in ("365", "252"):
            raise ValueError("`base_dias` deve ser '365' ou '252'.")
        if self.juros_carencia not in ("capitalizar", "pagar"):
            raise ValueError("`juros_carencia` deve ser 'capitalizar' ou 'pagar'.")
        if self.fee_tipo not in ("percentual", "fixo"):
            raise ValueError("`fee_tipo` deve ser 'percentual' ou 'fixo'.")
        if self.fee_modo not in ("deduzir", "financiar"):
            raise ValueError("`fee_modo` deve ser 'deduzir' ou 'financiar'.")
        if self.correcao_modo not in ("pagar", "incorporar"):
            raise ValueError("`correcao_modo` deve ser 'pagar' ou 'incorporar'.")
        if self.fee_inicial < 0:
            raise ValueError("`fee_inicial` não pode ser negativo.")
        if self.fee_modo == "deduzir" and self.valor_fee >= self.valor_contrato:
            raise ValueError("O fee inicial não pode consumir todo o desembolso.")

    # ------------------------------------------------------------------ #
    # Indexador
    # ------------------------------------------------------------------ #
    def definir_indice(
        self,
        nome: Optional[str] = None,
        taxa_anual: Optional[float] = None,
        serie: Optional[TipoSerie] = None,
        funcao: Optional[TipoFuncaoIndice] = None,
    ) -> "CalculadoraEmprestimo":
        """Injeta (ou substitui) a fonte de dados do indexador.

        nome ........ rótulo do índice (ex.: 'IPCA', 'CDI', 'IGPM').
        taxa_anual .. taxa anual esperada, em decimal. Sozinha, funciona como
                      projeção fixa; junto de `serie`, extrapola a série após
                      o último ponto disponível.
        serie ....... dict ou pandas.Series {data: número-índice}.
        funcao ...... f(data_inicio, data_fim) -> fator de correção do período.
                      Tem precedência sobre `serie` e `taxa_anual`.
        """
        if nome is not None:
            self.nome_indice = str(nome).strip().upper()
        if taxa_anual is not None:
            if taxa_anual <= -1:
                raise ValueError("`taxa_anual` do índice deve ser maior que -100%.")
            self._taxa_indice = float(taxa_anual)
        if serie is not None:
            s = pd.Series(serie, dtype="float64")
            s.index = pd.to_datetime(s.index)
            s = s.sort_index()
            if s.empty or (s <= 0).any() or s.isna().any():
                raise ValueError(
                    "A série do índice deve conter números-índice positivos."
                )
            self._serie_indice = s
        if funcao is not None:
            if not callable(funcao):
                raise TypeError("`funcao` deve ser chamável.")
            self._funcao_indice = funcao
        self._tem_indice = True
        return self

    def _nivel_indice(self, data: pd.Timestamp) -> float:
        """Número-índice na data (último valor disponível até a data);
        após o fim da série, extrapola com a taxa projetada."""
        s = self._serie_indice
        if data < s.index[0]:
            raise ValueError(
                f"A série do índice começa em {s.index[0].date()}, "
                f"posterior à data {data.date()} exigida pelo contrato."
            )
        ultima = s.index[-1]
        if data <= ultima:
            return float(s.asof(data))
        if self._taxa_indice is None:
            raise ValueError(
                f"A série do índice termina em {ultima.date()} e o fluxo exige "
                f"{data.date()}. Informe `taxa_anual` em `definir_indice()` "
                "para projetar o restante."
            )
        return float(s.iloc[-1]) * (1.0 + self._taxa_indice) ** self._fracao_ano(
            ultima, data
        )

    def _fator_correcao(self, ini: pd.Timestamp, fim: pd.Timestamp) -> float:
        """Fator de correção monetária entre duas datas."""
        if not self._tem_indice:
            return 1.0
        if self._funcao_indice is not None:
            fator = float(self._funcao_indice(ini, fim))
            if not np.isfinite(fator) or fator <= 0:
                raise ValueError(
                    "A função do índice deve retornar um fator positivo "
                    "(ex.: 1.0045 para 0,45% no período)."
                )
            return fator
        if self._serie_indice is not None:
            return self._nivel_indice(fim) / self._nivel_indice(ini)
        if self._taxa_indice is not None:
            return (1.0 + self._taxa_indice) ** self._fracao_ano(ini, fim)
        raise ValueError(
            f"Índice {self.nome_indice!r} informado sem dados. Passe "
            "`taxa_indice_projetada` no construtor ou chame `definir_indice()` "
            "com `taxa_anual`, `serie` ou `funcao`."
        )

    # ------------------------------------------------------------------ #
    # Datas e contagem de dias
    # ------------------------------------------------------------------ #
    def _obter_calendario(self):
        """Carrega (uma única vez) o calendário ANBIMA da biblioteca `bizdays`."""
        if self._calendario is None:
            try:
                from bizdays import Calendar
            except ImportError as exc:  # pragma: no cover
                raise ImportError(
                    "A base '252' exige a biblioteca `bizdays` "
                    "(pip install bizdays) ou um objeto `calendario` com o "
                    "método .bizdays(data_ini, data_fim)."
                ) from exc
            # Argumento nomeado: funciona nas versões 1.x e 2.x da bizdays.
            self._calendario = Calendar.load(name="ANBIMA")
        return self._calendario

    def _dias(self, ini: pd.Timestamp, fim: pd.Timestamp) -> int:
        """Dias entre duas datas, conforme a base do contrato."""
        if self.base_dias == "365":
            return int((fim - ini).days)
        cal = self._obter_calendario()
        return int(cal.bizdays(ini.date(), fim.date()))

    def _fracao_ano(self, ini: pd.Timestamp, fim: pd.Timestamp) -> float:
        return self._dias(ini, fim) / float(self.base_dias)

    def datas_pagamento(self) -> list:
        """Datas de vencimento, sempre contadas a partir da data de contratação
        (evita o deslocamento de fim de mês ao somar períodos sucessivos)."""
        meses = self._MESES_POR_PERIODO[self.periodicidade]
        return [
            self.data_contratacao + relativedelta(months=meses * k)
            for k in range(1, self.prazo_total + 1)
        ]

    # ------------------------------------------------------------------ #
    # Fee e desembolso
    # ------------------------------------------------------------------ #
    @property
    def valor_fee(self) -> float:
        if self.fee_tipo == "percentual":
            return self.valor_contrato * self.fee_inicial
        return self.fee_inicial

    @property
    def saldo_devedor_inicial(self) -> float:
        """Saldo devedor na data zero (inclui o fee se ele for financiado)."""
        if self.fee_modo == "financiar":
            return self.valor_contrato + self.valor_fee
        return self.valor_contrato

    @property
    def desembolso_liquido(self) -> float:
        """Caixa efetivamente recebido pelo tomador na contratação."""
        if self.fee_modo == "deduzir":
            return self.valor_contrato - self.valor_fee
        return self.valor_contrato

    # ------------------------------------------------------------------ #
    # Cálculo principal
    # ------------------------------------------------------------------ #
    def gerar_tabela(self, casas_decimais: Optional[int] = 2) -> pd.DataFrame:
        """Gera a tabela de pagamentos.

        casas_decimais : arredondamento aplicado apenas na saída (o cálculo
        interno usa precisão plena). Use None para não arredondar.
        """
        datas = self.datas_pagamento()
        n = self.prazo_total

        # Fator de juros (spread) de cada período, pela contagem real de dias.
        inicios = [self.data_contratacao] + datas[:-1]
        fatores_juros = np.array(
            [
                (1.0 + self.spread) ** self._fracao_ano(ini, fim)
                for ini, fim in zip(inicios, datas)
            ]
        )

        saldo = self.saldo_devedor_inicial
        linhas = []

        for k in range(1, n + 1):
            ini, fim = inicios[k - 1], datas[k - 1]
            saldo_inicial = saldo

            # (1) Correção monetária sobre o saldo devedor
            correcao = saldo_inicial * (self._fator_correcao(ini, fim) - 1.0)
            saldo_corrigido = saldo_inicial + correcao

            # (2) Juros do spread sobre o saldo corrigido
            juros = saldo_corrigido * (fatores_juros[k - 1] - 1.0)

            # (3) Amortização e parcela
            paga_correcao = self.correcao_modo == "pagar"
            if k <= self.carencia:
                amortizacao = 0.0
                if self.juros_carencia == "capitalizar":
                    # Nada é pago: juros e correção vão para o saldo devedor.
                    parcela = 0.0
                    saldo_final = saldo_corrigido + juros
                elif paga_correcao:
                    parcela = juros + correcao
                    saldo_final = saldo_inicial
                else:
                    parcela = juros
                    saldo_final = saldo_corrigido
            else:
                # Base de amortização: principal (correção paga na parcela)
                # ou saldo corrigido (correção incorporada ao saldo).
                base = saldo_inicial if paga_correcao else saldo_corrigido
                restantes = n - k + 1  # parcelas de amortização, incluindo esta
                if restantes == 1:
                    amortizacao = base  # liquida o saldo
                elif self.sistema_amortizacao == "SAC":
                    amortizacao = base / restantes
                else:  # PRICE
                    # Prestação constante que zera a base, descontando cada
                    # parcela futura pelos fatores reais dos períodos restantes.
                    # Com dias iguais em todos os períodos, reduz-se à fórmula
                    # clássica PMT = PV * i / (1 - (1 + i) ** -n).
                    descontos = 1.0 / np.cumprod(fatores_juros[k - 1 :])
                    pmt = base / descontos.sum()
                    amortizacao = pmt - base * (fatores_juros[k - 1] - 1.0)
                parcela = amortizacao + juros + (correcao if paga_correcao else 0.0)
                saldo_final = base - amortizacao

            if abs(saldo_final) < 1e-6:
                saldo_final = 0.0

            linhas.append(
                {
                    "Data": fim,
                    "Saldo Devedor Inicial": saldo_inicial,
                    "Juros": juros,
                    "Amortização": amortizacao,
                    "Parcela": parcela,
                    "Correção Monetária": correcao,
                    "Saldo Devedor Final": saldo_final,
                }
            )
            saldo = saldo_final

        tabela = pd.DataFrame(linhas, columns=self.COLUNAS)
        tabela.index = pd.RangeIndex(1, n + 1, name="Nº")
        if casas_decimais is not None:
            numericas = self.COLUNAS[1:]
            tabela[numericas] = tabela[numericas].round(casas_decimais)
        return tabela

    # ------------------------------------------------------------------ #
    # Cálculo diário (pro rata die) e contabilização mensal
    # ------------------------------------------------------------------ #
    COLUNAS_DIARIO = [
        "Data",
        "Nº Parcela",
        "Saldo Devedor Inicial",
        "Correção Monetária",
        "Juros",
        "Amortização",
        "Juros Pagos",
        "Correção Paga",
        "Parcela",
        "Saldo Devedor Final",
        "Correção Acumulada no Período",
        "Juros Acumulados no Período",
    ]

    def _dias_acumulados(self, ini: pd.Timestamp, datas: pd.DatetimeIndex) -> np.ndarray:
        """Dias (corridos ou úteis, conforme a base) entre `ini` e cada data."""
        if self.base_dias == "365":
            return np.asarray((datas - ini).days, dtype=float)
        cal = self._obter_calendario()
        return np.array([float(cal.bizdays(ini.date(), d.date())) for d in datas])

    def gerar_tabela_diaria(self, casas_decimais: Optional[int] = 2) -> pd.DataFrame:
        """Apropriação diária (pro rata die) de correção monetária e juros.

        Uma linha por dia corrido, do dia seguinte à contratação até o último
        vencimento. Em cada período entre vencimentos, com saldo inicial S0,
        fator de correção do período Fc, D dias no período e d dias decorridos
        (corridos na base 365, úteis na base 252):

            correção acumulada = S0 * (Fc ** (d / D) - 1)
            juros acumulados   = S0 * Fc ** (d / D) * ((1 + spread) ** (d / base) - 1)

        O valor do dia é a diferença entre o acumulado do dia e o do dia
        anterior; assim, a soma dos dias de um período é exatamente o valor
        do período na tabela de pagamentos. Na base 252 só há apropriação
        quando decorre um dia útil (fins de semana e feriados ficam zerados).

        Em todas as linhas:
            Saldo Devedor Final = Saldo Devedor Inicial + Correção Monetária
                                  + Juros - Parcela
        onde o saldo devedor é a dívida total (principal + correção e juros
        apropriados e ainda não pagos). Os pagamentos aparecem na data do
        vencimento. As colunas "... Acumulada(os) no Período" mostram quanto
        do saldo final corresponde a correção e a juros apropriados desde o
        último vencimento (zeram no vencimento, quando são pagos ou
        incorporados ao principal).
        """
        bruta = self.gerar_tabela(casas_decimais=None)
        final = self.gerar_tabela(casas_decimais=casas_decimais)
        datas_venc = list(bruta["Data"])
        inicios = [self.data_contratacao] + datas_venc[:-1]
        base = float(self.base_dias)

        def arredondar(vetor: np.ndarray) -> np.ndarray:
            return vetor if casas_decimais is None else np.round(vetor, casas_decimais)

        blocos = []
        for k, (ini, fim) in enumerate(zip(inicios, datas_venc), start=1):
            dias = pd.date_range(ini + pd.Timedelta(days=1), fim, freq="D")
            d = self._dias_acumulados(ini, dias)
            total_dias = d[-1]
            s0 = float(bruta.at[k, "Saldo Devedor Inicial"])
            fator_corr = self._fator_correcao(ini, fim)
            fracao = d / total_dias if total_dias > 0 else np.ones_like(d)

            corr_acum = s0 * (fator_corr ** fracao - 1.0)
            juros_acum = s0 * fator_corr ** fracao * ((1.0 + self.spread) ** (d / base) - 1.0)
            corr_acum, juros_acum = arredondar(corr_acum), arredondar(juros_acum)
            # O último dia fecha exatamente com a tabela de pagamentos.
            corr_acum[-1] = final.at[k, "Correção Monetária"]
            juros_acum[-1] = final.at[k, "Juros"]

            corr_dia = np.diff(corr_acum, prepend=0.0)
            juros_dia = np.diff(juros_acum, prepend=0.0)

            s0_exibido = float(final.at[k, "Saldo Devedor Inicial"])
            saldo_fim = s0_exibido + corr_acum + juros_acum
            saldo_ini = np.concatenate(([s0_exibido], saldo_fim[:-1]))

            n_dias = len(dias)
            amortizacao = np.zeros(n_dias)
            juros_pagos = np.zeros(n_dias)
            corr_paga = np.zeros(n_dias)
            parcela = np.zeros(n_dias)

            # Eventos do vencimento (último dia do período)
            parcela_k = float(final.at[k, "Parcela"])
            amort_k = float(final.at[k, "Amortização"])
            if parcela_k > 0:
                juros_k = float(final.at[k, "Juros"])
                amortizacao[-1] = amort_k
                juros_pagos[-1] = juros_k
                corr_paga[-1] = parcela_k - amort_k - juros_k
                parcela[-1] = parcela_k
            saldo_fim[-1] = float(final.at[k, "Saldo Devedor Final"])
            corr_acum_exib, juros_acum_exib = corr_acum.copy(), juros_acum.copy()
            corr_acum_exib[-1] = 0.0   # pagos ou incorporados ao principal
            juros_acum_exib[-1] = 0.0

            blocos.append(
                pd.DataFrame(
                    {
                        "Data": dias,
                        "Nº Parcela": k,
                        "Saldo Devedor Inicial": saldo_ini,
                        "Correção Monetária": corr_dia,
                        "Juros": juros_dia,
                        "Amortização": amortizacao,
                        "Juros Pagos": juros_pagos,
                        "Correção Paga": corr_paga,
                        "Parcela": parcela,
                        "Saldo Devedor Final": saldo_fim,
                        "Correção Acumulada no Período": corr_acum_exib,
                        "Juros Acumulados no Período": juros_acum_exib,
                    }
                )
            )

        diaria = pd.concat(blocos, ignore_index=True)[self.COLUNAS_DIARIO]
        if casas_decimais is not None:
            numericas = self.COLUNAS_DIARIO[2:]
            diaria[numericas] = diaria[numericas].round(casas_decimais)
            diaria[numericas] = diaria[numericas] + 0.0  # elimina -0,00
        return diaria

    COLUNAS_MENSAL = [
        "Competência",
        "Saldo Devedor Inicial",
        "Correção Monetária",
        "Juros",
        "Amortização",
        "Juros Pagos",
        "Correção Paga",
        "Parcela",
        "Saldo Devedor Final",
        "Correção Acumulada no Período",
        "Juros Acumulados no Período",
    ]

    def gerar_resumo_mensal(self, casas_decimais: Optional[int] = 2) -> pd.DataFrame:
        """Resumo por mês de competência, para a contabilização mensal.

        Para cada mês: correção monetária e juros apropriados (soma dos dias
        do mês), pagamentos realizados e saldo devedor no último dia do mês
        (ou no último vencimento, no mês final). "Competência" traz a data de
        fechamento de cada mês.
        """
        diaria = self.gerar_tabela_diaria(casas_decimais=casas_decimais)
        grupos = diaria.groupby(diaria["Data"].dt.to_period("M"), sort=True)
        mensal = grupos.agg(
            **{
                "Competência": ("Data", "last"),
                "Saldo Devedor Inicial": ("Saldo Devedor Inicial", "first"),
                "Correção Monetária": ("Correção Monetária", "sum"),
                "Juros": ("Juros", "sum"),
                "Amortização": ("Amortização", "sum"),
                "Juros Pagos": ("Juros Pagos", "sum"),
                "Correção Paga": ("Correção Paga", "sum"),
                "Parcela": ("Parcela", "sum"),
                "Saldo Devedor Final": ("Saldo Devedor Final", "last"),
                "Correção Acumulada no Período": ("Correção Acumulada no Período", "last"),
                "Juros Acumulados no Período": ("Juros Acumulados no Período", "last"),
            }
        ).reset_index(drop=True)[self.COLUNAS_MENSAL]
        if casas_decimais is not None:
            numericas = self.COLUNAS_MENSAL[1:]
            mensal[numericas] = mensal[numericas].round(casas_decimais)
        return mensal

    # ------------------------------------------------------------------ #
    # Fluxo de caixa e custo efetivo
    # ------------------------------------------------------------------ #
    def fluxo_caixa(self) -> pd.DataFrame:
        """Fluxo de caixa na ótica do tomador: desembolso líquido (+) na data
        zero e parcelas (-) nos vencimentos."""
        tabela = self.gerar_tabela(casas_decimais=None)
        datas = [self.data_contratacao] + list(tabela["Data"])
        valores = [self.desembolso_liquido] + list(-tabela["Parcela"])
        return pd.DataFrame({"Data": datas, "Fluxo": valores})

    def custo_efetivo_anual(self) -> float:
        """Taxa interna de retorno anualizada do fluxo (dias corridos / 365),
        já considerando o fee inicial e a correção monetária projetada."""
        fluxo = self.fluxo_caixa()
        anos = (fluxo["Data"] - fluxo["Data"].iloc[0]).dt.days.to_numpy() / 365.0
        valores = fluxo["Fluxo"].to_numpy()

        def vpl(taxa: float) -> float:
            return float(np.sum(valores / (1.0 + taxa) ** anos))

        baixo, alto = -0.99, 10.0
        if vpl(baixo) * vpl(alto) > 0:
            raise ValueError("Não foi possível calcular o custo efetivo do fluxo.")
        for _ in range(200):  # bisseção
            meio = (baixo + alto) / 2.0
            if vpl(baixo) * vpl(meio) <= 0:
                alto = meio
            else:
                baixo = meio
        return (baixo + alto) / 2.0

    # ------------------------------------------------------------------ #
    # Exportação para Excel
    # ------------------------------------------------------------------ #
    def descricao_indice(self) -> str:
        """Texto legível sobre o indexador configurado."""
        if not self._tem_indice:
            return "Sem correção monetária"
        nome = self.nome_indice or "Índice"
        if self._funcao_indice is not None:
            return f"{nome} (função informada)"
        if self._serie_indice is not None:
            texto = f"{nome} (série histórica até {self._serie_indice.index[-1]:%d/%m/%Y}"
            if self._taxa_indice is not None:
                texto += f"; depois, projeção de {_fmt_br(self._taxa_indice * 100)}% a.a."
            return texto + ")"
        if self._taxa_indice is not None:
            return f"{nome} (projeção de {_fmt_br(self._taxa_indice * 100)}% a.a.)"
        return nome

    def resumo(self) -> list:
        """Lista de (descrição, valor, formato) com parâmetros e totais."""
        tabela = self.gerar_tabela(casas_decimais=None)
        try:
            custo = self.custo_efetivo_anual()
        except ValueError:
            custo = "n/d"
        moeda, pct, data, texto = "#,##0.00", "0.00%", "DD/MM/YYYY", None
        return [
            ("PARÂMETROS DO CONTRATO", None, None),
            ("Valor do contrato", self.valor_contrato, moeda),
            ("Índice de correção monetária", self.descricao_indice(), texto),
            (
                "Correção monetária",
                "Não se aplica"
                if not self._tem_indice
                else "Paga junto com a parcela"
                if self.correcao_modo == "pagar"
                else "Incorporada ao saldo devedor",
                texto,
            ),
            ("Spread (ao ano)", self.spread, pct),
            ("Data da contratação", self.data_contratacao.to_pydatetime(), data),
            ("Periodicidade", self.periodicidade.capitalize(), texto),
            ("Sistema de amortização", self.sistema_amortizacao, texto),
            ("Prazo total (parcelas, inclui carência)", self.prazo_total, "0"),
            ("Carência (períodos)", self.carencia, "0"),
            (
                "Juros na carência",
                "Capitalizados no saldo devedor"
                if self.juros_carencia == "capitalizar"
                else "Pagos a cada período",
                texto,
            ),
            ("Fee inicial", self.valor_fee, moeda),
            (
                "Cobrança do fee",
                "Deduzido do desembolso"
                if self.fee_modo == "deduzir"
                else "Somado ao saldo devedor",
                texto,
            ),
            (
                "Capitalização dos juros",
                "365 (dias corridos)"
                if self.base_dias == "365"
                else "252 (dias úteis - calendário ANBIMA)",
                texto,
            ),
            ("RESULTADOS", None, None),
            ("Desembolso líquido", self.desembolso_liquido, moeda),
            ("Primeiro vencimento", tabela["Data"].iloc[0].to_pydatetime(), data),
            ("Último vencimento", tabela["Data"].iloc[-1].to_pydatetime(), data),
            ("Total de juros", float(tabela["Juros"].sum()), moeda),
            ("Total de correção monetária", float(tabela["Correção Monetária"].sum()), moeda),
            ("Total amortizado", float(tabela["Amortização"].sum()), moeda),
            ("Total pago (soma das parcelas)", float(tabela["Parcela"].sum()), moeda),
            ("Custo efetivo anual (inclui fee e correção)", custo, pct),
        ]

    def exportar_excel(self, caminho: Union[str, Path]) -> Path:
        """Grava um arquivo .xlsx com quatro abas: 'Resumo' (parâmetros e
        totais), 'Tabela de Pagamentos' (por parcela), 'Contabilização Mensal'
        (apropriação por mês de competência) e 'Cálculo Diário' (pro rata die).
        Devolve o caminho completo do arquivo."""
        try:
            from openpyxl import Workbook
            from openpyxl.styles import Alignment, Font, PatternFill
            from openpyxl.utils import get_column_letter
        except ImportError as exc:  # pragma: no cover
            raise ImportError(
                "A exportação para Excel exige a biblioteca `openpyxl` "
                "(pip install openpyxl)."
            ) from exc

        caminho = Path(caminho).expanduser()
        if caminho.suffix.lower() != ".xlsx":
            caminho = caminho.with_name(caminho.name + ".xlsx")

        negrito = Font(bold=True)
        cabecalho_fonte = Font(bold=True, color="FFFFFF")
        cabecalho_fundo = PatternFill("solid", fgColor="1F3864")
        secao_fundo = PatternFill("solid", fgColor="D9E1F2")

        wb = Workbook()

        # --- Aba Resumo ---
        ws = wb.active
        ws.title = "Resumo"
        for descricao, valor, formato in self.resumo():
            ws.append([descricao, valor])
            linha = ws.max_row
            if valor is None:  # título de seção
                ws.cell(linha, 1).font = negrito
                ws.cell(linha, 1).fill = secao_fundo
                ws.cell(linha, 2).fill = secao_fundo
            else:
                celula = ws.cell(linha, 2)
                celula.alignment = Alignment(horizontal="right")
                if formato:
                    celula.number_format = formato
        ws.column_dimensions["A"].width = 46
        ws.column_dimensions["B"].width = 48

        def escrever_aba(titulo, dados, somar, congelar):
            """Grava um DataFrame em uma aba, com formatação e linha de totais."""
            aba = wb.create_sheet(titulo)
            colunas = list(dados.columns)
            aba.append(colunas)
            for celula in aba[1]:
                celula.font = cabecalho_fonte
                celula.fill = cabecalho_fundo
                celula.alignment = Alignment(
                    horizontal="center", vertical="center", wrap_text=True
                )
            aba.row_dimensions[1].height = 32

            eh_data = [pd.api.types.is_datetime64_any_dtype(dados[c]) for c in colunas]
            eh_inteiro = [pd.api.types.is_integer_dtype(dados[c]) for c in colunas]
            for valores in dados.itertuples(index=False, name=None):
                aba.append(
                    [
                        v.to_pydatetime() if data else (int(v) if inteiro else float(v))
                        for v, data, inteiro in zip(valores, eh_data, eh_inteiro)
                    ]
                )
            ultima = aba.max_row
            centro = Alignment(horizontal="center")
            for linha in aba.iter_rows(min_row=2, max_row=ultima):
                for celula, data, inteiro in zip(linha, eh_data, eh_inteiro):
                    if data:
                        celula.number_format = "DD/MM/YYYY"
                        celula.alignment = centro
                    elif inteiro:
                        celula.alignment = centro
                    else:
                        celula.number_format = "#,##0.00"

            # Linha de totais (fórmulas do próprio Excel)
            total = ultima + 1
            aba.cell(total, 1, "Total").font = negrito
            for nome in somar:
                col = colunas.index(nome) + 1
                letra = get_column_letter(col)
                celula = aba.cell(total, col, f"=SUM({letra}2:{letra}{ultima})")
                celula.number_format = "#,##0.00"
                celula.font = negrito
            for col in range(1, len(colunas) + 1):
                aba.cell(total, col).fill = secao_fundo
                largura = 13 if eh_data[col - 1] else 9 if eh_inteiro[col - 1] else 20
                aba.column_dimensions[get_column_letter(col)].width = largura
            aba.freeze_panes = congelar
            aba.auto_filter.ref = f"A1:{get_column_letter(len(colunas))}{ultima}"

        # --- Aba Tabela de Pagamentos (uma linha por parcela) ---
        pagamentos = self.gerar_tabela(casas_decimais=2).reset_index()
        escrever_aba(
            "Tabela de Pagamentos",
            pagamentos,
            somar=("Juros", "Amortização", "Parcela", "Correção Monetária"),
            congelar="C2",
        )

        # --- Aba Contabilização Mensal (uma linha por mês de competência) ---
        somas = ("Correção Monetária", "Juros", "Amortização", "Juros Pagos",
                 "Correção Paga", "Parcela")
        escrever_aba(
            "Contabilização Mensal",
            self.gerar_resumo_mensal(casas_decimais=2),
            somar=somas,
            congelar="B2",
        )

        # --- Aba Cálculo Diário (uma linha por dia) ---
        escrever_aba(
            "Cálculo Diário",
            self.gerar_tabela_diaria(casas_decimais=2),
            somar=somas,
            congelar="C2",
        )

        wb.save(caminho)
        return caminho.resolve()

    def __repr__(self) -> str:
        return (
            f"CalculadoraEmprestimo(valor={self.valor_contrato:,.2f}, "
            f"indice={self.nome_indice}, spread={self.spread:.4%}, "
            f"sistema={self.sistema_amortizacao}, prazo={self.prazo_total}, "
            f"carencia={self.carencia}, base={self.base_dias})"
        )


# ---------------------------------------------------------------------- #
# Modo interativo: perguntas ao gerente financeiro
# ---------------------------------------------------------------------- #
def _fmt_br(valor: float, casas: int = 2) -> str:
    """Número no padrão brasileiro: 1.234.567,89"""
    texto = f"{valor:,.{casas}f}"
    return texto.replace(",", "#").replace(".", ",").replace("#", ".")


def _interpretar_numero(texto: str, milhar_sem_virgula: bool = False) -> float:
    """Converte '10.000.000,00', '10000000', 'R$ 1.500,50' ou '5,5%' em float.

    Com vírgula: padrão brasileiro (ponto = milhar, vírgula = decimal).
    Sem vírgula: vários pontos são separadores de milhar; um único ponto é
    decimal, exceto em valores monetários do tipo '250.000' (milhar)."""
    t = texto.strip().replace("R$", "").replace("%", "").replace(" ", "")
    if not t:
        raise ValueError
    if "," in t:
        t = t.replace(".", "").replace(",", ".")
    elif t.count(".") > 1:
        t = t.replace(".", "")
    elif t.count(".") == 1 and milhar_sem_virgula:
        inteiro, fracao = t.split(".")
        if len(fracao) == 3 and inteiro.lstrip("-").isdigit():
            t = inteiro + fracao
    return float(t)


def _perguntar(pergunta: str, padrao: Optional[str] = None) -> str:
    sufixo = f" [Enter = {padrao}]" if padrao is not None else ""
    resposta = input(f"{pergunta}{sufixo}: ").strip()
    return resposta if resposta else (padrao or "")


def _ler_numero(
    pergunta: str,
    padrao: Optional[str] = None,
    minimo: Optional[float] = None,
    aceita_minimo: bool = True,
    monetario: bool = False,
) -> float:
    while True:
        resposta = _perguntar(pergunta, padrao)
        try:
            valor = _interpretar_numero(resposta, milhar_sem_virgula=monetario)
        except ValueError:
            print("   -> Não entendi. Digite apenas o número, por exemplo: 1.500.000,00")
            continue
        if minimo is not None and (valor < minimo or (valor == minimo and not aceita_minimo)):
            sinal = "maior ou igual a" if aceita_minimo else "maior que"
            print(f"   -> O valor deve ser {sinal} {_fmt_br(minimo)}.")
            continue
        return valor


def _ler_inteiro(pergunta: str, padrao: Optional[str] = None, minimo: int = 0,
                 maximo: Optional[int] = None) -> int:
    while True:
        resposta = _perguntar(pergunta, padrao)
        if not resposta.isdigit():
            print("   -> Digite um número inteiro, por exemplo: 12")
            continue
        valor = int(resposta)
        if valor < minimo or (maximo is not None and valor > maximo):
            limite = f"entre {minimo} e {maximo}" if maximo is not None else f"a partir de {minimo}"
            print(f"   -> O número deve estar {limite}.")
            continue
        return valor


def _ler_data(pergunta: str, padrao: Optional[str] = None) -> datetime:
    while True:
        resposta = _perguntar(pergunta, padrao)
        for formato in ("%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y", "%d/%m/%y"):
            try:
                return datetime.strptime(resposta, formato)
            except ValueError:
                pass
        print("   -> Data inválida. Use o formato dia/mês/ano, por exemplo: 15/01/2026")


def _ler_opcao(pergunta: str, opcoes: list, padrao: int = 1):
    """Mostra um menu numerado. `opcoes` é uma lista de (valor, descrição)."""
    print(pergunta)
    for i, (_, descricao) in enumerate(opcoes, start=1):
        print(f"   {i} - {descricao}")
    while True:
        resposta = _perguntar("   Escolha o número da opção", str(padrao))
        if resposta.isdigit() and 1 <= int(resposta) <= len(opcoes):
            return opcoes[int(resposta) - 1][0]
        print(f"   -> Digite um número de 1 a {len(opcoes)}.")


def _ler_sim_nao(pergunta: str, padrao: str = "S") -> bool:
    while True:
        resposta = _perguntar(f"{pergunta} (S/N)", padrao).strip().lower()
        if resposta in ("s", "sim"):
            return True
        if resposta in ("n", "nao", "não"):
            return False
        print("   -> Responda S para sim ou N para não.")


def _em_notebook() -> bool:
    """True quando o código roda no Google Colab ou em um Jupyter Notebook."""
    return "google.colab" in sys.modules or "ipykernel" in sys.modules


def _garantir_biblioteca(modulo: str, finalidade: str) -> bool:
    """Verifica se uma biblioteca está instalada; se não estiver, oferece a
    instalação automática (pip) e devolve True quando ela ficar disponível."""
    try:
        importlib.import_module(modulo)
        return True
    except ImportError:
        pass
    comando = [sys.executable, "-m", "pip", "install", modulo]
    print(f"\n   -> Falta instalar o componente '{modulo}' ({finalidade}).")
    if _ler_sim_nao("   Deseja instalar agora? (é preciso estar conectado à internet)", "S"):
        print("   Instalando, aguarde...")
        try:
            subprocess.run(comando, check=False)
        except OSError as erro:
            print(f"   -> A instalação não pôde ser iniciada ({erro}).")
        importlib.invalidate_caches()
        try:
            importlib.import_module(modulo)
            print(f"   -> '{modulo}' instalado com sucesso.")
            return True
        except ImportError:
            print("   -> A instalação automática não funcionou.")
    print("      Para instalar manualmente, peça à TI ou digite no Prompt de Comando:")
    print(f'      "{sys.executable}" -m pip install {modulo}')
    return False


def _ler_serie_de_arquivo() -> pd.Series:
    """Lê uma série de número-índice de um arquivo Excel ou CSV.
    1ª coluna: data (dia/mês/ano). 2ª coluna: número-índice."""
    while True:
        caminho = Path(_perguntar("   Caminho do arquivo (.xlsx ou .csv)").strip('"')).expanduser()
        try:
            if caminho.suffix.lower() in (".xlsx", ".xlsm", ".xls"):
                dados = pd.read_excel(caminho)
            else:
                dados = pd.read_csv(caminho, sep=None, engine="python", decimal=",")
            datas = pd.to_datetime(dados.iloc[:, 0], dayfirst=True)
            valores = pd.to_numeric(dados.iloc[:, 1])
            serie = pd.Series(valores.to_numpy(), index=datas).dropna()
            if serie.empty:
                raise ValueError("arquivo sem dados")
            print(f"   -> {len(serie)} pontos lidos, de {serie.index.min():%d/%m/%Y} "
                  f"a {serie.index.max():%d/%m/%Y}.")
            return serie
        except Exception as erro:  # arquivo inexistente, formato inesperado etc.
            print(f"   -> Não consegui ler o arquivo ({erro}).")
            print("      A 1ª coluna deve ter as datas e a 2ª o número-índice.")
            if not _ler_sim_nao("   Tentar outro arquivo?", "S"):
                raise ValueError("Série histórica não informada.") from erro


def coletar_parametros() -> dict:
    """Faz as perguntas ao usuário e devolve os argumentos da classe."""
    p: dict = {}

    print("\n1) VALOR DO CONTRATO")
    p["valor_contrato"] = _ler_numero(
        "   Valor financiado em R$ (ex.: 10.000.000,00)",
        minimo=0, aceita_minimo=False, monetario=True,
    )

    print("\n2) ÍNDICE DE CORREÇÃO MONETÁRIA")
    nome = _perguntar("   Nome do índice (IPCA, CDI, IGPM...) ou deixe em branco se não houver", "")
    if nome and nome.strip().lower() not in ("nenhum", "nao", "não", "sem", "-"):
        p["indice"] = nome
        forma = _ler_opcao(
            "   Como informar a variação do índice?",
            [
                ("taxa", "Taxa anual esperada (projeção fixa)"),
                ("arquivo", "Série histórica em arquivo Excel/CSV (data e número-índice)"),
            ],
        )
        if forma == "arquivo":
            p["indice"] = _ler_serie_de_arquivo()
            p["_nome_indice"] = nome
            if _ler_sim_nao("   Projetar o índice após o fim da série com uma taxa anual?", "S"):
                p["taxa_indice_projetada"] = _ler_numero(
                    f"   Taxa anual esperada do {nome.upper()}, em % (ex.: 4,5)", minimo=-99.99
                ) / 100.0
        else:
            p["taxa_indice_projetada"] = _ler_numero(
                f"   Taxa anual esperada do {nome.upper()}, em % (ex.: 4,5)", minimo=-99.99
            ) / 100.0
        p["correcao_modo"] = _ler_opcao(
            "   Como a correção monetária é cobrada?",
            [
                ("pagar", "Paga junto com a parcela (Parcela = Amortização + Juros + Correção)"),
                ("incorporar", "Incorporada ao saldo devedor e amortizada nas parcelas seguintes"),
            ],
        )
    else:
        p["indice"] = None

    print("\n3) SPREAD")
    p["spread"] = _ler_numero(
        "   Juros ao ano acima do índice, em % (ex.: 5)", minimo=-99.99
    ) / 100.0

    print("\n4) DATA DA CONTRATAÇÃO")
    p["data_contratacao"] = _ler_data("   Data no formato dia/mês/ano (ex.: 15/01/2026)")

    print("\n5) PERIODICIDADE DOS PAGAMENTOS")
    p["periodicidade"] = _ler_opcao(
        "   Frequência das parcelas:",
        [(nome_p, nome_p.capitalize()) for nome_p in CalculadoraEmprestimo._MESES_POR_PERIODO],
    )

    print("\n6) SISTEMA DE AMORTIZAÇÃO")
    p["sistema_amortizacao"] = _ler_opcao(
        "   Sistema:",
        [
            ("SAC", "SAC - amortização constante, parcelas decrescentes"),
            ("PRICE", "PRICE - parcelas constantes (Tabela Price / Francesa)"),
        ],
    )

    print("\n7) PRAZO TOTAL")
    p["prazo_total"] = _ler_inteiro(
        "   Número total de parcelas, já contando a carência (ex.: 60)", minimo=1
    )

    print("\n8) CARÊNCIA")
    p["carencia"] = _ler_inteiro(
        "   Períodos sem amortização do principal (0 se não houver)",
        padrao="0", minimo=0, maximo=p["prazo_total"] - 1,
    )
    if p["carencia"] > 0:
        p["juros_carencia"] = _ler_opcao(
            "   Durante a carência, os juros são:",
            [
                ("capitalizar", "Capitalizados (somados ao saldo devedor)"),
                ("pagar", "Pagos a cada período"),
            ],
        )

    print("\n9) FEE INICIAL")
    p["fee_tipo"] = _ler_opcao(
        "   O fee é cobrado como:",
        [("percentual", "Percentual do valor do contrato"), ("fixo", "Valor fixo em R$")],
    )
    if p["fee_tipo"] == "percentual":
        p["fee_inicial"] = _ler_numero(
            "   Fee em % (ex.: 1; 0 se não houver)", padrao="0", minimo=0
        ) / 100.0
    else:
        p["fee_inicial"] = _ler_numero(
            "   Fee em R$ (0 se não houver)", padrao="0", minimo=0, monetario=True
        )
    if p["fee_inicial"] > 0:
        p["fee_modo"] = _ler_opcao(
            "   Como o fee é cobrado?",
            [
                ("deduzir", "Deduzido do valor desembolsado"),
                ("financiar", "Somado ao saldo devedor"),
            ],
        )

    print("\n10) FORMA DE CAPITALIZAÇÃO DOS JUROS")
    while True:
        p["base_dias"] = _ler_opcao(
            "   Base de dias:",
            [("365", "365 - dias corridos"), ("252", "252 - dias úteis (calendário ANBIMA)")],
        )
        # A base 252 depende da biblioteca `bizdays` (feriados ANBIMA).
        if p["base_dias"] == "365" or _garantir_biblioteca(
            "bizdays", "contagem de dias úteis da base 252"
        ):
            break
        print("\n   Sem esse componente não é possível usar a base 252. Escolha "
              "novamente\n   (365 funciona sem ele) ou instale-o e repita a opção 2.")
    return p


def _montar_calculadora(parametros: dict) -> CalculadoraEmprestimo:
    p = dict(parametros)
    nome = p.pop("_nome_indice", None)
    calculadora = CalculadoraEmprestimo(**p)
    if nome:
        calculadora.definir_indice(nome=nome)
    return calculadora


def _imprimir_resumo(calculadora: CalculadoraEmprestimo, somente_parametros: bool = False) -> None:
    for descricao, valor, formato in calculadora.resumo():
        if valor is None:
            if somente_parametros and descricao == "RESULTADOS":
                break
            print(f"\n{descricao}")
            continue
        if isinstance(valor, datetime):
            texto = valor.strftime("%d/%m/%Y")
        elif formato == "0.00%" and not isinstance(valor, str):
            texto = _fmt_br(valor * 100, 2) + "%"
        elif formato == "#,##0.00":
            texto = "R$ " + _fmt_br(valor)
        else:
            texto = str(valor)
        print(f"   {descricao:.<46} {texto}")


def executar_interativo() -> None:
    print("=" * 70)
    print(" CALCULADORA DE EMPRÉSTIMO - TABELA DE PAGAMENTOS")
    print("=" * 70)
    print("Responda às perguntas e tecle Enter. Quando houver [Enter = ...],")
    print("basta teclar Enter para aceitar a sugestão.")
    if _em_notebook():
        print("\nATENÇÃO (Colab/Jupyter): a cada pergunta aparece uma caixinha de texto")
        print("logo abaixo dela. Clique na caixinha, digite a resposta e tecle Enter.")
        print("Enquanto o programa espera a resposta, a célula aparece como 'executando'.")

    if not _garantir_biblioteca("openpyxl", "gravação do arquivo Excel"):
        print("\nSem esse componente não é possível gerar o Excel. Instale-o e abra "
              "o programa novamente.")
        return

    while True:
        try:
            calculadora = _montar_calculadora(coletar_parametros())
            tabela = calculadora.gerar_tabela()
            print("\n" + "-" * 70 + "\nCONFIRA OS DADOS INFORMADOS")
            _imprimir_resumo(calculadora, somente_parametros=True)
        except (ValueError, ImportError) as erro:
            print(f"\n*** Não foi possível calcular: {erro}")
            if _ler_sim_nao("Deseja informar os dados novamente?", "S"):
                continue
            return
        if _ler_sim_nao("\nOs dados estão corretos?", "S"):
            break
        print("\nTudo bem, vamos começar de novo.")

    print("\n" + "-" * 70)
    _imprimir_resumo(calculadora)
    print("\nPRIMEIRAS 5 LINHAS DA TABELA")
    with pd.option_context(
        "display.width", 200, "display.max_columns", None,
        "display.float_format", lambda v: _fmt_br(v),
    ):
        exibicao = tabela.head(5).copy()
        exibicao["Data"] = exibicao["Data"].dt.strftime("%d/%m/%Y")
        print(exibicao)

        print("\nCONTABILIZAÇÃO MENSAL - 5 PRIMEIROS MESES (apropriação diária)")
        mensal = calculadora.gerar_resumo_mensal().head(5)[
            ["Competência", "Correção Monetária", "Juros", "Parcela", "Saldo Devedor Final"]
        ].copy()
        mensal["Competência"] = mensal["Competência"].dt.strftime("%d/%m/%Y")
        print(mensal.to_string(index=False))
    print("\nO Excel terá 4 abas: Resumo, Tabela de Pagamentos, Contabilização Mensal")
    print("e Cálculo Diário.")

    print("\n" + "-" * 70 + "\nEXPORTAÇÃO PARA EXCEL")
    sugestao = f"tabela_emprestimo_{datetime.now():%Y%m%d_%H%M}.xlsx"
    while True:
        nome_arquivo = _perguntar("Nome do arquivo Excel", sugestao)
        try:
            caminho = calculadora.exportar_excel(nome_arquivo)
        except PermissionError:
            print("   -> Não consegui gravar. O arquivo está aberto no Excel? "
                  "Feche-o ou escolha outro nome.")
            continue
        except (OSError, ImportError) as erro:
            print(f"   -> Não consegui gravar o arquivo ({erro}).")
            if not _ler_sim_nao("Tentar novamente?", "S"):
                return
            continue
        print(f"\nArquivo gravado em:\n   {caminho}")
        if "google.colab" in sys.modules:
            # No Colab o arquivo fica no servidor; dispara o download no navegador.
            try:
                from google.colab import files

                files.download(str(caminho))
                print("O download do arquivo foi iniciado no navegador.")
            except Exception:
                print("Para baixar: abra o ícone de pasta, à esquerda, clique nos "
                      "três pontos do arquivo e escolha 'Fazer download'.")
        print("\nFIM - cálculo concluído.")
        return


def executar_exemplo() -> None:
    """Exemplo pronto, sem perguntas (python calculadora_emprestimo.py --exemplo)."""
    pd.set_option("display.width", 200)
    pd.set_option("display.max_columns", None)
    pd.set_option("display.float_format", lambda v: f"{v:,.2f}")

    emprestimo = CalculadoraEmprestimo(
        valor_contrato=10_000_000.00,
        indice="IPCA",
        taxa_indice_projetada=0.045,   # projeção de 4,5% a.a. para o IPCA
        spread=0.05,                   # 5% a.a. acima do índice
        data_contratacao=datetime(2026, 1, 15),
        periodicidade="mensal",
        sistema_amortizacao="SAC",
        prazo_total=60,                # 60 parcelas, incluindo a carência
        carencia=6,                    # 6 meses com juros capitalizados
        fee_inicial=0.01,              # 1% deduzido do desembolso
        base_dias="365",               # troque por "252" (requer `bizdays`)
    )
    print(emprestimo)
    print(emprestimo.gerar_tabela().head(5))
    print("Excel gravado em:", emprestimo.exportar_excel("tabela_emprestimo_exemplo.xlsx"))


def _em_streamlit() -> bool:
    """True quando este arquivo está sendo executado pelo Streamlit."""
    try:
        from streamlit.runtime.scriptrunner import get_script_run_ctx

        return get_script_run_ctx() is not None
    except Exception:
        return False


if __name__ == "__main__":
    if _em_streamlit():
        # Este arquivo foi indicado como arquivo principal no Streamlit. As
        # perguntas no terminal não funcionam ali (a página ficaria em branco),
        # então abre-se a interface web, que está em app.py na mesma pasta.
        import runpy

        import streamlit as st

        _app = Path(__file__).with_name("app.py")
        if _app.exists():
            runpy.run_path(str(_app), run_name="__main__")
        else:
            st.error(
                "O arquivo app.py não foi encontrado no repositório. Envie o "
                "app.py para a mesma pasta deste arquivo e indique-o como "
                "arquivo principal (Main file path) do aplicativo."
            )
    elif "--exemplo" in sys.argv:
        executar_exemplo()
    else:
        try:
            executar_interativo()
        except (KeyboardInterrupt, EOFError):
            print("\n\nOperação cancelada.")
        # Mantém a janela aberta quando o programa é aberto com duplo clique.
        # (No Colab/Jupyter isso deixaria a célula "executando" à toa.)
        if not _em_notebook():
            try:
                input("\nPressione Enter para sair...")
            except (KeyboardInterrupt, EOFError):
                pass
