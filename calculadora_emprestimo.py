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

Cronograma de pagamento informado
---------------------------------
Em vez da periodicidade fixa, é possível informar o cronograma das parcelas
(parâmetro `cronograma`): as datas de vencimento e, opcionalmente, a
amortização de cada parcela, em R$ ou em % do principal.
* Só as datas: a amortização é calculada por SAC ou PRICE sobre essas datas.
* Datas + amortização: vale a amortização informada. As parcelas iniciais com
  amortização zero são tratadas como carência; parcelas posteriores com
  amortização zero pagam apenas os encargos.

CDI/Selic: taxas realizadas e curva futura
------------------------------------------
* Histórico: `ler_serie_indice()` lê o arquivo de taxas diárias do Banco
  Central e `buscar_serie_bcb()` consulta a mesma informação pela internet.
* Projeção: em vez de uma taxa anual fixa, pode-se usar a curva DI x pré
  (taxas referenciais da B3): `ler_curva_b3()` lê o arquivo baixado no site da
  B3 e `buscar_curva_b3()` tenta consultá-la pela internet. A curva é
  interpolada (taxa a termo constante entre os vértices, base 252 dias úteis)
  e projeta o índice a partir do fim do histórico.

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
import io
import subprocess
import unicodedata
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
    cronograma : lista de datas | DataFrame | lista de pares, opcional
        Cronograma de pagamento informado pelo usuário. Substitui
        `periodicidade` e `prazo_total`. Formatos aceitos:
        - lista de datas ............ só os vencimentos (amortização por SAC/PRICE,
                                      respeitando `carencia`);
        - DataFrame com 1 coluna .... idem;
        - DataFrame com 2 colunas ... 1ª = data, 2ª = amortização de cada parcela
                                      (substitui `sistema_amortizacao` e `carencia`);
        - lista de pares (data, amortização) ... idem.
    curva_projecao : DataFrame, opcional
        Curva de juros futuros (DI x pré da B3) usada para projetar o índice
        após o fim do histórico, no lugar de `taxa_indice_projetada`. Colunas
        "Dias Úteis" e "Taxa (% a.a.)", como devolvem `ler_curva_b3()` e
        `buscar_curva_b3()`.
    data_curva : data, opcional
        Data de referência da curva (padrão: a registrada na própria curva).
    cronograma_tipo : str
        Como ler a amortização do cronograma: 'valor' (R$, padrão - a soma deve
        ser igual ao saldo devedor inicial) ou 'percentual' (% do principal -
        a soma deve ser 100).
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
        periodicidade: str = "mensal",
        sistema_amortizacao: str = "SAC",
        prazo_total: Optional[int] = None,
        carencia: int = 0,
        fee_inicial: float = 0.0,
        base_dias: Union[str, int] = "365",
        *,
        taxa_indice_projetada: Optional[float] = None,
        juros_carencia: str = "capitalizar",
        fee_tipo: str = "percentual",
        fee_modo: str = "deduzir",
        correcao_modo: str = "pagar",
        cronograma=None,
        cronograma_tipo: str = "valor",
        curva_projecao=None,
        data_curva: Optional[TipoData] = None,
        calendario=None,
    ) -> None:
        self.valor_contrato = float(valor_contrato)
        self.spread = float(spread)
        self.data_contratacao = pd.Timestamp(data_contratacao).normalize()
        self.periodicidade = str(periodicidade).strip().lower()
        self.sistema_amortizacao = str(sistema_amortizacao).strip().upper()
        self.carencia = int(carencia)
        self.fee_inicial = float(fee_inicial)
        self.base_dias = str(base_dias).strip()
        self.juros_carencia = str(juros_carencia).strip().lower()
        self.fee_tipo = str(fee_tipo).strip().lower()
        self.fee_modo = str(fee_modo).strip().lower()
        self.correcao_modo = str(correcao_modo).strip().lower()
        self._calendario = calendario

        # Cronograma informado (datas e, opcionalmente, amortizações)
        self.cronograma_tipo = str(cronograma_tipo).strip().lower()
        self._datas_cronograma: Optional[list] = None
        self._amortizacao_cronograma: Optional[np.ndarray] = None
        self._pesos: Optional[np.ndarray] = None
        if cronograma is not None:
            self._carregar_cronograma(cronograma)
            self.prazo_total = len(self._datas_cronograma)
        elif prazo_total is None:
            raise ValueError("Informe `prazo_total` ou um `cronograma` de pagamento.")
        else:
            self.prazo_total = int(prazo_total)

        self._validar_parametros()
        if self._amortizacao_cronograma is not None:
            self._definir_pesos()

        # Estado do indexador
        self.nome_indice: Optional[str] = None
        self._taxa_indice: Optional[float] = None
        self._serie_indice: Optional[pd.Series] = None
        self._serie_diaria = False
        self._curva_x: Optional[np.ndarray] = None   # dias úteis dos vértices
        self._curva_y: Optional[np.ndarray] = None   # log do fator acumulado
        self._data_curva: Optional[pd.Timestamp] = None
        self._du_curva_cache: dict = {}
        self._funcao_indice: Optional[TipoFuncaoIndice] = None
        self._tem_indice = (
            indice is not None
            or taxa_indice_projetada is not None
            or curva_projecao is not None
        )
        if curva_projecao is not None:
            self.definir_indice(curva=curva_projecao, data_curva=data_curva)

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
        if self._datas_cronograma is None and self.periodicidade not in self._MESES_POR_PERIODO:
            raise ValueError(
                f"`periodicidade` inválida: {self.periodicidade!r}. "
                f"Use uma de {sorted(self._MESES_POR_PERIODO)}."
            )
        if self._amortizacao_cronograma is None and self.sistema_amortizacao not in ("SAC", "PRICE"):
            raise ValueError("`sistema_amortizacao` deve ser 'SAC' ou 'PRICE'.")
        if self.prazo_total <= 0:
            raise ValueError("`prazo_total` deve ser um inteiro positivo.")
        if self._amortizacao_cronograma is None and not 0 <= self.carencia < self.prazo_total:
            raise ValueError(
                f"A carência ({self.carencia}) deve ser menor que o número de "
                f"parcelas ({self.prazo_total}): é preciso ao menos uma parcela "
                "de amortização."
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
        curva=None,
        data_curva: Optional[TipoData] = None,
    ) -> "CalculadoraEmprestimo":
        """Injeta (ou substitui) a fonte de dados do indexador.

        nome ........ rótulo do índice (ex.: 'IPCA', 'CDI', 'IGPM').
        taxa_anual .. taxa anual esperada, em decimal. Sozinha, funciona como
                      projeção fixa; junto de `serie`, extrapola a série após
                      o último ponto disponível.
        serie ....... dict ou pandas.Series {data: número-índice}.
        funcao ...... f(data_inicio, data_fim) -> fator de correção do período.
                      Tem precedência sobre `serie` e `taxa_anual`.
        curva ....... curva de juros futuros (DataFrame com "Dias Úteis" e
                      "Taxa (% a.a.)"), usada para projetar o índice após o
                      fim da série. Tem precedência sobre `taxa_anual`.
        data_curva .. data de referência da curva.
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
            # Série com pontos diários (ex.: CDI/Selic acumulados): o cálculo
            # diário usa o nível real de cada dia, em vez do rateio pro rata.
            espacos = s.index.to_series().diff().dt.days.dropna()
            self._serie_diaria = bool(len(espacos) and espacos.median() <= 4)
        if funcao is not None:
            if not callable(funcao):
                raise TypeError("`funcao` deve ser chamável.")
            self._funcao_indice = funcao
        if curva is not None:
            self._carregar_curva(curva, data_curva)
        self._tem_indice = True
        return self

    def _carregar_curva(self, curva: pd.DataFrame, data_curva) -> None:
        """Prepara a curva de juros para interpolação.

        Guarda, para cada vértice, o prazo em dias úteis e o logaritmo do
        fator acumulado, ln[(1 + taxa) ** (du / 252)]. A interpolação linear
        desse logaritmo equivale a taxa a termo constante entre os vértices.
        """
        referencia = data_curva if data_curva is not None else curva.attrs.get("data_referencia")
        if referencia is None:
            raise ValueError("Informe a data de referência da curva de juros (`data_curva`).")
        colunas = {_sem_acento(c): c for c in curva.columns}
        col_du = next((c for k, c in colunas.items() if "uteis" in k), None)
        col_taxa = next((c for k, c in colunas.items() if "taxa" in k or "preco" in k), None)
        if col_du is None or col_taxa is None:
            raise ValueError("A curva deve ter as colunas 'Dias Úteis' e 'Taxa (% a.a.)'.")
        dados = pd.DataFrame({
            "du": pd.to_numeric(curva[col_du], errors="coerce"),
            "taxa": pd.to_numeric(curva[col_taxa], errors="coerce"),
        }).dropna()
        dados = dados[dados["du"] > 0].drop_duplicates("du").sort_values("du")
        if dados.empty or (dados["taxa"] <= -100).any():
            raise ValueError("A curva de juros não tem vértices válidos.")
        du = dados["du"].to_numpy(dtype=float)
        log_fator = du / 252.0 * np.log1p(dados["taxa"].to_numpy(dtype=float) / 100.0)
        self._curva_x = np.concatenate(([0.0], du))
        self._curva_y = np.concatenate(([0.0], log_fator))
        self._data_curva = pd.Timestamp(referencia).normalize()
        self._du_curva_cache = {}

    def _log_fator_curva(self, data: pd.Timestamp) -> float:
        """Logaritmo do fator acumulado pela curva entre a data de referência
        e `data` (negativo para datas anteriores à referência)."""
        du = self._du_curva_cache.get(data)
        if du is None:
            if data >= self._data_curva:
                du = float(self._dias_uteis(self._data_curva, data))
            else:
                du = -float(self._dias_uteis(data, self._data_curva))
            self._du_curva_cache[data] = du
        x, y = self._curva_x, self._curva_y
        if du < 0:  # antes da referência: taxa do primeiro vértice
            return du * y[1] / x[1]
        if du > x[-1]:  # além do último vértice: última taxa a termo
            if len(x) > 2:
                inclinacao = (y[-1] - y[-2]) / (x[-1] - x[-2])
            else:
                inclinacao = y[-1] / x[-1]
            return float(y[-1] + (du - x[-1]) * inclinacao)
        return float(np.interp(du, x, y))

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
        if self._curva_x is not None:
            # Projeção pela curva de juros a partir do último ponto da série.
            return float(s.iloc[-1]) * float(
                np.exp(self._log_fator_curva(data) - self._log_fator_curva(ultima))
            )
        if self._taxa_indice is None:
            raise ValueError(
                f"A série do índice termina em {ultima.date()} e o fluxo exige "
                f"{data.date()}. Informe uma taxa anual ou uma curva de juros "
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
        if self._curva_x is not None:
            return float(np.exp(self._log_fator_curva(fim) - self._log_fator_curva(ini)))
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
        return self._dias_uteis(ini, fim)

    def _dias_uteis(self, ini: pd.Timestamp, fim: pd.Timestamp) -> int:
        """Dias úteis no intervalo [ini, fim): inclui a data inicial e exclui a final.

        Datas em fim de semana ou feriado são levadas ao dia útil seguinte
        antes da contagem. Assim nenhum dia útil é perdido nem contado em dobro
        quando um vencimento cai em dia não útil (a sexta-feira anterior a um
        vencimento no sábado, por exemplo, pertence ao período que se encerra).
        """
        cal = self._obter_calendario()
        ajustar = getattr(cal, "following", None)
        d1, d2 = ini.date(), fim.date()
        if ajustar is not None:
            d1 = pd.Timestamp(ajustar(d1)).date()
            d2 = pd.Timestamp(ajustar(d2)).date()
        return int(cal.bizdays(d1, d2))

    def _fracao_ano(self, ini: pd.Timestamp, fim: pd.Timestamp) -> float:
        return self._dias(ini, fim) / float(self.base_dias)

    def _carregar_cronograma(self, cronograma) -> None:
        """Normaliza o cronograma informado: datas e, se houver, amortizações."""
        if self.cronograma_tipo not in ("valor", "percentual"):
            raise ValueError("`cronograma_tipo` deve ser 'valor' ou 'percentual'.")

        if isinstance(cronograma, pd.DataFrame):
            dados = cronograma.copy()
        elif isinstance(cronograma, pd.Series):
            dados = cronograma.to_frame()
        else:
            itens = list(cronograma)
            if itens and isinstance(itens[0], (list, tuple)):
                dados = pd.DataFrame(itens)
            else:
                dados = pd.DataFrame({"Data": itens})
        dados = dados.dropna(how="all")
        if dados.empty:
            raise ValueError("O cronograma de pagamento está vazio.")

        coluna_data = dados.iloc[:, 0]
        if coluna_data.isna().any():
            raise ValueError("O cronograma tem linha sem data de vencimento.")
        try:
            if pd.api.types.is_datetime64_any_dtype(coluna_data):
                datas = pd.to_datetime(coluna_data)
            else:
                datas = pd.to_datetime(coluna_data, dayfirst=True)
        except (ValueError, TypeError) as exc:
            raise ValueError(
                "O cronograma tem data inválida. Use o formato dia/mês/ano."
            ) from exc
        datas = [pd.Timestamp(d).normalize() for d in datas]

        if datas[0] <= self.data_contratacao:
            raise ValueError(
                f"O primeiro vencimento do cronograma ({datas[0]:%d/%m/%Y}) deve ser "
                f"posterior à data da contratação ({self.data_contratacao:%d/%m/%Y})."
            )
        for anterior, atual in zip(datas[:-1], datas[1:]):
            if atual <= anterior:
                raise ValueError(
                    "As datas do cronograma devem estar em ordem crescente e sem "
                    f"repetição (verifique {anterior:%d/%m/%Y} e {atual:%d/%m/%Y})."
                )
        self._datas_cronograma = datas

        if dados.shape[1] >= 2:
            try:
                valores = pd.to_numeric(dados.iloc[:, 1]).fillna(0.0).to_numpy(dtype=float)
            except (ValueError, TypeError) as exc:
                raise ValueError(
                    "A coluna de amortização do cronograma deve conter apenas números."
                ) from exc
            if (valores < 0).any():
                raise ValueError("O cronograma não pode ter amortização negativa.")
            if valores.sum() <= 0:
                raise ValueError("O cronograma não tem nenhuma amortização informada.")
            self._amortizacao_cronograma = valores

    def _definir_pesos(self) -> None:
        """Converte a amortização informada em pesos (fração do principal) e
        define a carência como as parcelas iniciais sem amortização."""
        valores = self._amortizacao_cronograma
        total = float(valores.sum())
        if self.cronograma_tipo == "percentual":
            if abs(total - 100.0) > 0.01:
                raise ValueError(
                    "Os percentuais de amortização do cronograma devem somar 100%. "
                    f"A soma informada é {_fmt_br(total, 4)}%."
                )
        else:
            esperado = self.saldo_devedor_inicial
            if abs(total - esperado) > 0.01:
                raise ValueError(
                    "A soma das amortizações do cronograma (R$ "
                    f"{_fmt_br(total)}) deve ser igual ao saldo devedor inicial "
                    f"(R$ {_fmt_br(esperado)}). Diferença: R$ {_fmt_br(total - esperado)}."
                )
        self._pesos = valores / total
        self.carencia = int(np.argmax(valores > 0))  # zeros iniciais = carência

    @property
    def tem_cronograma(self) -> bool:
        return self._datas_cronograma is not None

    @property
    def amortizacao_informada(self) -> bool:
        return self._pesos is not None

    def datas_pagamento(self) -> list:
        """Datas de vencimento: as do cronograma informado ou, na falta dele,
        as geradas pela periodicidade a partir da data de contratação (evita o
        deslocamento de fim de mês ao somar períodos sucessivos)."""
        if self._datas_cronograma is not None:
            return list(self._datas_cronograma)
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
                if self._pesos is not None:
                    # Amortização informada no cronograma: proporção do peso
                    # desta parcela sobre os pesos ainda a amortizar. Sem
                    # capitalização, resulta exatamente no valor informado;
                    # se o saldo cresceu (carência capitalizada ou correção
                    # incorporada), o acréscimo é distribuído na mesma proporção.
                    peso_restante = float(self._pesos[k - 1 :].sum())
                    if restantes == 1 or peso_restante - self._pesos[k - 1] < 1e-12:
                        amortizacao = base if self._pesos[k - 1] > 0 else 0.0
                    else:
                        amortizacao = base * self._pesos[k - 1] / peso_restante
                elif restantes == 1:
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
        return np.array([float(self._dias_uteis(ini, d)) for d in datas])

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

            if self._funcao_indice is None and (
                (self._serie_diaria and self._serie_indice is not None)
                or self._curva_x is not None
            ):
                # Série diária ou curva de juros: variação do índice até cada dia.
                fator_acum = np.array([self._fator_correcao(ini, dia) for dia in dias])
            else:
                # Demais casos: variação do período distribuída pro rata die.
                fator_acum = fator_corr ** fracao
            corr_acum = s0 * (fator_acum - 1.0)
            juros_acum = s0 * fator_acum * ((1.0 + self.spread) ** (d / base) - 1.0)
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
        if self._curva_x is not None:
            projecao = f"curva DI x pré de {self._data_curva:%d/%m/%Y}"
        elif self._taxa_indice is not None:
            projecao = f"projeção de {_fmt_br(self._taxa_indice * 100)}% a.a."
        else:
            projecao = None
        if self._serie_indice is not None:
            texto = f"{nome} (série histórica até {self._serie_indice.index[-1]:%d/%m/%Y}"
            if projecao:
                texto += f"; depois, {projecao}"
            return texto + ")"
        if projecao:
            return f"{nome} ({projecao})"
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
            (
                "Vencimentos",
                "Conforme cronograma informado"
                if self.tem_cronograma
                else f"Periodicidade {self.periodicidade}",
                texto,
            ),
            (
                "Sistema de amortização",
                (
                    "Amortização informada no cronograma (R$)"
                    if self.cronograma_tipo == "valor"
                    else "Amortização informada no cronograma (% do principal)"
                )
                if self.amortizacao_informada
                else self.sistema_amortizacao,
                texto,
            ),
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
            f"sistema={'CRONOGRAMA' if self.amortizacao_informada else self.sistema_amortizacao}, "
            f"prazo={self.prazo_total}, "
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


def _sem_acento(texto) -> str:
    texto = unicodedata.normalize("NFKD", str(texto))
    return "".join(c for c in texto if not unicodedata.combining(c)).strip().lower()


def _para_numero(coluna: pd.Series) -> pd.Series:
    """Converte textos como '1.234,56', '13,65' ou '5%' em número."""
    if pd.api.types.is_numeric_dtype(coluna):
        return pd.to_numeric(coluna, errors="coerce")

    def converter(valor):
        if valor is None or (isinstance(valor, float) and np.isnan(valor)):
            return np.nan
        if isinstance(valor, (int, float, np.number)):
            return float(valor)
        t = str(valor).strip().replace("R$", "").replace("%", "").replace(" ", "")
        if not t:
            return np.nan
        if "," in t:
            t = t.replace(".", "").replace(",", ".")
        try:
            return float(t)
        except ValueError:
            return np.nan

    return coluna.map(converter).astype(float)


def _para_data(coluna: pd.Series) -> pd.Series:
    if pd.api.types.is_datetime64_any_dtype(coluna):
        return pd.to_datetime(coluna, errors="coerce")
    texto = coluna.map(lambda v: v if isinstance(v, (datetime, date, pd.Timestamp)) else str(v).strip())
    return pd.to_datetime(texto, dayfirst=True, errors="coerce", format="mixed")


def _ler_tabela(arquivo, nome_arquivo: Optional[str] = None):
    """Lê um arquivo Excel ou CSV e localiza onde começam os dados.

    Aceita arquivos com linhas de título antes do cabeçalho (como os
    exportados pelo Banco Central), CSV separado por ponto e vírgula, vírgula
    ou tabulação, e números no padrão brasileiro. Devolve (dados, cabecalhos):
    `dados` só com as linhas cuja 1ª coluna é uma data, e `cabecalhos` com o
    nome de cada coluna (em minúsculas e sem acento; vazio se não houver).
    """
    nome = str(nome_arquivo or getattr(arquivo, "name", arquivo)).lower()
    if nome.endswith((".xlsx", ".xlsm", ".xls")):
        bruto = pd.read_excel(arquivo, header=None)
    else:
        if hasattr(arquivo, "read"):
            conteudo = arquivo.read()
        else:
            conteudo = Path(arquivo).read_bytes()
        if isinstance(conteudo, bytes):
            try:
                conteudo = conteudo.decode("utf-8-sig")
            except UnicodeDecodeError:
                conteudo = conteudo.decode("latin-1")
        linhas = [linha for linha in conteudo.splitlines() if linha.strip()]
        if not linhas:
            raise ValueError("O arquivo está vazio.")
        amostra = linhas[-20:]  # as últimas linhas são sempre de dados
        separador = max(
            (";", "\t", ","), key=lambda sep: sum(linha.count(sep) for linha in amostra)
        )
        campos = [linha.split(separador) for linha in linhas]
        largura = max(len(c) for c in campos)
        bruto = pd.DataFrame(
            [[(v.strip() or None) if v is not None else None for v in c + [None] * (largura - len(c))]
             for c in campos],
            dtype=object,
        )

    bruto = bruto.dropna(how="all").reset_index(drop=True)
    if bruto.empty:
        raise ValueError("O arquivo não tem dados.")
    datas = _para_data(bruto.iloc[:, 0])
    eh_dado = datas.notna()
    if not eh_dado.any():
        raise ValueError(
            "Não encontrei datas na 1ª coluna do arquivo. Ela deve trazer as "
            "datas no formato dia/mês/ano."
        )
    primeira = int(eh_dado.idxmax())
    if primeira > 0:
        cabecalhos = [_sem_acento(v) if pd.notna(v) else "" for v in bruto.iloc[primeira - 1]]
    else:
        cabecalhos = [""] * bruto.shape[1]
    # Monta a tabela de dados: datas já convertidas + demais colunas não vazias.
    colunas = {0: datas[eh_dado].reset_index(drop=True)}
    nomes = [cabecalhos[0]]
    for posicao in range(1, bruto.shape[1]):
        coluna = bruto.iloc[:, posicao][eh_dado].reset_index(drop=True)
        if coluna.notna().any():
            colunas[len(colunas)] = coluna.astype(object)
            nomes.append(cabecalhos[posicao])
    return pd.DataFrame(colunas), nomes


def _indice_de_fatores(datas, fatores, origem: str) -> pd.Series:
    """Acumula fatores diários em um número-índice.

    Nível 1 no primeiro dia; o fator de cada dia útil vale a partir do dia
    seguinte (corrige o saldo desse dia até o próximo dia útil).
    """
    tabela = pd.DataFrame({"data": pd.to_datetime(datas), "fator": fatores}).dropna()
    tabela = tabela.drop_duplicates("data").sort_values("data")
    if tabela.empty or (tabela["fator"] <= 0).any():
        raise ValueError("A série de fatores diários é inválida ou está vazia.")
    acumulado = tabela["fator"].cumprod().to_numpy()
    indice_datas = [tabela["data"].iloc[0]] + list(tabela["data"] + pd.Timedelta(days=1))
    serie = pd.Series(np.concatenate(([1.0], acumulado)), index=pd.DatetimeIndex(indice_datas))
    serie.attrs["origem"] = origem
    serie.attrs["periodo"] = (tabela["data"].iloc[0], tabela["data"].iloc[-1])
    # Taxa anual equivalente ao último fator (sugestão para a projeção)
    serie.attrs["taxa_anual_final"] = float(tabela["fator"].iloc[-1]) ** 252 - 1.0
    return serie


def _texto_do_arquivo(arquivo) -> str:
    """Conteúdo de um arquivo de texto (caminho ou objeto enviado), em UTF-8
    ou, se não for possível, em Latin-1 (padrão dos arquivos da B3)."""
    conteudo = arquivo.read() if hasattr(arquivo, "read") else Path(arquivo).read_bytes()
    if isinstance(conteudo, bytes):
        try:
            conteudo = conteudo.decode("utf-8-sig")
        except UnicodeDecodeError:
            conteudo = conteudo.decode("latin-1")
    return conteudo


COLUNAS_CURVA = ["Dias Úteis", "Dias Corridos", "Taxa (% a.a.)"]


def ler_curva_b3(arquivo, nome_arquivo: Optional[str] = None,
                 data_referencia: Optional[TipoData] = None) -> pd.DataFrame:
    """Lê o arquivo de taxas referenciais da B3 (curva DI x pré).

    Formato do arquivo baixado no site da B3 (TaxaReferencia_PRE_AAAAMMDD.csv):
    "Descrição da Taxa; Dias Úteis; Dias Corridos; Preço/Taxa", com a taxa em
    % ao ano, base 252 dias úteis. A data de referência é lida do nome do
    arquivo (AAAAMMDD), a menos que `data_referencia` seja informada; fica em
    `curva.attrs["data_referencia"]` (None se não for possível identificá-la).
    """
    nome = str(nome_arquivo or getattr(arquivo, "name", arquivo))
    if nome.lower().endswith((".xlsx", ".xlsm", ".xls")):
        bruto = pd.read_excel(arquivo, header=None).astype(object)
        linhas = [list(linha) for linha in bruto.itertuples(index=False, name=None)]
    else:
        texto = _texto_do_arquivo(arquivo)
        linhas = [linha.split(";") for linha in texto.splitlines() if linha.strip()]
    posicao = None
    for numero, linha in enumerate(linhas):
        titulos = [_sem_acento(v) for v in linha]
        if any("uteis" in t for t in titulos):
            posicao = numero
            col_du = next(i for i, t in enumerate(titulos) if "uteis" in t)
            col_dc = next((i for i, t in enumerate(titulos) if "corridos" in t), None)
            col_taxa = next(
                (i for i, t in enumerate(titulos) if "taxa" in t and "descri" not in t), None
            )
            break
    if posicao is None or col_taxa is None:
        raise ValueError(
            "Não reconheci o arquivo da curva. Ele deve ter as colunas "
            "'Dias Úteis' e 'Preço/Taxa', como o arquivo de taxas referenciais da B3."
        )
    dados = pd.DataFrame(linhas[posicao + 1:])
    curva = pd.DataFrame({
        "Dias Úteis": _para_numero(dados[col_du]),
        "Dias Corridos": _para_numero(dados[col_dc]) if col_dc is not None else np.nan,
        "Taxa (% a.a.)": _para_numero(dados[col_taxa]),
    }).dropna(subset=["Dias Úteis", "Taxa (% a.a.)"])
    curva = curva.sort_values("Dias Úteis").reset_index(drop=True)
    if curva.empty:
        raise ValueError("O arquivo da curva não tem vértices.")

    referencia = pd.Timestamp(data_referencia) if data_referencia is not None else None
    if referencia is None:
        import re

        achado = re.search(r"(20\d{2})(\d{2})(\d{2})", Path(nome).name)
        if achado:
            try:
                referencia = pd.Timestamp(
                    int(achado.group(1)), int(achado.group(2)), int(achado.group(3))
                )
            except ValueError:
                referencia = None
    curva.attrs["data_referencia"] = referencia
    curva.attrs["origem"] = "arquivo da B3"
    return curva


# ---------------------------------------------------------------------- #
# Consultas pela internet (Banco Central e B3)
# ---------------------------------------------------------------------- #
class ErroDeConsulta(RuntimeError):
    """Falha ao consultar uma fonte de dados externa."""


def _baixar(url: str, timeout: int = 20) -> bytes:
    import urllib.error
    import urllib.request

    requisicao = urllib.request.Request(
        url,
        headers={
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                          "(KHTML, like Gecko) Chrome/120.0 Safari/537.36",
            "Accept": "*/*",
        },
    )
    try:
        with urllib.request.urlopen(requisicao, timeout=timeout) as resposta:
            return resposta.read()
    except (urllib.error.URLError, OSError, ValueError) as erro:
        raise ErroDeConsulta(str(erro)) from erro


# Séries do SGS (Banco Central): taxa diária, em % ao dia
CODIGOS_SGS = {"CDI": 12, "SELIC": 11}


def buscar_serie_bcb(indice: str, data_inicial: TipoData,
                     data_final: Optional[TipoData] = None, timeout: int = 20) -> pd.Series:
    """Consulta no Banco Central (SGS) as taxas diárias realizadas do CDI ou
    da Selic e devolve o número-índice acumulado, no mesmo formato de
    `ler_serie_indice()`.

    Séries usadas: 12 (CDI, % ao dia) e 11 (Selic, % ao dia).
    """
    import json

    chave = _sem_acento(indice).upper()
    if chave not in CODIGOS_SGS:
        raise ValueError("A consulta automática ao Banco Central está disponível para CDI e Selic.")
    inicio = pd.Timestamp(data_inicial).normalize()
    fim = pd.Timestamp(data_final).normalize() if data_final is not None else pd.Timestamp.today().normalize()
    if fim < inicio:
        raise ErroDeConsulta("Não há taxas realizadas: a data inicial é posterior a hoje.")
    url = (
        f"https://api.bcb.gov.br/dados/serie/bcdata.sgs.{CODIGOS_SGS[chave]}/dados"
        f"?formato=json&dataInicial={inicio:%d/%m/%Y}&dataFinal={fim:%d/%m/%Y}"
    )
    try:
        registros = json.loads(_baixar(url, timeout).decode("utf-8"))
        datas = [pd.to_datetime(r["data"], format="%d/%m/%Y") for r in registros]
        fatores = [1.0 + float(str(r["valor"]).replace(",", ".")) / 100.0 for r in registros]
    except ErroDeConsulta:
        raise
    except Exception as erro:
        raise ErroDeConsulta(f"resposta inesperada do Banco Central ({erro})") from erro
    if not datas:
        raise ErroDeConsulta("o Banco Central não devolveu taxas para o período.")
    return _indice_de_fatores(datas, fatores, f"taxas diárias do Banco Central (SGS {CODIGOS_SGS[chave]})")


def _linhas_de_tabela_html(html: str) -> list:
    """Extrai as linhas (listas de textos das células) das tabelas de um HTML."""
    from html.parser import HTMLParser

    class _Leitor(HTMLParser):
        def __init__(self):
            super().__init__()
            self.linhas, self.linha, self.celula = [], None, None

        def handle_starttag(self, tag, attrs):
            if tag == "tr":
                self.linha = []
            elif tag in ("td", "th") and self.linha is not None:
                self.celula = []

        def handle_endtag(self, tag):
            if tag in ("td", "th") and self.celula is not None and self.linha is not None:
                self.linha.append(" ".join("".join(self.celula).split()))
                self.celula = None
            elif tag == "tr" and self.linha is not None:
                self.linhas.append(self.linha)
                self.linha = None

        def handle_data(self, data):
            if self.celula is not None:
                self.celula.append(data)

    leitor = _Leitor()
    leitor.feed(html)
    return leitor.linhas


URL_CURVA_B3 = (
    "https://www2.bmf.com.br/pages/portal/bmfbovespa/lumis/"
    "lum-taxas-referenciais-bmf-ptBR.asp?Data={data:%d/%m/%Y}&Data1={data:%Y%m%d}&slcTaxa=PRE"
)


def _curva_do_html_b3(html: str, referencia: pd.Timestamp) -> Optional[pd.DataFrame]:
    """Monta a curva a partir da página de taxas referenciais da B3, cuja
    tabela traz: dias corridos | taxa base 252 | taxa base 360."""
    import re

    vertices = []
    for linha in _linhas_de_tabela_html(html):
        if len(linha) >= 2 and re.fullmatch(r"\d+", linha[0]) and re.fullmatch(r"-?\d+,\d+", linha[1]):
            vertices.append((int(linha[0]), float(linha[1].replace(",", "."))))
    if len(vertices) < 5:
        return None
    try:
        from bizdays import Calendar
    except ImportError as exc:  # pragma: no cover
        raise ImportError("A leitura da curva da B3 exige a biblioteca `bizdays`.") from exc
    calendario = Calendar.load(name="ANBIMA")
    inicio = pd.Timestamp(calendario.following(referencia.date())).date()
    linhas = []
    for dias_corridos, taxa in vertices:
        vencimento = pd.Timestamp(
            calendario.following((referencia + pd.Timedelta(days=dias_corridos)).date())
        ).date()
        linhas.append((int(calendario.bizdays(inicio, vencimento)), dias_corridos, taxa))
    curva = pd.DataFrame(linhas, columns=COLUNAS_CURVA)
    curva = curva[curva["Dias Úteis"] > 0].drop_duplicates("Dias Úteis").reset_index(drop=True)
    curva.attrs["data_referencia"] = referencia
    curva.attrs["origem"] = "consulta ao site da B3"
    return curva


def buscar_curva_b3(data_referencia: Optional[TipoData] = None, dias_anteriores: int = 7,
                    timeout: int = 20) -> pd.DataFrame:
    """Tenta consultar no site da B3 a curva DI x pré (taxas referenciais).

    Sem `data_referencia`, procura a curva mais recente, voltando até
    `dias_anteriores` dias a partir de hoje (a curva do dia só é publicada
    após o fechamento do pregão). Se o site não responder ou o formato da
    página tiver mudado, levanta `ErroDeConsulta`; nesse caso, baixe o arquivo
    no site da B3 e use `ler_curva_b3()`.
    """
    if data_referencia is not None:
        candidatas = [pd.Timestamp(data_referencia).normalize()]
    else:
        hoje = pd.Timestamp.today().normalize()
        candidatas = [hoje - pd.Timedelta(days=k) for k in range(dias_anteriores + 1)]
        candidatas = [d for d in candidatas if d.weekday() < 5]
    ultimo_erro = "a página não trouxe a tabela de taxas"
    for referencia in candidatas:
        try:
            html = _baixar(URL_CURVA_B3.format(data=referencia), timeout).decode("latin-1")
        except ErroDeConsulta as erro:
            ultimo_erro = str(erro)
            break  # site inacessível: não adianta tentar outras datas
        curva = _curva_do_html_b3(html, referencia)
        if curva is not None:
            return curva
    raise ErroDeConsulta(f"não foi possível obter a curva no site da B3 ({ultimo_erro}).")


def ler_serie_indice(arquivo, nome_arquivo: Optional[str] = None) -> pd.Series:
    """Lê a série histórica de um índice e devolve um NÚMERO-ÍNDICE por data.

    Formatos reconhecidos (pelo cabeçalho das colunas):
    * coluna "Fator diário" (ex.: arquivo de Selic/CDI diários do Banco
      Central): os fatores diários são acumulados em um número-índice;
    * coluna "Taxa (% a.a.)" sem fator: cada taxa anual é convertida em fator
      diário na base 252 dias úteis e acumulada;
    * caso contrário: a 2ª coluna é lida como número-índice.

    Em séries de fatores ou taxas, o fator de um dia útil corrige o saldo
    desse dia até o dia útil seguinte. O texto em `serie.attrs["origem"]`
    descreve como o arquivo foi interpretado.
    """
    dados, cabecalhos = _ler_tabela(arquivo, nome_arquivo)
    if dados.shape[1] < 2:
        raise ValueError("O arquivo da série precisa de ao menos duas colunas (data e valor).")
    datas = pd.to_datetime(dados[0])

    coluna_fator = next((i for i, c in enumerate(cabecalhos) if i > 0 and "fator" in c), None)
    coluna_taxa = next(
        (i for i, c in enumerate(cabecalhos)
         if i > 0 and "taxa" in c and ("a.a" in c or "ano" in c or "anual" in c)),
        None,
    )

    if coluna_fator is not None or coluna_taxa is not None:
        if coluna_fator is not None:
            fatores = _para_numero(dados[coluna_fator])
            origem = "fatores diários acumulados"
        else:
            fatores = (1.0 + _para_numero(dados[coluna_taxa]) / 100.0) ** (1.0 / 252.0)
            origem = "taxas anuais convertidas em fator diário (base 252) e acumuladas"
        return _indice_de_fatores(datas, fatores, origem)

    valores = _para_numero(dados[1])
    serie = pd.Series(valores.to_numpy(), index=pd.DatetimeIndex(datas)).dropna().sort_index()
    serie = serie[~serie.index.duplicated()]
    if serie.empty:
        raise ValueError("O arquivo da série histórica não tem valores numéricos na 2ª coluna.")
    if (serie <= 0).any():
        raise ValueError(
            "A 2ª coluna deve trazer o número-índice (valores positivos). Se o "
            "arquivo traz taxas ou fatores, o cabeçalho da coluna deve conter "
            "'Fator diário' ou 'Taxa (% a.a.)'."
        )
    serie.attrs["origem"] = "número-índice"
    serie.attrs["periodo"] = (serie.index[0], serie.index[-1])
    return serie


def ler_cronograma(arquivo, nome_arquivo: Optional[str] = None) -> pd.DataFrame:
    """Lê um cronograma de pagamento de um arquivo Excel ou CSV.

    1ª coluna: data de vencimento (dia/mês/ano).
    2ª coluna (opcional): amortização da parcela, em R$ ou em % do principal.
    Devolve um DataFrame com as colunas "Data" e, se houver, "Amortização".
    """
    dados, _ = _ler_tabela(arquivo, nome_arquivo)
    cronograma = pd.DataFrame({"Data": pd.to_datetime(dados[0]).to_numpy()})
    if dados.shape[1] >= 2:
        cronograma["Amortização"] = _para_numero(dados[1]).fillna(0.0).to_numpy()
    return cronograma


def _ler_cronograma_de_arquivo() -> pd.DataFrame:
    while True:
        caminho = Path(_perguntar("   Caminho do arquivo (.xlsx ou .csv)").strip('"')).expanduser()
        try:
            cronograma = ler_cronograma(caminho)
            print(f"   -> {len(cronograma)} parcelas lidas, de "
                  f"{cronograma['Data'].min():%d/%m/%Y} a {cronograma['Data'].max():%d/%m/%Y}.")
            return cronograma
        except Exception as erro:
            print(f"   -> Não consegui ler o arquivo ({erro}).")
            print("      A 1ª coluna deve ter as datas e a 2ª (opcional) a amortização.")
            if not _ler_sim_nao("   Tentar outro arquivo?", "S"):
                raise ValueError("Cronograma não informado.") from erro


def _digitar_cronograma(com_amortizacao: bool, rotulo_valor: str) -> pd.DataFrame:
    """Digitação do cronograma parcela a parcela (Enter em branco encerra)."""
    print("   Digite as parcelas uma a uma. Deixe a data em branco para encerrar.")
    datas, valores = [], []
    while True:
        numero = len(datas) + 1
        resposta = _perguntar(f"   Parcela {numero} - data (dia/mês/ano)", "")
        if not resposta:
            if datas:
                break
            print("   -> Informe ao menos uma parcela.")
            continue
        data = None
        for formato in ("%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y", "%d/%m/%y"):
            try:
                data = datetime.strptime(resposta, formato)
                break
            except ValueError:
                pass
        if data is None:
            print("   -> Data inválida. Use o formato dia/mês/ano, por exemplo: 15/01/2026")
            continue
        if datas and data <= datas[-1]:
            print("   -> A data deve ser posterior à da parcela anterior.")
            continue
        datas.append(data)
        if com_amortizacao:
            valores.append(_ler_numero(
                f"   Parcela {numero} - amortização {rotulo_valor} (0 se não houver)",
                padrao="0", minimo=0, monetario=rotulo_valor == "em R$",
            ))
    cronograma = pd.DataFrame({"Data": datas})
    if com_amortizacao:
        cronograma["Amortização"] = valores
    return cronograma


def _ler_serie_de_arquivo() -> pd.Series:
    """Pergunta o caminho do arquivo da série histórica e o lê."""
    while True:
        caminho = Path(_perguntar("   Caminho do arquivo (.xlsx ou .csv)").strip('"')).expanduser()
        try:
            serie = ler_serie_indice(caminho)
            inicio, fim = serie.attrs["periodo"]
            print(f"   -> Série lida de {inicio:%d/%m/%Y} a {fim:%d/%m/%Y} "
                  f"({serie.attrs['origem']}).")
            return serie
        except Exception as erro:  # arquivo inexistente, formato inesperado etc.
            print(f"   -> Não consegui ler o arquivo ({erro}).")
            print("      A 1ª coluna deve ter as datas; a 2ª, o número-índice (ou use o")
            print("      arquivo de fatores diários do Banco Central).")
            if not _ler_sim_nao("   Tentar outro arquivo?", "S"):
                raise ValueError("Série histórica não informada.") from erro


def _ler_curva_de_arquivo() -> pd.DataFrame:
    """Pergunta o caminho do arquivo da curva DI x pré da B3 e o lê."""
    while True:
        caminho = Path(_perguntar("   Caminho do arquivo da curva (.csv da B3)").strip('"')).expanduser()
        try:
            curva = ler_curva_b3(caminho)
            if curva.attrs.get("data_referencia") is None:
                curva.attrs["data_referencia"] = pd.Timestamp(
                    _ler_data("   Data de referência da curva (dia/mês/ano)")
                )
            print(f"   -> Curva de {curva.attrs['data_referencia']:%d/%m/%Y} lida, "
                  f"com {len(curva)} vértices.")
            return curva
        except Exception as erro:
            print(f"   -> Não consegui ler o arquivo ({erro}).")
            if not _ler_sim_nao("   Tentar outro arquivo?", "S"):
                raise ValueError("Curva de juros não informada.") from erro


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
        def perguntar_taxa():
            p["taxa_indice_projetada"] = _ler_numero(
                f"   Taxa anual esperada do {nome.upper()}, em % (ex.: 4,5)", minimo=-99.99
            ) / 100.0

        forma = _ler_opcao(
            "   Como informar a variação do índice?",
            [
                ("taxa", "Taxa anual esperada (projeção fixa)"),
                ("arquivo", "Série histórica em arquivo Excel/CSV (ex.: taxas diárias do Banco Central)"),
                ("curva", "Curva DI x pré da B3 em arquivo (projeção pela curva futura)"),
            ],
        )
        if forma == "arquivo":
            p["indice"] = _ler_serie_de_arquivo()
            p["_nome_indice"] = nome
            projecao = _ler_opcao(
                "   Como projetar o índice após o fim da série?",
                [
                    ("taxa", "Taxa anual fixa"),
                    ("curva", "Curva DI x pré da B3 (arquivo)"),
                    ("nenhuma", "Não projetar (a série cobre todo o contrato)"),
                ],
            )
            if projecao == "taxa":
                perguntar_taxa()
            elif projecao == "curva":
                p["curva_projecao"] = _ler_curva_de_arquivo()
        elif forma == "curva":
            p["curva_projecao"] = _ler_curva_de_arquivo()
        else:
            perguntar_taxa()
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

    print("\n5) VENCIMENTOS DAS PARCELAS")
    forma_venc = _ler_opcao(
        "   Como são definidos os vencimentos?",
        [
            ("periodicidade", "Periodicidade fixa (mensal, semestral...)"),
            ("arquivo", "Cronograma próprio, em arquivo Excel/CSV"),
            ("digitar", "Cronograma próprio, digitado parcela a parcela"),
        ],
    )
    if forma_venc == "periodicidade":
        p["periodicidade"] = _ler_opcao(
            "   Frequência das parcelas:",
            [(nome_p, nome_p.capitalize()) for nome_p in CalculadoraEmprestimo._MESES_POR_PERIODO],
        )

    print("\n6) SISTEMA DE AMORTIZAÇÃO")
    opcoes_sistema = [
        ("SAC", "SAC - amortização constante, parcelas decrescentes"),
        ("PRICE", "PRICE - parcelas constantes (Tabela Price / Francesa)"),
    ]
    if forma_venc != "periodicidade":
        opcoes_sistema += [
            ("valor", "Amortização de cada parcela informada no cronograma, em R$"),
            ("percentual", "Amortização de cada parcela informada no cronograma, em % do principal"),
        ]
    sistema = _ler_opcao("   Sistema:", opcoes_sistema)
    informada = sistema in ("valor", "percentual")
    if informada:
        p["cronograma_tipo"] = sistema
    else:
        p["sistema_amortizacao"] = sistema

    if forma_venc == "periodicidade":
        print("\n7) PRAZO TOTAL")
        p["prazo_total"] = _ler_inteiro(
            "   Número total de parcelas, já contando a carência (ex.: 60)", minimo=1
        )
        n_parcelas = p["prazo_total"]
    else:
        print("\n7) CRONOGRAMA DE PAGAMENTO")
        if forma_venc == "arquivo":
            print("   No arquivo, a 1ª coluna traz as datas de vencimento"
                  + (" e a 2ª a amortização de cada parcela." if informada else "."))
            cronograma = _ler_cronograma_de_arquivo()
            if informada and "Amortização" not in cronograma.columns:
                raise ValueError(
                    "O arquivo do cronograma não tem a 2ª coluna, com a amortização de cada parcela."
                )
        else:
            cronograma = _digitar_cronograma(
                informada, "em R$" if sistema == "valor" else "em %"
            )
        if not informada:
            cronograma = cronograma[["Data"]]
        p["cronograma"] = cronograma
        n_parcelas = len(cronograma)

    print("\n8) CARÊNCIA")
    if informada:
        amortizacoes = p["cronograma"]["Amortização"].to_numpy()
        carencia = int(np.argmax(amortizacoes > 0)) if (amortizacoes > 0).any() else 0
        print(f"   Pelo cronograma, há {carencia} parcela(s) inicial(is) sem amortização.")
    else:
        carencia = p["carencia"] = _ler_inteiro(
            "   Períodos sem amortização do principal (0 se não houver)",
            padrao="0", minimo=0, maximo=max(n_parcelas - 1, 0),
        )
    if carencia > 0:
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
    if "streamlit" not in sys.modules:  # no terminal o Streamlit não está carregado
        return False
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
