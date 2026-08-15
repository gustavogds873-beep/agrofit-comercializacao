#!/usr/bin/env python
# -*- coding: utf-8 -*-
# =============================================================================
#  ANALISE INTEGRADA AGROFIT x COMERCIALIZACAO DE AGROTOXICOS
# =============================================================================
#
#  Script completo, executavel diretamente (Windows / Linux / macOS), que
#  cruza a base de produtos formulados do Agrofit com as bases anuais de
#  comercializacao de agrotoxicos, gerando uma estrutura analitica (star
#  schema) com granularidades separadas para evitar duplicacao/multiplicacao
#  artificial dos volumes comercializados.
#
#  Abas geradas (duas perspectivas: produto consolidado x IA separada):
#    00_Resumo              - indicadores e rankings
#    01_Produtos            - dimensao: 1 linha por registro MAPA
#    02_Agrofit_Uso         - dimensao: Registro x Cultura x Praga
#    03_Comerc_Produto      - fato PRODUTO: 1 linha por Registro x Ano
#                             (somente volumes de produto formulado / PF)
#    04_Comerc_IA           - fato IA: 1 linha por Registro x Ano x IA
#                             (ponderacao individual de cada ingrediente ativo)
#    05_Comerc_UF           - fato PRODUTO: vendas por UF (Registro x Ano x UF)
#    06_IA_Anual            - agregacao IA: volumes por ingrediente ativo e ano
#    07_IA_UF               - agregacao IA: vendas estimadas de IA por UF
#    08_Registro_Resumo     - resumo por registro (venda total PF no periodo)
#    09_IA_Resumo           - resumo por IA (a partir da comercializacao por IA)
#    10_Evol_Produtos_Anual - tabela WIDE: 1 linha por produto (Registro),
#                             1 coluna por ANO com o volume vendido (PF) e
#                             coluna TOTAL_PERIODO com a soma de todos os anos.
#                             Pronta para graficos (tabela dinamica/grafico de
#                             linhas) de evolucao anual de uso por produto.
#    11_Evol_IA_Anual       - tabela WIDE: 1 linha por ingrediente ativo,
#                             1 coluna por ANO com a venda ponderada/estimada
#                             de IA e coluna TOTAL_PERIODO com a soma de todos
#                             os anos. Pronta para graficos de evolucao anual
#                             de uso por IA.
#    12_Agrofit_Sem_Vendas  - registros Agrofit sem comercializacao
#    13_Vendas_Sem_Agrofit  - comercializacao sem match no Agrofit
#    14_Validacao           - metricas de correspondencia
#
#  Multi-IA na Tabela Bruta: o mesmo produto formulado aparece em N linhas
#  (uma por ingrediente ativo), repetindo os volumes de PF. O script trata:
#    - perspectiva PRODUTO -> first/max nos volumes PF (nao soma)
#    - perspectiva IA     -> mantem cada linha e aplica PONDERACAO_IA da linha
#
#  Volumes de comercializacao:
#    - ORIGINAIS (auditoria): VENDAS_TOTAIS, VENDAS_CLIENTE, VENDAS_INDUSTRIA,
#      PRODUCAO_NACIONAL, IMPORTACAO, EXPORTACAO, ESTOQUE_* — valores da planilha,
#      sem ponderacao pelo teor de IA.
#    - PONDERADOS (analise de IA): VENDA_PONDERADA_IA / VENDA_ESTIMADA_IA =
#      volume original x PONDERACAO_IA — quantidade efetiva de ingrediente ativo.
#  As colunas originais nunca sao sobrescritas pelo calculo ponderado.
#
#  Uso:
#      python analise_integrada_agrofit_comercializacao.py
#
#  Estrutura de pastas esperada (criada automaticamente se nao existir):
#      dados/      -> coloque aqui agrofitprodutosformulados.csv e os
#                      arquivos anuais de comercializacao (.xlsx ou .csv)
#      saida/      -> aqui sera gravado Relatorio_Comercializacao_Agrofit.xlsx
#      logs/       -> aqui sera gravado analise_comercializacao.log
#
#  Consulte o README.md para detalhes de instalacao, execucao e o
#  significado de cada aba gerada.
# =============================================================================

from __future__ import annotations

import os
import sys
import re
import glob
import time
import logging
import unicodedata
import warnings
import traceback
from pathlib import Path
from datetime import datetime
from typing import Optional, List, Dict, Tuple, Any

# -----------------------------------------------------------------------
# Garantia de dependencias (instala automaticamente se faltar)
# -----------------------------------------------------------------------
def _ensure_packages() -> None:
    required = {
        "pandas": "pandas",
        "numpy": "numpy",
        "openpyxl": "openpyxl",
    }
    missing = []
    for module_name, pip_name in required.items():
        try:
            __import__(module_name)
        except ImportError:
            missing.append(pip_name)
    if missing:
        print(f"Instalando pacotes faltantes: {', '.join(missing)} ...")
        import subprocess
        subprocess.check_call(
            [sys.executable, "-m", "pip", "install", "-q"] + missing
        )
        print("Pacotes instalados.\n")


_ensure_packages()

import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import LineChart, BarChart, Reference
from openpyxl.worksheet.table import Table, TableStyleInfo

warnings.filterwarnings("ignore", category=UserWarning)
warnings.filterwarnings("ignore", category=FutureWarning)

pd.set_option("mode.chained_assignment", None)

# =============================================================================
# CONFIGURACAO GERAL
# =============================================================================

SCRIPT_DIR = Path(__file__).resolve().parent
# Pastas padrao (ao lado do script). Podem ser sobrescritas por variaveis de
# ambiente — util no Google Colab:
#   AGROFIT_DADOS_DIR=/content/dados
#   AGROFIT_SAIDA_DIR=/content/resultados
#   AGROFIT_LOGS_DIR=/content/logs
DADOS_DIR = Path(os.environ["AGROFIT_DADOS_DIR"]) if os.environ.get("AGROFIT_DADOS_DIR") else SCRIPT_DIR / "dados"
SAIDA_DIR = Path(os.environ["AGROFIT_SAIDA_DIR"]) if os.environ.get("AGROFIT_SAIDA_DIR") else SCRIPT_DIR / "saida"
LOGS_DIR = Path(os.environ["AGROFIT_LOGS_DIR"]) if os.environ.get("AGROFIT_LOGS_DIR") else SCRIPT_DIR / "logs"

ARQUIVO_AGROFIT_PADRAO = "agrofitprodutosformulados.csv"
ARQUIVO_SAIDA = "Relatorio_Comercializacao_Agrofit.xlsx"
ARQUIVO_LOG = "analise_comercializacao.log"

# Palavras-chave usadas para localizar automaticamente o arquivo do Agrofit
# dentro da pasta dados/, caso o nome exato nao seja encontrado.
PISTAS_NOME_AGROFIT = ["agrofit"]

# As 27 siglas de UF esperadas nas bases de comercializacao
UFS = [
    "RO", "AC", "AM", "RR", "PA", "AP", "TO", "MA", "PI", "CE", "RN", "PB",
    "PE", "AL", "SE", "BA", "MG", "ES", "RJ", "SP", "PR", "SC", "RS", "MS",
    "MT", "GO", "DF",
]

REGIAO_POR_UF = {
    "RO": "Norte", "AC": "Norte", "AM": "Norte", "RR": "Norte", "PA": "Norte",
    "AP": "Norte", "TO": "Norte",
    "MA": "Nordeste", "PI": "Nordeste", "CE": "Nordeste", "RN": "Nordeste",
    "PB": "Nordeste", "PE": "Nordeste", "AL": "Nordeste", "SE": "Nordeste",
    "BA": "Nordeste",
    "MG": "Sudeste", "ES": "Sudeste", "RJ": "Sudeste", "SP": "Sudeste",
    "PR": "Sul", "SC": "Sul", "RS": "Sul",
    "MS": "Centro-Oeste", "MT": "Centro-Oeste", "GO": "Centro-Oeste",
    "DF": "Centro-Oeste",
}

# Unidades reconhecidas como "de massa/volume" (compativeis para o calculo
# de venda estimada de IA por ponderacao percentual). Unidades de contagem
# (ex.: "UN", "PC") sao consideradas incompativeis para esse calculo.
UNIDADES_COMPATIVEIS = {
    "KG", "L", "T", "TON", "TONELADA", "TONELADAS", "LITRO", "LITROS",
    "QUILOGRAMA", "QUILOGRAMAS", "G", "GRAMA", "GRAMAS", "ML", "M3",
}

# Na Tabela Bruta, a coluna "Unidade" refere-se a CONCENTRACAO (g/l, g/kg),
# nao a unidade do volume comercializado. Essas unidades nao devem bloquear
# o calculo de volumes ponderados.
UNIDADES_CONCENTRACAO = {
    "G_L", "G/L", "GL", "G_KG", "G/KG", "GKG", "MG_L", "MG/L",
    "PERCENT", "PERCENTUAL", "%", "PPM", "PPB",
}

LOGGER = logging.getLogger("agrofit_comercializacao")


# =============================================================================
# LOGGING E MENSAGENS DE PROGRESSO
# =============================================================================

class Cores:
    HEADER = "\033[95m"
    AZUL = "\033[94m"
    VERDE = "\033[92m"
    AMARELO = "\033[93m"
    VERMELHO = "\033[91m"
    FIM = "\033[0m"
    NEGRITO = "\033[1m"


if sys.platform == "win32":
    try:
        import ctypes
        kernel32 = ctypes.windll.kernel32
        kernel32.SetConsoleMode(kernel32.GetStdHandle(-11), 7)
    except Exception:
        for _attr in ["HEADER", "AZUL", "VERDE", "AMARELO", "VERMELHO", "FIM", "NEGRITO"]:
            setattr(Cores, _attr, "")


def configurar_logging() -> None:
    LOGS_DIR.mkdir(parents=True, exist_ok=True)
    caminho_log = LOGS_DIR / ARQUIVO_LOG
    LOGGER.setLevel(logging.DEBUG)
    LOGGER.handlers.clear()

    fh = logging.FileHandler(caminho_log, mode="a", encoding="utf-8")
    fh.setLevel(logging.DEBUG)
    fh.setFormatter(logging.Formatter(
        "%(asctime)s | %(levelname)-8s | %(message)s", "%Y-%m-%d %H:%M:%S"
    ))
    LOGGER.addHandler(fh)


def titulo(texto: str) -> None:
    linha = "=" * 78
    print(f"\n{Cores.HEADER}{Cores.NEGRITO}{linha}{Cores.FIM}")
    print(f"{Cores.NEGRITO}{texto.center(78)}{Cores.FIM}")
    print(f"{Cores.HEADER}{linha}{Cores.FIM}\n")
    LOGGER.info("=== %s ===", texto)


def ok(msg: str) -> None:
    print(f"{Cores.VERDE}[OK] {msg}{Cores.FIM}")
    LOGGER.info("OK: %s", msg)


def aviso(msg: str) -> None:
    print(f"{Cores.AMARELO}[AVISO] {msg}{Cores.FIM}")
    LOGGER.warning(msg)


def erro(msg: str) -> None:
    print(f"{Cores.VERMELHO}[ERRO] {msg}{Cores.FIM}")
    LOGGER.error(msg)


def passo(msg: str) -> None:
    print(f"{Cores.AZUL}-> {msg}{Cores.FIM}")
    LOGGER.info("PASSO: %s", msg)


def tempo_legivel(segundos: float) -> str:
    if segundos < 60:
        return f"{segundos:.1f}s"
    if segundos < 3600:
        return f"{segundos / 60:.1f} min"
    return f"{segundos / 3600:.2f} h"


class Cronometro:
    """Cronometro simples para medir e logar a duracao de cada etapa."""

    def __init__(self, nome: str):
        self.nome = nome
        self.inicio = None

    def __enter__(self):
        self.inicio = time.time()
        passo(f"Iniciando: {self.nome}")
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        duracao = time.time() - self.inicio
        if exc_type is None:
            ok(f"Concluido: {self.nome} ({tempo_legivel(duracao)})")
        else:
            erro(f"Falhou: {self.nome} ({tempo_legivel(duracao)}) -> {exc_val}")
        LOGGER.info("TEMPO | %s | %.2fs", self.nome, duracao)
        return False


# =============================================================================
# NORMALIZACAO DE TEXTO E NOMES DE COLUNA
# =============================================================================

def remover_acentos(texto: str) -> str:
    if texto is None:
        return ""
    nfkd = unicodedata.normalize("NFKD", str(texto))
    return "".join(c for c in nfkd if not unicodedata.combining(c))


def normalizar_nome_coluna(nome: str) -> str:
    """Padrao unico de nomes de coluna: MAIUSCULAS, SEM ACENTOS, SEM ESPACOS.

    Trata cabecalhos tipicos da Tabela Bruta IBAMA/MAPA, incluindo tags HTML
    de quebra de linha (ex.: 'Estoque<br>Inicial', 'Vendas a<br>Industria').
    """
    if nome is None:
        return ""
    s = str(nome)
    # Remove tags HTML / quebras embutidas nos cabecalhos do Excel
    s = re.sub(r"<br\s*/?>", " ", s, flags=re.IGNORECASE)
    s = s.replace("\n", " ").replace("\r", " ")
    s = remover_acentos(s).strip().upper()
    s = re.sub(r"[^\w]+", "_", s)
    s = re.sub(r"_+", "_", s).strip("_")
    return s


def normalizar_texto_valor(valor: Any) -> Any:
    """Normaliza um valor de texto para comparacoes (chaves de cruzamento),
    mas NAO altera os dados exibidos no relatorio final."""
    if pd.isna(valor):
        return valor
    s = remover_acentos(str(valor)).strip().upper()
    s = re.sub(r"\s+", " ", s)
    return s


def extrair_chave_registro(valor: Any) -> str:
    """Normaliza NR_REGISTRO para uma chave comparavel: remove prefixos
    (MAPA/IBAMA), pontuacao e zeros a esquerda."""
    if pd.isna(valor):
        return ""
    s = str(valor).strip().upper()
    s = re.sub(r"^(MAPA|IBAMA|N[R°º]?\.?)[\s\-/:]*", "", s)
    s = re.sub(r"\D", "", s)
    return s.lstrip("0") or ("0" if s else "")


# =============================================================================
# DICIONARIOS DE ALIASES DE COLUNA (mapeiam nomes "crus" -> nomes padronizados)
# =============================================================================
# As chaves sao os nomes FINAIS padronizados (MAIUSCULAS_SEM_ACENTO_SEM_ESPACO).
# Os valores sao listas de variantes conhecidas, ja normalizadas pela mesma
# funcao normalizar_nome_coluna(), que podem aparecer nos arquivos de origem.
# Se um arquivo novo trouxer uma variante nao mapeada, o programa registra um
# aviso no log e mantem o nome original (normalizado) para nao perder dados.

ALIASES_AGROFIT: Dict[str, List[str]] = {
    "NR_REGISTRO": ["NR_REGISTRO", "NUMERO_REGISTRO", "REGISTRO", "N_REGISTRO",
                     "REGISTRO_MAPA", "NR_REGISTRO_MAPA", "NO_REGISTRO",
                     "NUMERO_DO_REGISTRO", "NO_DE_REGISTRO", "N_DE_REGISTRO"],
    "MARCA_COMERCIAL": ["MARCA_COMERCIAL", "MARCA", "NOME_COMERCIAL", "PRODUTO"],
    "FORMULACAO": ["FORMULACAO", "TIPO_FORMULACAO"],
    "INGREDIENTE_ATIVO": ["INGREDIENTE_ATIVO", "INGREDIENTES_ATIVOS", "IA",
                           "PRINCIPIO_ATIVO"],
    "TITULAR_DE_REGISTRO": ["TITULAR_DE_REGISTRO", "TITULAR_REGISTRO", "TITULAR",
                             "EMPRESA_TITULAR", "DETENTOR_DO_REGISTRO"],
    "CLASSE": ["CLASSE", "CLASSE_DO_PRODUTO", "CLASSE_AGRONOMICA"],
    "MODO_DE_ACAO": ["MODO_DE_ACAO", "MODO_ACAO", "MECANISMO_DE_ACAO"],
    "CULTURA": ["CULTURA", "CULTURAS"],
    "PRAGA_NOME_CIENTIFICO": ["PRAGA_NOME_CIENTIFICO", "NOME_CIENTIFICO_PRAGA",
                               "NOME_CIENTIFICO"],
    "PRAGA_NOME_COMUM": ["PRAGA_NOME_COMUM", "NOME_COMUM_PRAGA", "NOME_COMUM"],
    "EMPRESA_PAIS_TIPO": ["EMPRESA_PAIS_TIPO", "PAIS_TITULAR", "PAIS", "TIPO_EMPRESA"],
    "CLASSE_TOXICOLOGICA": ["CLASSE_TOXICOLOGICA", "CLASSIFICACAO_TOXICOLOGICA",
                             "CLASSE_TOXICOLOGICA_ANVISA"],
    "CLASSE_AMBIENTAL": ["CLASSE_AMBIENTAL", "CLASSIFICACAO_AMBIENTAL",
                          "CLASSE_AMBIENTAL_IBAMA"],
    "ORGANICOS": ["ORGANICOS", "USO_EM_AGRICULTURA_ORGANICA", "AGRICULTURA_ORGANICA"],
    "SITUACAO": ["SITUACAO", "SITUACAO_DO_PRODUTO", "STATUS"],
}

ALIASES_COMERC: Dict[str, List[str]] = {
    "NR_REGISTRO": ["NR_REGISTRO", "NUMERO_REGISTRO", "REGISTRO", "N_REGISTRO",
                     "REGISTRO_MAPA", "REGISTRO_DO_PRODUTO", "N_REGISTRO_MAPA",
                     "NO_REGISTRO", "NUMERO_DO_REGISTRO", "NO_DE_REGISTRO", "N_DE_REGISTRO"],
    "MARCA_COMERCIAL": ["MARCA_COMERCIAL", "PRODUTO_COMERCIAL", "MARCA",
                         "NOME_DO_PRODUTO", "MARCA_COMERCIAL_VENDAS"],
    "INGREDIENTE_ATIVO": ["INGREDIENTE_ATIVO", "INGREDIENTES_ATIVOS", "IA",
                           "PRINCIPIO_ATIVO", "INGREDIENTE_ATIVO_VENDAS"],
    "DETENTOR": ["DETENTOR", "DETENTOR_DO_REGISTRO", "TITULAR_DE_REGISTRO",
                  "TITULAR", "EMPRESA"],
    "PRODUTO_TECNICO": ["PRODUTO_TECNICO", "TIPO_DE_PRODUTO", "CATEGORIA"],
    "CONCENTRACAO": ["CONCENTRACAO", "CONCENTRACAO_DO_IA", "TEOR"],
    "UNIDADE": ["UNIDADE", "UNIDADE_DE_MEDIDA", "UM", "UNID", "UNIDADE_CONCENTRACAO"],
    "CLASSES_USO": ["CLASSES_USO", "CLASSE_DE_USO", "CLASSES_DE_USO"],
    "PONDERACAO_IA": [
        "PONDERACAO_IA", "PONDERACAO", "PONDERACAO_PERCENT_IA_NO_PF",
        "PONDERACAO_DO_IA", "PERCENTUAL_IA", "PONDERACAO_IA_NO_PF",
        "PONDERACAO_IA_NO_PF", "PONDERACAO_PERCENT_IA_NO_PF",
    ],
    # Cabecalhos da Tabela Bruta (apos limpar <br>): Estoque Inicial, Producao Nacional,
    # Importacao, Exportacao, Vendas a Industria, Vendas a Cliente, Estoque Final
    "ESTOQUE_INICIAL": [
        "ESTOQUE_INICIAL", "ESTOQUE_INICIAL_PF", "ESTOQUE_INI",
    ],
    "PRODUCAO_NACIONAL": [
        "PRODUCAO_NACIONAL", "PRODUCAO", "PRODUCAO_NACIONAL_PF",
    ],
    "IMPORTACAO": [
        "IMPORTACAO", "IMPORTACAO_PF", "IMPORTAC",
    ],
    "EXPORTACAO": [
        "EXPORTACAO", "EXPORTACAO_PF", "EXPORTAC",
    ],
    "VENDAS_INDUSTRIA": [
        "VENDAS_INDUSTRIA", "VENDA_INDUSTRIA", "VENDAS_A_INDUSTRIA",
        "VENDAS_INDUSTRIA_PF", "VENDAS_A_INDUSTRIA_PF", "VENDAS_INDUST",
    ],
    "VENDAS_CLIENTE": [
        "VENDAS_CLIENTE", "VENDA_CLIENTE", "VENDAS_A_CLIENTE",
        "VENDAS_CLIENTE_PF", "VENDAS_A_CLIENTE_PF", "VENDAS_CLIEN",
    ],
    "VENDAS_TOTAIS": [
        "VENDAS_TOTAIS", "VENDA_TOTAL", "TOTAL_DE_VENDAS", "VENDAS",
        "VENDAS_PF_TOTAL", "VENDAS_TOTAIS_PF",
    ],
    "ESTOQUE_FINAL": [
        "ESTOQUE_FINAL", "ESTOQUE_FINAL_PF", "ESTOQUE_FIN",
    ],
    "ANO": ["ANO", "ANO_BASE", "ANO_REFERENCIA", "ANO_COMERCIALIZACAO"],
}


def padronizar_colunas(
    df: pd.DataFrame, aliases: Dict[str, List[str]], origem: str = ""
) -> pd.DataFrame:
    """Renomeia as colunas de df para o padrao definido em `aliases`.
    Colunas nao mapeadas sao mantidas (apenas normalizadas) e reportadas
    no log para acompanhamento."""
    df = df.copy()
    df.columns = [normalizar_nome_coluna(c) for c in df.columns]

    # inverte o dicionario: variante -> nome padronizado
    mapa_reverso: Dict[str, str] = {}
    for padrao, variantes in aliases.items():
        for v in variantes:
            mapa_reverso[v] = padrao

    novas_colunas = {}
    nao_mapeadas = []
    for col in df.columns:
        if col in mapa_reverso:
            novas_colunas[col] = mapa_reverso[col]
        else:
            novas_colunas[col] = col
            if col not in aliases:  # coluna nao reconhecida
                nao_mapeadas.append(col)

    df = df.rename(columns=novas_colunas)

    # se o rename gerar colunas duplicadas, mantem a primeira nao-nula por linha
    if df.columns.duplicated().any():
        dup_names = df.columns[df.columns.duplicated()].unique().tolist()
        for nome in dup_names:
            sub = df.loc[:, df.columns == nome]
            combinado = sub.bfill(axis=1).iloc[:, 0]
            df = df.loc[:, df.columns != nome]
            df[nome] = combinado

    if nao_mapeadas:
        LOGGER.info(
            "Colunas nao mapeadas em %s (mantidas como estao): %s",
            origem, ", ".join(sorted(set(nao_mapeadas))),
        )
    return df


# =============================================================================
# CONVERSAO NUMERICA (padrao brasileiro: 1.234,56 -> 1234.56)
# =============================================================================

def converter_numero_br(serie: pd.Series) -> pd.Series:
    """Converte uma serie (texto ou numero) para float, tratando o padrao
    brasileiro de separador decimal (virgula) e milhar (ponto)."""
    if serie is None:
        return serie
    if pd.api.types.is_numeric_dtype(serie):
        return pd.to_numeric(serie, errors="coerce")

    def _conv(v):
        if pd.isna(v):
            return np.nan
        if isinstance(v, (int, float, np.integer, np.floating)):
            return float(v)
        s = str(v).strip()
        if s == "" or s.upper() in {"NA", "N/A", "-", "ND", "NAO_INFORMADO"}:
            return np.nan
        s = re.sub(r"[^\d,.\-]", "", s)
        if s == "":
            return np.nan
        # se tem ambos separadores, o ultimo encontrado eh o decimal
        if "," in s and "." in s:
            if s.rfind(",") > s.rfind("."):
                s = s.replace(".", "").replace(",", ".")
            else:
                s = s.replace(",", "")
        elif "," in s:
            s = s.replace(".", "").replace(",", ".")
        try:
            return float(s)
        except ValueError:
            return np.nan

    return serie.apply(_conv)


def converter_percentual(serie: pd.Series) -> pd.Series:
    """Converte coluna de ponderacao/percentual para fracao entre 0 e 1.
    Aceita '50%', '50', '0,5', 0.5, 50 -> todos viram 0.5."""
    valores = converter_numero_br(serie)

    def _norm(v):
        if pd.isna(v):
            return np.nan
        v = float(v)
        if v > 1.0:
            v = v / 100.0
        if v < 0 or v > 1:
            return np.nan
        return v

    return valores.apply(_norm)


# =============================================================================
# CARGA: BASE AGROFIT
# =============================================================================

def localizar_arquivo_agrofit() -> Optional[Path]:
    caminho_padrao = DADOS_DIR / ARQUIVO_AGROFIT_PADRAO
    if caminho_padrao.exists():
        return caminho_padrao
    for arq in DADOS_DIR.glob("*.csv"):
        nome_norm = remover_acentos(arq.stem).lower()
        if any(pista in nome_norm for pista in PISTAS_NOME_AGROFIT):
            return arq
    return None


def _ler_csv_robusto(caminho: Path) -> pd.DataFrame:
    """Tenta combinacoes comuns de separador/encoding de arquivos .gov.br."""
    tentativas = [
        dict(sep=";", encoding="utf-8-sig"),
        dict(sep=";", encoding="latin1"),
        dict(sep=",", encoding="utf-8-sig"),
        dict(sep=",", encoding="latin1"),
        dict(sep="\t", encoding="utf-8-sig"),
    ]
    ultimo_erro = None
    for opcoes in tentativas:
        try:
            df = pd.read_csv(caminho, dtype=str, low_memory=False, **opcoes)
            if df.shape[1] > 1:
                return df
        except Exception as e:  # noqa: BLE001
            ultimo_erro = e
            continue
    raise RuntimeError(f"Nao foi possivel ler o CSV {caminho.name}: {ultimo_erro}")


def carregar_agrofit(caminho: Path) -> pd.DataFrame:
    df = _ler_csv_robusto(caminho)
    linhas_brutas = len(df)
    df = padronizar_colunas(df, ALIASES_AGROFIT, origem=f"Agrofit ({caminho.name})")

    for col in ["NR_REGISTRO", "MARCA_COMERCIAL", "INGREDIENTE_ATIVO", "CULTURA",
                "PRAGA_NOME_CIENTIFICO", "PRAGA_NOME_COMUM", "TITULAR_DE_REGISTRO",
                "FORMULACAO", "SITUACAO"]:
        if col not in df.columns:
            df[col] = np.nan
            aviso(f"Coluna esperada do Agrofit ausente no arquivo: {col}")

    df["NR_REGISTRO"] = df["NR_REGISTRO"].astype(str).str.strip()
    df["_CHAVE_REGISTRO"] = df["NR_REGISTRO"].apply(extrair_chave_registro)
    df = df[df["_CHAVE_REGISTRO"] != ""]

    LOGGER.info(
        "Agrofit carregado: %s linhas brutas, %s linhas validas, %s registros unicos",
        linhas_brutas, len(df), df["_CHAVE_REGISTRO"].nunique(),
    )
    return df


# =============================================================================
# CARGA: BASES DE COMERCIALIZACAO (multiplos anos)
# =============================================================================

PADRAO_ANO = re.compile(r"(19|20)\d{2}")


def extrair_ano_do_nome(nome_arquivo: str) -> Optional[int]:
    achados = PADRAO_ANO.findall(nome_arquivo)
    # findall com grupo retorna so o prefixo; refazer busca sem grupo de captura
    achados = re.findall(r"(?:19|20)\d{2}", nome_arquivo)
    if not achados:
        return None
    # se houver mais de um numero de 4 digitos, usa o ultimo (mais provavel de ser o ano)
    return int(achados[-1])


def detectar_arquivos_comercializacao(caminho_agrofit: Optional[Path]) -> List[Path]:
    extensoes = ["*.xlsx", "*.xls", "*.csv"]
    candidatos: List[Path] = []
    for ext in extensoes:
        candidatos.extend(DADOS_DIR.glob(ext))
    candidatos = [
        c for c in candidatos
        if not (caminho_agrofit is not None and c.resolve() == caminho_agrofit.resolve())
        and not c.name.startswith("~$")
    ]
    return sorted(candidatos)


def _escolher_aba_excel(caminho: Path) -> Optional[str]:
    """Escolhe a aba mais provavel de conter os dados de comercializacao
    (evita abas de metadados/instrucoes quando houver mais de uma)."""
    try:
        xls = pd.ExcelFile(caminho)
    except Exception:
        return None
    abas = xls.sheet_names
    if len(abas) == 1:
        return abas[0]
    prioridade = ["TODOS_AGROTOXICOS", "TODOS", "BASE", "DADOS", "TABELA_BRUTA", "COMERCIALIZACAO"]
    for pista in prioridade:
        for aba in abas:
            if pista in remover_acentos(aba).upper().replace(" ", "_"):
                return aba
    # senao, escolhe a aba com mais linhas
    melhor_aba, melhor_linhas = abas[0], -1
    for aba in abas:
        try:
            n = pd.read_excel(caminho, sheet_name=aba, nrows=0).shape[1]
            df_teste = pd.read_excel(caminho, sheet_name=aba, nrows=5)
            if df_teste.shape[1] > melhor_linhas:
                melhor_linhas = df_teste.shape[1]
                melhor_aba = aba
        except Exception:
            continue
    return melhor_aba



def _limpar_cabecalho_tabela_bruta(nome: str) -> str:
    """Limpa tags <br> e espacos de cabecalhos da Tabela Bruta IBAMA/MAPA."""
    if nome is None:
        return ""
    s = str(nome).replace("\xa0", " ").strip()
    s = re.sub(r"<br\s*/?>", " ", s, flags=re.IGNORECASE)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def _mapear_colunas_tabela_bruta(df: pd.DataFrame) -> pd.DataFrame:
    """Mapeia explicitamente a estrutura da Tabela Bruta de Comercializacao.

    Layout tipico (aba TODOS_Agrotóxicos):
      [metadados] ... Classificacao Ambiental PT |
      BLOCO PF (produto formulado / original):
        Estoque Inicial, Producao Nacional, Importacao, Exportacao,
        Vendas a Industria, Vendas a Cliente, Estoque Final, RO..DF
      Linha | Ponderacao (% IA no PF) |
      BLOCO IA (ja ponderado na planilha de origem; pandas sufixa .1):
        Estoque Inicial.1, ... Vendas a Cliente.1, Vendas Totais, Estoque Final.1,
        RO.1..DF.1, Total UF, Brasil

    - Colunas do BLOCO PF viram ESTOQUE_INICIAL, VENDAS_INDUSTRIA, ... (originais)
    - Colunas do BLOCO IA viram ESTOQUE_INICIAL_PONDERADO_IA, ... e tambem
      alimentam VENDA_PONDERADA_IA a partir de 'Vendas Totais' / 'Brasil'
    - VENDAS_TOTAIS (original) = Vendas a Industria + Vendas a Cliente do bloco PF
    """
    df = df.copy()
    # Limpa nomes
    df.columns = [_limpar_cabecalho_tabela_bruta(c) for c in df.columns]
    # Descarta colunas vazias / so espaco
    df = df.loc[:, [c for c in df.columns if c and not str(c).startswith("Unnamed")]]

    # Mapa posicional por nome limpo (primeira ocorrencia = PF, .1 = IA)
    rename: Dict[str, str] = {}
    cols = list(df.columns)

    # Metadados
    meta = {
        "Detentor": "DETENTOR",
        "Marca comercial": "MARCA_COMERCIAL",
        "Tipo": "PRODUTO_TECNICO",
        "Registro Mapa": "NR_REGISTRO",
        "Classes Uso": "CLASSES_USO",
        "Produto Tecnico": "PRODUTO_TECNICO_DETALHE",
        "Produto Técnico": "PRODUTO_TECNICO_DETALHE",
        "Ingrediente Ativo": "INGREDIENTE_ATIVO",
        "Concentracao": "CONCENTRACAO",
        "Concentração": "CONCENTRACAO",
        "Unidade": "UNIDADE",
        "Ponderacao (% IA no PF)": "PONDERACAO_IA",
        "Ponderação (% IA no PF)": "PONDERACAO_IA",
        "Linha": "LINHA",
        "Densidade": "DENSIDADE",
        "Total UF": "TOTAL_UF_IA",
        "Brasil": "BRASIL_IA",
        "Vendas Totais": "VENDAS_TOTAIS_IA_ORIGEM",  # ja e IA na planilha
    }
    for c in cols:
        if c in meta:
            rename[c] = meta[c]
        # variante sem acento apos limpeza parcial
        c_sem = remover_acentos(c)
        if c_sem in meta and c not in rename:
            rename[c] = meta[c_sem]

    # Volumes PF (sem sufixo .1)
    pf_map = {
        "Estoque Inicial": "ESTOQUE_INICIAL",
        "Producao Nacional": "PRODUCAO_NACIONAL",
        "Produção Nacional": "PRODUCAO_NACIONAL",
        "Importacao": "IMPORTACAO",
        "Importação": "IMPORTACAO",
        "Exportacao": "EXPORTACAO",
        "Exportação": "EXPORTACAO",
        "Vendas a Industria": "VENDAS_INDUSTRIA",
        "Vendas a Indústria": "VENDAS_INDUSTRIA",
        "Vendas a Cliente": "VENDAS_CLIENTE",
        "Estoque Final": "ESTOQUE_FINAL",
    }
    # Volumes IA (sufixo .1 do pandas)
    ia_map = {
        "Estoque Inicial.1": "ESTOQUE_INICIAL_PONDERADO_IA",
        "Producao Nacional.1": "PRODUCAO_NACIONAL_PONDERADA_IA",
        "Produção Nacional.1": "PRODUCAO_NACIONAL_PONDERADA_IA",
        "Importacao.1": "IMPORTACAO_PONDERADA_IA",
        "Importação.1": "IMPORTACAO_PONDERADA_IA",
        "Exportacao.1": "EXPORTACAO_PONDERADA_IA",
        "Exportação.1": "EXPORTACAO_PONDERADA_IA",
        "Vendas a Industria.1": "VENDAS_INDUSTRIA_PONDERADA_IA",
        "Vendas a Indústria.1": "VENDAS_INDUSTRIA_PONDERADA_IA",
        "Vendas a Cliente.1": "VENDAS_CLIENTE_PONDERADA_IA",
        "Estoque Final.1": "ESTOQUE_FINAL_PONDERADO_IA",
    }

    for c in cols:
        if c in pf_map:
            rename[c] = pf_map[c]
        elif c in ia_map:
            rename[c] = ia_map[c]
        else:
            # UF do bloco PF (sigla pura)
            if c in UFS:
                rename[c] = c  # mantem UF para formato longo PF
            # UF do bloco IA
            elif c.endswith(".1") and c[:-2] in UFS:
                rename[c] = f"UF_{c[:-2]}_IA"

    df = df.rename(columns=rename)

    # Se ainda houver colunas duplicadas apos rename, combina
    if df.columns.duplicated().any():
        dup_names = df.columns[df.columns.duplicated()].unique().tolist()
        for nome in dup_names:
            sub = df.loc[:, df.columns == nome]
            combinado = sub.bfill(axis=1).iloc[:, 0]
            df = df.loc[:, df.columns != nome]
            df[nome] = combinado

    return df


def carregar_arquivo_comercializacao(caminho: Path) -> pd.DataFrame:
    ano_detectado = extrair_ano_do_nome(caminho.name)

    if caminho.suffix.lower() == ".csv":
        df = _ler_csv_robusto(caminho)
        df = padronizar_colunas(df, ALIASES_COMERC, origem=f"Comercializacao ({caminho.name})")
    else:
        aba = _escolher_aba_excel(caminho)
        df = pd.read_excel(caminho, sheet_name=aba, dtype=str)
        # Detecta Tabela Bruta (colunas com <br> ou Ponderacao / Vendas Totais + Brasil)
        cols_raw = [str(c) for c in df.columns]
        is_tabela_bruta = any(
            ("<br>" in c) or ("Pondera" in c) or (c.strip() == "Brasil")
            for c in cols_raw
        )
        if is_tabela_bruta:
            LOGGER.info("Formato Tabela Bruta detectado em %s (aba=%s)", caminho.name, aba)
            df = _mapear_colunas_tabela_bruta(df)
            # Garante nomes padronizados restantes via aliases (sem sobrescrever ja mapeados)
            df = padronizar_colunas(df, ALIASES_COMERC, origem=f"Comercializacao/TB ({caminho.name})")
        else:
            df = padronizar_colunas(df, ALIASES_COMERC, origem=f"Comercializacao ({caminho.name})")

    if "ANO" in df.columns and df["ANO"].notna().any():
        df["ANO"] = converter_numero_br(df["ANO"]).astype("Int64")
        if ano_detectado is not None and df["ANO"].isna().all():
            df["ANO"] = ano_detectado
    elif ano_detectado is not None:
        df["ANO"] = ano_detectado
    else:
        aviso(f"Nao foi possivel identificar o ANO do arquivo {caminho.name}; linhas descartadas para fins de serie historica.")
        df["ANO"] = np.nan

    df["_ARQUIVO_ORIGEM"] = caminho.name

    if "NR_REGISTRO" not in df.columns:
        # tentativa final: qualquer coluna que contenha "REGISTRO" e nao seja
        # claramente outra coisa (situacao, titular, detentor, data)
        candidatas = [
            c for c in df.columns
            if "REGISTRO" in c and not any(p in c for p in ["SITUACAO", "TITULAR", "DETENTOR", "DATA"])
        ]
        if candidatas:
            df = df.rename(columns={candidatas[0]: "NR_REGISTRO"})
            LOGGER.info("Coluna '%s' assumida como NR_REGISTRO em %s (correspondencia aproximada).", candidatas[0], caminho.name)
        else:
            aviso(f"Arquivo {caminho.name} nao possui coluna de NR_REGISTRO identificavel; sera ignorado.")
            return pd.DataFrame()

    df["NR_REGISTRO"] = df["NR_REGISTRO"].astype(str).str.strip()
    df["_CHAVE_REGISTRO"] = df["NR_REGISTRO"].apply(extrair_chave_registro)
    df = df[df["_CHAVE_REGISTRO"] != ""]

    LOGGER.info(
        "Arquivo de comercializacao carregado: %s | linhas=%s | ano=%s",
        caminho.name, len(df), sorted(df["ANO"].dropna().unique().tolist()),
    )
    return df


def carregar_todas_comercializacoes(arquivos: List[Path]) -> pd.DataFrame:
    partes = []
    for caminho in arquivos:
        try:
            df = carregar_arquivo_comercializacao(caminho)
            if not df.empty:
                partes.append(df)
                ok(f"Processado: {caminho.name} ({len(df):,} linhas)")
            else:
                aviso(f"Arquivo ignorado (sem dados aproveitaveis): {caminho.name}")
        except Exception as e:  # noqa: BLE001
            erro(f"Falha ao processar {caminho.name}: {e}")
            LOGGER.error("Falha ao processar %s: %s\n%s", caminho.name, e, traceback.format_exc())

    if not partes:
        return pd.DataFrame()

    df_final = pd.concat(partes, ignore_index=True, sort=False)
    return df_final


# =============================================================================
# PREPARACAO NUMERICA DAS COLUNAS DE COMERCIALIZACAO
# =============================================================================

COLUNAS_NUMERICAS_COMERC = [
    "ESTOQUE_INICIAL", "PRODUCAO_NACIONAL", "IMPORTACAO", "EXPORTACAO",
    "VENDAS_INDUSTRIA", "VENDAS_CLIENTE", "VENDAS_TOTAIS", "ESTOQUE_FINAL",
]


def preparar_numericos_comercializacao(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    for col in COLUNAS_NUMERICAS_COMERC:
        if col in df.columns:
            df[col] = converter_numero_br(df[col])
        else:
            df[col] = np.nan

    # Converte tambem colunas ja ponderadas vindas da Tabela Bruta e auxiliares
    extras_num = [
        "PONDERACAO_IA", "VENDAS_TOTAIS_IA_ORIGEM", "BRASIL_IA", "TOTAL_UF_IA",
        "CONCENTRACAO", "DENSIDADE",
    ] + list(MAPA_VOLUMES_PONDERADOS.values())
    for col in extras_num:
        if col in df.columns:
            if col == "PONDERACAO_IA":
                df[col] = converter_percentual(df[col])
            else:
                df[col] = converter_numero_br(df[col])

    if "PONDERACAO_IA" not in df.columns:
        df["PONDERACAO_IA"] = np.nan

    # VENDAS_TOTAIS ORIGINAL = produto formulado (industria + cliente).
    # Na Tabela Bruta nao existe "Vendas Totais" no bloco PF — so no bloco IA.
    soma_ind_cli = df[["VENDAS_INDUSTRIA", "VENDAS_CLIENTE"]].sum(axis=1, min_count=1)
    if "VENDAS_TOTAIS" not in df.columns or df["VENDAS_TOTAIS"].isna().all():
        df["VENDAS_TOTAIS"] = soma_ind_cli
    else:
        # Nao usar Vendas Totais do bloco IA como original
        df["VENDAS_TOTAIS"] = df["VENDAS_TOTAIS"].where(
            df["VENDAS_TOTAIS"].notna(), soma_ind_cli
        )
        # Se VENDAS_TOTAIS veio mapeado de "Vendas Totais" (IA), prefere soma PF
        if "VENDAS_TOTAIS_IA_ORIGEM" in df.columns:
            df["VENDAS_TOTAIS"] = soma_ind_cli

    if "UNIDADE" not in df.columns:
        df["UNIDADE"] = np.nan
    df["_UNIDADE_NORM"] = df["UNIDADE"].apply(
        lambda v: normalizar_texto_valor(v) if pd.notna(v) else np.nan
    )

    return df


# Colunas de volume da planilha de comercializacao (produto formulado / original)
# e os nomes das respectivas colunas ponderadas pelo teor de IA.
MAPA_VOLUMES_PONDERADOS = {
    "ESTOQUE_INICIAL": "ESTOQUE_INICIAL_PONDERADO_IA",
    "PRODUCAO_NACIONAL": "PRODUCAO_NACIONAL_PONDERADA_IA",
    "IMPORTACAO": "IMPORTACAO_PONDERADA_IA",
    "EXPORTACAO": "EXPORTACAO_PONDERADA_IA",
    "VENDAS_INDUSTRIA": "VENDAS_INDUSTRIA_PONDERADA_IA",
    "VENDAS_CLIENTE": "VENDAS_CLIENTE_PONDERADA_IA",
    "VENDAS_TOTAIS": "VENDA_PONDERADA_IA",  # alias principal de vendas ponderadas
    "ESTOQUE_FINAL": "ESTOQUE_FINAL_PONDERADO_IA",
}


def calcular_venda_estimada_ia(df: pd.DataFrame) -> pd.DataFrame:
    """Preserva volumes ORIGINAIS e cria colunas PONDERADAS pelo teor de IA.

    Para cada coluna de volume da planilha (Estoque Inicial, Producao Nacional,
    Importacao, Exportacao, Vendas a Industria, Vendas a Cliente, Vendas Totais,
    Estoque Final):

      coluna_original  -> permanece intacta (auditoria)
      coluna_ponderada = coluna_original * PONDERACAO_IA  (quantidade efetiva de IA)

    VENDA_PONDERADA_IA e VENDA_ESTIMADA_IA sao aliases de VENDAS_TOTAIS * teor.
    STATUS_CALCULO_IA indica se o ponderado foi calculado ou o motivo da falha.
    Nunca sobrescreve colunas originais.
    """
    df = df.copy()

    tem_ponderacao = "PONDERACAO_IA" in df.columns and df["PONDERACAO_IA"].notna().any()
    tem_vendas = "VENDAS_TOTAIS" in df.columns

    def _status(row):
        # Status baseado principalmente em vendas totais (metrica central)
        vendas = row.get("VENDAS_TOTAIS", np.nan)
        if pd.isna(vendas) and all(
            pd.isna(row.get(c, np.nan)) for c in COLUNAS_NUMERICAS_COMERC if c != "VENDAS_TOTAIS"
        ):
            return "DADO_AUSENTE"
        if pd.isna(row.get("PONDERACAO_IA", np.nan)):
            return "SEM_PONDERACAO"
        unidade = row.get("_UNIDADE_NORM")
        if pd.notna(unidade):
            u = str(unidade).replace("/", "_").replace(" ", "_")
            # Unidade de concentracao (g/l, g/kg) da Tabela Bruta: nao bloqueia
            if u in UNIDADES_CONCENTRACAO or any(
                u.startswith(p) or u.endswith(p) for p in ("G_L", "G_KG", "MG_L")
            ):
                return "CALCULADO"
            if u not in UNIDADES_COMPATIVEIS and u not in {"", "NAN"}:
                # Unidades de contagem (UN, PC, CX...) sao incompativeis para massa/volume
                if u in {"UN", "UND", "UNIDADE", "PC", "PÇ", "CX", "FR", "SC", "DOSE"}:
                    return "UNIDADE_INCOMPATIVEL"
        return "CALCULADO"

    df["STATUS_CALCULO_IA"] = df.apply(_status, axis=1)
    pode_calcular = df["STATUS_CALCULO_IA"] == "CALCULADO"

    for col_orig, col_pond in MAPA_VOLUMES_PONDERADOS.items():
        if col_orig not in df.columns:
            df[col_orig] = np.nan
        # Prefer values already present from Tabela Bruta (bloco IA)
        if col_pond in df.columns and df[col_pond].notna().any():
            # Completa lacunas com calculo PF * ponderacao
            calculado = np.where(
                pode_calcular & df[col_orig].notna(),
                df[col_orig] * df["PONDERACAO_IA"],
                np.nan,
            )
            df[col_pond] = df[col_pond].where(df[col_pond].notna(), calculado)
        else:
            df[col_pond] = np.where(
                pode_calcular & df.get(col_orig, pd.Series(np.nan, index=df.index)).notna(),
                df[col_orig] * df["PONDERACAO_IA"],
                np.nan,
            )

    # Venda ponderada: prefere Vendas Totais / Brasil da planilha (ja em IA)
    if "VENDAS_TOTAIS_IA_ORIGEM" in df.columns and df["VENDAS_TOTAIS_IA_ORIGEM"].notna().any():
        df["VENDA_PONDERADA_IA"] = df["VENDAS_TOTAIS_IA_ORIGEM"].where(
            df["VENDAS_TOTAIS_IA_ORIGEM"].notna(),
            df.get("VENDA_PONDERADA_IA"),
        )
    elif "BRASIL_IA" in df.columns and df["BRASIL_IA"].notna().any():
        df["VENDA_PONDERADA_IA"] = df["BRASIL_IA"].where(
            df["BRASIL_IA"].notna(),
            df.get("VENDA_PONDERADA_IA"),
        )

    if "VENDA_PONDERADA_IA" in df.columns:
        df["VENDA_ESTIMADA_IA"] = df["VENDA_PONDERADA_IA"]
    else:
        df["VENDA_ESTIMADA_IA"] = np.nan
        df["VENDA_PONDERADA_IA"] = np.nan

    return df


# =============================================================================
# NORMALIZACAO DAS COLUNAS DE UF (27 colunas -> formato longo)
# =============================================================================

def identificar_colunas_uf(df: pd.DataFrame) -> List[str]:
    """Localiza quais das 27 siglas de UF existem como colunas no dataframe."""
    presentes = [uf for uf in UFS if uf in df.columns]
    return presentes


def montar_venda_para_uf_longo(
    df_vendas_bruto: pd.DataFrame, colunas_id: List[str]
) -> pd.DataFrame:
    """Transforma as colunas de UF (uma por estado) em formato longo:
    NR_REGISTRO | ANO | UF | REGIAO | VENDA."""
    colunas_uf = identificar_colunas_uf(df_vendas_bruto)
    if not colunas_uf:
        return pd.DataFrame(columns=["NR_REGISTRO", "ANO", "UF", "REGIAO", "VENDA"])

    df = df_vendas_bruto.copy()
    for uf in colunas_uf:
        df[uf] = converter_numero_br(df[uf])

    cols_manter = [c for c in colunas_id if c in df.columns] + colunas_uf
    df_long = df[cols_manter].melt(
        id_vars=[c for c in colunas_id if c in df.columns],
        value_vars=colunas_uf,
        var_name="UF",
        value_name="VENDA",
    )
    df_long = df_long[df_long["VENDA"].notna() & (df_long["VENDA"] != 0)]
    df_long["REGIAO"] = df_long["UF"].map(REGIAO_POR_UF)
    return df_long


# =============================================================================
# CONSTRUCAO DAS TABELAS ANALITICAS
# =============================================================================

# ---------- Tabela 1: Produtos (1 linha = 1 registro/produto) ----------
def montar_tabela_produtos(df_agrofit: pd.DataFrame) -> pd.DataFrame:
    colunas = [
        "NR_REGISTRO", "MARCA_COMERCIAL", "FORMULACAO", "INGREDIENTE_ATIVO",
        "TITULAR_DE_REGISTRO", "CLASSE", "MODO_DE_ACAO", "EMPRESA_PAIS_TIPO",
        "CLASSE_TOXICOLOGICA", "CLASSE_AMBIENTAL", "ORGANICOS", "SITUACAO",
    ]
    presentes = [c for c in colunas if c in df_agrofit.columns]

    def _juntar_unicos(serie: pd.Series) -> str:
        vals = [str(v).strip() for v in serie.dropna().unique() if str(v).strip()]
        return "; ".join(sorted(set(vals)))

    agregadores = {}
    for c in presentes:
        if c == "NR_REGISTRO":
            continue
        if c in ("INGREDIENTE_ATIVO", "CULTURA"):
            agregadores[c] = _juntar_unicos
        else:
            agregadores[c] = "first"

    df = df_agrofit.copy()
    df_prod = df.groupby(["_CHAVE_REGISTRO", "NR_REGISTRO"], as_index=False).agg(agregadores)
    df_prod = df_prod.drop_duplicates(subset=["_CHAVE_REGISTRO"]).reset_index(drop=True)
    return df_prod


# ---------- Tabela 2: Agrofit_Uso (1 linha = Registro x Cultura x Praga) ----------
def montar_tabela_agrofit_uso(df_agrofit: pd.DataFrame) -> pd.DataFrame:
    colunas = [
        "NR_REGISTRO", "_CHAVE_REGISTRO", "CULTURA", "PRAGA_NOME_CIENTIFICO",
        "PRAGA_NOME_COMUM", "MARCA_COMERCIAL", "INGREDIENTE_ATIVO", "FORMULACAO",
    ]
    presentes = [c for c in colunas if c in df_agrofit.columns]
    df_uso = df_agrofit[presentes].drop_duplicates().reset_index(drop=True)
    return df_uso


def _enriquecer_com_agrofit(
    df: pd.DataFrame, df_produtos: pd.DataFrame, campos: Optional[List[str]] = None
) -> pd.DataFrame:
    """Preenche campos de identificacao ausentes com dados do Agrofit (nao toca volumes)."""
    if df.empty or df_produtos.empty:
        return df
    if campos is None:
        campos = ["MARCA_COMERCIAL", "INGREDIENTE_ATIVO", "TITULAR_DE_REGISTRO", "FORMULACAO"]
    campos = [c for c in campos if c in df_produtos.columns]
    if not campos or "_CHAVE_REGISTRO" not in df.columns:
        return df
    df_prod_aux = df_produtos[["_CHAVE_REGISTRO"] + campos].drop_duplicates("_CHAVE_REGISTRO")
    df_prod_aux = df_prod_aux.rename(columns={c: f"{c}_AGROFIT" for c in campos})
    out = df.merge(df_prod_aux, on="_CHAVE_REGISTRO", how="left")
    for c in campos:
        col_agro = f"{c}_AGROFIT"
        if c in out.columns:
            out[c] = out[c].where(
                out[c].notna() & (out[c].astype(str).str.strip() != ""),
                out[col_agro],
            )
        else:
            out[c] = out[col_agro]
        out.drop(columns=[col_agro], inplace=True, errors="ignore")
    return out


# ---------- Tabela 3: Comercializacao_Produto (1 linha = Registro x Ano) ----------
# Perspectiva PRODUTO: consolida multi-IA em uma linha. Somente volumes de
# produto formulado (PF). Colunas ponderadas de IA NAO entram aqui — use
# 04_Comerc_IA / 06_IA_Anual para analise por ingrediente ativo.
def montar_comercializacao_produto(
    df_vendas: pd.DataFrame, df_produtos: pd.DataFrame
) -> pd.DataFrame:
    if df_vendas.empty:
        return pd.DataFrame()

    # Volumes ja devem estar numericos (preparar_numericos no main).
    # Nao calcula nem expõe colunas ponderadas de IA nesta aba.
    df = df_vendas.copy()
    for col in COLUNAS_NUMERICAS_COMERC:
        if col in df.columns and not pd.api.types.is_numeric_dtype(df[col]):
            df[col] = converter_numero_br(df[col])
    if "VENDAS_TOTAIS" not in df.columns:
        df["VENDAS_TOTAIS"] = np.nan
    if df["VENDAS_TOTAIS"].isna().all() and {"VENDAS_INDUSTRIA", "VENDAS_CLIENTE"}.issubset(df.columns):
        df["VENDAS_TOTAIS"] = df[["VENDAS_INDUSTRIA", "VENDAS_CLIENTE"]].sum(axis=1, min_count=1)

    campos_volume_pf = [c for c in COLUNAS_NUMERICAS_COMERC if c in df.columns]
    campos_primeiro = [
        c for c in [
            "MARCA_COMERCIAL", "DETENTOR", "PRODUTO_TECNICO",
            "UNIDADE", "CLASSES_USO", "_UNIDADE_NORM", "_ARQUIVO_ORIGEM",
        ] if c in df.columns
    ]

    agregadores: Dict[str, Any] = {c: "first" for c in campos_volume_pf}
    agregadores.update({c: "first" for c in campos_primeiro})
    if "INGREDIENTE_ATIVO" in df.columns:
        def _join_ias(s):
            vals = sorted({str(v).strip() for v in s.dropna() if str(v).strip()})
            return "; ".join(vals) if vals else np.nan
        agregadores["INGREDIENTE_ATIVO"] = _join_ias
    if "CONCENTRACAO" in df.columns:
        def _join_conc(s):
            vals = [str(v).strip() for v in s.dropna() if str(v).strip()]
            return "; ".join(vals) if vals else np.nan
        agregadores["CONCENTRACAO"] = _join_conc
    if "NR_REGISTRO" in df.columns:
        agregadores["NR_REGISTRO"] = "first"

    group_keys = ["_CHAVE_REGISTRO", "ANO"]
    if "NR_REGISTRO" not in agregadores and "NR_REGISTRO" in df.columns:
        group_keys = ["_CHAVE_REGISTRO", "NR_REGISTRO", "ANO"]

    df_anual = df.groupby(group_keys, as_index=False).agg(agregadores)

    df_anual = _enriquecer_com_agrofit(df_anual, df_produtos)
    if "DETENTOR" not in df_anual.columns:
        df_anual["DETENTOR"] = np.nan
    if "INGREDIENTE_ATIVO" in df_anual.columns:
        df_anual["NUM_IAS"] = df_anual["INGREDIENTE_ATIVO"].apply(
            lambda v: len([x for x in str(v).split(";") if x.strip()]) if pd.notna(v) else 0
        )

    colunas_finais = [
        "NR_REGISTRO", "ANO", "MARCA_COMERCIAL", "INGREDIENTE_ATIVO", "NUM_IAS",
        "DETENTOR", "PRODUTO_TECNICO", "CONCENTRACAO", "UNIDADE", "CLASSES_USO",
        # Somente volumes de produto formulado (auditoria / analise de produto)
        "ESTOQUE_INICIAL", "PRODUCAO_NACIONAL", "IMPORTACAO", "EXPORTACAO",
        "VENDAS_INDUSTRIA", "VENDAS_CLIENTE", "VENDAS_TOTAIS", "ESTOQUE_FINAL",
        "_CHAVE_REGISTRO",
    ]
    colunas_finais = [c for c in colunas_finais if c in df_anual.columns]
    return df_anual[colunas_finais].sort_values(["NR_REGISTRO", "ANO"]).reset_index(drop=True)


# Alias legado usado internamente por funcoes que ainda esperam o nome antigo
def montar_comercializacao_anual(
    df_vendas: pd.DataFrame, df_produtos: pd.DataFrame
) -> pd.DataFrame:
    return montar_comercializacao_produto(df_vendas, df_produtos)


# ---------- Tabela 4: Comercializacao_IA (1 linha = Registro x Ano x IA) ----------
# Perspectiva IA: NAO consolida. Cada ingrediente ativo permanece em linha propria
# com sua PONDERACAO_IA e volumes ponderados individuais.
def montar_comercializacao_ia(
    df_vendas: pd.DataFrame, df_produtos: pd.DataFrame
) -> pd.DataFrame:
    if df_vendas.empty:
        return pd.DataFrame()

    df = calcular_venda_estimada_ia(df_vendas.copy())
    df = _enriquecer_com_agrofit(
        df, df_produtos, campos=["MARCA_COMERCIAL", "TITULAR_DE_REGISTRO", "FORMULACAO"]
    )
    if "DETENTOR" not in df.columns:
        df["DETENTOR"] = np.nan

    if "VENDA_PONDERADA_IA" in df.columns:
        df["VENDA_ESTIMADA_IA"] = df["VENDA_PONDERADA_IA"]
    elif "VENDA_ESTIMADA_IA" in df.columns:
        df["VENDA_PONDERADA_IA"] = df["VENDA_ESTIMADA_IA"]

    colunas_finais = [
        "NR_REGISTRO", "ANO", "MARCA_COMERCIAL", "INGREDIENTE_ATIVO",
        "DETENTOR", "PRODUTO_TECNICO", "CONCENTRACAO", "UNIDADE", "CLASSES_USO",
        "PONDERACAO_IA",
        # Volumes PF da linha (iguais entre IAs do mesmo produto — nao somar entre IAs
        # se o objetivo for volume do produto; use 03_Comerc_Produto para isso)
        "ESTOQUE_INICIAL", "PRODUCAO_NACIONAL", "IMPORTACAO", "EXPORTACAO",
        "VENDAS_INDUSTRIA", "VENDAS_CLIENTE", "VENDAS_TOTAIS", "ESTOQUE_FINAL",
        # Volumes ponderados DESTA IA
        "ESTOQUE_INICIAL_PONDERADO_IA", "PRODUCAO_NACIONAL_PONDERADA_IA",
        "IMPORTACAO_PONDERADA_IA", "EXPORTACAO_PONDERADA_IA",
        "VENDAS_INDUSTRIA_PONDERADA_IA", "VENDAS_CLIENTE_PONDERADA_IA",
        "VENDA_PONDERADA_IA", "VENDA_ESTIMADA_IA", "ESTOQUE_FINAL_PONDERADO_IA",
        "STATUS_CALCULO_IA", "_CHAVE_REGISTRO",
    ]
    colunas_finais = [c for c in colunas_finais if c in df.columns]
    return (
        df[colunas_finais]
        .sort_values(["NR_REGISTRO", "ANO", "INGREDIENTE_ATIVO"])
        .reset_index(drop=True)
    )


# ---------- Tabela 5: Comercializacao_UF (produto; 1 linha = Registro x Ano x UF) ----------
# Volumes de UF sao de produto formulado. Multi-IA: deduplica com max para nao
# multiplicar o volume do produto pelo numero de ingredientes ativos.
def montar_comercializacao_uf(
    df_vendas: pd.DataFrame, df_produtos: Optional[pd.DataFrame] = None
) -> pd.DataFrame:
    if df_vendas.empty:
        return pd.DataFrame(columns=[
            "ANO", "NR_REGISTRO", "MARCA_COMERCIAL", "INGREDIENTE_ATIVO",
            "UF", "REGIAO", "VENDA",
        ])
    df_long = montar_venda_para_uf_longo(df_vendas, ["_CHAVE_REGISTRO", "NR_REGISTRO", "ANO"])
    if df_long.empty:
        return df_long
    # max (nao sum): multi-IA repete o mesmo volume PF em cada linha de IA
    df_long = df_long.groupby(
        ["_CHAVE_REGISTRO", "NR_REGISTRO", "ANO", "UF", "REGIAO"], as_index=False
    )["VENDA"].max()

    if df_produtos is not None and not df_produtos.empty:
        campos_id = [c for c in ["MARCA_COMERCIAL", "INGREDIENTE_ATIVO"] if c in df_produtos.columns]
        if campos_id:
            df_prod_aux = df_produtos[["_CHAVE_REGISTRO"] + campos_id].drop_duplicates("_CHAVE_REGISTRO")
            df_long = df_long.merge(df_prod_aux, on="_CHAVE_REGISTRO", how="left")

    ordem = [
        "ANO", "NR_REGISTRO", "MARCA_COMERCIAL", "INGREDIENTE_ATIVO",
        "UF", "REGIAO", "VENDA", "_CHAVE_REGISTRO",
    ]
    ordem = [c for c in ordem if c in df_long.columns]
    return df_long[ordem].sort_values(["ANO", "NR_REGISTRO", "UF"]).reset_index(drop=True)


# ---------- Tabela 6: IA_Anual (agrega a partir da perspectiva IA) ----------
def montar_ia_anual(df_com_ia: pd.DataFrame) -> pd.DataFrame:
    if df_com_ia.empty or "INGREDIENTE_ATIVO" not in df_com_ia.columns:
        return pd.DataFrame()
    df = df_com_ia.copy()
    df = df[df["INGREDIENTE_ATIVO"].notna() & (df["INGREDIENTE_ATIVO"].astype(str).str.strip() != "")]
    # Evita contar o mesmo produto N vezes no volume PF ao agregar por IA:
    # para VENDA_TOTAL (PF) usamos o volume da linha (cada linha = um produto-IA);
    # a metrica correta de "quanto de IA foi comercializado" e VENDA_PONDERADA_IA.
    agg_kwargs = dict(
        NUM_REGISTROS=("NR_REGISTRO", "nunique"),
        VENDA_TOTAL_PF=("VENDAS_TOTAIS", "sum"),
        VENDA_PONDERADA_IA=("VENDA_PONDERADA_IA", "sum"),
        VENDA_ESTIMADA_IA=("VENDA_ESTIMADA_IA", "sum"),
        PRODUCAO_NACIONAL_PONDERADA_IA=("PRODUCAO_NACIONAL_PONDERADA_IA", "sum"),
        IMPORTACAO_PONDERADA_IA=("IMPORTACAO_PONDERADA_IA", "sum"),
        EXPORTACAO_PONDERADA_IA=("EXPORTACAO_PONDERADA_IA", "sum"),
        VENDAS_INDUSTRIA_PONDERADA_IA=("VENDAS_INDUSTRIA_PONDERADA_IA", "sum"),
        VENDAS_CLIENTE_PONDERADA_IA=("VENDAS_CLIENTE_PONDERADA_IA", "sum"),
        ESTOQUE_INICIAL_PONDERADO_IA=("ESTOQUE_INICIAL_PONDERADO_IA", "sum"),
        ESTOQUE_FINAL_PONDERADO_IA=("ESTOQUE_FINAL_PONDERADO_IA", "sum"),
    )
    if "MARCA_COMERCIAL" in df.columns:
        agg_kwargs["NUM_PRODUTOS"] = ("MARCA_COMERCIAL", "nunique")
    else:
        agg_kwargs["NUM_PRODUTOS"] = ("NR_REGISTRO", "nunique")
    agg_kwargs = {
        k: v for k, v in agg_kwargs.items()
        if v[0] in df.columns or k in ("NUM_REGISTROS", "NUM_PRODUTOS")
    }
    grp = df.groupby(["ANO", "INGREDIENTE_ATIVO"], as_index=False).agg(**agg_kwargs)
    # Alias principal de ordenacao = quantidade efetiva de IA
    if "VENDA_PONDERADA_IA" in grp.columns:
        grp["VENDA_TOTAL"] = grp["VENDA_PONDERADA_IA"]
    elif "VENDA_TOTAL_PF" in grp.columns:
        grp["VENDA_TOTAL"] = grp["VENDA_TOTAL_PF"]
    return grp.sort_values(["ANO", "VENDA_TOTAL"], ascending=[True, False]).reset_index(drop=True)


# ---------- Tabela 10: Evolucao_Produtos_Anual (WIDE: produto x ano) ----------
# 1 linha por produto (Registro), 1 coluna por ANO com o volume total
# comercializado (VENDAS_TOTAIS, perspectiva PRODUTO/PF) + coluna TOTAL_PERIODO
# com a soma de todos os anos. Formato ideal para levantar graficos de
# evolucao anual de uso por produto (linhas = produtos, colunas = anos) —
# basta selecionar as colunas de ano e montar um grafico de linhas/tabela
# dinamica diretamente no Excel.
def montar_evolucao_produtos_anual(df_com_produto: pd.DataFrame) -> pd.DataFrame:
    if (
        df_com_produto.empty
        or "ANO" not in df_com_produto.columns
        or "VENDAS_TOTAIS" not in df_com_produto.columns
    ):
        return pd.DataFrame()

    df = df_com_produto.copy()
    if not pd.api.types.is_numeric_dtype(df["ANO"]):
        df["ANO"] = converter_numero_br(df["ANO"])
    df["ANO"] = df["ANO"].astype("Int64")
    df["VENDAS_TOTAIS"] = pd.to_numeric(df["VENDAS_TOTAIS"], errors="coerce")
    df = df[df["ANO"].notna()]
    if df.empty:
        return pd.DataFrame()

    campos_id = [c for c in ["NR_REGISTRO", "MARCA_COMERCIAL"] if c in df.columns]
    if not campos_id:
        return pd.DataFrame()

    # 1 linha por produto x ano (soma caso ainda exista alguma granularidade extra)
    agg = df.groupby(campos_id + ["ANO"], as_index=False)["VENDAS_TOTAIS"].sum()

    pivot = agg.pivot_table(
        index=campos_id, columns="ANO", values="VENDAS_TOTAIS", aggfunc="sum", fill_value=0
    )
    anos_ordenados = sorted(pivot.columns.tolist())
    pivot = pivot[anos_ordenados]
    colunas_ano = [str(int(a)) for a in anos_ordenados]
    pivot.columns = colunas_ano
    pivot["TOTAL_PERIODO"] = pivot[colunas_ano].sum(axis=1)
    pivot = pivot.reset_index()
    return pivot.sort_values("TOTAL_PERIODO", ascending=False).reset_index(drop=True)


# ---------- Tabela 11: Evolucao_IA_Anual (WIDE: ingrediente ativo x ano) ----------
# 1 linha por ingrediente ativo, 1 coluna por ANO com a venda ponderada/
# estimada de IA (perspectiva IA) + coluna TOTAL_PERIODO com a soma de todos
# os anos. Construida a partir de 06_IA_Anual (ja agregada por ANO x
# INGREDIENTE_ATIVO), evitando somar novamente o volume de produto (PF).
def montar_evolucao_ia_anual(df_ia_anual: pd.DataFrame) -> pd.DataFrame:
    if (
        df_ia_anual.empty
        or "ANO" not in df_ia_anual.columns
        or "INGREDIENTE_ATIVO" not in df_ia_anual.columns
    ):
        return pd.DataFrame()

    valor_col = next(
        (
            c for c in ["VENDA_PONDERADA_IA", "VENDA_ESTIMADA_IA", "VENDA_TOTAL", "VENDA_TOTAL_PF"]
            if c in df_ia_anual.columns
        ),
        None,
    )
    if valor_col is None:
        return pd.DataFrame()

    df = df_ia_anual.copy()
    if not pd.api.types.is_numeric_dtype(df["ANO"]):
        df["ANO"] = converter_numero_br(df["ANO"])
    df["ANO"] = df["ANO"].astype("Int64")
    df[valor_col] = pd.to_numeric(df[valor_col], errors="coerce")
    df = df[df["ANO"].notna() & df["INGREDIENTE_ATIVO"].notna()]
    df = df[df["INGREDIENTE_ATIVO"].astype(str).str.strip() != ""]
    if df.empty:
        return pd.DataFrame()

    pivot = df.pivot_table(
        index="INGREDIENTE_ATIVO", columns="ANO", values=valor_col, aggfunc="sum", fill_value=0
    )
    anos_ordenados = sorted(pivot.columns.tolist())
    pivot = pivot[anos_ordenados]
    colunas_ano = [str(int(a)) for a in anos_ordenados]
    pivot.columns = colunas_ano
    pivot["TOTAL_PERIODO"] = pivot[colunas_ano].sum(axis=1)
    pivot = pivot.reset_index()
    return pivot.sort_values("TOTAL_PERIODO", ascending=False).reset_index(drop=True)


# ---------- Tabela 7: IA_UF ----------
# Venda estimada de cada IA por UF = volume PF da UF do produto x ponderacao da IA.
def montar_ia_uf(df_com_uf: pd.DataFrame, df_com_ia: pd.DataFrame) -> pd.DataFrame:
    if df_com_uf.empty or df_com_ia.empty or "INGREDIENTE_ATIVO" not in df_com_ia.columns:
        return pd.DataFrame()
    cols_ia = ["_CHAVE_REGISTRO", "ANO", "INGREDIENTE_ATIVO"]
    if "PONDERACAO_IA" in df_com_ia.columns:
        cols_ia.append("PONDERACAO_IA")
    df_ia = df_com_ia[cols_ia].drop_duplicates()
    df_ia = df_ia[df_ia["INGREDIENTE_ATIVO"].notna()]

    # Remove colunas que tambem existem em df_ia para evitar sufixos _x/_y no merge
    cols_uf = [
        c for c in df_com_uf.columns
        if c not in {"INGREDIENTE_ATIVO", "PONDERACAO_IA", "MARCA_COMERCIAL"}
    ]
    # Mantem identificadores e volume de UF (produto formulado)
    essenciais = ["_CHAVE_REGISTRO", "ANO", "UF", "REGIAO", "VENDA"]
    for c in essenciais:
        if c in df_com_uf.columns and c not in cols_uf:
            cols_uf.append(c)
    df_uf = df_com_uf[cols_uf].copy()

    df = df_uf.merge(df_ia, on=["_CHAVE_REGISTRO", "ANO"], how="inner")
    if "PONDERACAO_IA" in df.columns:
        df["VENDA"] = df["VENDA"] * df["PONDERACAO_IA"]
    if "INGREDIENTE_ATIVO" not in df.columns or "VENDA" not in df.columns:
        return pd.DataFrame()
    grp = df.groupby(["ANO", "INGREDIENTE_ATIVO", "UF", "REGIAO"], as_index=False)["VENDA"].sum()
    return grp.sort_values(["ANO", "INGREDIENTE_ATIVO", "VENDA"], ascending=[True, True, False]).reset_index(drop=True)


# ---------- Tabela 8: Registro_Resumo ----------
# Resumo por registro/produto: apenas metricas do produto (venda PF no periodo).
# Nao inclui venda ponderada de IA (use 09_IA_Resumo / 06_IA_Anual).
def montar_registro_resumo(
    df_produtos: pd.DataFrame, df_agrofit_uso: pd.DataFrame, df_com_produto: pd.DataFrame
) -> pd.DataFrame:
    base_cols = [c for c in ["_CHAVE_REGISTRO", "NR_REGISTRO", "MARCA_COMERCIAL",
                              "INGREDIENTE_ATIVO", "TITULAR_DE_REGISTRO", "FORMULACAO",
                              "SITUACAO"] if c in df_produtos.columns]
    df = df_produtos[base_cols].copy()

    if not df_agrofit_uso.empty:
        uso_agg = df_agrofit_uso.groupby("_CHAVE_REGISTRO").agg(
            NUM_CULTURAS=("CULTURA", "nunique") if "CULTURA" in df_agrofit_uso.columns else ("_CHAVE_REGISTRO", "count"),
            NUM_PRAGAS=("PRAGA_NOME_CIENTIFICO", "nunique") if "PRAGA_NOME_CIENTIFICO" in df_agrofit_uso.columns else ("_CHAVE_REGISTRO", "count"),
        ).reset_index()
        df = df.merge(uso_agg, on="_CHAVE_REGISTRO", how="left")
    else:
        df["NUM_CULTURAS"] = 0
        df["NUM_PRAGAS"] = 0

    if not df_com_produto.empty and "VENDAS_TOTAIS" in df_com_produto.columns:
        com_agg = df_com_produto.groupby("_CHAVE_REGISTRO").agg(
            PRIMEIRO_ANO_COMERCIALIZACAO=("ANO", "min"),
            ULTIMO_ANO_COMERCIALIZACAO=("ANO", "max"),
            VENDA_TOTAL_PERIODO=("VENDAS_TOTAIS", "sum"),
        ).reset_index()
        df = df.merge(com_agg, on="_CHAVE_REGISTRO", how="left")
    else:
        for c in ["PRIMEIRO_ANO_COMERCIALIZACAO", "ULTIMO_ANO_COMERCIALIZACAO", "VENDA_TOTAL_PERIODO"]:
            df[c] = np.nan

    df["NUM_CULTURAS"] = df["NUM_CULTURAS"].fillna(0).astype(int)
    df["NUM_PRAGAS"] = df["NUM_PRAGAS"].fillna(0).astype(int)
    df = df.drop(columns=["_CHAVE_REGISTRO"])
    return df.sort_values("VENDA_TOTAL_PERIODO", ascending=False, na_position="last").reset_index(drop=True)


# ---------- Tabela 9: IA_Resumo ----------
# Base = perspectiva de comercializacao por IA (04_Comerc_IA), nao o catalogo
# Agrofit. Contagens regulatórias do Agrofit entram so como enriquecimento opcional.
def montar_ia_resumo(
    df_agrofit: pd.DataFrame, df_agrofit_uso: pd.DataFrame, df_com_ia: pd.DataFrame
) -> pd.DataFrame:
    if df_com_ia.empty or "INGREDIENTE_ATIVO" not in df_com_ia.columns:
        return pd.DataFrame()

    df_ia = df_com_ia[
        df_com_ia["INGREDIENTE_ATIVO"].notna()
        & (df_com_ia["INGREDIENTE_ATIVO"].astype(str).str.strip() != "")
    ].copy()
    if df_ia.empty:
        return pd.DataFrame()

    agg_map: Dict[str, Any] = {
        "PRIMEIRO_ANO": ("ANO", "min"),
        "ULTIMO_ANO": ("ANO", "max"),
        "NUM_REGISTROS_COM_VENDAS": ("NR_REGISTRO", "nunique"),
    }
    if "MARCA_COMERCIAL" in df_ia.columns:
        agg_map["NUM_PRODUTOS_COM_VENDAS"] = ("MARCA_COMERCIAL", "nunique")
    if "VENDA_ESTIMADA_IA" in df_ia.columns:
        agg_map["VENDA_ESTIMADA_IA_PERIODO"] = ("VENDA_ESTIMADA_IA", "sum")
    elif "VENDA_PONDERADA_IA" in df_ia.columns:
        agg_map["VENDA_ESTIMADA_IA_PERIODO"] = ("VENDA_PONDERADA_IA", "sum")
    if "PRODUCAO_NACIONAL_PONDERADA_IA" in df_ia.columns:
        agg_map["PRODUCAO_NACIONAL_PONDERADA_IA_PERIODO"] = ("PRODUCAO_NACIONAL_PONDERADA_IA", "sum")
    if "IMPORTACAO_PONDERADA_IA" in df_ia.columns:
        agg_map["IMPORTACAO_PONDERADA_IA_PERIODO"] = ("IMPORTACAO_PONDERADA_IA", "sum")
    if "EXPORTACAO_PONDERADA_IA" in df_ia.columns:
        agg_map["EXPORTACAO_PONDERADA_IA_PERIODO"] = ("EXPORTACAO_PONDERADA_IA", "sum")

    base = df_ia.groupby("INGREDIENTE_ATIVO", as_index=False).agg(**agg_map)

    # Enriquecimento opcional com contagens do Agrofit (mesmo nome de IA)
    if not df_agrofit.empty and "INGREDIENTE_ATIVO" in df_agrofit.columns:
        agro_agg = df_agrofit.groupby("INGREDIENTE_ATIVO").agg(
            NUM_PRODUTOS_AGROFIT=("MARCA_COMERCIAL", "nunique") if "MARCA_COMERCIAL" in df_agrofit.columns else ("_CHAVE_REGISTRO", "nunique"),
            NUM_REGISTROS_AGROFIT=("_CHAVE_REGISTRO", "nunique"),
            NUM_CULTURAS_AGROFIT=("CULTURA", "nunique") if "CULTURA" in df_agrofit.columns else ("_CHAVE_REGISTRO", "nunique"),
            NUM_PRAGAS_AGROFIT=("PRAGA_NOME_CIENTIFICO", "nunique") if "PRAGA_NOME_CIENTIFICO" in df_agrofit.columns else ("_CHAVE_REGISTRO", "nunique"),
        ).reset_index()
        base = base.merge(agro_agg, on="INGREDIENTE_ATIVO", how="left")

    sort_col = "VENDA_ESTIMADA_IA_PERIODO"
    if sort_col in base.columns:
        return base.sort_values(sort_col, ascending=False, na_position="last").reset_index(drop=True)
    return base.reset_index(drop=True)


# ---------- Tabela 9: Agrofit_Sem_Comercializacao ----------
def montar_agrofit_sem_comercializacao(
    df_produtos: pd.DataFrame, df_com_anual: pd.DataFrame
) -> pd.DataFrame:
    chaves_com = set(df_com_anual["_CHAVE_REGISTRO"].unique()) if not df_com_anual.empty else set()
    colunas = [c for c in ["NR_REGISTRO", "MARCA_COMERCIAL", "INGREDIENTE_ATIVO",
                            "TITULAR_DE_REGISTRO", "FORMULACAO", "SITUACAO",
                            "_CHAVE_REGISTRO"] if c in df_produtos.columns]
    df = df_produtos[colunas].copy()
    df = df[~df["_CHAVE_REGISTRO"].isin(chaves_com)]
    df["TEM_COMERCIALIZACAO"] = "NAO"
    return df.drop(columns=["_CHAVE_REGISTRO"]).reset_index(drop=True)


# ---------- Tabela 10: Comercializacao_Sem_Agrofit ----------
def montar_comercializacao_sem_agrofit(
    df_com_anual: pd.DataFrame, df_produtos: pd.DataFrame
) -> pd.DataFrame:
    if df_com_anual.empty:
        return pd.DataFrame()
    chaves_agro = set(df_produtos["_CHAVE_REGISTRO"].unique())
    df = df_com_anual[~df_com_anual["_CHAVE_REGISTRO"].isin(chaves_agro)].copy()
    df["ENCONTRADO_AGROFIT"] = "NAO"
    return df.drop(columns=["_CHAVE_REGISTRO"]).reset_index(drop=True)


# ---------- Tabela 11: Validacao ----------
def montar_validacao(
    df_agrofit: pd.DataFrame, df_produtos: pd.DataFrame, df_com_anual: pd.DataFrame,
    df_agrofit_sem_com: pd.DataFrame, df_com_sem_agro: pd.DataFrame,
) -> pd.DataFrame:
    total_agrofit = df_produtos["_CHAVE_REGISTRO"].nunique() if "_CHAVE_REGISTRO" in df_produtos.columns else 0
    total_com = df_com_anual["_CHAVE_REGISTRO"].nunique() if not df_com_anual.empty else 0
    chaves_agro = set(df_produtos["_CHAVE_REGISTRO"].unique()) if "_CHAVE_REGISTRO" in df_produtos.columns else set()
    chaves_com = set(df_com_anual["_CHAVE_REGISTRO"].unique()) if not df_com_anual.empty else set()

    encontrados = len(chaves_agro & chaves_com)
    apenas_agro = len(chaves_agro - chaves_com)
    apenas_com = len(chaves_com - chaves_agro)
    pct = round(100 * encontrados / total_agrofit, 2) if total_agrofit else 0.0

    registros_dup_agrofit = int(df_agrofit["_CHAVE_REGISTRO"].duplicated().sum()) if "_CHAVE_REGISTRO" in df_agrofit.columns else 0

    linhas = [
        ("TOTAL_REGISTROS_AGROFIT", total_agrofit),
        ("TOTAL_REGISTROS_COMERCIALIZACAO", total_com),
        ("REGISTROS_ENCONTRADOS", encontrados),
        ("REGISTROS_APENAS_AGROFIT", apenas_agro),
        ("REGISTROS_APENAS_COMERCIALIZACAO", apenas_com),
        ("PERCENTUAL_CORRESPONDENCIA", pct),
        ("LINHAS_AGROFIT_DUPLICADAS_MESMO_REGISTRO_CULTURA_PRAGA", registros_dup_agrofit),
        ("REGISTROS_COM_INFORMACOES_INCOMPLETAS_PRODUTOS",
         int(df_produtos.isna().any(axis=1).sum()) if not df_produtos.empty else 0),
    ]
    return pd.DataFrame(linhas, columns=["INDICADOR", "VALOR"])


# =============================================================================
# RESUMO EXECUTIVO (aba 00_Resumo)
# =============================================================================

def montar_resumo_executivo(
    df_agrofit: pd.DataFrame, df_produtos: pd.DataFrame, df_agrofit_uso: pd.DataFrame,
    df_com_anual: pd.DataFrame, df_com_uf: pd.DataFrame, df_ia_anual: pd.DataFrame,
    df_evol_produtos: Optional[pd.DataFrame] = None,
) -> Tuple[pd.DataFrame, Dict[str, pd.DataFrame]]:
    indicadores = []

    def add(nome, valor):
        indicadores.append((nome, valor))

    add("Numero de ingredientes ativos", df_agrofit["INGREDIENTE_ATIVO"].nunique() if "INGREDIENTE_ATIVO" in df_agrofit.columns else 0)
    add("Numero de produtos (marcas comerciais)", df_agrofit["MARCA_COMERCIAL"].nunique() if "MARCA_COMERCIAL" in df_agrofit.columns else 0)
    add("Numero de registros MAPA", df_produtos["_CHAVE_REGISTRO"].nunique() if "_CHAVE_REGISTRO" in df_produtos.columns else 0)
    add("Numero de culturas", df_agrofit["CULTURA"].nunique() if "CULTURA" in df_agrofit.columns else 0)
    add("Numero de pragas", df_agrofit["PRAGA_NOME_CIENTIFICO"].nunique() if "PRAGA_NOME_CIENTIFICO" in df_agrofit.columns else 0)
    add("Numero de registros comercializados", df_com_anual["_CHAVE_REGISTRO"].nunique() if not df_com_anual.empty else 0)
    add("Volume total comercializado PF (soma VENDAS_TOTAIS)", round(df_com_anual["VENDAS_TOTAIS"].sum(), 2) if not df_com_anual.empty and "VENDAS_TOTAIS" in df_com_anual.columns else 0)
    vol_ia = 0
    if not df_ia_anual.empty:
        col_ia = "VENDA_PONDERADA_IA" if "VENDA_PONDERADA_IA" in df_ia_anual.columns else (
            "VENDA_ESTIMADA_IA" if "VENDA_ESTIMADA_IA" in df_ia_anual.columns else (
                "VENDA_TOTAL" if "VENDA_TOTAL" in df_ia_anual.columns else None
            )
        )
        if col_ia:
            vol_ia = round(df_ia_anual[col_ia].sum(), 2)
    add("Volume estimado de IA comercializado (soma ponderada)", vol_ia)
    add("Numero de UFs com comercializacao", df_com_uf["UF"].nunique() if not df_com_uf.empty else 0)
    add("Ano inicial", int(df_com_anual["ANO"].min()) if not df_com_anual.empty and df_com_anual["ANO"].notna().any() else None)
    add("Ano final", int(df_com_anual["ANO"].max()) if not df_com_anual.empty and df_com_anual["ANO"].notna().any() else None)

    df_resumo = pd.DataFrame(indicadores, columns=["INDICADOR", "VALOR"])

    rankings: Dict[str, pd.DataFrame] = {}

    if not df_ia_anual.empty:
        top_ia = (
            df_ia_anual.groupby("INGREDIENTE_ATIVO", as_index=False)["VENDA_TOTAL"].sum()
            .sort_values("VENDA_TOTAL", ascending=False).head(20).reset_index(drop=True)
        )
        rankings["Top20_Ingredientes_Ativos"] = top_ia

        # crescimento: compara primeiro e ultimo ano disponiveis por IA
        anos = sorted(df_ia_anual["ANO"].dropna().unique())
        if len(anos) >= 2:
            ano_ini, ano_fim = anos[0], anos[-1]
            base_ini = df_ia_anual[df_ia_anual["ANO"] == ano_ini].set_index("INGREDIENTE_ATIVO")["VENDA_TOTAL"]
            base_fim = df_ia_anual[df_ia_anual["ANO"] == ano_fim].set_index("INGREDIENTE_ATIVO")["VENDA_TOTAL"]
            cresc = pd.DataFrame({"VENDA_INICIAL": base_ini, "VENDA_FINAL": base_fim}).fillna(0)
            cresc = cresc[cresc["VENDA_INICIAL"] > 0]
            cresc["CRESCIMENTO_PERCENTUAL"] = round(
                100 * (cresc["VENDA_FINAL"] - cresc["VENDA_INICIAL"]) / cresc["VENDA_INICIAL"], 2
            )
            cresc = cresc.reset_index().sort_values("CRESCIMENTO_PERCENTUAL", ascending=False).head(10)
            rankings["Top10_IA_Crescimento"] = cresc

    # Top20_Produtos: unifica o que antes eram dois rankings praticamente
    # identicos (por MARCA_COMERCIAL e por NR_REGISTRO). Reaproveita a tabela
    # WIDE de evolucao anual (10_Evol_Produtos_Anual) para trazer, quando
    # houver mais de um ano de comercializacao, a quebra por ano, o numero de
    # anos com venda e o crescimento % entre o primeiro e o ultimo ano.
    if df_evol_produtos is None and not df_com_anual.empty:
        df_evol_produtos = montar_evolucao_produtos_anual(df_com_anual)

    if df_evol_produtos is not None and not df_evol_produtos.empty:
        top_prod = df_evol_produtos.head(20).copy()

        if "DETENTOR" in df_com_anual.columns and "NR_REGISTRO" in top_prod.columns:
            detentor_map = (
                df_com_anual.dropna(subset=["DETENTOR"])
                .drop_duplicates("NR_REGISTRO")
                .set_index("NR_REGISTRO")["DETENTOR"]
            )
            top_prod.insert(
                min(2, len(top_prod.columns)), "DETENTOR", top_prod["NR_REGISTRO"].map(detentor_map)
            )

        colunas_ano = [c for c in top_prod.columns if c.isdigit()]
        if colunas_ano:
            top_prod["NUM_ANOS_COM_VENDA"] = (top_prod[colunas_ano] > 0).sum(axis=1)
        if len(colunas_ano) >= 2:
            ano_ini, ano_fim = colunas_ano[0], colunas_ano[-1]
            venda_ini = top_prod[ano_ini].astype(float)
            venda_fim = top_prod[ano_fim].astype(float)
            with np.errstate(divide="ignore", invalid="ignore"):
                cresc = np.where(venda_ini > 0, 100 * (venda_fim - venda_ini) / venda_ini, np.nan)
            top_prod[f"CRESCIMENTO_%_{ano_ini}_{ano_fim}"] = np.round(cresc, 2)

        rankings["Top20_Produtos"] = top_prod

    if not df_com_uf.empty:
        top_uf = (
            df_com_uf.groupby("UF", as_index=False)["VENDA"].sum()
            .sort_values("VENDA", ascending=False).head(10).reset_index(drop=True)
        )
        rankings["Top10_UF"] = top_uf

    return df_resumo, rankings


# =============================================================================
# GERACAO DO EXCEL FINAL
# =============================================================================

ESTILO_CABECALHO_FONTE = Font(bold=True, color="FFFFFF", size=11)
ESTILO_CABECALHO_FUNDO = PatternFill("solid", fgColor="2E5B34")
ESTILO_AVISO_FONTE = Font(bold=True, color="9C4221", italic=True)
ESTILO_AVISO_FUNDO = PatternFill("solid", fgColor="FFF3CD")

LIMITE_LINHAS_TABELA_EXCEL = 900_000  # margem de seguranca abaixo do limite do Excel


def _escrever_dataframe_em_aba(ws, df: pd.DataFrame, aviso_topo: Optional[str] = None, nome_tabela: Optional[str] = None):
    linha_inicio = 1
    if aviso_topo:
        ws.cell(row=1, column=1, value=aviso_topo)
        ws.cell(row=1, column=1).font = ESTILO_AVISO_FONTE
        ws.cell(row=1, column=1).fill = ESTILO_AVISO_FUNDO
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max(len(df.columns), 1))
        ws.row_dimensions[1].height = 30
        linha_inicio = 3

    if df is None or df.empty:
        ws.cell(row=linha_inicio, column=1, value="(sem dados para os arquivos fornecidos)")
        return

    df_out = df.copy()
    if len(df_out) > LIMITE_LINHAS_TABELA_EXCEL:
        df_out = df_out.head(LIMITE_LINHAS_TABELA_EXCEL)

    for j, col in enumerate(df_out.columns, start=1):
        c = ws.cell(row=linha_inicio, column=j, value=str(col))
        c.font = ESTILO_CABECALHO_FONTE
        c.fill = ESTILO_CABECALHO_FUNDO
        c.alignment = Alignment(horizontal="center", vertical="center")

    for i, (_, row) in enumerate(df_out.iterrows(), start=linha_inicio + 1):
        for j, val in enumerate(row, start=1):
            if pd.isna(val):
                val = None
            elif isinstance(val, (np.integer,)):
                val = int(val)
            elif isinstance(val, (np.floating,)):
                val = float(val)
            ws.cell(row=i, column=j, value=val)

    ws.freeze_panes = ws.cell(row=linha_inicio + 1, column=1)

    ultima_linha = linha_inicio + len(df_out)
    ultima_coluna = len(df_out.columns)
    if nome_tabela and len(df_out) > 0 and ultima_linha > linha_inicio:
        try:
            ref = f"{get_column_letter(1)}{linha_inicio}:{get_column_letter(ultima_coluna)}{ultima_linha}"
            tabela = Table(displayName=nome_tabela, ref=ref)
            tabela.tableStyleInfo = TableStyleInfo(
                name="TableStyleMedium2", showRowStripes=True, showFirstColumn=False
            )
            ws.add_table(tabela)
        except Exception as e:  # noqa: BLE001
            LOGGER.warning("Nao foi possivel criar tabela formatada em %s: %s", nome_tabela, e)

    for j, col in enumerate(df_out.columns, start=1):
        try:
            maior = max(
                [len(str(col))] + [len(str(v)) for v in df_out.iloc[:, j - 1].head(200)]
            )
        except Exception:
            maior = 15
        ws.column_dimensions[get_column_letter(j)].width = min(max(maior + 2, 10), 45)


def gerar_excel(
    caminho_saida: Path,
    tabelas: Dict[str, pd.DataFrame],
    df_resumo: pd.DataFrame,
    rankings: Dict[str, pd.DataFrame],
) -> None:
    wb = Workbook()
    wb.remove(wb.active)

    # ---- 00_Resumo ----
    ws = wb.create_sheet("00_Resumo")
    _escrever_dataframe_em_aba(ws, df_resumo, nome_tabela="TabResumo")
    linha_atual = len(df_resumo) + 4
    for nome_ranking, df_rank in rankings.items():
        ws.cell(row=linha_atual, column=1, value=nome_ranking.replace("_", " ")).font = Font(bold=True, size=12)
        linha_atual += 1
        ws_range_inicio = linha_atual
        for j, col in enumerate(df_rank.columns, start=1):
            c = ws.cell(row=linha_atual, column=j, value=str(col))
            c.font = ESTILO_CABECALHO_FONTE
            c.fill = ESTILO_CABECALHO_FUNDO
        linha_atual += 1
        for _, row in df_rank.iterrows():
            for j, val in enumerate(row, start=1):
                if pd.isna(val):
                    val = None
                ws.cell(row=linha_atual, column=j, value=val)
            linha_atual += 1
        linha_atual += 2

    # ---- Abas numeradas principais ----
    abas_numeradas = [
        ("01_Produtos", "Produtos", "TabProdutos"),
        ("02_Agrofit_Uso", "Agrofit_Uso", "TabAgrofitUso"),
        ("03_Comerc_Produto", "Comercializacao_Produto", "TabComercProduto"),
        ("04_Comerc_IA", "Comercializacao_IA", "TabComercIA"),
        ("05_Comerc_UF", "Comercializacao_UF", "TabComercUF"),
        ("06_IA_Anual", "IA_Anual", "TabIAAnual"),
        ("07_IA_UF", "IA_UF", "TabIAUF"),
        ("08_Registro_Resumo", "Registro_Resumo", "TabRegistroResumo"),
        ("09_IA_Resumo", "IA_Resumo", "TabIAResumo"),
        ("10_Evol_Produtos_Anual", "Evolucao_Produtos_Anual", "TabEvolProdutosAnual"),
        ("11_Evol_IA_Anual", "Evolucao_IA_Anual", "TabEvolIAAnual"),
        ("12_Agrofit_Sem_Vendas", "Agrofit_Sem_Comercializacao", "TabAgrofitSemVendas"),
        ("13_Vendas_Sem_Agrofit", "Comercializacao_Sem_Agrofit", "TabVendasSemAgrofit"),
        ("14_Validacao", "Validacao", "TabValidacao"),
    ]
    for nome_aba, chave_tabela, nome_tab in abas_numeradas:
        ws = wb.create_sheet(nome_aba)
        _escrever_dataframe_em_aba(ws, tabelas.get(chave_tabela, pd.DataFrame()), nome_tabela=nome_tab)

    # Graficos desabilitados a pedido (usar tabelas agregadas para visualizacao externa)
    wb.save(caminho_saida)


def inserir_graficos(wb: Workbook, ws_graf, tabelas: Dict[str, pd.DataFrame]) -> None:
    """Cria graficos baseados SOMENTE nas tabelas agregadas (nunca na
    tabela Agrofit detalhada), conforme a regra de negocio do projeto."""
    ws_graf.cell(row=1, column=1, value="Graficos gerados a partir das tabelas agregadas (nunca da tabela detalhada).").font = Font(bold=True, italic=True)

    ancora_linha = 3
    largura_bloco = 15

    def _prep_aba_dados(nome_aba: str, df: pd.DataFrame) -> Optional[Any]:
        if df is None or df.empty:
            return None
        ws = wb.create_sheet(nome_aba)
        for j, col in enumerate(df.columns, start=1):
            ws.cell(row=1, column=j, value=str(col))
        for i, (_, row) in enumerate(df.iterrows(), start=2):
            for j, val in enumerate(row, start=1):
                if pd.isna(val):
                    val = None
                ws.cell(row=i, column=j, value=val)
        ws.sheet_state = "hidden"
        return ws

    # 1) Evolucao das vendas totais por ano (produto PF)
    df_com_anual = tabelas.get("Comercializacao_Anual", pd.DataFrame())
    if not df_com_anual.empty and "ANO" in df_com_anual.columns and "VENDAS_TOTAIS" in df_com_anual.columns:
        evol = df_com_anual.groupby("ANO", as_index=False)[["VENDAS_TOTAIS"]].sum().sort_values("ANO")
        ws_dados = _prep_aba_dados("_dados_evol_anual", evol)
        if ws_dados is not None:
            chart = LineChart()
            chart.title = "Evolucao das vendas totais por ano (produto PF)"
            chart.y_axis.title = "Volume"
            chart.x_axis.title = "Ano"
            dados = Reference(ws_dados, min_col=2, max_col=2, min_row=1, max_row=len(evol) + 1)
            categorias = Reference(ws_dados, min_col=1, min_row=2, max_row=len(evol) + 1)
            chart.add_data(dados, titles_from_data=True)
            chart.set_categories(categorias)
            ws_graf.add_chart(chart, f"A{ancora_linha}")
            ancora_linha += largura_bloco

    # Evolucao da venda estimada de IA (a partir de IA_Anual)
    df_ia_chart = tabelas.get("IA_Anual", pd.DataFrame())
    col_ia_vol = next(
        (c for c in ["VENDA_PONDERADA_IA", "VENDA_ESTIMADA_IA", "VENDA_TOTAL"] if c in df_ia_chart.columns),
        None,
    )
    if not df_ia_chart.empty and "ANO" in df_ia_chart.columns and col_ia_vol:
        evol_ia = df_ia_chart.groupby("ANO", as_index=False)[[col_ia_vol]].sum().sort_values("ANO")
        ws_dados = _prep_aba_dados("_dados_evol_ia", evol_ia)
        if ws_dados is not None:
            chart2 = LineChart()
            chart2.title = "Evolucao da venda estimada de IA"
            dados2 = Reference(ws_dados, min_col=2, max_col=2, min_row=1, max_row=len(evol_ia) + 1)
            categorias = Reference(ws_dados, min_col=1, min_row=2, max_row=len(evol_ia) + 1)
            chart2.add_data(dados2, titles_from_data=True)
            chart2.set_categories(categorias)
            ws_graf.add_chart(chart2, f"A{ancora_linha}")
            ancora_linha += largura_bloco

    # 2) Top 20 ingredientes ativos
    df_ia_anual = tabelas.get("IA_Anual", pd.DataFrame())
    if not df_ia_anual.empty:
        top_ia = df_ia_anual.groupby("INGREDIENTE_ATIVO", as_index=False)["VENDA_TOTAL"].sum().sort_values("VENDA_TOTAL", ascending=False).head(20)
        ws_dados = _prep_aba_dados("_dados_top_ia", top_ia)
        if ws_dados is not None:
            chart = BarChart()
            chart.title = "Top 20 ingredientes ativos"
            chart.type = "bar"
            dados = Reference(ws_dados, min_col=2, min_row=1, max_row=len(top_ia) + 1)
            categorias = Reference(ws_dados, min_col=1, min_row=2, max_row=len(top_ia) + 1)
            chart.add_data(dados, titles_from_data=True)
            chart.set_categories(categorias)
            ws_graf.add_chart(chart, f"A{ancora_linha}")
            ancora_linha += largura_bloco

        # 4) Evolucao dos principais IA (top 5) ao longo dos anos
        top5_nomes = top_ia["INGREDIENTE_ATIVO"].head(5).tolist()
        evol_top5 = df_ia_anual[df_ia_anual["INGREDIENTE_ATIVO"].isin(top5_nomes)]
        if not evol_top5.empty:
            pivot = evol_top5.pivot_table(index="ANO", columns="INGREDIENTE_ATIVO", values="VENDA_TOTAL", aggfunc="sum").fillna(0).reset_index().sort_values("ANO")
            ws_dados = _prep_aba_dados("_dados_evol_top5_ia", pivot)
            if ws_dados is not None:
                chart = LineChart()
                chart.title = "Evolucao dos principais ingredientes ativos"
                dados = Reference(ws_dados, min_col=2, max_col=pivot.shape[1], min_row=1, max_row=len(pivot) + 1)
                categorias = Reference(ws_dados, min_col=1, min_row=2, max_row=len(pivot) + 1)
                chart.add_data(dados, titles_from_data=True)
                chart.set_categories(categorias)
                ws_graf.add_chart(chart, f"A{ancora_linha}")
                ancora_linha += largura_bloco

    # 3) Top 20 produtos
    if not df_com_anual.empty and "MARCA_COMERCIAL" in df_com_anual.columns:
        top_prod = df_com_anual.groupby("MARCA_COMERCIAL", as_index=False)["VENDAS_TOTAIS"].sum().sort_values("VENDAS_TOTAIS", ascending=False).head(20)
        ws_dados = _prep_aba_dados("_dados_top_prod", top_prod)
        if ws_dados is not None:
            chart = BarChart()
            chart.title = "Top 20 produtos"
            chart.type = "bar"
            dados = Reference(ws_dados, min_col=2, min_row=1, max_row=len(top_prod) + 1)
            categorias = Reference(ws_dados, min_col=1, min_row=2, max_row=len(top_prod) + 1)
            chart.add_data(dados, titles_from_data=True)
            chart.set_categories(categorias)
            ws_graf.add_chart(chart, f"A{ancora_linha}")
            ancora_linha += largura_bloco

    # 5) Vendas por regiao / 6) Vendas por UF
    df_com_uf = tabelas.get("Comercializacao_UF", pd.DataFrame())
    if not df_com_uf.empty:
        por_regiao = df_com_uf.groupby("REGIAO", as_index=False)["VENDA"].sum().sort_values("VENDA", ascending=False)
        ws_dados = _prep_aba_dados("_dados_regiao", por_regiao)
        if ws_dados is not None:
            chart = BarChart()
            chart.title = "Vendas por regiao"
            dados = Reference(ws_dados, min_col=2, min_row=1, max_row=len(por_regiao) + 1)
            categorias = Reference(ws_dados, min_col=1, min_row=2, max_row=len(por_regiao) + 1)
            chart.add_data(dados, titles_from_data=True)
            chart.set_categories(categorias)
            ws_graf.add_chart(chart, f"A{ancora_linha}")
            ancora_linha += largura_bloco

        por_uf = df_com_uf.groupby("UF", as_index=False)["VENDA"].sum().sort_values("VENDA", ascending=False).head(10)
        ws_dados = _prep_aba_dados("_dados_uf", por_uf)
        if ws_dados is not None:
            chart = BarChart()
            chart.title = "Vendas por UF (Top 10)"
            dados = Reference(ws_dados, min_col=2, min_row=1, max_row=len(por_uf) + 1)
            categorias = Reference(ws_dados, min_col=1, min_row=2, max_row=len(por_uf) + 1)
            chart.add_data(dados, titles_from_data=True)
            chart.set_categories(categorias)
            ws_graf.add_chart(chart, f"A{ancora_linha}")
            ancora_linha += largura_bloco

    # 7) Participacao dos principais IA (pizza aproximada via barra, mais legivel em bases grandes)
    if not df_ia_anual.empty:
        participacao = df_ia_anual.groupby("INGREDIENTE_ATIVO", as_index=False)["VENDA_TOTAL"].sum().sort_values("VENDA_TOTAL", ascending=False).head(10)
        ws_dados = _prep_aba_dados("_dados_participacao_ia", participacao)
        if ws_dados is not None:
            chart = BarChart()
            chart.title = "Participacao dos principais IA (Top 10)"
            chart.type = "bar"
            dados = Reference(ws_dados, min_col=2, min_row=1, max_row=len(participacao) + 1)
            categorias = Reference(ws_dados, min_col=1, min_row=2, max_row=len(participacao) + 1)
            chart.add_data(dados, titles_from_data=True)
            chart.set_categories(categorias)
            ws_graf.add_chart(chart, f"A{ancora_linha}")
            ancora_linha += largura_bloco


# =============================================================================
# ORQUESTRACAO PRINCIPAL
# =============================================================================

def preparar_pastas() -> None:
    for pasta in [DADOS_DIR, SAIDA_DIR, LOGS_DIR]:
        pasta.mkdir(parents=True, exist_ok=True)


def main() -> None:
    tempo_total_inicio = time.time()
    preparar_pastas()
    configurar_logging()

    titulo("ANALISE INTEGRADA AGROFIT x COMERCIALIZACAO DE AGROTOXICOS")
    LOGGER.info("Execucao iniciada em %s", datetime.now().isoformat())

    # ---------------- 1) Localizar e carregar Agrofit ----------------
    caminho_agrofit = localizar_arquivo_agrofit()
    if caminho_agrofit is None:
        erro(
            f"Arquivo do Agrofit nao encontrado em {DADOS_DIR}. "
            f"Coloque o arquivo '{ARQUIVO_AGROFIT_PADRAO}' dentro da pasta 'dados/' e execute novamente."
        )
        if os.environ.get("AGROFIT_NO_INPUT"):
            return
        input("\nPressione Enter para sair...")
        return

    with Cronometro(f"Carregar base Agrofit ({caminho_agrofit.name})"):
        df_agrofit = carregar_agrofit(caminho_agrofit)

    if df_agrofit.empty:
        erro("A base do Agrofit foi carregada, porem esta vazia ou sem registros validos. Encerrando.")
        if os.environ.get("AGROFIT_NO_INPUT"):
            return
        input("\nPressione Enter para sair...")
        return

    # ---------------- 2) Detectar e carregar comercializacao ----------------
    arquivos_com = detectar_arquivos_comercializacao(caminho_agrofit)
    if not arquivos_com:
        aviso(
            f"Nenhum arquivo de comercializacao encontrado em {DADOS_DIR}. "
            "O relatorio sera gerado apenas com as informacoes regulatorias do Agrofit."
        )
    else:
        print(f"\n{len(arquivos_com)} arquivo(s) de comercializacao detectado(s):")
        for a in arquivos_com:
            print(f"   - {a.name}")

    with Cronometro("Carregar bases de comercializacao"):
        df_vendas_bruto = carregar_todas_comercializacoes(arquivos_com)

    if not df_vendas_bruto.empty:
        with Cronometro("Preparar campos numericos de comercializacao"):
            df_vendas_bruto = preparar_numericos_comercializacao(df_vendas_bruto)

    # ---------------- 3) Construcao das tabelas analiticas ----------------
    # Perspectiva PRODUTO (consolidado) e perspectiva IA (separada) lado a lado.
    with Cronometro("Montar Tabela 1 - Produtos"):
        df_produtos = montar_tabela_produtos(df_agrofit)
        ok(f"{len(df_produtos):,} registros/produtos unicos")

    with Cronometro("Montar Tabela 2 - Agrofit_Uso"):
        df_agrofit_uso = montar_tabela_agrofit_uso(df_agrofit)
        ok(f"{len(df_agrofit_uso):,} combinacoes Registro x Cultura x Praga")

    with Cronometro("Montar Tabela 3 - Comercializacao_Produto (Registro x Ano)"):
        df_com_produto = montar_comercializacao_produto(df_vendas_bruto, df_produtos)
        ok(f"{len(df_com_produto):,} linhas (produto consolidado)")

    with Cronometro("Montar Tabela 4 - Comercializacao_IA (Registro x Ano x IA)"):
        df_com_ia = montar_comercializacao_ia(df_vendas_bruto, df_produtos)
        ok(f"{len(df_com_ia):,} linhas (IA separada)")

    with Cronometro("Montar Tabela 5 - Comercializacao_UF"):
        df_com_uf = montar_comercializacao_uf(df_vendas_bruto, df_produtos)
        ok(f"{len(df_com_uf):,} linhas (Registro x Ano x UF)")

    with Cronometro("Montar Tabela 6 - IA_Anual"):
        df_ia_anual = montar_ia_anual(df_com_ia)

    with Cronometro("Montar Tabela 10 - Evolucao_Produtos_Anual"):
        df_evol_produtos = montar_evolucao_produtos_anual(df_com_produto)
        ok(f"{len(df_evol_produtos):,} produtos com serie historica anual")

    with Cronometro("Montar Tabela 11 - Evolucao_IA_Anual"):
        df_evol_ia = montar_evolucao_ia_anual(df_ia_anual)
        ok(f"{len(df_evol_ia):,} ingredientes ativos com serie historica anual")

    with Cronometro("Montar Tabela 7 - IA_UF"):
        df_ia_uf = montar_ia_uf(df_com_uf, df_com_ia)

    with Cronometro("Montar Tabela 8 - Registro_Resumo"):
        df_registro_resumo = montar_registro_resumo(df_produtos, df_agrofit_uso, df_com_produto)

    with Cronometro("Montar Tabela 9 - IA_Resumo"):
        df_ia_resumo = montar_ia_resumo(df_agrofit, df_agrofit_uso, df_com_ia)

    with Cronometro("Montar Tabela 10 - Agrofit_Sem_Comercializacao"):
        df_agrofit_sem_com = montar_agrofit_sem_comercializacao(df_produtos, df_com_produto)
        ok(f"{len(df_agrofit_sem_com):,} registros sem comercializacao encontrada")

    with Cronometro("Montar Tabela 11 - Comercializacao_Sem_Agrofit"):
        df_com_sem_agro = montar_comercializacao_sem_agrofit(df_com_produto, df_produtos)
        if len(df_com_sem_agro) > 0:
            aviso(f"{len(df_com_sem_agro):,} registros de comercializacao nao encontrados no Agrofit")

    with Cronometro("Montar Tabela 12 - Validacao"):
        df_validacao = montar_validacao(df_agrofit, df_produtos, df_com_produto, df_agrofit_sem_com, df_com_sem_agro)

    with Cronometro("Montar resumo executivo e rankings"):
        df_resumo, rankings = montar_resumo_executivo(
            df_agrofit, df_produtos, df_agrofit_uso, df_com_produto, df_com_uf, df_ia_anual,
            df_evol_produtos,
        )

    tabelas = {
        "Produtos": df_produtos.drop(columns=["_CHAVE_REGISTRO"], errors="ignore"),
        "Agrofit_Uso": df_agrofit_uso.drop(columns=["_CHAVE_REGISTRO"], errors="ignore"),
        "Comercializacao_Produto": df_com_produto.drop(columns=["_CHAVE_REGISTRO"], errors="ignore"),
        "Comercializacao_IA": df_com_ia.drop(columns=["_CHAVE_REGISTRO"], errors="ignore"),
        "Comercializacao_UF": df_com_uf.drop(columns=["_CHAVE_REGISTRO"], errors="ignore"),
        "IA_Anual": df_ia_anual,
        "Evolucao_Produtos_Anual": df_evol_produtos,
        "Evolucao_IA_Anual": df_evol_ia,
        "IA_UF": df_ia_uf,
        "Registro_Resumo": df_registro_resumo,
        "IA_Resumo": df_ia_resumo,
        "Agrofit_Sem_Comercializacao": df_agrofit_sem_com,
        "Comercializacao_Sem_Agrofit": df_com_sem_agro,
        "Validacao": df_validacao,
        # alias interno para graficos/resumo que ainda referenciam o nome antigo
        "Comercializacao_Anual": df_com_produto.drop(columns=["_CHAVE_REGISTRO"], errors="ignore"),
    }

    # ---------------- 4) Geracao do Excel ----------------
    caminho_saida = SAIDA_DIR / ARQUIVO_SAIDA
    with Cronometro(f"Gerar arquivo Excel final ({ARQUIVO_SAIDA})"):
        gerar_excel(caminho_saida, tabelas, df_resumo, rankings)

    tempo_total = time.time() - tempo_total_inicio
    LOGGER.info("Execucao concluida em %.2fs", tempo_total)

    titulo("PROCESSAMENTO CONCLUIDO")
    print(f"Arquivo gerado : {caminho_saida}")
    print(f"Log detalhado  : {LOGS_DIR / ARQUIVO_LOG}")
    print(f"Tempo total    : {tempo_legivel(tempo_total)}\n")

    pct = df_validacao.loc[df_validacao['INDICADOR'] == 'PERCENTUAL_CORRESPONDENCIA', 'VALOR']
    if len(pct):
        print(f"Percentual de correspondencia Agrofit x Comercializacao: {pct.iloc[0]}%")
    print()


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        aviso("Interrompido pelo usuario.")
    except Exception as e:  # noqa: BLE001
        erro(f"Erro inesperado: {e}")
        LOGGER.error("Erro inesperado: %s\n%s", e, traceback.format_exc())
        traceback.print_exc()
        if not os.environ.get("AGROFIT_NO_INPUT"):
            input("\nPressione Enter para sair...")
