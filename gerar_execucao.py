"""
gerar_execucao.py
------------------
Cruza os arquivos EMENDASESTADUAIS<sufixo>.xlsx com a planilha de referência
VW_SG_V2_EP_INDIC_RECURSOS_TW.xlsx (coluna NUMERO_SIAFI) e gera, para cada um,
um arquivo execucao<sufixo>.xlsx contendo apenas as linhas cujo número SIAFI
também aparece na referência.

REQUISITOS (instalar se faltar em algum computador):
    pip install pandas openpyxl

Como usar:
    python gerar_execucao.py

O que o script faz, em ordem:
    1. Localiza todos os arquivos "EMENDASESTADUAIS*.xlsx" na pasta de upload.
    2. Lê "VW_SG_V2_EP_INDIC_RECURSOS_TW.xlsx" e monta o conjunto de números
       válidos a partir da coluna NUMERO_SIAFI.
    3. Para cada EMENDASESTADUAIS<sufixo>.xlsx:
       a. Remove a primeira linha e a primeira coluna.
       b. Remove linhas e colunas totalmente vazias.
       c. Usa a linha seguinte como cabeçalho.
       d. Filtra mantendo apenas as linhas cujo número SIAFI está na referência.
       e. Salva como execucao<sufixo>.xlsx na mesma pasta.
    4. Apaga o arquivo EMENDASESTADUAIS<sufixo>.xlsx original (somente os que
       foram processados com sucesso).
    5. Imprime um relatório com a quantidade de linhas que coincidiram.
"""

import re
import sys
from pathlib import Path

try:
    import pandas as pd
except ImportError:
    print("Faltam dependências. Rode: pip install pandas openpyxl")
    sys.exit(1)

# ---------------------------------------------------------------------------
# CONFIGURAÇÃO — ajuste aqui se necessário
# ---------------------------------------------------------------------------
PASTA_UPLOAD = Path(r"G:\Meu Drive\CGE\bi_atualizacao\portal_emendas_estaduais\upload")
ARQUIVO_REFERENCIA_NOME = "VW_SG_V2_EP_INDIC_RECURSOS_TW_202608041638.xlsx"
COLUNA_REFERENCIA = "NUMERO_SIAFI"
PADRAO_ENTRADA = "EMENDASESTADUAIS*.xlsx"
PREFIXO_SAIDA = "execucao"


def normalizar_siafi(valor):
    """Normaliza um número SIAFI para comparação (só dígitos, sem .0 de float)."""
    if pd.isna(valor):
        return None
    texto = str(valor).strip()
    if texto.endswith(".0"):
        texto = texto[:-2]
    texto = re.sub(r"\D", "", texto)
    return texto or None


def encontrar_coluna_siafi(df):
    """Localiza a coluna que contém o número SIAFI na planilha de emendas."""
    for col in df.columns:
        if "SIAFI" in str(col).upper():
            return col
    raise ValueError(
        "Não encontrei nenhuma coluna com 'SIAFI' no nome. "
        f"Colunas disponíveis: {list(df.columns)}"
    )


def limpar_planilha(caminho_arquivo):
    """Lê a planilha bruta, remove 1ª linha/coluna e linhas/colunas vazias."""
    bruto = pd.read_excel(caminho_arquivo, header=None)

    # remove a primeira linha e a primeira coluna
    bruto = bruto.iloc[1:, 1:]

    # remove linhas e colunas totalmente vazias
    bruto = bruto.dropna(how="all", axis=0)
    bruto = bruto.dropna(how="all", axis=1)

    # a próxima linha restante vira o cabeçalho
    bruto.columns = bruto.iloc[0]
    df = bruto.iloc[1:].reset_index(drop=True)
    df.columns = [str(c).strip() for c in df.columns]

    return df


def carregar_siafi_referencia(pasta):
    caminho_ref = pasta / ARQUIVO_REFERENCIA_NOME
    if not caminho_ref.exists():
        raise FileNotFoundError(f"Arquivo de referência não encontrado: {caminho_ref}")

    df_ref = pd.read_excel(caminho_ref)
    if COLUNA_REFERENCIA not in df_ref.columns:
        raise ValueError(
            f"A coluna '{COLUNA_REFERENCIA}' não foi encontrada em "
            f"{ARQUIVO_REFERENCIA_NOME}. Colunas disponíveis: {list(df_ref.columns)}"
        )

    siafi_validos = {normalizar_siafi(v) for v in df_ref[COLUNA_REFERENCIA]}
    siafi_validos.discard(None)
    return siafi_validos


def extrair_sufixo(nome_arquivo):
    """Extrai o sufixo (ex.: ano) do nome EMENDASESTADUAIS<sufixo>.xlsx"""
    m = re.match(r"EMENDASESTADUAIS(.*)\.xlsx$", nome_arquivo, re.IGNORECASE)
    return m.group(1) if m else Path(nome_arquivo).stem


def processar_arquivo(caminho_arquivo, siafi_validos, pasta_saida):
    df = limpar_planilha(caminho_arquivo)
    coluna_siafi = encontrar_coluna_siafi(df)

    df["_siafi_norm"] = df[coluna_siafi].apply(normalizar_siafi)
    total_linhas = len(df)

    df_filtrado = df[df["_siafi_norm"].isin(siafi_validos)].drop(columns=["_siafi_norm"])
    qtd_match = len(df_filtrado)

    sufixo = extrair_sufixo(caminho_arquivo.name)
    caminho_saida = pasta_saida / f"{PREFIXO_SAIDA}{sufixo}.xlsx"
    df_filtrado.to_excel(caminho_saida, index=False)

    return {
        "arquivo_origem": caminho_arquivo.name,
        "arquivo_saida": caminho_saida.name,
        "total_linhas_origem": total_linhas,
        "linhas_coincidentes": qtd_match,
        "coluna_siafi_usada": coluna_siafi,
    }


def main():
    if not PASTA_UPLOAD.exists():
        print(f"ERRO: pasta não encontrada -> {PASTA_UPLOAD}")
        sys.exit(1)

    print(f"Pasta de trabalho: {PASTA_UPLOAD}\n")

    try:
        siafi_validos = carregar_siafi_referencia(PASTA_UPLOAD)
    except (FileNotFoundError, ValueError) as e:
        print(f"ERRO: {e}")
        sys.exit(1)

    print(f"Referência carregada: {len(siafi_validos)} números SIAFI únicos "
          f"em {ARQUIVO_REFERENCIA_NOME}\n")

    arquivos = sorted(PASTA_UPLOAD.glob(PADRAO_ENTRADA))
    if not arquivos:
        print(f"Nenhum arquivo encontrado com o padrão '{PADRAO_ENTRADA}' em {PASTA_UPLOAD}")
        sys.exit(0)

    relatorios = []
    for arquivo in arquivos:
        try:
            resultado = processar_arquivo(arquivo, siafi_validos, PASTA_UPLOAD)
            relatorios.append(resultado)
        except Exception as e:
            print(f"ERRO ao processar {arquivo.name}: {e}")
            continue

    # apaga os arquivos originais EMENDASESTADUAIS* processados com sucesso
    nomes_processados = {r["arquivo_origem"] for r in relatorios}
    for arquivo in arquivos:
        if arquivo.name in nomes_processados:
            arquivo.unlink()

    # ------------------------------------------------------------------
    # RELATÓRIO
    # ------------------------------------------------------------------
    print("=" * 70)
    print("RELATÓRIO DE EXECUÇÃO")
    print("=" * 70)
    for r in relatorios:
        pct = (r["linhas_coincidentes"] / r["total_linhas_origem"] * 100
               if r["total_linhas_origem"] else 0)
        print(f"\nArquivo origem : {r['arquivo_origem']}")
        print(f"Arquivo salvo  : {r['arquivo_saida']}")
        print(f"Coluna SIAFI   : {r['coluna_siafi_usada']}")
        print(f"Total de linhas lidas : {r['total_linhas_origem']}")
        print(f"Linhas coincidentes   : {r['linhas_coincidentes']} ({pct:.1f}%)")

    print("\n" + "=" * 70)
    total_geral = sum(r["linhas_coincidentes"] for r in relatorios)
    print(f"TOTAL GERAL DE LINHAS COINCIDENTES: {total_geral}")
    print("=" * 70)


if __name__ == "__main__":
    main()
