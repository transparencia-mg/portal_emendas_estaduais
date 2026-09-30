"""
gerar_execucao.py
------------------
Cruza os arquivos EMENDASESTADUAIS<sufixo>.xlsx com a planilha de referência
VW_SG_V2_EP_INDIC_RECURSOS_TW.xlsx (coluna NUMERO_SIAFI) e gera, para cada um,
um arquivo execucao<sufixo>.xlsx contendo apenas as linhas cujo número SIAFI
também aparece na referência.

Funciona em qualquer computador (Windows, Mac ou Linux) com Python 3.8+.
As dependências (pandas e openpyxl) são instaladas automaticamente se faltarem.

Como usar:
    python gerar_execucao.py                      -> encontra a pasta sozinho
    python gerar_execucao.py "C:\\caminho\\upload"  -> usa a pasta informada

Como o script encontra a pasta de trabalho (nesta ordem):
    1. Caminho passado na linha de comando.
    2. A própria pasta onde o script está salvo (se o arquivo de referência
       estiver lá).
    3. A pasta "CGE/bi_atualizacao/portal_emendas_estaduais/upload" dentro do
       Google Drive, em qualquer letra de unidade ("Meu Drive" ou "My Drive"),
       inclusive no Mac.
    4. Se nada der certo, abre uma janela para você escolher a pasta
       (ou pede para digitar o caminho).

O que o script faz, em ordem:
    1. Localiza todos os arquivos "EMENDASESTADUAIS*.xlsx" na pasta.
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

import os
import re
import string
import subprocess
import sys
from pathlib import Path

# Garante acentos corretos no console do Windows
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8")
    except Exception:
        pass


# ---------------------------------------------------------------------------
# DEPENDÊNCIAS — instala automaticamente se faltar
# ---------------------------------------------------------------------------
def garantir_dependencias():
    faltando = []
    for modulo in ("pandas", "openpyxl"):
        try:
            __import__(modulo)
        except ImportError:
            faltando.append(modulo)

    if not faltando:
        return

    print(f"Instalando dependências que faltam: {', '.join(faltando)} ...")
    try:
        subprocess.check_call(
            [sys.executable, "-m", "pip", "install", "--user", *faltando]
        )
    except Exception as e:
        print(f"\nNão consegui instalar automaticamente ({e}).")
        print(f"Rode manualmente:  {sys.executable} -m pip install {' '.join(faltando)}")
        pausar_e_sair(1)

    # recarrega o caminho de pacotes do usuário recém-instalados
    import site
    import importlib
    try:
        site.addsitedir(site.getusersitepackages())
    except Exception:
        pass
    importlib.invalidate_caches()


def pausar_e_sair(codigo=0):
    """Mantém a janela aberta quando o script é aberto com duplo clique."""
    if "--sem-pausa" not in sys.argv:
        try:
            input("\nPressione Enter para fechar...")
        except EOFError:
            pass
    sys.exit(codigo)


garantir_dependencias()
import pandas as pd  # noqa: E402


# ---------------------------------------------------------------------------
# CONFIGURAÇÃO — ajuste aqui se necessário
# ---------------------------------------------------------------------------
SUBPASTA_NO_DRIVE = Path("CGE") / "bi_atualizacao" / "portal_emendas_estaduais" / "upload"
NOMES_RAIZ_DRIVE = ("Meu Drive", "My Drive")
ARQUIVO_REFERENCIA_NOME = "VW_SG_V2_EP_INDIC_RECURSOS_TW.xlsx"
COLUNA_REFERENCIA = "NUMERO_SIAFI"
PADRAO_ENTRADA = "EMENDASESTADUAIS*.xlsx"
PREFIXO_SAIDA = "execucao"


# ---------------------------------------------------------------------------
# LOCALIZAÇÃO DA PASTA DE TRABALHO
# ---------------------------------------------------------------------------
def pasta_valida(pasta):
    return pasta is not None and (Path(pasta) / ARQUIVO_REFERENCIA_NOME).exists()


def candidatas_google_drive():
    """Lista possíveis locais da pasta de upload dentro do Google Drive."""
    raizes = []

    # Windows: Google Drive para computador monta como unidade (G:, H:, ...)
    if os.name == "nt":
        for letra in string.ascii_uppercase:
            for nome in NOMES_RAIZ_DRIVE:
                raizes.append(Path(f"{letra}:\\") / nome)

    home = Path.home()
    # Mac: ~/Library/CloudStorage/GoogleDrive-<email>/Meu Drive
    cloud = home / "Library" / "CloudStorage"
    if cloud.exists():
        for conta in cloud.glob("GoogleDrive-*"):
            for nome in NOMES_RAIZ_DRIVE:
                raizes.append(conta / nome)

    # Instalações antigas / outros sistemas
    for nome in NOMES_RAIZ_DRIVE:
        raizes.append(home / "Google Drive" / nome)
    raizes.append(home / "Google Drive")

    for raiz in raizes:
        try:
            if raiz.exists():
                yield raiz / SUBPASTA_NO_DRIVE
        except OSError:
            # unidades vazias (ex.: leitor de cartão) podem gerar erro
            continue


def escolher_pasta_manualmente():
    """Abre uma janela para escolher a pasta; se não houver interface, pede no console."""
    try:
        import tkinter as tk
        from tkinter import filedialog

        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        escolhida = filedialog.askdirectory(
            title=f"Selecione a pasta que contém {ARQUIVO_REFERENCIA_NOME}"
        )
        root.destroy()
        if escolhida:
            return Path(escolhida)
    except Exception:
        pass

    try:
        texto = input("Digite o caminho da pasta de upload: ").strip().strip('"')
        return Path(texto) if texto else None
    except EOFError:
        return None


def localizar_pasta_trabalho():
    # 1. argumento na linha de comando
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    if args:
        pasta = Path(args[0]).expanduser()
        if pasta_valida(pasta):
            return pasta
        print(f"AVISO: '{ARQUIVO_REFERENCIA_NOME}' não está em {pasta}")

    # 2. pasta do próprio script
    pasta_script = Path(__file__).resolve().parent
    if pasta_valida(pasta_script):
        return pasta_script

    # 3. Google Drive
    for pasta in candidatas_google_drive():
        if pasta_valida(pasta):
            return pasta

    # 4. escolha manual
    print("Não encontrei a pasta de upload automaticamente.")
    pasta = escolher_pasta_manualmente()
    if pasta_valida(pasta):
        return pasta

    return None


# ---------------------------------------------------------------------------
# PROCESSAMENTO
# ---------------------------------------------------------------------------
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
    bruto = pd.read_excel(caminho_arquivo, header=None, engine="openpyxl")

    # remove a primeira linha e a primeira coluna
    bruto = bruto.iloc[1:, 1:]

    # remove linhas e colunas totalmente vazias
    bruto = bruto.dropna(how="all", axis=0)
    bruto = bruto.dropna(how="all", axis=1)

    if bruto.empty:
        raise ValueError("a planilha ficou vazia após a limpeza")

    # a próxima linha restante vira o cabeçalho
    bruto.columns = bruto.iloc[0]
    df = bruto.iloc[1:].reset_index(drop=True)
    df.columns = [str(c).strip() for c in df.columns]

    return df


def carregar_siafi_referencia(pasta):
    caminho_ref = pasta / ARQUIVO_REFERENCIA_NOME
    if not caminho_ref.exists():
        raise FileNotFoundError(f"Arquivo de referência não encontrado: {caminho_ref}")

    df_ref = pd.read_excel(caminho_ref, engine="openpyxl")
    df_ref.columns = [str(c).strip() for c in df_ref.columns]
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
    df_filtrado.to_excel(caminho_saida, index=False, engine="openpyxl")

    return {
        "arquivo_origem": caminho_arquivo.name,
        "arquivo_saida": caminho_saida.name,
        "total_linhas_origem": total_linhas,
        "linhas_coincidentes": qtd_match,
        "coluna_siafi_usada": coluna_siafi,
    }


def main():
    pasta = localizar_pasta_trabalho()
    if pasta is None:
        print(f"ERRO: não encontrei nenhuma pasta com o arquivo {ARQUIVO_REFERENCIA_NOME}.")
        print('Dica: rode  python gerar_execucao.py "caminho\\da\\pasta"')
        pausar_e_sair(1)

    print(f"Pasta de trabalho: {pasta}\n")

    try:
        siafi_validos = carregar_siafi_referencia(pasta)
    except PermissionError:
        print(f"ERRO: não consegui abrir {ARQUIVO_REFERENCIA_NOME}. "
              "Feche o arquivo no Excel e tente de novo.")
        pausar_e_sair(1)
    except (FileNotFoundError, ValueError) as e:
        print(f"ERRO: {e}")
        pausar_e_sair(1)

    print(f"Referência carregada: {len(siafi_validos)} números SIAFI únicos "
          f"em {ARQUIVO_REFERENCIA_NOME}\n")

    # ignora arquivos temporários do Excel (~$EMENDAS...)
    arquivos = sorted(
        a for a in pasta.glob(PADRAO_ENTRADA) if not a.name.startswith("~$")
    )
    if not arquivos:
        print(f"Nenhum arquivo encontrado com o padrão '{PADRAO_ENTRADA}' em {pasta}")
        pausar_e_sair(0)

    relatorios = []
    for arquivo in arquivos:
        try:
            resultado = processar_arquivo(arquivo, siafi_validos, pasta)
            relatorios.append(resultado)
        except PermissionError:
            print(f"ERRO ao processar {arquivo.name}: arquivo aberto em outro programa "
                  "(feche o Excel e rode de novo).")
        except Exception as e:
            print(f"ERRO ao processar {arquivo.name}: {e}")

    # apaga os arquivos originais EMENDASESTADUAIS* processados com sucesso
    nomes_processados = {r["arquivo_origem"] for r in relatorios}
    for arquivo in arquivos:
        if arquivo.name in nomes_processados:
            try:
                arquivo.unlink()
            except OSError as e:
                print(f"AVISO: não consegui apagar {arquivo.name}: {e}")

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

    pausar_e_sair(0)


if __name__ == "__main__":
    main()
