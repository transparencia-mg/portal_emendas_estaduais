"""
gerar_execucao.py
------------------
Cruza os arquivos EMENDAS_ESTADUAIS<sufixo>.xlsx com a planilha de referência
VW_SG_V2_EP_INDIC_RECURSOS_TW.xlsx (coluna NUMERO_SIAFI) e gera, para cada um,
um arquivo execucao<sufixo>.xlsx contendo apenas as linhas cujo número SIAFI
também aparece na referência. No final, envia tudo para o GitHub.

FUNCIONA EM QUALQUER COMPUTADOR — não há caminho fixo no código.
    Coloque este arquivo na RAIZ do repositório portal_emendas_estaduais
    (ao lado da pasta "upload") ou dentro da própria pasta "upload".
    O script descobre sozinho onde está a pasta upload.

REQUISITOS:
    - Python 3.
    - pandas e openpyxl (se faltarem, o script tenta instalar sozinho).
    - Para enviar ao GitHub: Git instalado, o repositório clonado com
      "git clone" e login do GitHub feito nesse computador.

Como usar:
    python gerar_execucao.py
    python gerar_execucao.py "C:/outra/pasta/upload"   (opcional: força a pasta)

O que o script faz, em ordem:
    1. git pull — sincroniza com o GitHub antes de mexer em qualquer arquivo.
    2. Lê "VW_SG_V2_EP_INDIC_RECURSOS_TW.xlsx" e monta o conjunto de números
       válidos a partir da coluna NUMERO_SIAFI.
    3. Para cada EMENDAS_ESTADUAIS<sufixo>.xlsx da pasta upload:
       a. Remove a primeira linha e a primeira coluna.
       b. Remove linhas e colunas totalmente vazias.
       c. Usa a linha seguinte como cabeçalho.
       d. Filtra mantendo apenas as linhas cujo número SIAFI está na referência.
       e. Salva como execucao<sufixo>.xlsx na mesma pasta.
    4. Apaga os EMENDAS_ESTADUAIS<sufixo>.xlsx originais.
    5. Imprime um relatório com a quantidade de linhas que coincidiram.
    6. git add / commit / push.

Se algum arquivo der erro, NADA é apagado e NADA é enviado ao GitHub
(os execucao*.xlsx gerados ficam na pasta para conferência).
"""

import re
import subprocess
import sys
from datetime import datetime
from pathlib import Path

def garantir_dependencias():
    """Instala pandas/openpyxl automaticamente se faltarem neste computador."""
    faltando = []
    for modulo in ("pandas", "openpyxl"):
        try:
            __import__(modulo)
        except ImportError:
            faltando.append(modulo)
    if not faltando:
        return
    print(f"Instalando dependências que faltam: {', '.join(faltando)}...")
    r = subprocess.run([sys.executable, "-m", "pip", "install", *faltando])
    if r.returncode != 0:
        sys.exit(f"Não consegui instalar. Rode manualmente: "
                 f"{sys.executable} -m pip install {' '.join(faltando)}")


garantir_dependencias()
import pandas as pd  # noqa: E402

# ---------------------------------------------------------------------------
# CONFIGURAÇÃO — ajuste aqui se necessário
# ---------------------------------------------------------------------------
ARQUIVO_REFERENCIA_NOME = "VW_SG_V2_EP_INDIC_RECURSOS_TW.xlsx"
COLUNA_REFERENCIA = "NUMERO_SIAFI"
PADRAO_ENTRADA = "EMENDAS_ESTADUAIS*.xlsx"
PREFIXO_SAIDA = "execucao"

APAGAR_ORIGINAIS = True     # etapa 4
ATUALIZAR_GITHUB = True     # etapas 1 e 6
MENSAGEM_COMMIT = "Atualização da execução das emendas estaduais"
NOME_PASTA_UPLOAD = "upload"


def localizar_pasta_upload():
    """
    Descobre a pasta upload sem caminho fixo, na seguinte ordem:
      1. pasta passada na linha de comando;
      2. <pasta do script>/upload;
      3. a própria pasta do script (se o script estiver dentro de upload).
    Aceita a pasta que tiver a planilha de referência ou arquivos EMENDAS_ESTADUAIS.
    """
    if len(sys.argv) > 1:
        return Path(sys.argv[1]).expanduser().resolve()

    pasta_script = Path(__file__).resolve().parent
    candidatas = [pasta_script / NOME_PASTA_UPLOAD, pasta_script]
    for pasta in candidatas:
        if pasta.is_dir() and ((pasta / ARQUIVO_REFERENCIA_NOME).exists()
                               or any(pasta.glob(PADRAO_ENTRADA))):
            return pasta
    # nenhuma tem os arquivos ainda: usa a pasta upload se existir
    return candidatas[0] if candidatas[0].is_dir() else pasta_script


PASTA_UPLOAD = localizar_pasta_upload()


# ---------------------------------------------------------------------------
# GIT
# ---------------------------------------------------------------------------
def git(*args, mostrar=True):
    """Roda um comando git dentro da pasta do repositório."""
    try:
        r = subprocess.run(["git", *args], cwd=PASTA_UPLOAD, capture_output=True,
                           text=True, encoding="utf-8", errors="replace")
    except FileNotFoundError:
        # git não está instalado neste computador
        return subprocess.CompletedProcess(args, 1, "", "git não encontrado")
    if mostrar and r.stdout.strip():
        print("    " + r.stdout.strip().replace("\n", "\n    "))
    if r.returncode != 0 and r.stderr.strip():
        print("    " + r.stderr.strip().replace("\n", "\n    "))
    return r


def parar(msg):
    print(f"\nERRO: {msg}")
    sys.exit("Processo interrompido. Os arquivos de origem NÃO foram apagados "
             "e nada foi enviado ao GitHub.")


# ---------------------------------------------------------------------------
# PROCESSAMENTO DAS PLANILHAS
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
    """Extrai o sufixo (ex.: ano) do nome EMENDAS_ESTADUAIS<sufixo>.xlsx"""
    m = re.match(r"EMENDAS_ESTADUAIS(.*)\.xlsx$", nome_arquivo, re.IGNORECASE)
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


def imprimir_relatorio(relatorios):
    print("\n" + "=" * 70)
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


# ---------------------------------------------------------------------------
# EXECUÇÃO
# ---------------------------------------------------------------------------
def main():
    if not PASTA_UPLOAD.exists():
        print(f"ERRO: pasta não encontrada -> {PASTA_UPLOAD}")
        print("Coloque o script na raiz do repositório (ao lado da pasta 'upload') "
              "ou informe a pasta: python gerar_execucao.py \"caminho\\da\\upload\"")
        sys.exit(1)

    print(f"Pasta de trabalho: {PASTA_UPLOAD}")

    usar_git = ATUALIZAR_GITHUB and git(
        "rev-parse", "--is-inside-work-tree", mostrar=False).returncode == 0
    if ATUALIZAR_GITHUB and not usar_git:
        print("  AVISO: esta pasta não está num repositório git (ou o git não está "
              "instalado); os arquivos serão gerados, mas não enviados ao GitHub.")

    # 1) Sincroniza antes de alterar qualquer arquivo
    if usar_git:
        print("\n[1/6] Sincronizando com o GitHub (git pull)...")
        if git("pull", "--rebase", "--autostash").returncode != 0:
            sys.exit("Falha no git pull. Resolva (ex.: conflito ou sem internet) e "
                     "rode de novo. Nenhum arquivo foi alterado.")

    # 2) Referência
    print("\n[2/6] Carregando a referência...")
    try:
        siafi_validos = carregar_siafi_referencia(PASTA_UPLOAD)
    except (FileNotFoundError, ValueError) as e:
        parar(e)
    print(f"  {len(siafi_validos)} números SIAFI únicos em {ARQUIVO_REFERENCIA_NOME}")

    # 3) Processa os arquivos
    print("\n[3/6] Gerando os arquivos de execução...")
    arquivos = sorted(PASTA_UPLOAD.glob(PADRAO_ENTRADA))
    relatorios = []
    erros = []
    if not arquivos:
        print(f"  Nenhum arquivo '{PADRAO_ENTRADA}' na pasta; "
              "os arquivos de execução atuais serão mantidos.")
    for arquivo in arquivos:
        try:
            relatorios.append(processar_arquivo(arquivo, siafi_validos, PASTA_UPLOAD))
            print(f"  ok: {arquivo.name} -> {relatorios[-1]['arquivo_saida']}")
        except Exception as e:
            print(f"  ERRO ao processar {arquivo.name}: {e}")
            erros.append(arquivo.name)

    if erros:
        if relatorios:
            imprimir_relatorio(relatorios)
        parar(f"falha em {len(erros)} arquivo(s): {', '.join(erros)}. "
              "Corrija e rode de novo.")

    # 4) Apaga os originais
    if APAGAR_ORIGINAIS and arquivos:
        print("\n[4/6] Apagando os arquivos de origem...")
        nao_apagados = []
        for arquivo in arquivos:
            try:
                arquivo.unlink()
                print(f"  apagado: {arquivo.name}")
            except PermissionError:
                print(f"  AVISO: não foi possível apagar {arquivo.name} "
                      "(está aberto no Excel?).")
                nao_apagados.append(arquivo.name)
        if nao_apagados and usar_git:
            sys.exit("Feche o arquivo no Excel, apague-o da pasta upload e rode de novo "
                     "(ele não pode ir para o GitHub). Nada foi enviado.")

    # 5) Relatório
    if relatorios:
        print("\n[5/6] Relatório")
        imprimir_relatorio(relatorios)

    # 6) Commit e push
    if not usar_git:
        print("\nConcluído (sem envio ao GitHub).")
        return

    print("\n[6/6] Enviando alterações para o GitHub...")
    git("add", "-A")
    status = git("status", "--porcelain", mostrar=False).stdout.strip()
    if not status:
        print("  Nenhuma alteração para enviar (tudo já estava igual ao GitHub).")
        print("\nConcluído.")
        return
    print("  Alterações:")
    print("    " + status.replace("\n", "\n    "))

    if not (git("config", "user.name", mostrar=False).stdout.strip()
            and git("config", "user.email", mostrar=False).stdout.strip()):
        git("reset", mostrar=False)
        sys.exit("O Git deste computador não tem nome/e-mail configurados. Rode uma vez:\n"
                 '  git config --global user.name "Seu Nome"\n'
                 '  git config --global user.email "seu@email"\n'
                 "e depois rode o script de novo (os arquivos já gerados serão enviados).")
    if git("commit", "-m", f"{MENSAGEM_COMMIT} ({datetime.now():%d/%m/%Y %H:%M})").returncode != 0:
        sys.exit("Falha no commit.")
    if git("push").returncode != 0:
        # O GitHub pode ter commits novos (ex.: automação): sincroniza e tenta de novo
        print("  Push recusado; sincronizando e tentando de novo...")
        if git("pull", "--rebase").returncode != 0 or git("push").returncode != 0:
            sys.exit("Falha no git push. O commit foi feito localmente.\n"
                     "Se for o primeiro uso neste computador, confira se o login do "
                     "GitHub está feito (ex.: rode 'git push' na pasta do repositório "
                     "e entre com sua conta). Depois rode 'git push'.")
    print("  GitHub atualizado.")
    print("\nConcluído.")


if __name__ == "__main__":
    main()
