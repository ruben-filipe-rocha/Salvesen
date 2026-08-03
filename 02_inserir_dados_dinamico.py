# ==============================================================================
# PASSO 3 — Inserir dados nas tabelas DINAMICAMENTE
# (funciona com qualquer ano, sem hardcoding)
# ==============================================================================
# Este script substitui o notebook original, com a vantagem de:
#   - Suportar QUALQUER ano (não está limitado a 2025/2026)
#   - Gerar automaticamente as instruções SQL para cada ano
#   - Escalar facilmente a novos anos sem modificação de código
# ==============================================================================

import platform
import sqlite3
import os
import glob
import xml.etree.ElementTree as ET
from datetime import datetime

# ==========================
# 1. Definir caminhos conforme o SO (igual ao script anterior)
# ==========================
if platform.system() == 'Windows':
    DB_PATH = r"C:\Users\LISARR\Documents\python\01.Financeiro\inform_27.db"
    PASTA_FICHEIROS = r"C:\Users\LISARR\Documents\python\01.Financeiro\inform_27"
elif platform.system() == 'Darwin':
    DB_PATH = "/Volumes/RR/DB/inform_27.db"
    PASTA_FICHEIROS = "/Volumes/RR/DB/inform_27"
else:
    DB_PATH = "inform_27.db"
    PASTA_FICHEIROS = "inform_27"

print("DB_PATH:", DB_PATH)
print("PASTA_FICHEIROS:", PASTA_FICHEIROS)
print()

# ==========================
# 2. Definir colunas e campos (igual aos scripts anteriores)
# ==========================

COLUNAS_FICHEIRO = [
    "PROPIETARIO", "TRAYECTO", "TRANSPORTISTA", "TRACTORA", "REMOLQUE",
    "INGRESODT", "COSTEDT", "RENTADT", "PALETSDT", "PESO_BRUTO",
    "CODEUT", "ESTADO_UT", "RANGO_UT", "FCARGA", "ACTIVIDAD",
    "CODEDT", "ESTADO_DT", "REFERENCIA", "CODACT", "LOCORIGEN",
    "PROV_ORIGEN", "PAISORIGEN", "CPOSTAL", "LOCDESTINO", "PROV_DESTINO",
    "PAISDESTINO", "CPOSTAD", "KM", "FENTREGA", "ORIGEN",
    "ENTREGAR", "PROV_ENTREGAR", "PAISENTREGAR", "DESTINO", "PALETS",
    "PREFAC", "RUTA", "COBROREAL", "GESTION", "DEPART",
    "USCODE", "USUARIO", "TIPOCLIENTE", "TIPOFLUJO", "WMSCODRGT",
    "LOCCAR", "LUGARCARGA", "LOCDES", "LUGARDESCARGA", "TEMP_MERC_PED",
    "TIPOPALETA", "CAMION_TIPO", "CAMION_CAPACIDAD", "TIPO_COMBUSTIBLE", "KMREALES",
    "ALBARAN",
]

COLUNA_ORIGEM = "ficheiro_origem"
CAMPO_DATA = "FENTREGA"

# ==========================
# 3. Funções auxiliares
# ==========================

NS = {"ss": "urn:schemas-microsoft-com:office:spreadsheet"}
SS_INDEX = "{urn:schemas-microsoft-com:office:spreadsheet}Index"

def ler_linhas_xml_spreadsheet(caminho_ficheiro):
    """Lê um ficheiro Excel XML Spreadsheet 2003."""
    try:
        tree = ET.parse(caminho_ficheiro)
        root = tree.getroot()

        worksheet = root.find("ss:Worksheet", NS)
        if worksheet is None:
            return None, None

        tabela = worksheet.find("ss:Table", NS)
        if tabela is None:
            return None, None

        linhas_xml = tabela.findall("ss:Row", NS)
        if not linhas_xml:
            return [], []

        def extrair_linha(linha_xml):
            valores = []
            proximo_indice = 1
            for cell in linha_xml.findall("ss:Cell", NS):
                idx = cell.get(SS_INDEX)
                idx = int(idx) if idx is not None else proximo_indice
                while len(valores) < idx - 1:
                    valores.append(None)
                data_el = cell.find("ss:Data", NS)
                valor = data_el.text if data_el is not None else None
                valores.append(valor)
                proximo_indice = idx + 1
            return valores

        cabecalho = extrair_linha(linhas_xml[0])
        cabecalho = [c.strip() if c else f"COLUNA_{i+1}" for i, c in enumerate(cabecalho)]

        linhas = []
        n_colunas = len(cabecalho)
        for linha_xml in linhas_xml[1:]:
            valores = extrair_linha(linha_xml)
            if len(valores) < n_colunas:
                valores += [None] * (n_colunas - len(valores))
            elif len(valores) > n_colunas:
                valores = valores[:n_colunas]
            linhas.append(valores)

        return cabecalho, linhas
    except Exception as e:
        print(f"    [ERRO ao ler] {e}")
        return None, None


def listar_ficheiros_excel(pasta):
    """Procura todos os ficheiros .xls na pasta (incluindo subpastas)."""
    encontrados = glob.glob(os.path.join(pasta, "**", "*.xls"), recursive=True)
    vistos = set()
    ficheiros = []
    for caminho in encontrados:
        chave = os.path.normcase(os.path.abspath(caminho))
        if chave not in vistos:
            vistos.add(chave)
            ficheiros.append(caminho)
    return sorted(ficheiros)


def detectar_anos_na_bd(db_path):
    """
    Detecta quais são as tabelas de anos presentes na BD.
    Devolve uma lista de anos (ex: ['2023', '2024', '2025', '2026'])
    """
    con = sqlite3.connect(db_path)
    cur = con.cursor()
    
    cur.execute("SELECT name FROM sqlite_master WHERE type='table'")
    tabelas = [nome for (nome,) in cur.fetchall()]
    
    con.close()
    
    # Filtrar apenas tabelas que começam com "dados_" e têm um ano válido
    anos = []
    for tabela in tabelas:
        if tabela.startswith("dados_"):
            ano = tabela.replace("dados_", "")
            if ano.isdigit() and len(ano) == 4:
                anos.append(ano)
    
    return sorted(anos)


# ==========================
# 4. MAIN: Lógica principal
# ==========================

print("=" * 70)
print("PASSO 3 — INSERIR DADOS DINAMICAMENTE")
print("=" * 70)
print()

# Detectar anos na BD
anos_na_bd = detectar_anos_na_bd(DB_PATH)
print(f"Anos com tabelas na BD: {', '.join(anos_na_bd)}")
print()

# Localizar ficheiros Excel
ficheiros = listar_ficheiros_excel(PASTA_FICHEIROS)
print(f"Ficheiros .xls encontrados: {len(ficheiros)}")
for f in ficheiros[:20]:
    print(f"  - {os.path.basename(f)}")
if len(ficheiros) > 20:
    print(f"  ... e mais {len(ficheiros) - 20} ficheiros")
print()

if not ficheiros:
    print("[AVISO] Nenhum ficheiro .xls encontrado. Abortando.")
    exit(1)

# ==========================
# 5. Preparar as instruções SQL DINAMICAMENTE
# ==========================

cols_sql = ", ".join(f'"{c}"' for c in COLUNAS_FICHEIRO)
placeholders = ", ".join(["?"] * len(COLUNAS_FICHEIRO))
idx_data = COLUNAS_FICHEIRO.index(CAMPO_DATA)

# Gerar uma instrução SQL INSERT OR IGNORE para cada ano
sql_por_ano = {}
for ano in anos_na_bd:
    tabela_nome = f"dados_{ano}"
    sql = f'INSERT OR IGNORE INTO {tabela_nome} ({cols_sql}, "{COLUNA_ORIGEM}") VALUES ({placeholders}, ?)'
    sql_por_ano[ano] = sql

print(f"Instruções SQL geradas para {len(sql_por_ano)} ano(s):")
for ano in sorted(sql_por_ano.keys()):
    print(f"  - dados_{ano}")
print()

# ==========================
# 6. Criar índices únicos (evita duplicatas)
# ==========================

print("A criar índices únicos em CODEDT...")
con = sqlite3.connect(DB_PATH)
cur = con.cursor()

for ano in anos_na_bd:
    tabela_nome = f"dados_{ano}"
    indice_nome = f"idx_{tabela_nome}_codedt"
    try:
        cur.execute(f'''
            CREATE UNIQUE INDEX IF NOT EXISTS {indice_nome}
            ON {tabela_nome}("CODEDT")
        ''')
    except sqlite3.OperationalError:
        # O índice pode já existir de uma execução anterior
        pass

con.commit()
con.close()
print("✓ Índices prontos.")
print()

# ==========================
# 7. Loop principal: ler cada ficheiro e inserir dados
# ==========================

print("=" * 70)
print("A processar ficheiros...")
print("=" * 70)
print()

con = sqlite3.connect(DB_PATH)
cur = con.cursor()

# Dicionário para contar inserções por ano
totais_inseridos = {ano: 0 for ano in anos_na_bd}
totais_duplicados = 0
totais_sem_data = 0

for i, caminho in enumerate(ficheiros, start=1):
    nome_ficheiro = os.path.basename(caminho)
    print(f"[{i}/{len(ficheiros)}] A processar: {nome_ficheiro}")

    cabecalho, linhas = ler_linhas_xml_spreadsheet(caminho)
    if cabecalho is None:
        print(f"  [ERRO] Falha a ler o ficheiro. A saltar.")
        continue

    if CAMPO_DATA not in cabecalho:
        print(f"  [AVISO] Coluna '{CAMPO_DATA}' não encontrada. A saltar.")
        continue

    # Organizar as linhas por ano em batches
    batches_por_ano = {ano: [] for ano in anos_na_bd}
    sem_data = 0

    for valores in linhas:
        # Ignorar linhas completamente vazias
        if all(v is None or str(v).strip() == "" for v in valores):
            continue

        # Remapaer a linha do ficheiro para a ordem fixa
        mapa = dict(zip(cabecalho, valores))
        valores_ordenados = [mapa.get(c) for c in COLUNAS_FICHEIRO]

        # Extrair e validar a data
        valor_data = valores_ordenados[idx_data]
        if not valor_data or not str(valor_data).strip()[:4].isdigit():
            sem_data += 1
            continue

        texto_data = str(valor_data).strip()
        ano = texto_data[:4]

        # Se o ano está definido na BD, adicionar ao batch correspondente
        if ano in batches_por_ano:
            batches_por_ano[ano].append(valores_ordenados + [nome_ficheiro])
        else:
            sem_data += 1  # ano não tem tabela definida

    # Inserir os batches
    for ano in anos_na_bd:
        batch = batches_por_ano[ano]
        if not batch:
            continue

        antes = con.total_changes
        cur.executemany(sql_por_ano[ano], batch)
        con.commit()
        inseridos = con.total_changes - antes
        totais_inseridos[ano] += inseridos

        # Contar duplicatas
        duplicados_neste_ano = len(batch) - inseridos
        totais_duplicados += duplicados_neste_ano

    totais_sem_data += sem_data

    # Resumo desta linha
    resumo_insercoes = ", ".join(
        f"{totais_inseridos[ano]} novo(s) em {ano}"
        for ano in sorted(anos_na_bd)
    )
    agora = datetime.now().isoformat(timespec="seconds")
    print(f"  -> {resumo_insercoes}, {totais_duplicados} duplicata(s), "
          f"{totais_sem_data} sem data | {agora}")

con.close()

# ==========================
# 8. Resumo final
# ==========================

print()
print("=" * 70)
print("RESUMO FINAL")
print("=" * 70)
print(f"Ficheiros processados: {len(ficheiros)}")
print()
for ano in sorted(anos_na_bd):
    print(f"  Linhas novas inseridas em {ano}: {totais_inseridos[ano]}")
print()
print(f"Linhas já existentes (duplicadas por CODEDT): {totais_duplicados}")
print(f"Linhas sem FENTREGA válido: {totais_sem_data}")
print("=" * 70)
print()
print("✓ Inserção de dados concluída!")
