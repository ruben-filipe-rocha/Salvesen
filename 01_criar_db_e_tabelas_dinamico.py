# ==============================================================================
# PASSO 1 — Criar a base de dados e tabelas DINAMICAMENTE
# (detecta automaticamente quais são os anos nos dados)
# ==============================================================================
# Este script substitui o notebook original, com a vantagem de:
#   - Detectar automaticamente quais anos estão nos ficheiros .xls
#   - Criar tabelas para TODOS os anos encontrados (não apenas 2025/2026)
#   - Não precisar de modificar código para adicionar novos anos
# ==============================================================================

import platform
import sqlite3
import os
import glob
import xml.etree.ElementTree as ET

# ==========================
# 1. Definir caminhos conforme o SO
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
# 2. Criar a base de dados (garante que a pasta existe)
# ==========================
pasta = os.path.dirname(DB_PATH)
if pasta and not os.path.exists(pasta):
    os.makedirs(pasta, exist_ok=True)
    print(f"Pasta criada: {pasta}")

con = sqlite3.connect(DB_PATH)
con.close()
print(f"Base de dados criada/aberta em: {DB_PATH}")
print()

# ==========================
# 3. Detectar automaticamente quais são os anos presentes nos dados
# ==========================

# Definição de colunas (igual ao notebook original)
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

CAMPO_DATA = "FENTREGA"  # campo que contém o ano (formato AAAAMMDD)
COLUNA_ORIGEM = "ficheiro_origem"

# Função auxiliar para ler ficheiros Excel XML (igual ao notebook de inserção)
NS = {"ss": "urn:schemas-microsoft-com:office:spreadsheet"}
SS_INDEX = "{urn:schemas-microsoft-com:office:spreadsheet}Index"

def ler_linhas_xml_spreadsheet(caminho_ficheiro):
    """
    Lê um ficheiro no formato 'Excel XML Spreadsheet 2003' e devolve
    (cabecalho, lista_de_linhas).
    """
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
    except Exception:
        return None, None


def detectar_anos_nos_ficheiros(pasta):
    """
    Procura todos os ficheiros .xls na pasta e detecta quais são
    os anos presentes nos dados (através do campo FENTREGA).
    Devolve um conjunto (set) com os anos encontrados.
    """
    ficheiros = glob.glob(os.path.join(pasta, "**", "*.xls"), recursive=True)
    anos = set()
    
    for caminho in ficheiros:
        cabecalho, linhas = ler_linhas_xml_spreadsheet(caminho)
        if cabecalho is None or CAMPO_DATA not in cabecalho:
            continue
        
        idx_data = cabecalho.index(CAMPO_DATA)
        for linha in linhas:
            if linha and idx_data < len(linha):
                valor_data = linha[idx_data]
                if valor_data and str(valor_data).strip()[:4].isdigit():
                    ano = str(valor_data).strip()[:4]
                    anos.add(ano)
    
    return sorted(anos)


# Detectar os anos
anos_detectados = detectar_anos_nos_ficheiros(PASTA_FICHEIROS)
print(f"Anos detectados nos ficheiros: {', '.join(anos_detectados)}")
print()

# ==========================
# 4. Criar as tabelas para cada ano detectado
# ==========================

colunas_sql = ",\n    ".join(f'"{c}" TEXT' for c in COLUNAS_FICHEIRO)

con = sqlite3.connect(DB_PATH)
cur = con.cursor()

for ano in anos_detectados:
    tabela_nome = f"dados_{ano}"
    cur.execute(f'''
        CREATE TABLE IF NOT EXISTS {tabela_nome} (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            {colunas_sql},
            "{COLUNA_ORIGEM}" TEXT
        )
    ''')
    print(f"Tabela {tabela_nome} criada (ou já existia).")

con.commit()
con.close()
print()
print("✓ OK - tabelas prontas para todos os anos detectados.")
print()

# ==========================
# 5. Confirmar: listar as tabelas e colunas criadas
# ==========================

con = sqlite3.connect(DB_PATH)
cur = con.cursor()

cur.execute("SELECT name FROM sqlite_master WHERE type='table'")
tabelas_bd = [nome for (nome,) in cur.fetchall() if not nome.startswith('sqlite_')]

print("Tabelas na base de dados:")
for tabela in sorted(tabelas_bd):
    print(f" - {tabela}")

print()
for ano in anos_detectados:
    tabela_nome = f"dados_{ano}"
    if tabela_nome in tabelas_bd:
        cur.execute(f"PRAGMA table_info({tabela_nome})")
        colunas = cur.fetchall()
        print(f"Colunas de {tabela_nome}: {len(colunas)}")
        print(f"  ({', '.join([c[1] for c in colunas[:5]])}...)")

con.close()
print()
print("=" * 70)
print("PRÓXIMO PASSO: executar o script 03_inserir_dados_dinamico.py")
print("=" * 70)
