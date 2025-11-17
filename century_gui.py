#Construtor de Arquivos de Entrada CENTURY
#
# Este script Python (century_gui.py) fornece uma Interface Gráfica Interativa (GUI) baseada em FreeSimpleGUI para simplificar a criação dos arquivos de agendamento (.SCH) e a extração automatizada de dados geoespaciais de ponto único ou em lote (.WTH), incluindo:
#
# * Uso e Cobertura do Solo (LULC - MapBiomas): Extração de séries temporais de classes de uso.
# * Dados de Solo (Embrapa): Extração de propriedades físicas (areia, silte, argila, densidade, pH).
# * Dados Climáticos (INMET): Busca das estações mais próximas e cálculo das médias mensais e       anuais para formatos CSV e .WTH.
#
# Autor: Marcos Cardoso; Equipe Carbono LAPIG-UFG
# 2025-11-17 (v28)
# ---

#Para começar, execute esta célula abaixo para instalar os pacotes necessários (instalação única) e a próxima (código) para iniciar a aplicação GUI.

import sys
!{sys.executable} -m pip install FreeSimpleGUI pandas rasterio


# ---


import FreeSimpleGUI as sg
import os
import re
from pathlib import Path
import math
import pandas as pd
import sys

try:
    import rasterio
    import rasterio.sample
    LIBS_INSTALADAS = True
except ImportError:
    LIBS_INSTALADAS = False

MAPBIOMAS_LEGEND = {
    1 : 'Floresta', 3 : 'Formação Florestal', 4 : 'Formação Savânica',
    5 : 'Mangue', 6 : 'Floresta Alagável', 49 : 'Restinga Arbórea',
    10 : 'Vegetação Herbácea e Arbustiva', 11 : 'Campo Alagado e Área Pantanosa',
    12 : 'Formação Campestre', 32 : 'Apicum', 29 : 'Afloramento Rochoso',
    50 : 'Restinga Herbácea', 14 : 'Agropecuária', 15 : 'Pastagem',
    18 : 'Agricultura', 19 : 'Lavoura Temporária', 39 : 'Soja',
    20 : 'Cana', 40 : 'Arroz', 62 : 'Algodão (beta)',
    41 : 'Outras Lavouras Temporárias', 36 : 'Lavoura Perene',
    46 : 'Café', 47 : 'Citrus', 35 : 'Dendê',
    48 : 'Outras Lavouras Perenes', 9 : 'Silvicultura',
    21 : 'Mosaico de Usos', 22 : 'Área não Vegetada', 23 : 'Praia, Duna e Areal',
    24 : 'Área Urbanizada', 30 : 'Mineração', 75 : 'Usina Fotovoltaica (beta)',
    25 : 'Outras Áreas não Vegetadas', 26 : 'Corpo D\'água',
    33 : 'Rio, Lago e Oceano', 31 : 'Aquicultura', 27 : 'Não observado'
}

# Tipos de evento atualizados com as descrições fornecidas, mais a adição de IRRI e EROD
TIPOS_DE_EVENTO_COM_CODIGO = {
    'CROP': 'CROP: Seleciona cultura', 
    'PLTM': 'PLTM: Marca plantio', 
    'HARV': 'HARV: Agenda colheita', 
    'FERT': 'FERT: Agenda fertilização', 
    'CULT': 'CULT: Agenda cultivo', 
    'OMAD': 'OMAD: Adiciona matéria',
    'GRAZ': 'GRAZ: Agenda pastejo', 
    'FIRE': 'FIRE: Agenda fogo',
    'TREE': 'TREE: Seleciona árvore', 
    'TREM': 'TREM: Remove árvore',
    'IRRI': 'IRRI: Agenda irrigação',
    'EROD': 'EROD: Agenda erosão'
}
TIPOS_DE_EVENTO_SEM_CODIGO = {
    'FRST': 'FRST: Início crescimento Cultivo', 
    'LAST': 'LAST: Fim crescimento Cultivo',
    'SENM': 'SENM: Marca senescência',
    'TFST': 'TFST: Início floresta',
    'TLST': 'TLST: Fim floresta'
}

# Configurações de Weather
WEATHER_CHOICES_MAP = {
    'M': 'M: valores médios do site.100',
    'S': 'S: do site.100, mas prec estocástica',
    'F': 'F: do inicio arquivo WTH',
    'C': 'C: continuação WTH, sem retrocesso'
}
WEATHER_CHOICES = list(WEATHER_CHOICES_MAP.values())
WEATHER_DESC_TO_CODE = {v: k for k, v in WEATHER_CHOICES_MAP.items()}


# Variáveis auxiliares para a interface gráfica
ALL_EVENT_TYPES = list(TIPOS_DE_EVENTO_COM_CODIGO.keys()) + list(TIPOS_DE_EVENTO_SEM_CODIGO.keys())
ALL_EVENT_DESCRIPTIONS = list(TIPOS_DE_EVENTO_COM_CODIGO.values()) + list(TIPOS_DE_EVENTO_SEM_CODIGO.values())
EVENT_DESC_TO_CODE = {v: k for d in [TIPOS_DE_EVENTO_COM_CODIGO, TIPOS_DE_EVENTO_SEM_CODIGO] for k, v in d.items()}

CODIGOS_E_DESCRICOES = {
    'CROP': {'HER': 'Herbáceas Cerrado', 'BE8': 'Pastagem tradicional (médio vigor)', 'MEL': 'Pastagem produtiva', 'DEG': 'Pastagem degradada', 'SJ': 'Soja', 'CAN': 'Cana', 'MLH': 'Milho (safra principal)', 'MSF': 'Milho (safrinha)'},
    'CULT': {'P': 'Revolvimento (renovação pastagem)', 'S': 'Revolvimento (pós-conversão)', 'S1D': 'Plantio direto (soja/milho)', 'S1C': 'Plantio convencional (soja/milho)', 'AP': 'Aração + gradagem (cana)'},
    'FERT': {'A': 'Manutenção Automática (mínimo)', 'MED': 'Concentrações médias automáticas', 'N45': 'Nitrogênio (4.5 gN/m2)', 'N150': 'Nitrogênio (renovação pastagem)', 'NP1': 'NPK (plantio cana)', 'FMP': 'Fertilização (plantio milho)', 'FMC': 'Fertilização (cobertura milho)'},
    'FIRE': {'CER': 'Após desmatamento Cerrado', 'H': 'Fogo alta intensidade (Hot)', 'M': 'Fogo média intensidade', 'PHF': 'Fogo pré-colheita (cana)'},
    'HARV': {'HS': 'Colheita soja', 'SF': 'Colheita cana (sem fogo)', 'CF': 'Colheita cana (com fogo)', 'GMLH': 'Colheita grão milho (90% palha)', 'GMSF': 'Colheita grão milho (50% palha)'},
    'OMAD': {'M': 'Palha e esterco', 'F': 'Torta de filtro', 'V': 'Vinhaça', 'FV': 'Torta de filtro e vinhaça'},
    'TREM': {'CCER3': 'Remoção CCER3'},
    'GRAZ': {'GM': 'Pastejo Baixa Intensidade (GM)'},
    'TREE': {'CER': 'Árvore (Cerrado)'}
}

INITIAL_CROP_OPTIONS = list(CODIGOS_E_DESCRICOES['CROP'].keys())
INITIAL_TREE_OPTIONS = list(CODIGOS_E_DESCRICOES['TREE'].keys())
DISPLAY_PARA_CODIGO = {}

BLOCO_PADRAO_SAVANA_EVENTS = [
    {'tipo': 'EVENT', 'event_type': 'CROP', 'code': 'HER', 'month': '1', 'block_num': '1'},
    {'tipo': 'EVENT', 'event_type': 'FRST', 'code': None, 'month': '1', 'block_num': '1'},
    {'tipo': 'EVENT', 'event_type': 'TREE', 'code': 'CER', 'month': '1', 'block_num': '1'},
    {'tipo': 'EVENT', 'event_type': 'TFST', 'code': None, 'month': '1', 'block_num': '1'},
    {'tipo': 'EVENT', 'event_type': 'LAST', 'code': None, 'month': '5', 'block_num': '1'},
    {'tipo': 'EVENT', 'event_type': 'SENM', 'code': None, 'month': '5', 'block_num': '1'},
    {'tipo': 'EVENT', 'event_type': 'FRST', 'code': None, 'month': '9', 'block_num': '1'},
    {'tipo': 'EVENT', 'event_type': 'TLST', 'code': None, 'month': '12', 'block_num': '1'},
    {'tipo': 'EVENT', 'event_type': 'CROP', 'code': 'HER', 'month': '1', 'block_num': '2'},
    {'tipo': 'EVENT', 'event_type': 'TREE', 'code': 'CER', 'month': '1', 'block_num': '2'},
    {'tipo': 'EVENT', 'event_type': 'TFST', 'code': None, 'month': '1', 'block_num': '2'},
    {'tipo': 'EVENT', 'event_type': 'LAST', 'code': None, 'month': '5', 'block_num': '2'},
    {'tipo': 'EVENT', 'event_type': 'SENM', 'code': None, 'month': '5', 'block_num': '2'},
    {'tipo': 'EVENT', 'event_type': 'FRST', 'code': None, 'month': '9', 'block_num': '2'},
    {'tipo': 'EVENT', 'event_type': 'TLST', 'code': None, 'month': '12', 'block_num': '2'},
    {'tipo': 'EVENT', 'event_type': 'CROP', 'code': 'HER', 'month': '1', 'block_num': '3'},
    {'tipo': 'EVENT', 'event_type': 'TREE', 'code': 'CER', 'month': '1', 'block_num': '3'},
    {'tipo': 'EVENT', 'event_type': 'TFST', 'code': None, 'month': '1', 'block_num': '3'},
    {'tipo': 'EVENT', 'event_type': 'LAST', 'code': None, 'month': '5', 'block_num': '3'},
    {'tipo': 'EVENT', 'event_type': 'SENM', 'code': None, 'month': '5', 'block_num': '3'},
    {'tipo': 'EVENT', 'event_type': 'FRST', 'code': None, 'month': '9', 'block_num': '3'},
    {'tipo': 'EVENT', 'event_type': 'TLST', 'code': None, 'month': '12', 'block_num': '3'},
    {'tipo': 'EVENT', 'event_type': 'CROP', 'code': 'HER', 'month': '1', 'block_num': '4'},
    {'tipo': 'EVENT', 'event_type': 'TREE', 'code': 'CER', 'month': '1', 'block_num': '4'},
    {'tipo': 'EVENT', 'event_type': 'TFST', 'code': None, 'month': '1', 'block_num': '4'},
    {'tipo': 'EVENT', 'event_type': 'LAST', 'code': None, 'month': '5', 'block_num': '4'},
    {'tipo': 'EVENT', 'event_type': 'SENM', 'code': None, 'month': '5', 'block_num': '4'},
    {'tipo': 'EVENT', 'event_type': 'FRST', 'code': None, 'month': '9', 'block_num': '4'},
    {'tipo': 'EVENT', 'event_type': 'TLST', 'code': None, 'month': '12', 'block_num': '4'},
    {'tipo': 'EVENT', 'event_type': 'CROP', 'code': 'HER', 'month': '1', 'block_num': '5'},
    {'tipo': 'EVENT', 'event_type': 'TREE', 'code': 'CER', 'month': '1', 'block_num': '5'},
    {'tipo': 'EVENT', 'event_type': 'TFST', 'code': None, 'month': '1', 'block_num': '5'},
    {'tipo': 'EVENT', 'event_type': 'LAST', 'code': None, 'month': '5', 'block_num': '5'},
    {'tipo': 'EVENT', 'event_type': 'SENM', 'code': None, 'month': '5', 'block_num': '5'},
    {'tipo': 'EVENT', 'event_type': 'FIRE', 'code': 'CER', 'month': '7', 'block_num': '5'},
    {'tipo': 'EVENT', 'event_type': 'FRST', 'code': None, 'month': '9', 'block_num': '5'},
    {'tipo': 'EVENT', 'event_type': 'TLST', 'code': None, 'month': '12', 'block_num': '5'}
]

# NOVO BLOCO: Desmatamento + Pastagem Tradicional (2 repetições) - Bloco 2 Manual
BLOCO_DESMATAMENTO_PASTAGEM_EVENTS = [
    # Ano 1 (Repetição 1): Manutenção/Savana
    {'tipo': 'EVENT', 'event_type': 'CROP', 'code': 'HER', 'month': '1', 'block_num': '1'},
    {'tipo': 'EVENT', 'event_type': 'FRST', 'code': None, 'month': '1', 'block_num': '1'},
    {'tipo': 'EVENT', 'event_type': 'TREE', 'code': 'CER', 'month': '1', 'block_num': '1'},
    {'tipo': 'EVENT', 'event_type': 'TFST', 'code': None, 'month': '1', 'block_num': '1'},
    {'tipo': 'EVENT', 'event_type': 'LAST', 'code': None, 'month': '5', 'block_num': '1'},
    {'tipo': 'EVENT', 'event_type': 'SENM', 'code': None, 'month': '5', 'block_num': '1'},
    {'tipo': 'EVENT', 'event_type': 'FRST', 'code': None, 'month': '9', 'block_num': '1'},
    {'tipo': 'EVENT', 'event_type': 'TLST', 'code': None, 'month': '12', 'block_num': '1'},
    
    # Ano 2 (Repetição 2): Desmatamento/Pastagem (Com CROP DEG)
    {'tipo': 'EVENT', 'event_type': 'TREE', 'code': 'CER', 'month': '1', 'block_num': '2'},
    {'tipo': 'EVENT', 'event_type': 'TFST', 'code': None, 'month': '1', 'block_num': '2'},
    {'tipo': 'EVENT', 'event_type': 'LAST', 'code': None, 'month': '5', 'block_num': '2'},
    {'tipo': 'EVENT', 'event_type': 'SENM', 'code': None, 'month': '5', 'block_num': '2'},
    {'tipo': 'EVENT', 'event_type': 'TLST', 'code': None, 'month': '6', 'block_num': '2'},
    {'tipo': 'EVENT', 'event_type': 'TREM', 'code': 'CCER3', 'month': '6', 'block_num': '2'},
    {'tipo': 'EVENT', 'event_type': 'FIRE', 'code': 'CER', 'month': '6', 'block_num': '2'},
    {'tipo': 'EVENT', 'event_type': 'CULT', 'code': 'S', 'month': '8', 'block_num': '2'},
    {'tipo': 'EVENT', 'event_type': 'CROP', 'code': 'DEG', 'month': '9', 'block_num': '2'}, # Código DEG (Pastagem degradada)
    {'tipo': 'EVENT', 'event_type': 'PLTM', 'code': None, 'month': '9', 'block_num': '2'},
    {'tipo': 'EVENT', 'event_type': 'FRST', 'code': None, 'month': '9', 'block_num': '2'},
]

# Bloco para Transição Automática SAVANA -> PASTAGEM (Idêntico ao Bloco 2)
BLOCO_SAVANA_TO_PASTAGEM_EVENTS = BLOCO_DESMATAMENTO_PASTAGEM_EVENTS

# Bloco de Pastagem (Mantendo Pastagem Tradicional - CROP BE8)
BLOCO_PASTAGEM_EVENTS = [
    {'tipo': 'EVENT', 'event_type': 'CROP', 'code': 'BE8', 'month': '1', 'block_num': '1'},
    {'tipo': 'EVENT', 'event_type': 'GRAZ', 'code': 'GM', 'month': '1', 'block_num': '1'},
    {'tipo': 'EVENT', 'event_type': 'GRAZ', 'code': 'GM', 'month': '2', 'block_num': '1'},
    {'tipo': 'EVENT', 'event_type': 'GRAZ', 'code': 'GM', 'month': '3', 'block_num': '1'},
    {'tipo': 'EVENT', 'event_type': 'GRAZ', 'code': 'GM', 'month': '4', 'block_num': '1'},
    {'tipo': 'EVENT', 'event_type': 'GRAZ', 'code': 'GM', 'month': '5', 'block_num': '1'},
    {'tipo': 'EVENT', 'event_type': 'LAST', 'code': None, 'month': '6', 'block_num': '1'},
    {'tipo': 'EVENT', 'event_type': 'SENM', 'code': None, 'month': '6', 'block_num': '1'},
    {'tipo': 'EVENT', 'event_type': 'GRAZ', 'code': 'GM', 'month': '6', 'block_num': '1'},
    {'tipo': 'EVENT', 'event_type': 'GRAZ', 'code': 'GM', 'month': '7', 'block_num': '1'},
    {'tipo': 'EVENT', 'event_type': 'GRAZ', 'code': 'GM', 'month': '8', 'block_num': '1'},
    {'tipo': 'EVENT', 'event_type': 'GRAZ', 'code': 'GM', 'month': '9', 'block_num': '1'},
    {'tipo': 'EVENT', 'event_type': 'FRST', 'code': None, 'month': '9', 'block_num': '1'},
    {'tipo': 'EVENT', 'event_type': 'GRAZ', 'code': 'GM', 'month': '10', 'block_num': '1'},
    {'tipo': 'EVENT', 'event_type': 'GRAZ', 'code': 'GM', 'month': '11', 'block_num': '1'},
    {'tipo': 'EVENT', 'event_type': 'GRAZ', 'code': 'GM', 'month': '12', 'block_num': '1'},
]

# Bloco de Soja (Soja + Pastagem, Simplificado para 1 repetição anual)
BLOCO_SOJA_EVENTS = [
    # CROP BE8 (Janeiro a Setembro)
    {'tipo': 'EVENT', 'event_type': 'CROP', 'code': 'BE8', 'month': '1', 'block_num': '1'},
    {'tipo': 'EVENT', 'event_type': 'GRAZ', 'code': 'GM', 'month': '1', 'block_num': '1'},
    {'tipo': 'EVENT', 'event_type': 'GRAZ', 'code': 'GM', 'month': '2', 'block_num': '1'},
    {'tipo': 'EVENT', 'event_type': 'LAST', 'code': None, 'month': '2', 'block_num': '1'},
    {'tipo': 'EVENT', 'event_type': 'SENM', 'code': None, 'month': '3', 'block_num': '1'},
    {'tipo': 'EVENT', 'event_type': 'GRAZ', 'code': 'GM', 'month': '3', 'block_num': '1'},
    {'tipo': 'EVENT', 'event_type': 'GRAZ', 'code': 'GM', 'month': '4', 'block_num': '1'},
    {'tipo': 'EVENT', 'event_type': 'GRAZ', 'code': 'GM', 'month': '5', 'block_num': '1'},
    {'tipo': 'EVENT', 'event_type': 'GRAZ', 'code': 'GM', 'month': '6', 'block_num': '1'},
    {'tipo': 'EVENT', 'event_type': 'GRAZ', 'code': 'GM', 'month': '7', 'block_num': '1'},
    {'tipo': 'EVENT', 'event_type': 'GRAZ', 'code': 'GM', 'month': '8', 'block_num': '1'},
    {'tipo': 'EVENT', 'event_type': 'GRAZ', 'code': 'GM', 'month': '9', 'block_num': '1'},
    
    # Preparação para Soja (Outubro)
    {'tipo': 'EVENT', 'event_type': 'LAST', 'code': None, 'month': '9', 'block_num': '1'}, # Fim de Pastagem/Cerrado (BE8)
    {'tipo': 'EVENT', 'event_type': 'CULT', 'code': 'S1C', 'month': '10', 'block_num': '1'},
    {'tipo': 'EVENT', 'event_type': 'FERT', 'code': 'MED', 'month': '10', 'block_num': '1'},
    
    # Plantio de Soja (Novembro)
    {'tipo': 'EVENT', 'event_type': 'CROP', 'code': 'SJ', 'month': '11', 'block_num': '1'},
    {'tipo': 'EVENT', 'event_type': 'PLTM', 'code': None, 'month': '11', 'block_num': '1'},
    {'tipo': 'EVENT', 'event_type': 'FRST', 'code': None, 'month': '11', 'block_num': '1'}, # Início crescimento Soja
    
    # Colheita de Soja (Março do ano seguinte, mas no mesmo ano de repetição 1)
    {'tipo': 'EVENT', 'event_type': 'LAST', 'code': None, 'month': '3', 'block_num': '1'},
    {'tipo': 'EVENT', 'event_type': 'SENM', 'code': None, 'month': '3', 'block_num': '1'},
    {'tipo': 'EVENT', 'event_type': 'HARV', 'code': 'HS', 'month': '3', 'block_num': '1'},
    
    # Replantio de Pastagem/Cerrado (Abril)
    {'tipo': 'EVENT', 'event_type': 'CROP', 'code': 'BE8', 'month': '4', 'block_num': '1'},
    {'tipo': 'EVENT', 'event_type': 'FRST', 'code': None, 'month': '4', 'block_num': '1'},
]

BLOCO_SAVANICA_EVENTS = BLOCO_PADRAO_SAVANA_EVENTS


def extrair_dados_mapbiomas(folder_path, lat_str, lon_str, nome_sitio):
    if not LIBS_INSTALADAS:
        return {'status': 'erro', 'message': "Erro Crítico: Bibliotecas 'rasterio' e 'pandas' não encontradas."}
    
    try:
        lat = float(lat_str)
        lon = float(lon_str)
        coords = [(lon, lat)]
    except ValueError:
        return {'status': 'erro', 'message': "Erro: Latitude e Longitude devem ser números válidos."}

    if not os.path.isdir(folder_path):
        return {'status': 'erro', 'message': f"Erro: Pasta não encontrada no caminho:\n{folder_path}"}

    rasters_encontrados = []
    for f in os.listdir(folder_path):
        if f.endswith(('.tif', '.tiff')):
            match = re.match(r'.*(\d{4}).*\.tif.*', f, re.IGNORECASE)
            if match:
                ano = int(match.group(1))
                if 1980 < ano < 2050:
                    rasters_encontrados.append((ano, os.path.join(folder_path, f)))
    
    if not rasters_encontrados:
        return {'status': 'erro', 'message': "Erro: Nenhum raster .tif contendo um ano (ex: 1985) foi encontrado na pasta."}

    rasters_encontrados = sorted(list(set(rasters_encontrados)))
    data_rows = []
    
    try:
        for ano, raster_path in rasters_encontrados:
            with rasterio.open(raster_path) as src:
                if not (src.bounds.left <= lon <= src.bounds.right and src.bounds.bottom <= lat <= src.bounds.top):
                    continue
                
                valor_pixel_codigo = list(src.sample(coords))[0][0]
                valor_pixel_nome = MAPBIOMAS_LEGEND.get(valor_pixel_codigo, f"Código Desconhecido ({valor_pixel_codigo})")
                
                data_rows.append({
                    'Ano': ano,
                    'Codigo_MapBiomas': valor_pixel_codigo,
                    'Classe_MapBiomas': valor_pixel_nome,
                    'ponto': nome_sitio
                })
        
        if not data_rows:
            return {'status': 'aviso', 'message': "Aviso: Nenhum dado extraído MapBiomas para o ponto."}

        df = pd.DataFrame(data_rows)
        return {'status': 'ok', 'data': df}

    except Exception as e:
        return {'status': 'erro', 'message': f"Erro durante a extração do raster MapBiomas: {e}"}

def find_raster_file(directory, prefix):
    for f in os.listdir(directory):
        f_lower = f.lower()
        if f_lower.startswith(prefix) and f_lower.endswith(('.tif', '.tiff')):
            return os.path.join(directory, f)
    return None

def extrair_dados_solo(base_folder_path, profundidade, lat_str, lon_str, nome_sitio):
    if not LIBS_INSTALADAS:
        return {'status': 'erro', 'message': "Erro Crítico: Bibliotecas 'rasterio' e 'pandas' não encontradas."}

    try:
        lat = float(lat_str)
        lon = float(lon_str)
        coords = [(lon, lat)]
    except ValueError:
        return {'status': 'erro', 'message': "Erro: Latitude e Longitude devem ser números válidos."}

    target_folder = os.path.join(base_folder_path, profundidade)
    if not os.path.isdir(target_folder):
        return {'status': 'erro', 'message': f"Erro: Pasta de profundidade não encontrada no caminho:\n{target_folder}"}

    variaveis_map = {
        'areia': 'sand',
        'silte': 'silt',
        'argila': 'clay',
        'densidade': 'bkd',
        'pH': 'ph'
    }
    
    raster_paths = {}
    for var_nome, var_prefix in variaveis_map.items():
        path = find_raster_file(target_folder, var_prefix)
        if path:
            raster_paths[var_nome] = path
        else:
            return {'status': 'erro', 'message': f"Erro: Não foi possível encontrar o raster para '{var_nome}' (prefixo '{var_prefix}')\nna pasta: {target_folder}"}
    
    try:
        extracted_values = {}
        for var_nome, path in raster_paths.items():
            with rasterio.open(path) as src:
                if not (src.bounds.left <= lon <= src.bounds.right and src.bounds.bottom <= lat <= src.bounds.top):
                    return {'status': 'erro', 'message': f"Erro: Coordenadas ({lat}, {lon}) estão fora dos limites do raster:\n{path}"}
                
                value = list(src.sample(coords))[0][0]
                extracted_values[var_nome] = value
        
        def get_value(var_nome, divide=False):
            val = extracted_values.get(var_nome)
            if val is None:
                return None
            try:
                numeric_val = float(val)
                if divide:
                    return numeric_val / 1000.0
                return numeric_val
            except (ValueError, TypeError):
                return None

        data_for_csv = {
            'ponto': nome_sitio,
            'lat': lat,
            'long': lon,
            'profundidade': profundidade,
            'areia': get_value('areia', divide=True),
            'silte': get_value('silte', divide=True),
            'argila': get_value('argila', divide=True),
            'densidade': get_value('densidade'),
            'pH': get_value('pH')
        }
        
        df = pd.DataFrame([data_for_csv])
        return {'status': 'ok', 'data': df}

    except Exception as e:
        return {'status': 'erro', 'message': f"Erro durante a extração do raster de solo:\n{e}"}

def haversine(lon1, lat1, lon2, lat2):
    lon1, lat1, lon2, lat2 = map(math.radians, [lon1, lat1, lon2, lat2])
    dlon = lon2 - lon1
    dlat = lat2 - lat1
    a = math.sin(dlat/2)**2 + math.cos(lat1) * math.cos(lat2) * math.sin(dlon/2)**2
    c = 2 * math.asin(math.sqrt(a))
    r = 6371
    return c * r

def encontrar_estacoes_proximas(folder_path, target_lat_str, target_lon_str, num_estacoes_desejadas, is_batch=False):
    try:
        target_lat = float(target_lat_str)
        target_lon = float(target_lon_str)
    except ValueError:
        return {'status': 'erro', 'message': "Erro: Latitude e Longitude devem ser números válidos (ex: -16.5, -49.2)"}

    if not os.path.isdir(folder_path):
        return {'status': 'erro', 'message': f"Erro: Pasta de estações INMET não encontrada:\n{folder_path}"}

    if not is_batch:
        sg.popup_quick_message("Lendo estações INMET...", non_blocking=True, background_color='gray')
    
    estacoes_encontradas = []
    try:
        csv_files = [f for f in os.listdir(folder_path) if f.lower().endswith('.csv')]
        if not csv_files:
            return {'status': 'erro', 'message': "Erro: Nenhum arquivo .csv encontrado na pasta do INMET."}

        for f in csv_files:
            filepath = os.path.join(folder_path, f)
            try:
                df_check = pd.read_csv(filepath, nrows=2, encoding='latin1')
                
                lat_col = next((col for col in df_check.columns if 'lat' in col.lower()), None)
                lon_col = next((col for col in df_check.columns if 'lon' in col.lower()), None)
                
                data_ini_col_name = next((col for col in df_check.columns if col.lower() == 'data_inicial'), None)
                data_fin_col_name = next((col for col in df_check.columns if col.lower() == 'data_final'), None)
                data_col_name = next((col for col in df_check.columns if 'data' in col.lower() and col.lower() not in ['data_inicial', 'data_final']), None)

                if lat_col and lon_col:
                    lat_estacao = df_check.iloc[0][lat_col]
                    lon_estacao = df_check.iloc[0][lon_col]
                    dist = haversine(target_lon, target_lat, lon_estacao, lat_estacao)
                    estacoes_encontradas.append({
                        'filepath': filepath,
                        'distancia': dist,
                        'data_ini_col': data_ini_col_name,
                        'data_fin_col': data_fin_col_name,
                        'data_col': data_col_name
                    })
            except Exception as e:
                print(f"Erro ao ler cabeçalho de {f}: {e}")

    except Exception as e:
        return {'status': 'erro', 'message': f"Erro ao varrer a pasta do INMET: {e}"}

    if not estacoes_encontradas:
        return {'status': 'erro', 'message': "Erro: Nenhuma estação com Lat/Lon foi lida com sucesso na pasta."}

    estacoes_encontradas.sort(key=lambda x: x['distancia'])
    top_estacoes = estacoes_encontradas[:num_estacoes_desejadas]

    if not is_batch:
        popup_message = f"{len(top_estacoes)} estação(ões) mais próxima(s) encontrada(s):\n\n"
        for i, est in enumerate(top_estacoes):
            date_str = ""
            try:
                df_dates = pd.read_csv(est['filepath'], encoding='latin1')
                
                if est['data_ini_col'] and est['data_fin_col']:
                    data_ini_series = df_dates[est['data_ini_col']].dropna()
                    data_fin_series = df_dates[est['data_fin_col']].dropna()
                    
                    if not data_ini_series.empty and not data_fin_series.empty:
                        data_ini = str(data_ini_series.iloc[0])
                        data_fin = str(data_fin_series.iloc[0])
                        date_str = f"({data_ini} a {data_fin})"
                    else:
                        date_str = "(Colunas 'data_inicial' ou 'data_final' vazias/inválidas)"
                
                elif est['data_col']:
                    data_col = est['data_col']
                    df_dates['datetime'] = pd.to_datetime(df_dates[data_col], errors='coerce')
                    df_valid_dates = df_dates.dropna(subset=['datetime'])

                    if not df_valid_dates.empty:
                        data_ini = df_valid_dates['datetime'].min().strftime('%Y-%m-%d')
                        data_fin = df_valid_dates['datetime'].max().strftime('%Y-%m-%d')
                        date_str = f"(Histórico: {data_ini} a {data_fin})"
                    else:
                        date_str = "(Data Histórica não encontrada/inválida)"
                else:
                    date_str = "(Coluna de Data não encontrada)"
                        
            except Exception as e:
                print(f"Erro ao ler datas da estação {est['filepath']}: {e}")
                date_str = f"(Erro interno ao ler datas)"
            
            popup_message += f"#{i+1}: {os.path.basename(est['filepath'])}\n"
            popup_message += f"     Distância: {est['distancia']:.2f} km {date_str}\n"
        
        popup_message += "\nDeseja continuar e processar os dados destas estações?"
        return {'status': 'ok', 'top_estacoes': top_estacoes, 'popup_message': popup_message}

    return {'status': 'ok', 'top_estacoes': top_estacoes}

def processar_medias_estacoes(top_estacoes, nome_sitio, is_batch=False):
    if not is_batch:
        sg.popup_quick_message(f"Processando dados das {len(top_estacoes)} estações...", non_blocking=True, background_color='gray')

    prec_series_list = []
    tmin_series_list = []
    tmax_series_list = []

    for estacao in top_estacoes:
        try:
            try:
                df = pd.read_csv(estacao['filepath'])
            except UnicodeDecodeError:
                df = pd.read_csv(estacao['filepath'], encoding='latin1')

            col_map = {}
            for col in df.columns:
                col_lower = str(col).lower()
                if 'data' in col_lower: col_map['data'] = col
                elif 'prec' in col_lower: col_map['prec'] = col
                elif 'min' in col_lower: col_map['tmin'] = col
                elif 'max' in col_lower: col_map['tmax'] = col
            
            required_cols = ['data', 'prec', 'tmin', 'tmax']
            if not all(k in col_map for k in required_cols):
                print(f"Aviso: Estação {os.path.basename(estacao['filepath'])} ignorada (colunas faltando).")
                continue

            df = df[list(col_map.values())]
            df.columns = required_cols
            df['data'] = pd.to_datetime(df['data'], errors='coerce')
            df['prec'] = pd.to_numeric(df['prec'], errors='coerce')
            df['tmin'] = pd.to_numeric(df['tmin'], errors='coerce')
            df['tmax'] = pd.to_numeric(df['tmax'], errors='coerce')
            df = df.dropna(subset=['data'])
            df['month'] = df['data'].dt.month
            
            prec_series_list.append(df.groupby('month')['prec'].mean())
            tmin_series_list.append(df.groupby('month')['tmin'].mean())
            tmax_series_list.append(df.groupby('month')['tmax'].mean())
        except Exception as e:
            print(f"Erro ao processar dados da estação {os.path.basename(estacao['filepath'])}:\n{e}")
            
    if not prec_series_list:
        return {'status': 'erro', 'message': "Erro: Nenhuma das estações mais próximas pôde ser processada com sucesso."}
    
    df_prec_all = pd.concat(prec_series_list, axis=1)
    df_tmin_all = pd.concat(tmin_series_list, axis=1)
    df_tmax_all = pd.concat(tmax_series_list, axis=1)
    
    final_prec = df_prec_all.mean(axis=1)
    final_tmin = df_tmin_all.mean(axis=1)
    final_tmax = df_tmax_all.mean(axis=1)
    
    df_out = pd.DataFrame({
        'ponto': nome_sitio,
        'mes': final_prec.index,
        'ppt': (final_prec / 10).values,
        'tmin': (final_tmin / 10).values,
        'tmax': (final_tmax / 10).values
    })
    
    df_out['ppt'] = df_out['ppt'].round(4)
    df_out['tmin'] = df_out['tmin'].round(4)
    df_out['tmax'] = df_out['tmax'].round(4)
    
    if not is_batch:
        downloads_path = str(Path.home() / "Downloads")
        nome_arquivo_csv = f"{nome_sitio}_inmet_clima_media.csv"
        output_csv = os.path.join(downloads_path, nome_arquivo_csv)
        df_out.to_csv(output_csv, index=False)
        
        return {'status': 'ok', 'message': (f"Sucesso! Dados climáticos processados (Média Mensal).\n"
                                             f"Média de {len(prec_series_list)} estação(ões) calculada.\n\n"
                                             f"Arquivo salvo em:\n{output_csv}")}
    
    return {'status': 'ok', 'data': df_out}

def gerar_csv_clima_anual(top_estacoes, nome_sitio, is_batch=False):
    if not is_batch:
        sg.popup_quick_message(f"Processando dados anuais das {len(top_estacoes)} estações...", non_blocking=True, background_color='gray')

    all_dataframes = []
    for estacao in top_estacoes:
        try:
            try:
                df = pd.read_csv(estacao['filepath'])
            except UnicodeDecodeError:
                df = pd.read_csv(estacao['filepath'], encoding='latin1')
                
            col_map = {}
            for col in df.columns:
                col_lower = str(col).lower()
                if 'data' in col_lower: col_map['data'] = col
                elif 'prec' in col_lower: col_map['prec'] = col
                elif 'min' in col_lower: col_map['tmin'] = col
                elif 'max' in col_lower: col_map['tmax'] = col
            
            required_cols = ['data', 'prec', 'tmin', 'tmax']
            if not all(k in col_map for k in required_cols):
                print(f"Aviso: Estação {os.path.basename(estacao['filepath'])} ignorada (colunas faltando).")
                continue
            
            df = df[list(col_map.values())]
            df.columns = required_cols
            df['data'] = pd.to_datetime(df['data'], errors='coerce')
            df['prec'] = pd.to_numeric(df['prec'], errors='coerce')
            df['tmin'] = pd.to_numeric(df['tmin'], errors='coerce')
            df['tmax'] = pd.to_numeric(df['tmax'], errors='coerce')
            
            df = df.dropna(subset=['data', 'prec', 'tmin', 'tmax'])
            all_dataframes.append(df)

        except Exception as e:
            print(f"Erro ao processar dados da estação {os.path.basename(estacao['filepath'])}:\n{e}")

    if not all_dataframes:
        return {'status': 'erro', 'message': "Erro: Nenhuma das estações mais próximas pôde ser processada com sucesso."}

    df_combined = pd.concat(all_dataframes)
    
    df_combined['year'] = df_combined['data'].dt.year
    df_combined['month'] = df_combined['data'].dt.month
    
    df_final_agg = df_combined.groupby(['year', 'month']).agg(
        prec_sum=('prec', 'sum'),
        tmin_mean=('tmin', 'mean'),
        tmax_mean=('tmax', 'mean')
    ).reset_index()
    
    df_final_agg['ppt'] = (df_final_agg['prec_sum'] / 10.0).round(4)
    df_final_agg['tmin'] = (df_final_agg['tmin_mean'] / 10.0).round(4)
    df_final_agg['tmax'] = (df_final_agg['tmax_mean'] / 10.0).round(4)
    
    df_final = df_final_agg[['year', 'month', 'ppt', 'tmin', 'tmax']]
    df_final = df_final.rename(columns={'month': 'mes'})
    df_final = df_final.sort_values(by=['year', 'mes'])

    if not is_batch:
        downloads_path = str(Path.home() / "Downloads")
        nome_arquivo_csv = f"{nome_sitio}_inmet_clima_anual.csv"
        output_csv = os.path.join(downloads_path, nome_arquivo_csv)
        
        df_final.to_csv(output_csv, index=False, float_format='%.4f')
        
        return {'status': 'ok', 'message': (f"Sucesso! CSV de clima anual (real) gerado.\n"
                                             f"Dados de {len(all_dataframes)} estação(ões) combinados.\n\n"
                                             f"Arquivo salvo em:\n{output_csv}")}

    return {'status': 'ok', 'data': df_final}

def gerar_arquivo_wth(top_estacoes, nome_sitio, is_batch=False):
    result = gerar_csv_clima_anual(top_estacoes, nome_sitio, is_batch=True)

    if result['status'] == 'erro':
        return result

    df_wth = result['data']

    df_wth['ppt_cm'] = df_wth['ppt'] / 10.0
    
    min_year = df_wth['year'].min()
    max_year = df_wth['year'].max()
    
    idx = pd.MultiIndex.from_product([range(min_year, max_year + 1), range(1, 13)], names=['year', 'mes'])
    df_wth_full = df_wth.set_index(['year', 'mes']).reindex(idx, fill_value=0.0).reset_index()
    
    df_wide = df_wth_full.pivot(index='year', columns='mes', values=['ppt_cm', 'tmin', 'tmax'])

    df_final = pd.DataFrame()
    for var in ['ppt_cm', 'tmin', 'tmax']:
        temp_df = df_wide[var].reset_index()
        temp_df['Var'] = var.split('_')[0].upper()
        temp_df = temp_df[['Var', 'year'] + list(range(1, 13))]
        df_final = pd.concat([df_final, temp_df])

    wth_content = []
    
    for index, row in df_final.iterrows():
        var_str = str(row['Var']).upper().ljust(4)
        year_str = str(int(row['year'])).ljust(4)
        
        line_parts = [var_str, year_str]
        
        for month in range(1, 13):
            value = row[month]
            val_str = f"{value:.2f}"
            
            formatted_wth_value = val_str.rjust(5)
            line_parts.append(formatted_wth_value)

        wth_content.append(" ".join(line_parts))

    downloads_path = str(Path.home() / "Downloads")
    nome_arquivo_wth = f"{nome_sitio}.wth"
    output_wth = os.path.join(downloads_path, nome_arquivo_wth)

    try:
        with open(output_wth, "w") as f:
            f.write("\n".join(wth_content))
    except Exception as e:
        return {'status': 'erro', 'message': f"Erro ao salvar arquivo .WTH:\n{e}"}
    
    return {'status': 'ok', 'message': (f"Sucesso! Arquivo .WTH gerado com dados anuais/mensais.\n"
                                         f"Intervalo de anos: {min_year} a {max_year}\n\n"
                                         f"Arquivo salvo em:\n{output_wth}")}

def gerar_site_100(solo_file_path, clima_file_path, template_file_path, nome_sitio, lat_str, lon_str):
    try:
        solo_df = pd.read_csv(solo_file_path)
        clima_df = pd.read_csv(clima_file_path)
        
        solo_df_filtered = solo_df[solo_df['ponto'] == nome_sitio]
        if solo_df_filtered.empty:
              return f"Erro: Ponto '{nome_sitio}' não encontrado no arquivo de solo."
        
        clima_df_filtered = clima_df[clima_df['ponto'] == nome_sitio]
        if clima_df_filtered.empty:
              return f"Erro: Ponto '{nome_sitio}' não encontrado no arquivo de clima."

        soilData = solo_df_filtered.iloc[0][['areia', 'silte', 'argila', 'densidade', 'pH']].values
        soilData = [f"{float(x):.5f}" for x in soilData]

        clima_df = clima_df_filtered.sort_values(by='mes')
        clima_map = clima_df.set_index('mes')[['ppt', 'tmin', 'tmax']].to_dict('index')

        if not os.path.exists(template_file_path):
            return f"Erro: Arquivo template não encontrado no caminho:\n{template_file_path}"

        with open(template_file_path, 'r') as f:
            site_100_lines = f.readlines()
        
        for mes in range(1, 13):
            line_index = mes + 2
            prec_val = clima_map.get(mes, {'ppt': 0.0})['ppt']
            formatted_prec = f"{round(prec_val / 10.0, 5):.5f}"
            site_100_lines[line_index] = site_100_lines[line_index].split()[0] + '  ' + formatted_prec + '\n'

        for mes in range(1, 13):
            line_index = mes + 38
            tmin_val = clima_map.get(mes, {'tmin': 0.0})['tmin']
            formatted_tmin = f"{round(tmin_val, 5):.5f}"
            site_100_lines[line_index] = site_100_lines[line_index].split()[0] + '  ' + formatted_tmin + '\n'

        for mes in range(1, 13):
            line_index = mes + 50
            tmax_val = clima_map.get(mes, {'tmax': 0.0})['tmax']
            formatted_tmax = f"{round(tmax_val, 5):.5f}"
            site_100_lines[line_index] = site_100_lines[line_index].split()[0] + '  ' + formatted_tmax + '\n'
            
        site_100_lines[67] = site_100_lines[67].split()[0] + '  ' + soilData[0] + '\n'
        site_100_lines[68] = site_100_lines[68].split()[0] + '  ' + soilData[1] + '\n'
        site_100_lines[69] = site_100_lines[69].split()[0] + '  ' + soilData[2] + '\n'
        
        site_100_lines[71] = site_100_lines[71].split()[0] + '  ' + soilData[3] + '\n'
        
        site_100_lines[100] = site_100_lines[100].split()[0] + '  ' + soilData[4] + '\n'
        
        formatted_lat = f"{float(lat_str):.6f}"
        formatted_lon = f"{float(lon_str):.6f}"
        
        site_100_lines[65] = site_100_lines[65].replace(site_100_lines[65].split()[0], formatted_lat)
        site_100_lines[66] = site_100_lines[66].replace(site_100_lines[66].split()[0], formatted_lon)


        downloads_path = str(Path.home() / "Downloads")
        nome_arquivo_saida = f"{nome_sitio}_site.100"
        output_path = os.path.join(downloads_path, nome_arquivo_saida)
        
        with open(output_path, 'w') as f:
            f.writelines(site_100_lines)

        return f"Sucesso! Arquivo '{nome_arquivo_saida}' criado e preenchido.\n\nSalvo em:\n{output_path}"

    except ValueError:
        return "Erro: Latitude e Longitude devem ser números válidos."
    except KeyError as e:
        return f"Erro: Coluna {e} faltando nos arquivos CSV de entrada. Verifique a formatação do solo e do clima médio."
    except Exception as e:
        return f"Erro durante a geração do site.100:\n{e}"

def processar_lote_dados(csv_pontos_path, mb_folder, solo_folder, solo_prof, inmet_folder, inmet_n_estacoes, inmet_mode):
    if not LIBS_INSTALADAS:
        return f"Erro Crítico: Bibliotecas ausentes (rasterio/pandas)."

    downloads_path = str(Path.home() / "Downloads")

    try:
        try:
            pontos_df = pd.read_csv(csv_pontos_path)
        except UnicodeDecodeError:
            pontos_df = pd.read_csv(csv_pontos_path, encoding='latin1')

        required_cols_map = {'sitio': 'ponto', 'latitude': 'lat', 'longitude': 'lon'}
        
        for required_col, internal_name in required_cols_map.items():
            if required_col not in pontos_df.columns:
                return f"Erro: O CSV de pontos deve conter a coluna exata '{required_col}'."
            pontos_df.rename(columns={required_col: internal_name}, inplace=True)
        
        total_pontos = len(pontos_df)
        log_messages = [f"--- Início do Processamento de Lote ({total_pontos} Pontos) ---"]


        log_messages.append("\n--- Processando LULC (Passo 1/3) ---")
        for index, row in pontos_df.iterrows():
            nome_sitio = str(row['ponto'])
            lat_str = str(row['lat'])
            lon_str = str(row['lon'])
            
            mb_result = extrair_dados_mapbiomas(mb_folder, lat_str, lon_str, nome_sitio)
            
            if mb_result['status'] == 'ok':
                output_mb = os.path.join(downloads_path, f"{nome_sitio}_mapbiomas_extracao.csv")
                mb_result['data'].to_csv(output_mb, index=False)
                log_messages.append(f" Ponto {index+1}/{total_pontos} ({nome_sitio}): LULC OK. Salvo CSV.")
            elif mb_result['status'] == 'aviso':
                log_messages.append(f" Ponto {index+1}/{total_pontos} ({nome_sitio}): LULC AVISO - {mb_result['message']}")
            else:
                log_messages.append(f" Ponto {index+1}/{total_pontos} ({nome_sitio}): LULC ERRO - {mb_result['message']}")


        log_messages.append("\n--- Processando Solo (Passo 2/3) ---")
        for index, row in pontos_df.iterrows():
            nome_sitio = str(row['ponto'])
            lat_str = str(row['lat'])
            lon_str = str(row['lon'])
            
            solo_result = extrair_dados_solo(solo_folder, solo_prof, lat_str, lon_str, nome_sitio)

            if solo_result['status'] == 'ok':
                output_solo = os.path.join(downloads_path, f"{nome_sitio}_solo_extracao_{solo_prof.replace('-', '')}.csv")
                solo_result['data'].to_csv(output_solo, index=False, float_format='%.6f')
                log_messages.append(f" Ponto {index+1}/{total_pontos} ({nome_sitio}): SOLO OK. Salvo CSV.")
            else:
                log_messages.append(f" Ponto {index+1}/{total_pontos} ({nome_sitio}): SOLO ERRO - {solo_result['message']}")


        log_messages.append("\n--- Processando Clima (Passo 3/3) ---")
        for index, row in pontos_df.iterrows():
            nome_sitio = str(row['ponto'])
            lat_str = str(row['lat'])
            lon_str = str(row['lon'])

            inmet_search = encontrar_estacoes_proximas(inmet_folder, lat_str, lon_str, inmet_n_estacoes, is_batch=True)
            
            if inmet_search['status'] == 'erro':
                log_messages.append(f" Ponto {index+1}/{total_pontos} ({nome_sitio}): Clima BUSCA ERRO - {inmet_search['message']}")
                continue
            
            top_estacoes = inmet_search['top_estacoes']
            
            media_result = processar_medias_estacoes(top_estacoes, nome_sitio, is_batch=True)
            if media_result['status'] == 'ok':
                output_media = os.path.join(downloads_path, f"{nome_sitio}_inmet_clima_media.csv")
                media_result['data'].to_csv(output_media, index=False)
                log_messages.append(f" Ponto {index+1}/{total_pontos} ({nome_sitio}): Clima Média OK. Salvo CSV.")
            else:
                log_messages.append(f" Ponto {index+1}/{total_pontos} ({nome_sitio}): Clima Média ERRO - {media_result['message']}")

            if inmet_mode in ['ambos', 'anual']:
                anual_result = gerar_csv_clima_anual(top_estacoes, nome_sitio, is_batch=True)
                if anual_result['status'] == 'ok':
                    output_anual = os.path.join(downloads_path, f"{nome_sitio}_inmet_clima_anual.csv")
                    anual_result['data'].to_csv(output_anual, index=False)
                    log_messages.append(f" Ponto {index+1}/{total_pontos} ({nome_sitio}): Clima Anual (CSV) OK. Salvo CSV.")
                    
                    if inmet_mode == 'ambos':
                        wth_result = gerar_arquivo_wth(top_estacoes, nome_sitio, is_batch=True)
                        if wth_result['status'] == 'ok':
                            log_messages.append(f" Ponto {index+1}/{total_pontos} ({nome_sitio}): Clima WTH OK. Salvo .WTH.")
                        else:
                            log_messages.append(f" Ponto {index+1}/{total_pontos} ({nome_sitio}): Clima WTH ERRO - {wth_result['message']}")
                else:
                    log_messages.append(f" Ponto {index+1}/{total_pontos} ({nome_sitio}): Clima Anual ERRO - {anual_result['message']}")
        
        log_messages.append("\n--- Processamento de Lote Concluído ---")
        return "\n".join(log_messages)

    except Exception as e:
        return f"Erro fatal durante o processamento em lote:\n{e}"

def get_next_available_year(timeline_data, values):
    # Extracts the simulation starting year from global settings first
    current_sim_start_year = 1958
    try:
        current_sim_start_year = int(values.get('-ANO_INICIO-') or 1958)
    except ValueError:
        pass
        
    # Find the maximum last year set in any block so far
    last_block_year = current_sim_start_year - 1
    
    for _, data in timeline_data:
        # Check if it's a block header (complete or manual)
        if data.get('tipo') in ['BLOCO_COMPLETO', 'HEADER']:
            header = data.get('header') if data.get('tipo') == 'BLOCO_COMPLETO' else data
            try:
                block_last_year = int(header['last_year'])
                if block_last_year > last_block_year:
                    last_block_year = block_last_year
            except (ValueError, KeyError):
                continue
                
    return str(last_block_year + 1)


def processar_mapbiomas_em_blocos(mb_csv_path, start_block_num_str, year_limit_str, timeline_data_current, values):
    try:
        start_block_num = int(start_block_num_str)
        year_limit = int(year_limit_str)
    except ValueError:
        return {'status': 'erro', 'message': "Erro: Número do Bloco Inicial e Ano Limite devem ser números inteiros válidos."}

    if not LIBS_INSTALADAS or 'pandas' not in sys.modules:
        return {'status': 'erro', 'message': "Erro Crítico: Bibliotecas 'pandas' ausentes."}

    try:
        try:
            df = pd.read_csv(mb_csv_path)
        except UnicodeDecodeError:
            df = pd.read_csv(mb_csv_path, encoding='latin1')
        
        required_cols = ['Ano', 'Classe_MapBiomas']
        if not all(col in df.columns for col in required_cols):
            return {'status': 'erro', 'message': f"CSV MapBiomas inválido. Requer colunas: {required_cols}"}

        # Determine next year based on current timeline
        next_available_year = int(get_next_available_year(timeline_data_current, values))
        
        # Filter data from the year *after* the last block added (next_available_year) up to the user limit (max 2015)
        df_filtered = df[(df['Ano'] >= next_available_year) & (df['Ano'] <= year_limit)].sort_values(by='Ano').reset_index(drop=True)
        
        if df_filtered.empty:
            return {'status': 'aviso', 'message': f"Nenhum dado MapBiomas encontrado no período de {next_available_year} até o ano limite {year_limit}. Verifique se o CSV cobre este período."}
        
        if df_filtered.iloc[0]['Ano'] > next_available_year:
             return {'status': 'aviso', 'message': f"O primeiro ano de MapBiomas encontrado ({df_filtered.iloc[0]['Ano']}) é posterior ao ano esperado ({next_available_year})."}


    except Exception as e:
        return {'status': 'erro', 'message': f"Erro ao ler ou processar CSV MapBiomas: {e}"}
        
    generated_blocks = []
    current_block_num = int(start_block_num)
    i = 0

    # --- NEW CODE START: Read weather choice from UI ---
    mb_weather_desc = values.get('-MB_WEATHER_COMBO-') or WEATHER_CHOICES_MAP['C']
    # --- NEW CODE END ---

    def get_lulc_category(class_name):
        if 'Formação Savânica' in class_name or 'Formação Campestre' in class_name or 'Vegetação Herbácea e Arbustiva' in class_name:
            return 'SAVANA'
        elif 'Soja' in class_name or 'Lavoura Temporária' in class_name or 'Agricultura' in class_name:
            return 'SOJA'
        elif 'Pastagem' in class_name or 'Agropecuária' in class_name or 'Lavoura Perene' in class_name or 'Mosaico de Usos' in class_name:
            return 'PASTAGEM'
        return 'OUTRO'


    while i < len(df_filtered):
        current_year = df_filtered.iloc[i]['Ano']
        current_class = df_filtered.iloc[i]['Classe_MapBiomas']
        
        if current_year > year_limit:
            break

        current_category = get_lulc_category(current_class)
        
        if current_category == 'OUTRO':
            i += 1
            continue


        # Start of grouping consecutive years of the same category
        j = i
        while j < len(df_filtered) and df_filtered.iloc[j]['Ano'] <= year_limit:
            next_category = get_lulc_category(df_filtered.iloc[j]['Classe_MapBiomas'])
            
            # Condição para quebrar o agrupamento:
            # 1. Mudança de Categoria.
            # 2. Descontinuidade no ano.
            if next_category != current_category or (j > i and df_filtered.iloc[j]['Ano'] != df_filtered.iloc[j-1]['Ano'] + 1):
                break
            j += 1
        
        consecutive_years = j - i
        
        block_start_year = current_year 
        block_last_year = df_filtered.iloc[j-1]['Ano']
        
        # Determine block events and parameters based on category
        block_events = None
        block_type_name = None
        weather_choice_desc = mb_weather_desc # USE THE READ VALUE
        repeats = str(consecutive_years)

        
        # --- SPECIAL TRANSITION CHECK: SAVANA -> PASTAGEM (Requires 2 years) ---
        is_transition_from_savana_to_pastagem = False
        
        if current_category == 'PASTAGEM' and consecutive_years >= 2:
            previous_year = block_start_year - 1
            
            # Check the original full df (not just df_filtered)
            prev_year_class_row = df[df['Ano'] == previous_year]

            if not prev_year_class_row.empty:
                prev_class_name = prev_year_class_row.iloc[0]['Classe_MapBiomas']
                prev_category = get_lulc_category(prev_class_name)
                
                if prev_category == 'SAVANA':
                    is_transition_from_savana_to_pastagem = True
        
        if is_transition_from_savana_to_pastagem:
            
            # 1. Generate the 2-year transition block (Block X)
            years_consumed = 2 # This block is a 2-year sequence
            transition_block_last_year = block_start_year + years_consumed - 1
            prev_class_name = df[df['Ano'] == block_start_year - 1].iloc[0]['Classe_MapBiomas']
            
            lulc_desc = f"{prev_class_name} -> {current_class} ({block_start_year} - {transition_block_last_year})"
            block_type_name = "Desmatamento/Plantio Pastagem (Transição SAVANA)"
            
            header = {
                'tipo': 'HEADER', 
                'num': str(current_block_num), 
                'last_year': str(transition_block_last_year), 
                'repeats': str(years_consumed),
                'out_year': str(block_start_year),
                'out_month': '1', 
                'out_interval': '1', 
                'weather': weather_choice_desc,
                'block_description': lulc_desc 
            }
            
            bloco_completo = {
                'tipo': 'BLOCO_COMPLETO',
                'header': header,
                'events': BLOCO_SAVANA_TO_PASTAGEM_EVENTS
            }
            
            generated_blocks.append((f"BLOCO {current_block_num}: {block_type_name} (Anos {block_start_year} - {transition_block_last_year}, Repete 2x)", bloco_completo))
            
            # Adiciona Terminador
            terminator_data = {'tipo': 'TERMINATOR'}
            linha_display_terminator = f"--- FIM DO BLOCO (-999 -999 X) do Bloco {current_block_num} ---"
            generated_blocks.append((linha_display_terminator, terminator_data))
            
            current_block_num += 1
            
            # 2. Update starting year and remaining consecutive years for potential standard blocks
            block_start_year = transition_block_last_year + 1
            consecutive_years -= years_consumed
            
            # If the remaining years are > 0, they will be handled by the next section below
        
        # --- End of Special Transition Check ---


        if current_category == 'SAVANA':
            block_events = BLOCO_SAVANICA_EVENTS
            block_type_name = "Savana Padrão (LULC)"
            
            years_to_process = consecutive_years
            current_idx = i

            while years_to_process > 0:
                
                if years_to_process >= 5:
                    num_repeats = '5'
                    years_in_block = 5
                    block_last_year_calc = df_filtered.iloc[current_idx + years_in_block - 1]['Ano']
                    display_type = ""
                else:
                    num_repeats = str(years_to_process)
                    years_in_block = years_to_process
                    block_last_year_calc = df_filtered.iloc[current_idx + years_in_block - 1]['Ano']
                    display_type = " (Ciclo Parcial)"

                lulc_desc = f"{current_class} ({df_filtered.iloc[current_idx]['Ano']} - {block_last_year_calc})"
                
                header = {
                    'tipo': 'HEADER', 
                    'num': str(current_block_num), 
                    'last_year': str(block_last_year_calc), 
                    'repeats': num_repeats,
                    'out_year': str(df_filtered.iloc[current_idx]['Ano']),
                    'out_month': '1', 
                    'out_interval': '1', 
                    'weather': weather_choice_desc,
                    'block_description': lulc_desc
                }
                
                bloco_completo = {
                    'tipo': 'BLOCO_COMPLETO',
                    'header': header,
                    'events': BLOCO_SAVANICA_EVENTS
                }
                
                weather_code = WEATHER_DESC_TO_CODE.get(weather_choice_desc, 'C')
                linha_display = f"BLOCO {current_block_num}: {block_type_name} (Anos {header['out_year']} - {header['last_year']}, Repete {num_repeats}x){display_type}"
                
                generated_blocks.append((linha_display, bloco_completo))
                terminator_data = {'tipo': 'TERMINATOR'}
                linha_display_terminator = f"--- FIM DO BLOCO (-999 -999 X) do Bloco {current_block_num} ---"
                generated_blocks.append((linha_display_terminator, terminator_data))
                
                current_block_num += 1
                current_idx += years_in_block
                years_to_process -= years_in_block
                
            i = j
            continue

        
        # Dynamic block generation for standard Pastagem/Soja (or remainder after transition)
        if current_category in ['PASTAGEM', 'SOJA'] and consecutive_years > 0:
            
            if current_category == 'PASTAGEM':
                block_events = BLOCO_PASTAGEM_EVENTS # Maintenance block uses BE8
                block_type_name = "Mantendo Pastagem Tradicional (LULC)"
                
            elif current_category == 'SOJA':
                block_events = BLOCO_SOJA_EVENTS
                block_type_name = "Soja (LULC) - Ciclo Anual"
                
            
            if consecutive_years > 0:
                
                final_block_last_year = block_start_year + consecutive_years - 1
                
                lulc_desc = f"{current_class} ({block_start_year} - {final_block_last_year})"

                header = {
                    'tipo': 'HEADER', 
                    'num': str(current_block_num), 
                    'last_year': str(final_block_last_year), 
                    'repeats': str(consecutive_years), 
                    'out_year': str(block_start_year),
                    'out_month': '1', 
                    'out_interval': '1', 
                    'weather': weather_choice_desc,
                    'block_description': lulc_desc
                }
                
                bloco_completo = {
                    'tipo': 'BLOCO_COMPLETO',
                    'header': header,
                    'events': block_events
                }
                
                weather_code = WEATHER_DESC_TO_CODE.get(weather_choice_desc, 'C')
                linha_display = f"BLOCO {current_block_num}: {block_type_name} (Anos {block_start_year} - {final_block_last_year}, Repete {consecutive_years}x, Clima: {weather_code})"
                
                generated_blocks.append((linha_display, bloco_completo))
                
                terminator_data = {'tipo': 'TERMINATOR'}
                linha_display_terminator = f"--- FIM DO BLOCO (-999 -999 X) do Bloco {current_block_num} ---"
                generated_blocks.append((linha_display_terminator, terminator_data))
                
                current_block_num += 1
            
            i = j
        
        else:
            i = j
        
    return {'status': 'ok', 'blocks': generated_blocks}

def gerar_texto_item(data):
    tipo = data.get('tipo', '')
    
    if tipo == 'BLOCO_COMPLETO':
        h = data['header']
        
        bloco_nome_extra = ""
        # 1. Blocos Padrão (Manuais e Automáticos) - Usa block_description para maior clareza
        if 'block_description' in h:
            bloco_nome_extra = f"({h['block_description']})"
        else:
            # Fallback para caso block_description falhe
            if h['num'] == '1':
                bloco_nome_extra = "(Padrão Savana)"
            elif h['num'] == '2':
                bloco_nome_extra = "(Desmatamento + Pastagem Trad)"

        weather_code = WEATHER_DESC_TO_CODE.get(h['weather'], h['weather']) # Pega o código curto

        texto = f"""{h['num']:<14}Block #             {bloco_nome_extra}
{h['last_year']:<14}Last year
{h['repeats']:<14}Repeats # years
{h['out_year']:<14}Output starting year
{h['out_month']:<14}Output month
{h['out_interval']:<14}Output interval
{weather_code:<14}Weather choice
"""
        for ev in data['events']:
            # Pega o código curto do evento para o arquivo .sch
            event_type_code = ev['event_type'].split(':')[0] if ':' in ev['event_type'] else ev['event_type']
            texto += f"      {ev['block_num']:<5}{ev['month']:<5}{event_type_code}\n"
            if ev['code']:
                texto += f"{ev['code']}\n"
        return texto.rstrip()

    elif tipo == 'HEADER':
        h = data
        weather_code = WEATHER_DESC_TO_CODE.get(h['weather'], h['weather']) # Pega o código curto
        
        bloco_nome_extra = "(Manual GUI)"
        if 'block_description' in h:
             bloco_nome_extra = f"({h['block_description']})"

        return f"""{h['num']:<14}Block #             {bloco_nome_extra}
{h['last_year']:<14}Last year
{h['repeats']:<14}Repeats # years
{h['out_year']:<14}Output starting year
{h['out_month']:<14}Output month
{h['out_interval']:<14}Output interval
{weather_code:<14}Weather choice
""".rstrip()
    
    elif tipo == 'EVENT':
        ev = data
        event_type_code = ev['event_type'].split(':')[0] if ':' in ev['event_type'] else ev['event_type']
        linha_evento = f"      {ev['block_num']:<5}{ev['month']:<5}{event_type_code}\n"
        if ev['code']:
            linha_evento += f"{ev['code']}\n"
        return linha_evento.rstrip()
    
    elif tipo == 'TERMINATOR':
        return "-999 -999 X\n"
        
    return "Não há preview para este item."


def update_full_preview(window, values, timeline_data):
    if not values:
        return ""
    
    try:
        conteudo_final = f"""{values['-ANO_INICIO-']:<14}Starting year
{values['-ANO_FIM-']:<14}Last year
{values['-SITE_FILE-']:<14}Site file name
0             Labeling type
-1            Labeling year
-1.00         Microcosm
-1            CO2 Systems
-1            pH shift
-1            Soil Warming
0             N input scalar option
0             OMAD scalar option
0             Climate scalar option
3             Initial system
{values['-INIT_CROP-']:<14}Initial crop
{values['-INIT_TREE-']:<14}Initial tree

"""
        
        for _, data in timeline_data:
            conteudo_final += gerar_texto_item(data) + "\n"
        
        window['-PREVIEW_COMPLETO-'].update(value=conteudo_final)
        return conteudo_final
    except Exception as e:
        print(f"Erro ao atualizar preview: {e}")
        return ""

# Define as opções padrão usando as chaves dos dicionários WEATHER_CHOICES_MAP para manter a consistência com o CENTURY.
DEFAULT_WEATHER_M = WEATHER_CHOICES_MAP['M']
DEFAULT_WEATHER_F = WEATHER_CHOICES_MAP['F']

# Layout dos parâmetros globais ajustado para alinhamento vertical
col_globais = [
    [
        sg.Column([
            [sg.Text("Nome do Sítio:", size=(16,1), justification='right')], 
            [sg.Text("Ano Início (Simulação):", size=(16,1), justification='right')], 
            [sg.Text("Initial crop:", size=(16,1), justification='right')]
        ], element_justification='right', pad=(0,0)), 
        sg.Column([
            [sg.Input("Lu_AFGO", key="-SITIO-", size=(14,1), enable_events=True)],
            [sg.Input("1958", key="-ANO_INICIO-", size=(8,1), enable_events=True)],
            [sg.Combo(INITIAL_CROP_OPTIONS, default_value='HER', key='-INIT_CROP-', size=(10,1), readonly=True, enable_events=True)]
        ], pad=(0,0)), 
        sg.Column([
            [sg.Text("Arquivo .site.100:", size=(16,1), justification='right')],
            [sg.Text("Ano Fim (Simulação):", size=(16,1), justification='right')],
            [sg.Text("Initial tree:", size=(16,1), justification='right')]
        ], element_justification='right', pad=(10,0)),
        sg.Column([
            [sg.Input("lu_site.100", key="-SITE_FILE-", size=(14,1), enable_events=True)],
            [sg.Input("2025", key="-ANO_FIM-", size=(8,1), enable_events=True)],
            [sg.Combo(INITIAL_TREE_OPTIONS, default_value='CER', key='-INIT_TREE-', size=(10,1), readonly=True, enable_events=True)]
        ], pad=(0,0))
    ]
]


col_bloco_padrao = [
    [sg.Text("Ano Final do Bloco (Last Year):", size=(25,1)), sg.Input("1982", key='-P_LAST_YEAR-', size=(10,1), enable_events=True)],
    [sg.Text("Ano Início da Saída (Output Year):", size=(25,1)), sg.Input("1958", key='-P_OUT_YEAR-', size=(10,1))],
    [sg.Text("Weather Choice:", size=(25,1)), 
     sg.Combo(WEATHER_CHOICES, default_value=DEFAULT_WEATHER_M, key='-P_WEATHER_COMBO-', readonly=True, size=(30, 1))],
    [sg.Button("Adicionar Bloco Padrão 1 (Cerrado)", key='-ADD_BLOCO_CERRADO-')]
]

col_bloco_desmatamento = [
    [sg.Text("Block #: 2", size=(25,1)), sg.Text("Repeats #: 2", size=(10,1))],
    [sg.Text("Ano Final do Bloco (Last Year):", size=(25,1)), sg.Input("1984", key='-D_LAST_YEAR-', size=(10,1), enable_events=True)],
    [sg.Text("Ano Início da Saída (Output Year):", size=(25,1)), sg.Input("1983", key='-D_OUT_YEAR-', size=(10,1))],
    [sg.Text("Weather Choice:", size=(25,1)), 
     sg.Combo(WEATHER_CHOICES, default_value=DEFAULT_WEATHER_F, key='-D_WEATHER_COMBO-', readonly=True, size=(30, 1))],
    [sg.Button("Adicionar Bloco 2 (Desmat. + Past. Trad)", key='-ADD_BLOCO_DESMATAMENTO-')]
]

col_mapbiomas_automatico = [
    [sg.Text("CSV MapBiomas Extraído:", size=(25,1)), 
     sg.Input(key='-MB_CSV_FILE-', size=(15,1)), 
     sg.FileBrowse("Procurar", target='-MB_CSV_FILE-', file_types=(("CSV Files", "*.csv"),))],
    [sg.Text("Bloco # Inicial (ex: 3):", size=(25,1)), sg.Input("3", key='-MB_START_BLOCK_NUM-', size=(5,1))],
    [sg.Text("Ano Limite de Análise (Máx 2015):", size=(25,1)), sg.Input("2015", key='-MB_YEAR_LIMIT-', size=(5,1))],
    [sg.Text("Weather Choice (LULC Auto):", size=(25,1)), 
     sg.Combo(WEATHER_CHOICES, default_value=WEATHER_CHOICES_MAP['C'], key='-MB_WEATHER_COMBO-', readonly=True, size=(30, 1))],
    [sg.Button("Gerar Blocos LULC (até 2015)", key='-GENERATE_LULC_BLOCKS-', size=(30, 1), button_color=('white', '#8A2BE2'))]
]

col_bloco_manual_layout = [
    [sg.Text("Block #:", size=(12,1)), sg.Input("3", size=(5,1), key='-B_NUM-')],
    [sg.Text("Last Year:", size=(12,1)), sg.Input(size=(10,1), key='-B_LAST_YEAR-')],
    [sg.Text("Repeats #:", size=(12,1)), sg.Input("1", size=(5,1), key='-B_REPEATS-')],
    [sg.Text("Output Start Year:", size=(12,1)), sg.Input(size=(10,1), key='-B_OUT_YEAR-')],
    [sg.Text("Output Month:", size=(12,1)), sg.Input("1", size=(5,1), key='-B_OUT_MONTH-')],
    [sg.Text("Output Interval:", size=(12,1)), sg.Input("1", size=(5,1), key='-B_OUT_INTERVAL-')],
    [sg.Text("Weather:", size=(12,1)), 
     sg.Combo(WEATHER_CHOICES, default_value=DEFAULT_WEATHER_M, key='-B_WEATHER_COMBO-', readonly=True, size=(30, 1))],
    [sg.Button("Adicionar Cabeçalho de Bloco Manual")]
]

col_evento_layout = [
    [sg.Text("Tipo de Evento:", size=(15,1)), 
     sg.Combo(ALL_EVENT_DESCRIPTIONS, key='-E_TIPO-', size=(20,1), readonly=True, enable_events=True)],
    [sg.Text("Código Específico:", size=(15,1), key='-E_CODIGO_TXT-', visible=False), 
     sg.Combo([], key='-E_CODIGO_COMBO-', size=(35,1), readonly=True, visible=False)],
    [sg.Text("Ano (1, 2, ...):", size=(15,1)), sg.Input(size=(5,1), key='-E_BLOCK_NUM-')],
    [sg.Text("Mês (1-12):", size=(15,1)), sg.Input(size=(5,1), key='-E_MES-')],
    [sg.Button("Adicionar Evento Manual"), sg.Button("Fechar Bloco (-999)", key="Adicionar Terminador de Bloco (-999)", button_color=('white', 'firebrick'))],
    [sg.Button("Gerar Arquivo .SCH", size=(15,2)), sg.Button("Sair", size=(10,2))]
]

# INSERÇÃO CORRIGIDA: Definição de coluna_timeline (estava faltando ou mal posicionada)
coluna_timeline = [
    [sg.Text("Preview do Arquivo .SCH", font=("Helvetica", 15))],
    [sg.Multiline(size=(45, 30), key='-PREVIEW_COMPLETO-', disabled=True, font=("Courier New", 9), background_color='#F0F0F0', text_color='black')],
    [sg.Text("Linha do Tempo (Itens Adicionados)", font=("Helvetica", 10))],
    [sg.Listbox(values=[], size=(45, 10), key='-TIMELINE-', enable_events=True)],
    [sg.Button("Selecionar Item", key="Carregar Item Selecionado"), sg.Button("Remover Selecionado"), sg.Button("Limpar Tudo")]
]


col_coordenadas = [
    [sg.Text("Latitude (ex: -16.5):", size=(20,1)), sg.Input(size=(15,1), key='-MB_LAT-')],
    [sg.Text("Longitude (ex: -49.2):", size=(20,1)), sg.Input(size=(15,1), key='-MB_LON-')],
]

col_mapbiomas = [
    [sg.Text("Pasta Origem Lulc (.tif):", size=(20,1)), 
     sg.Input(key='-MB_FOLDER-', size=(18,1)), 
     sg.FolderBrowse("Procurar", target='-MB_FOLDER-')],
    [sg.Button("Extrair Dados", key='-MB_EXTRACT-', size=(30, 1), button_color=('white', 'darkgreen'))]
]

col_solo = [
    [sg.Text("Pasta Origem Solo (.tif):", size=(20,1)), 
     sg.Input(key='-SOLO_FOLDER-', size=(18,1)), 
     sg.FolderBrowse("Procurar", target='-SOLO_FOLDER-')],
    [sg.Text("Profundidade:", size=(20,1)),
     sg.Combo(['0-20cm', '0-30cm'], default_value='0-20cm', key='-SOLO_PROF-', readonly=True, size=(10,1))],
    [sg.Button("Extrair Dados", key='-SOLO_EXTRACT-', size=(30, 1), button_color=('white', 'darkblue'))]
]

col_inmet = [
    [sg.Text("Pasta Estações INMET (.csv):", size=(20,1)), 
     sg.Input(key='-INMET_FOLDER-', size=(18,1)), 
     sg.FolderBrowse("Procurar", target='-INMET_FOLDER-')],
    [sg.Text("Nº de Estações p/ Média:", size=(20,1)), 
     sg.Combo(['1', '2', '3'], default_value='3', key='-INMET_NUM_ESTACOES-', readonly=True, size=(5,1))],
    [sg.Button("Processar Média Clima", key='-INMET_PROCESS-', size=(30, 1), button_color=('white', 'darkred'))],
    [sg.Button("Criar Histórico.csv (Ano/Mês)", key='-INMET_WTH_CSV-', size=(30, 1), button_color=('white', '#800080'))],
    [sg.Button("Criar Arquivo .WTH", key='-INMET_WTH_FILE-', size=(30, 1), button_color=('white', 'indigo'))]
]

col_site_100_creator = [
    [sg.Text("1. Arquivo de Solo (média):", size=(20,1)), 
     sg.Input(key='-SITE_SOLO_FILE-', size=(18,1)), 
     sg.FileBrowse("Procurar", target='-SITE_SOLO_FILE-', file_types=(("CSV Files", "*.csv"),))],
    [sg.Text("2. Arquivo de Clima (média):", size=(20,1)), 
     sg.Input(key='-SITE_CLIMA_FILE-', size=(18,1)), 
     sg.FileBrowse("Procurar", target='-SITE_CLIMA_FILE-', file_types=(("CSV Files", "*.csv"),))],
    [sg.Text("3. Template site.100:", size=(20,1)), 
     sg.Input(key='-SITE_TEMPLATE_FILE-', size=(18,1)), 
     sg.FileBrowse("Procurar", target='-SITE_TEMPLATE_FILE-', file_types=(("Template Files", "*.100"),))],
    [sg.Button("Criar Arquivo site.100", key='-SITE_100_CREATE-', size=(30, 1), button_color=('white', 'darkgreen'))]
]

layout_col_1 = [
    [sg.Image(filename='lapig_logo.png', size=(338, 111), key='-LOGO_LAPIG-'), sg.Text("Carbono", font=("Helvetica", 18, 'bold'))],
    [sg.Frame("Coordenadas (Lat/Lon) - Ponto Único", col_coordenadas, key='-F_COORD-')],
    [sg.Frame("Extrator de Uso do Solo (MapBiomas) - Ponto Único", col_mapbiomas, key='-F_MAPBIOMAS-')],
    [sg.Frame("Extrator de Dados de Solo (Embrapa) - Ponto Único", col_solo, key='-F_SOLO-')],
    [sg.Frame("Extrator de Clima (INMET) - Ponto Único", col_inmet, key='-F_INMET-')],
    [sg.Frame("Criar arquivo SITE.100", col_site_100_creator, key='-F_SITE_CREATOR-')],
]

col_lote = [
    [sg.Text("csv Pontos (sitio,latitude,longitude)", size=(25,1)), 
     sg.Input(key='-LOTE_CSV-', size=(10,1)), 
     sg.FileBrowse("Procurar", target='-LOTE_CSV-', file_types=(("CSV Files", "*.csv"),))],
    [sg.Text("Pasta LULC:", size=(12,1)), 
     sg.Input(key='-LOTE_MB_FOLDER-', size=(16,1)), 
     sg.FolderBrowse("Procurar", target='-LOTE_MB_FOLDER-')],
    [sg.Text("Pasta SOLO:", size=(12,1)),
     sg.Input(key='-LOTE_SOLO_FOLDER-', size=(16,1)), 
     sg.FolderBrowse("Procurar", target='-LOTE_SOLO_FOLDER-')],
    [sg.Text("Profundidade:", size=(12,1)),
     sg.Combo(['0-20cm', '0-30cm'], default_value='0-20cm', key='-LOTE_SOLO_PROF-', readonly=True, size=(8,1))],
    [sg.Text("Pasta CLIMA:", size=(12,1)), 
     sg.Input(key='-LOTE_INMET_FOLDER-', size=(16,1)), 
     sg.FolderBrowse("Procurar", target='-LOTE_INMET_FOLDER-')],
    [sg.Text("Nº Estações:", size=(12,1)), 
     sg.Combo(['1', '2', '3'], default_value='3', key='-LOTE_INMET_NUM_ESTACOES-', readonly=True, size=(8,1))],
    [sg.Text("Saída Clima:", size=(12,1)), 
     sg.Combo(['média', 'anual', 'ambos'], default_value='ambos', key='-LOTE_INMET_MODE-', readonly=True, size=(10,1))],
    [sg.Text("⚠️ Apenas pontos na mesma fazenda.", font=('Helvetica', 10, 'bold'), text_color='orange')],
    [sg.Button("EXECUTAR LOTE", key='-LOTE_EXECUTE-', size=(30, 2), button_color=('white', 'darkorange'))]
]

layout_col_2 = [
    [sg.Frame("LOTE: Extração de Dados (LULC, Solo, Clima)", col_lote, key='-F_LOTE_PROCESS-')]
]

col_bloco_manual_layout = [
    [sg.Text("Block #:", size=(12,1)), sg.Input("3", size=(5,1), key='-B_NUM-')],
    [sg.Text("Last Year:", size=(12,1)), sg.Input(size=(10,1), key='-B_LAST_YEAR-')],
    [sg.Text("Repeats #:", size=(12,1)), sg.Input("1", size=(5,1), key='-B_REPEATS-')],
    [sg.Text("Output Start Year:", size=(12,1)), sg.Input(size=(10,1), key='-B_OUT_YEAR-')],
    [sg.Text("Output Month:", size=(12,1)), sg.Input("1", size=(5,1), key='-B_OUT_MONTH-')],
    [sg.Text("Output Interval:", size=(12,1)), sg.Input("1", size=(5,1), key='-B_OUT_INTERVAL-')],
    [sg.Text("Weather:", size=(12,1)), 
     sg.Combo(WEATHER_CHOICES, default_value=DEFAULT_WEATHER_M, key='-B_WEATHER_COMBO-', readonly=True, size=(30, 1))],
    [sg.Button("Adicionar Cabeçalho de Bloco Manual")]
]

col_evento_layout = [
    [sg.Text("Tipo de Evento:", size=(15,1)), 
     sg.Combo(ALL_EVENT_DESCRIPTIONS, key='-E_TIPO-', size=(20,1), readonly=True, enable_events=True)],
    [sg.Text("Código Específico:", size=(15,1), key='-E_CODIGO_TXT-', visible=False), 
     sg.Combo([], key='-E_CODIGO_COMBO-', size=(35,1), readonly=True, visible=False)],
    [sg.Text("Ano (1, 2, ...):", size=(15,1)), sg.Input(size=(5,1), key='-E_BLOCK_NUM-')],
    [sg.Text("Mês (1-12):", size=(15,1)), sg.Input(size=(5,1), key='-E_MES-')],
    [sg.Button("Adicionar Evento Manual"), sg.Button("Fechar Bloco (-999)", key="Adicionar Terminador de Bloco (-999)", button_color=('white', 'firebrick'))],
    [sg.Button("Gerar Arquivo .SCH", size=(15,2)), sg.Button("Sair", size=(10,2))]
]

layout_col_3 = [
    [sg.Frame("Parâmetros Globais", col_globais, key='-F_GLOBAIS-')],
    [sg.Frame("Bloco 1: Padrão Savana", col_bloco_padrao, key='-F_PADRAO-')],
    [sg.Frame("Bloco 2: Desmatamento + Pastagem Trad", col_bloco_desmatamento, key='-F_DESMATAMENTO-')],
    [sg.Frame("Bloco X: Gerador Automático (MapBiomas)", col_mapbiomas_automatico, key='-F_LULC_AUTO-')],
    [sg.Frame("Construtor de Bloco (Manual)", col_bloco_manual_layout, key='-F_MANUAL-')],
    [sg.Frame("Construtor de Eventos (Manual)", col_evento_layout, key='-F_EVENTO-')]
]

layout_col_4 = [
    [sg.Column(coluna_timeline)]
]

layout = [
    [
        sg.Column(layout_col_1, vertical_alignment='top'),
        sg.VSeperator(),
        sg.Column(layout_col_2, vertical_alignment='top'),
        sg.VSeperator(),
        sg.Column(layout_col_3, vertical_alignment='top'),
        sg.VSeperator(),
        sg.Column(layout_col_4, vertical_alignment='top')
    ]
]

window = sg.Window("Construtor de Arquivos de Entrada CENTURY v28", layout, finalize=True)
timeline_data = []
global_keys = ['-SITIO-', '-SITE_FILE-', '-ANO_INICIO-', '-ANO_FIM-', '-INIT_CROP-', '-INIT_TREE-']
need_preview_update = True

if not LIBS_INSTALADAS:
    sg.popup_error("Erro de Dependência: 'rasterio' e/ou 'pandas' não encontrados.\n\nAs funcionalidades de extração de dados estão desabilitadas.\n\nPor favor, feche o app e instale com:\npip install rasterio pandas")
    window['-MB_EXTRACT-'].update(disabled=True)
    window['-MB_FOLDER-'].update(disabled=True)
    window['-SOLO_EXTRACT-'].update(disabled=True)
    window['-SOLO_FOLDER-'].update(disabled=True)
    window['-INMET_PROCESS-'].update(disabled=True)
    window['-INMET_FOLDER-'].update(disabled=True)
    window['-INMET_WTH_CSV-'].update(disabled=True)
    window['-INMET_WTH_FILE-'].update(disabled=True)
    window['-SITE_100_CREATE-'].update(disabled=True)
    window['-LOTE_EXECUTE-'].update(disabled=True)
    window['-LOTE_CSV-'].update(disabled=True)
    window['-LOTE_MB_FOLDER-'].update(disabled=True)
    window['-LOTE_SOLO_FOLDER-'].update(disabled=True)
    window['-LOTE_INMET_FOLDER-'].update(disabled=True)
    window['-GENERATE_LULC_BLOCKS-'].update(disabled=True)
    window['-MB_CSV_FILE-'].update(disabled=True)

else:
    try:
        import pandas as pd
    except ImportError:
        sg.popup_error("Biblioteca 'pandas' não encontrada. A extração INMET e a geração de SITE.100 serão desabilitadas.")
        window['-INMET_PROCESS-'].update(disabled=True)
        window['-INMET_FOLDER-'].update(disabled=True)
        window['-INMET_WTH_CSV-'].update(disabled=True)
        window['-INMET_WTH_FILE-'].update(disabled=True)
        window['-SITE_100_CREATE-'].update(disabled=True)
        window['-LOTE_EXECUTE-'].update(disabled=True)
        window['-GENERATE_LULC_BLOCKS-'].update(disabled=True)
        window['-MB_CSV_FILE-'].update(disabled=True)
        
while True:
    event, values = window.read()

    if need_preview_update:
        update_full_preview(window, values, timeline_data)
        need_preview_update = False

    if event == sg.WIN_CLOSED or event == "Sair":
        break

    if event == '-P_LAST_YEAR-':
        try:
            last_year_str = values[event]
            if last_year_str:
                next_year = int(last_year_str) + 1
                window['-D_OUT_YEAR-'].update(str(next_year)) # Linka Bloco 1 Last Year ao Bloco 2 Output Year
                window['-B_OUT_YEAR-'].update(str(next_year))
        except ValueError:
            pass
            
    if event == '-D_LAST_YEAR-':
        try:
            last_year_str = values[event]
            if last_year_str:
                next_year = int(last_year_str) + 1
                window['-B_OUT_YEAR-'].update(str(next_year)) # Linka Bloco 2 Last Year ao Bloco Manual Output Year
        except ValueError:
            pass

    if event in global_keys:
        need_preview_update = True

    if event == '-E_TIPO-':
        # Ao selecionar uma DESCRIÇÃO, precisamos encontrar o CÓDIGO (ex: 'CROP')
        descricao_selecionada = values['-E_TIPO-']
        # Remove o código curto da descrição (ex: 'CROP: Seleciona cultura' -> 'CROP')
        tipo_selecionado = EVENT_DESC_TO_CODE.get(descricao_selecionada) 
        
        if tipo_selecionado in CODIGOS_E_DESCRICOES:
            codigos_dict = CODIGOS_E_DESCRICOES.get(tipo_selecionado, {})
            DISPLAY_PARA_CODIGO.clear()
            lista_display = []
            for codigo, descricao in codigos_dict.items():
                display_text = f"{codigo} - {descricao}"
                lista_display.append(display_text)
                DISPLAY_PARA_CODIGO[display_text] = codigo
            
            if lista_display:
                window['-E_CODIGO_COMBO-'].update(values=lista_display, value=lista_display[0], visible=True)
            else:
                window['-E_CODIGO_COMBO-'].update(values=[], value='', visible=True)

            window['-E_CODIGO_TXT-'].update(visible=True)
        else:
            window['-E_CODIGO_COMBO-'].update(values=[], visible=False)
            window['-E_CODIGO_TXT-'].update(visible=False)
            
    if event == '-MB_EXTRACT-':
        folder = values['-MB_FOLDER-']
        lat = values['-MB_LAT-']
        lon = values['-MB_LON-']
        nome_sitio = values['-SITIO-']
        
        if not all([folder, lat, lon]):
            sg.popup_error("Por favor, preencha a Latitude, Longitude e selecione a Pasta de Rasters.")
            continue
        if not nome_sitio:
             sg.popup_error("Por favor, preencha o 'Nome do Sítio' primeiro.", "Ele será usado no nome do arquivo CSV.")
             continue
        
        sg.popup_no_buttons("Processando MapBiomas... Isso pode levar alguns segundos.", auto_close=True, auto_close_duration=1, non_blocking=True)
        
        result = extrair_dados_mapbiomas(folder, lat, lon, nome_sitio)

        if result['status'] == 'ok':
            df = result['data']
            downloads_path = str(Path.home() / "Downloads")
            nome_arquivo_csv = f"{nome_sitio}_mapbiomas_extracao.csv"
            output_csv = os.path.join(downloads_path, nome_arquivo_csv)
            df.to_csv(output_csv, index=False)
            message = f"Sucesso! {len(df)} anos extraídos.\n\nArquivo salvo em:\n{output_csv}"
            window['-MB_CSV_FILE-'].update(output_csv)
        elif result['status'] == 'aviso':
             message = result['message']
        else:
            message = result['message']

        sg.popup(message, title="Resultado da Extração MapBiomas")


    if event == '-SOLO_EXTRACT-':
        folder = values['-SOLO_FOLDER-']
        prof = values['-SOLO_PROF-']
        lat = values['-MB_LAT-']
        lon = values['-MB_LON-']
        nome_sitio = values['-SITIO-']

        if not all([folder, lat, lon]):
            sg.popup_error("Por favor, preencha a Latitude, Longitude e selecione a Pasta Origem (Solo).")
            continue
        if not nome_sitio:
             sg.popup_error("Por favor, preencha o 'Nome do Sítio' primeiro.", "Ele será usado no nome do arquivo CSV.")
             continue

        sg.popup_no_buttons("Processando Solo... Isso pode levar alguns segundos.", auto_close=True, auto_close_duration=1, non_blocking=True)
        
        result = extrair_dados_solo(folder, prof, lat, lon, nome_sitio)
        
        if result['status'] == 'ok':
            df = result['data']
            downloads_path = str(Path.home() / "Downloads")
            nome_arquivo_csv = f"{nome_sitio}_solo_extracao_{prof.replace('-', '')}.csv"
            output_csv = os.path.join(downloads_path, nome_arquivo_csv)
            df.to_csv(output_csv, index=False, float_format='%.6f')
            message = f"Sucesso! Dados de solo extraídos.\n\nArquivo salvo em:\n{output_csv}"
        else:
            message = result['message']

        sg.popup(message, title="Resultado da Extração de Solo")

    if event == '-INMET_PROCESS-':
        folder = values['-INMET_FOLDER-']
        lat = values['-MB_LAT-']
        lon = values['-MB_LON-']
        nome_sitio = values['-SITIO-']
        num_estacoes = int(values['-INMET_NUM_ESTACOES-'])

        if not all([folder, lat, lon]):
            sg.popup_error("Por favor, preencha a Latitude, Longitude e selecione a Pasta Estações INMET.")
            continue
        if not nome_sitio:
             sg.popup_error("Por favor, preencha o 'Nome do Sítio' primeiro.", "Ele será usado no nome do arquivo CSV.")
             continue
        
        resultado_busca = encontrar_estacoes_proximas(folder, lat, lon, num_estacoes, is_batch=False)
        
        if resultado_busca['status'] == 'erro':
            sg.popup_error(resultado_busca['message'])
            continue

        resposta = sg.popup_yes_no(resultado_busca['popup_message'], title="Estações Encontradas")
        
        if resposta == 'Yes':
            window.disable()
            sg.popup_quick_message("Processando Média Clima INMET... Por favor aguarde.", non_blocking=True, background_color='gray', text_color='white')
            window.refresh()
            
            result = processar_medias_estacoes(resultado_busca['top_estacoes'], nome_sitio, is_batch=False)
            
            window.enable()
            sg.popup(result['message'], title="Resultado do Processamento INMET")
        else:
            sg.popup("Processamento INMET cancelado pelo usuário.", title="Cancelado")

    if event == '-INMET_WTH_CSV-' or event == '-INMET_WTH_FILE-':
        folder = values['-INMET_FOLDER-']
        lat = values['-MB_LAT-']
        lon = values['-MB_LON-']
        nome_sitio = values['-SITIO-']
        num_estacoes = int(values['-INMET_NUM_ESTACOES-'])
        is_wth_file = (event == '-INMET_WTH_FILE-')
        file_type = ".WTH" if is_wth_file else "CSV Anual"
        
        if not all([folder, lat, lon, nome_sitio]):
            sg.popup_error(f"Por favor, preencha todos os campos necessários para gerar o arquivo {file_type}.")
            continue
        
        resultado_busca = encontrar_estacoes_proximas(folder, lat, lon, num_estacoes, is_batch=False)
        
        if resultado_busca['status'] == 'erro':
            sg.popup_error(resultado_busca['message'])
            continue

        resposta = sg.popup_yes_no(resultado_busca['popup_message'], title=f"Estações Encontradas para {file_type}")
        
        if resposta == 'Yes':
            window.disable()
            sg.popup_quick_message(f"Processando Arquivo {file_type}... Por favor aguarde.", non_blocking=True, background_color='gray', text_color='white')
            window.refresh()
            
            if is_wth_file:
                result = gerar_arquivo_wth(resultado_busca['top_estacoes'], nome_sitio, is_batch=False)
            else:
                result = gerar_csv_clima_anual(resultado_busca['top_estacoes'], nome_sitio, is_batch=False)
            
            window.enable()
            sg.popup(result['message'], title=f"Resultado da Geração de Arquivo {file_type}")
        else:
            sg.popup(f"Geração de Arquivo {file_type} cancelada pelo usuário.", title="Cancelado")


    if event == '-SITE_100_CREATE-':
        solo_file = values['-SITE_SOLO_FILE-']
        clima_file = values['-SITE_CLIMA_FILE-']
        template_file = values['-SITE_TEMPLATE_FILE-']
        nome_sitio = values['-SITIO-']
        lat = values['-MB_LAT-']
        lon = values['-MB_LON-']

        if not all([solo_file, clima_file, template_file, nome_sitio, lat, lon]):
            sg.popup_error("Por favor, preencha todos os campos obrigatórios (Arquivos de Solo/Clima/Template, Nome do Sítio, Lat/Lon).")
            continue
        
        sg.popup_no_buttons("Processando SITE.100...", auto_close=True, auto_close_duration=1, non_blocking=True)
        message = gerar_site_100(solo_file, clima_file, template_file, nome_sitio, lat, lon)
        sg.popup(message, title="Resultado da Geração de Site.100")

    if event == '-LOTE_EXECUTE-':
        csv_pontos_path = values['-LOTE_CSV-']
        mb_folder = values['-LOTE_MB_FOLDER-']
        solo_folder = values['-LOTE_SOLO_FOLDER-']
        solo_prof = values['-LOTE_SOLO_PROF-']
        inmet_folder = values['-LOTE_INMET_FOLDER-']
        inmet_n_estacoes = int(values['-LOTE_INMET_NUM_ESTACOES-'])
        inmet_mode = values['-LOTE_INMET_MODE-']

        if not all([csv_pontos_path, mb_folder, solo_folder, inmet_folder]):
            sg.popup_error("Preencha todos os caminhos de arquivo/pasta na seção 'LOTE' para executar.")
            continue
        
        window.disable()
        sg.popup_quick_message("EXECUTANDO LOTE... Isso pode levar vários minutos.", background_color='darkorange', text_color='white', non_blocking=True)
        window.refresh()
        
        log_message = processar_lote_dados(csv_pontos_path, mb_folder, solo_folder, solo_prof, inmet_folder, inmet_n_estacoes, inmet_mode)

        window.enable()
        sg.popup_scrolled(log_message, title="Resultado do Processamento em Lote", size=(80, 20))

    if event == '-GENERATE_LULC_BLOCKS-':
        mb_csv_file = values['-MB_CSV_FILE-']
        start_block = values['-MB_START_BLOCK_NUM-']
        year_limit = values['-MB_YEAR_LIMIT-']
        
        if not mb_csv_file or not os.path.exists(mb_csv_file):
            sg.popup_error("Por favor, selecione um arquivo CSV de MapBiomas existente.")
            continue
            
        window.disable()
        sg.popup_quick_message("Gerando Blocos LULC. Aguarde...", non_blocking=True, background_color='#8A2BE2', text_color='white')
        window.refresh()
        
        result = processar_mapbiomas_em_blocos(mb_csv_file, start_block, year_limit, timeline_data, values)
        
        window.enable()

        if result['status'] == 'erro':
            sg.popup_error(result['message'], title="Erro de Geração LULC")
            
        elif result['status'] == 'aviso':
            sg.popup_ok(result['message'], title="Aviso de Geração LULC")
            
        elif result['status'] == 'ok':
            # Adiciona os novos blocos ao final da lista
            new_blocks = result['blocks']
            timeline_data.extend(new_blocks)
            window['-TIMELINE-'].update(values=[item[0] for item in timeline_data])
            need_preview_update = True
            sg.popup_ok(f"Sucesso! {len(new_blocks)//2} Bloco(s) de LULC gerado(s) e adicionado(s) à linha do tempo.", title="Geração LULC Concluída")


    if event == "-ADD_BLOCO_CERRADO-":
        last_year = values['-P_LAST_YEAR-']
        out_year = values['-P_OUT_YEAR-']
        weather_desc = values['-P_WEATHER_COMBO-']
        
        try:
            last_year_int = int(last_year)
            out_year_int = int(out_year)
        except ValueError:
            sg.popup_error("Por favor, insira anos válidos (números inteiros) para o Bloco Padrão.")
            continue
            
        header_padrao = {
            'tipo': 'HEADER', 'num': '1', 'last_year': last_year, 'repeats': '5',
            'out_year': out_year, 'out_month': '1', 'out_interval': '1', 'weather': weather_desc,
            'block_description': f"Padrão Savana ({out_year_int}-{last_year_int})"
        }
        bloco_completo = {
            'tipo': 'BLOCO_COMPLETO',
            'header': header_padrao,
            'events': BLOCO_PADRAO_SAVANA_EVENTS
        }
        linha_display = f"BLOCO 1: Padrão Savana (Até {last_year}, Clima: {WEATHER_DESC_TO_CODE.get(weather_desc, 'M')})"
        timeline_data.append((linha_display, bloco_completo))
        
        terminator_data = {'tipo': 'TERMINATOR'}
        linha_display_terminator = "--- FIM DO BLOCO (-999 -999 X) ---"
        timeline_data.append((linha_display_terminator, terminator_data))
        
        window['-TIMELINE-'].update(values=[item[0] for item in timeline_data])
        need_preview_update = True

    if event == '-ADD_BLOCO_DESMATAMENTO-':
        last_year = values['-D_LAST_YEAR-']
        out_year = values['-D_OUT_YEAR-']
        weather_desc = values['-D_WEATHER_COMBO-']
        
        try:
            last_year_int = int(last_year)
            out_year_int = int(out_year)
        except ValueError:
            sg.popup_error("Por favor, insira anos válidos (números inteiros) para o Bloco de Desmatamento.")
            continue
            
        header_bloco2 = {
            'tipo': 'HEADER', 'num': '2', 'last_year': last_year, 'repeats': '2',
            'out_year': out_year, 'out_month': '1', 'out_interval': '1', 'weather': weather_desc,
            'block_description': f"Desmatamento + Pastagem Tradicional ({out_year_int}-{last_year_int})"
        }
        bloco_completo = {
            'tipo': 'BLOCO_COMPLETO',
            'header': header_bloco2,
            'events': BLOCO_DESMATAMENTO_PASTAGEM_EVENTS
        }
        linha_display = f"BLOCO 2: Desmatamento/Pastagem (Até {last_year}, Clima: {WEATHER_DESC_TO_CODE.get(weather_desc, 'F')})"
        timeline_data.append((linha_display, bloco_completo))
        
        terminator_data = {'tipo': 'TERMINATOR'}
        linha_display_terminator = "--- FIM DO BLOCO (-999 -999 X) ---"
        timeline_data.append((linha_display_terminator, terminator_data))
        
        window['-TIMELINE-'].update(values=[item[0] for item in timeline_data])
        need_preview_update = True

    if event == "Adicionar Cabeçalho de Bloco Manual":
        weather_desc = values['-B_WEATHER_COMBO-']

        try:
            int(values['-B_NUM-'])
            last_year = values['-B_LAST_YEAR-']
            out_year = values['-B_OUT_YEAR-']
            int(values['-B_REPEATS-'])
            int(out_year)
            int(values['-B_OUT_MONTH-'])
            int(values['-B_OUT_INTERVAL-'])
        except ValueError:
            sg.popup_error("Todos os campos do Bloco Manual (exceto Weather) devem ser números inteiros válidos.")
            continue
            
        bloco = {
            'tipo': 'HEADER',
            'num': values['-B_NUM-'], 'last_year': last_year,
            'repeats': values['-B_REPEATS-'], 'out_year': out_year,
            'out_month': values['-B_OUT_MONTH-'], 'out_interval': values['-B_OUT_INTERVAL-'],
            'weather': weather_desc,
            'block_description': f"Manual ({out_year}-{last_year})"
        }
        linha_display = f"BLOCO {bloco['num']} (Manual): LastYear={bloco['last_year']}, Clima: {WEATHER_DESC_TO_CODE.get(weather_desc, 'M')}"
        timeline_data.append((linha_display, bloco))
        window['-TIMELINE-'].update(values=[item[0] for item in timeline_data])
        need_preview_update = True

    if event == "Adicionar Evento Manual":
        descricao_selecionada = values['-E_TIPO-']
        # Obtém o código curto (ex: 'CROP')
        tipo_evento_completo = EVENT_DESC_TO_CODE.get(descricao_selecionada) 
        if tipo_evento_completo:
            tipo_evento = tipo_evento_completo.split(':')[0]
        else:
             tipo_evento = None

        if not tipo_evento:
            sg.popup_error("Selecione um 'Tipo de Evento/Opção' válido primeiro.")
            continue
            
        try:
            int(values['-E_MES-'])
            int(values['-E_BLOCK_NUM-'])
        except ValueError:
            sg.popup_error("O Mês e o Ano de Repetição devem ser números inteiros válidos.")
            continue
            
        codigo_real = None
        
        if tipo_evento in CODIGOS_E_DESCRICOES:
            display_selecionado = values['-E_CODIGO_COMBO-']
            codigo_real = DISPLAY_PARA_CODIGO.get(display_selecionado)
            # PLTM não tem código associado, IRRI/EROD não precisam de código específico, então relaxa a checagem para eles
            if tipo_evento not in ['IRRI', 'EROD', 'PLTM'] and not codigo_real:
                sg.popup_error(f"Erro: Nenhum Código Específico selecionado para '{tipo_evento}'.")
                continue
            linha_display = f"      EVENTO: {tipo_evento} -> ({codigo_real}), Mês={values['-E_MES-']}, RepetirAno={values['-E_BLOCK_NUM-']}"
        else:
            linha_display = f"      EVENTO: {tipo_evento}, Mês={values['-E_MES-']}, RepetirAno={values['-E_BLOCK_NUM-']}"

        evento = {
            'tipo': 'EVENT', 'event_type': tipo_evento, 'code': codigo_real, 
            'month': values['-E_MES-'], 'block_num': values['-E_BLOCK_NUM-']
        }
        timeline_data.append((linha_display, evento))
        window['-TIMELINE-'].update(values=[item[0] for item in timeline_data])
        need_preview_update = True

    if event == "Adicionar Terminador de Bloco (-999)":
        terminator_data = {'tipo': 'TERMINATOR'}
        linha_display = "--- FIM DO BLOCO (-999 -999 X) ---"
        timeline_data.append((linha_display, terminator_data))
        window['-TIMELINE-'].update(values=[item[0] for item in timeline_data])
        need_preview_update = True

    if event == "Carregar Item Selecionado":
        try:
            selected_display = values['-TIMELINE-'][0]
            selected_data = None
            for display, data in timeline_data:
                if display == selected_display:
                    selected_data = data
                    break
            
            if selected_data:
                tipo = selected_data.get('tipo')
                if tipo == 'HEADER':
                    weather_desc = selected_data.get('weather', DEFAULT_WEATHER_M)
                    window['-B_NUM-'].update(selected_data.get('num', ''))
                    window['-B_LAST_YEAR-'].update(selected_data.get('last_year', ''))
                    window['-B_REPEATS-'].update(selected_data.get('repeats', ''))
                    window['-B_OUT_YEAR-'].update(selected_data.get('out_year', ''))
                    window['-B_OUT_MONTH-'].update(selected_data.get('out_month', ''))
                    window['-B_OUT_INTERVAL-'].update(selected_data.get('out_interval', ''))
                    window['-B_WEATHER_COMBO-'].update(value=weather_desc)
                
                elif tipo == 'EVENT':
                    event_type = selected_data.get('event_type', '')
                    # Mapeia o código de volta para a descrição completa para preencher o combo
                    event_desc = TIPOS_DE_EVENTO_COM_CODIGO.get(event_type) or TIPOS_DE_EVENTO_SEM_CODIGO.get(event_type)
                    
                    window['-E_TIPO-'].update(event_desc)
                    # Força a atualização dos combos de código específico
                    window.write_event_value('-E_TIPO-', event_desc)
                    
                    window['-E_MES-'].update(selected_data.get('month', ''))
                    window['-E_BLOCK_NUM-'].update(selected_data.get('block_num', ''))

                    if selected_data.get('code'):
                        code = selected_data.get('code')
                        codigos_dict = CODIGOS_E_DESCRICOES.get(event_type, {})
                        desc = codigos_dict.get(code, '')
                        display_text = f"{code} - {desc}"
                        
                        window.refresh()
                        window['-E_CODIGO_COMBO-'].update(value=display_text)
                
                elif tipo == 'BLOCO_COMPLETO' or tipo == 'TERMINATOR':
                    sg.popup_ok("Não é possível editar este tipo de item.\nRemova e adicione novamente, se necessário.", title="Aviso")
        
        except IndexError:
            pass
        except Exception as e:
            print(f"Erro ao carregar item: {e}")


    if event == "Remover Selecionado":
        try:
            item_selecionado_display = values['-TIMELINE-'][0]
            for item in timeline_data:
                if item[0] == item_selecionado_display:
                    timeline_data.remove(item)
                    break
            window['-TIMELINE-'].update(values=[item[0] for item in timeline_data])
            need_preview_update = True
        except Exception:
            pass

    if event == "Limpar Tudo":
        timeline_data = []
        window['-TIMELINE-'].update(values=timeline_data)
        need_preview_update = True

    if event == "Gerar Arquivo .SCH":
        try:
            if not timeline_data:
                sg.popup_error("Linha do tempo está vazia!", "Adicione Blocos e Eventos primeiro.")
                continue
            
            nome_sitio = values["-SITIO-"]
            if not nome_sitio:
                sg.popup_error("O campo 'Nome do Sítio' está vazio.", "Por favor, defina um nome para o sítio (ex: Lu_AFGO).")
                continue
                
            nome_arquivo_saida = f"{nome_sitio}.SCH"
            
            conteudo_final = update_full_preview(window, values, timeline_data)
            
            if not conteudo_final:
                sg.popup_error("Erro ao gerar conteúdo final.")
                continue

            downloads_path = str(Path.home() / "Downloads")
            caminho_completo_saida = os.path.join(downloads_path, nome_arquivo_saida)

            with open(caminho_completo_saida, "w") as f:
                f.write(conteudo_final)

            sg.popup_ok(f"Arquivo '{nome_arquivo_saida}' gerado com sucesso!",
                        f"Salvo na sua pasta de Downloads:\n{caminho_completo_saida}")

        except Exception as e:
            sg.popup_error("Ocorreu um erro ao gerar o arquivo:", str(e))

window.close()