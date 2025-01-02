import json
import pandas as pd
import os
from datetime import datetime
import zipfile
import config as cfg

# Função para obter o nome do arquivo pbit
def obter_nome_arquivo_pbit(caminho_pasta):
    arquivos = os.listdir(caminho_pasta)
    arquivos_pbit = [arquivo for arquivo in arquivos if arquivo.endswith('.pbit')]
    if arquivos_pbit:
        return arquivos_pbit[0]
    else:
        return 'Nenhum arquivo .pbit encontrado na pasta'

# Função para verificar e renomear arquivos
def verificar_ou_renomear_arquivo(arquivo_pbit, arquivo_zip):
    if os.path.exists(arquivo_zip):
        print("Arquivo .zip já existe. Pulando para a próxima instrução.")
    else:
        os.rename(arquivo_pbit, arquivo_zip)

# Função para extrair arquivos do ZIP
def extrair_arquivos_zip(arquivo_zip, caminho_BI, arquivos_para_extrair):
    with zipfile.ZipFile(arquivo_zip, 'r') as zip_ref:
        for arquivo in arquivos_para_extrair:
            zip_ref.extract(arquivo, caminho_BI)

# Função auxiliar para carregar dados JSON
def carregar_dados_json(arquivo: str, encoding: str = 'utf-16-le') -> dict:
    try:
        with open(arquivo, 'r', encoding=encoding) as f:
            return json.load(f)
    except Exception as e:
        print(f"Erro ao carregar JSON: {arquivo} - {e}")
        return {}

# Funções para extração de dados
def extrair_dados_layout(layout: dict) -> pd.DataFrame:
    return pd.DataFrame([{'Páginas': section.get('displayName')} for section in layout.get('sections', [])])

def transformar_visuals(layout: dict) -> pd.DataFrame:
    visual_containers = []
    for section in layout.get('sections', []):
        for container in section.get("visualContainers", []):
            visual_data = processar_visual_container(section, container)
            if visual_data:
                visual_containers.append(visual_data)
    return pd.DataFrame(visual_containers)

def processar_visual_container(section: dict, container: dict) -> dict:
    config_data = json.loads(container.get("config", "{}"))
    visual_type = config_data.get("singleVisual", {}).get("visualType")
    position = next(iter(config_data.get("layouts", [])), {}).get("position", {})
    query_refs = [item.get("queryRef") for key, items in config_data.get("singleVisual", {}).get("projections", {}).items() for item in items if item.get("queryRef")]

    return {
        "Página": section.get('displayName', 'Sem Nome'),
        "X": int(position.get("x", 0)),
        "Y": int(position.get("y", 0)),
        "Altura": int(position.get("height", 0)),
        "Largura": int(position.get("width", 0)),
        "Tipo de visual": visual_type,
        "Medidas utilizadas": query_refs or "Não há medidas utilizadas no visual"
    }

def extrair_tabelas(model_data: dict) -> pd.DataFrame:
    tables = []
    for table in model_data.get('model', {}).get('tables', []):
        table_name = table.get("name", "")
        if table_name.startswith("DateTableTemplate") or table_name.startswith("LocalDateTable"):
            continue

        for column in table.get('columns', []):
            column_name = column.get("name", "")
            data_type = column.get('dataType', "")
            column_type = column.get('type', "")
            is_calculated = column_type in ['calculatedTableColumn', 'calculated']

            tables.append({
                'Tabela': table_name,
                'Coluna': column_name,
                'Tipo de dados': data_type,
                'Coluna calculada?': 'Sim' if is_calculated else 'Não'
            })
    return pd.DataFrame(tables)

def salvar_html_com_versao(salvar_path):
    if not os.path.exists(salvar_path):
        return salvar_path
    
    base, ext = os.path.splitext(salvar_path)
    versao = 2
    
    while os.path.exists(f"{base}_versão_{versao:02}{ext}"):
        versao += 1
    
    return f"{base}_versão_{versao:02}{ext}"

def gerar_documento_html(cfg, dataframes, salvar_path):
    # Inicializa o conteúdo HTML
    html_content = "<html><head><title>Documentação Power BI</title><style>"
    html_content += "table {width: 100%; border-collapse: collapse;} "
    html_content += "th, td {border: 1px solid #ddd; padding: 8px; text-align: left;} "
    html_content += "th {background-color: #4CAF50; color: white;} "
    html_content += "tr:nth-child(even) {background-color: #f2f2f2;} </style></head><body>"
    
    # Cabeçalho do documento
    html_content += "<h1>Documentação de Relatório Power BI</h1>"
    html_content += f"<p><strong>Data da documentação:</strong> {datetime.now().strftime('%d/%m/%Y')}</p>"
    html_content += f"<p><strong>Nome do Relatório:</strong> {cfg.nome_BI}</p>"
    
    # Inserir as tabelas extraídas
    for titulo, df in dataframes.items():
        html_content += f"<h2>{titulo.capitalize()}</h2>"
        html_content += df.to_html(index=False, escape=False)
    
    html_content += "</body></html>"

    # Salvar o arquivo HTML
    caminho_final = salvar_html_com_versao(salvar_path)
    with open(caminho_final, "w") as f:
        f.write(html_content)
    print(f'Documentação gerada com sucesso em: {caminho_final}')

# Função principal para execução do processo
def main():
    caminho_BI = cfg.caminho_BI  # Caminho onde os arquivos .pbit estão localizados
    nome_arquivo = obter_nome_arquivo_pbit(caminho_BI)  # Pega o nome do arquivo .pbit
    
    # Caminhos para os arquivos de saída
    modelo_path = os.path.join(cfg.caminho_modelo_word, cfg.nome_modelo_word)
    salvar_path = os.path.join(cfg.caminho_documentação, f'{nome_arquivo}_doc.html')
    
    arquivo_pbit = os.path.join(caminho_BI, f'{nome_arquivo}')
    arquivo_zip = os.path.join(caminho_BI, f'{nome_arquivo}.zip')
    
    verificar_ou_renomear_arquivo(arquivo_pbit, arquivo_zip)
    extrair_arquivos_zip(arquivo_zip, caminho_BI, ['Report/Layout', 'DataModelSchema'])
    
    layout_data = carregar_dados_json(os.path.join(caminho_BI, 'Report/Layout'))
    model_data = carregar_dados_json(os.path.join(caminho_BI, 'DataModelSchema'))
    
    # Extrai os dados
    layout_df = extrair_dados_layout(layout_data)
    visuals_df = transformar_visuals(layout_data)
    tables_df = extrair_tabelas(model_data)
    
    # Organize os dataframes
    dataframes = {
        'Layout': layout_df,
        'Visuals': visuals_df,
        'Tabelas': tables_df,
    }
    
    # Gera o documento HTML
    gerar_documento_html(cfg, dataframes, salvar_path)

# Executa o processo
if __name__ == "__main__":
    main()
