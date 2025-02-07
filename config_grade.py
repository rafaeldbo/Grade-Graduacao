import numpy as np
import pandas as pd

import warnings
from dotenv import load_dotenv
from os import getenv, path

load_dotenv(override=True)
ABS_PATH = path.abspath(path.dirname(__file__))
warnings.filterwarnings('ignore')

from utils import success, warning, info, cleaner

filename = getenv('config_file', 'Config_Grade.xlsx')
config_filepath = path.join(ABS_PATH, filename)
if not path.isfile(config_filepath):
    raise FileNotFoundError(f'Arquivo de configuração [{filename}] não encontrado no diretório atual!')    

def load_manual_data() -> pd.DataFrame:
    info('Carregando dados adicionados manualmente')
    data_manual_raw = pd.read_excel(config_filepath, 'Adição Horários Manuais', header=1)
    columns_to_check = ['Curso', 'Série', 'Turma', 'Nome Disciplina', 'Tipo Atividade', 'Dia da Semana', 'Hora início', 'Hora fim']
    for i, row in data_manual_raw.iterrows():
        if row[columns_to_check].isnull().any():
            empty_columns = row[columns_to_check].index[row[columns_to_check].isnull()].tolist()
            warning(f'Linha [{i+3}] da aba [Adição Horários Manuais] será ignorada devido as colunas {empty_columns} estarem vazias')
    data_manual_raw = data_manual_raw.dropna(subset=columns_to_check)
    data_manual_raw.rename(columns={
        'Curso': 'curso',
        'Série': 'serie',
        'Turma': 'turma',
        'Nome Disciplina': 'nome_disciplina',
        'Tipo Atividade': 'tipo_atividade',
        'Dia da Semana': 'dia_semana',
        'Hora início': 'hora_inicio',
        'Hora fim': 'hora_fim',
        'Docente': 'docentes',
    }, inplace=True)
    data_manual = data_manual_raw.copy()
    data_manual = data_manual.fillna('')
    data_manual = data_manual.map(cleaner)
    data_manual = data_manual.drop(['Observação'], axis=1)
    data_manual['hora_inicio'] = pd.to_datetime(data_manual['hora_inicio'], format='%H:%M:%S').dt.time
    data_manual['hora_fim'] = pd.to_datetime(data_manual['hora_fim'], format='%H:%M:%S').dt.time
    data_manual['titular'] = data_manual['docentes']
    data_manual['Origem'] = 'Manual'
    data_manual['cod_turma'] = data_manual.apply(lambda row: f"{row['curso']}_{row['serie']}{row['turma']}", axis=1)
    
    success(f'Dados adicionados manualmente carregados com sucesso!')
    return data_manual

def load_configs() -> dict:
    info('Configurando automação a paritir do arquivo de configuração')
    data_config = pd.read_excel(config_filepath, 'Dados Configuráveis', header=1)
    columns_to_check = ['Tipo', 'Dado']
    for i, row in data_config.iterrows():
        if row[columns_to_check].isnull().any():
            empty_columns = row[columns_to_check].index[row[columns_to_check].isnull()].tolist()
            warning(f'Linha [{i+3}] da aba [Dados Configuráveis] será ignorada devido a coluna [{empty_columns[-1]}] estar vazia')
            
    configs = {}
    configs['DISCIPLINES_2_SLOTS_ATTENDANCE'] = data_config[data_config['Tipo'] == 'Disciplina com 2 Atendimentos']['Dado'].apply(cleaner).to_list()
    configs['DISCIPLINES_4_SLOTS_CLASS'] = data_config[data_config['Tipo'] == 'Disciplina de 110 Horas']['Dado'].apply(cleaner).to_list()
    configs['DISCIPLINES_GRUPED_CLASS'] = data_config[data_config['Tipo'] == 'Disciplina com Turmas Unidas']['Dado'].apply(cleaner).to_list()
    configs['ASSISTANT_PROFESSORS'] = data_config[data_config['Tipo'] == 'Professor Auxiliar / Assistente de Ensino']['Dado'].apply(cleaner).to_list()
    configs['NINJA_MONITORIES'] = data_config[data_config['Tipo'] == 'Monitoria Ninja']['Dado'].apply(cleaner).to_list()
    configs['SPECIAL_CLASSES'] = data_config[data_config['Tipo'] == 'Aula Especial']['Dado'].apply(cleaner).to_list()
    configs['EXCLUSIVE_TIMETABLE'] = data_config[data_config['Tipo'] == 'Grade Separada']['Dado'].apply(cleaner).to_list()
    
    devlife = data_config[data_config['Tipo'] == 'Nome da Vida do Desenvolvedor']['Dado'].apply(cleaner).unique()
    if len(devlife) > 0 and devlife[-1] == '':
        warning(f'Nome da vida do desenvolvedor de software não foi encontrado, será utilizado o padrão: VIDA DE DESENVOLVEDOR DE SOFTWARE - DEVELOPER LIFE')
    if len(devlife) > 1:
        warning(f'Foram encontrados mais de um nome para a vida do desenvolvedor de software, será utilizado o último encontrado: {devlife[-1]}')
    configs['DEVELOPER_LIFE_NAME'] = devlife[-1] if len(devlife) > 0 else 'VIDA DE DESENVOLVEDOR DE SOFTWARE - DEVELOPER LIFE'
    
    timetable = data_config[data_config['Tipo'] == 'Nova Grade']['Dado'].apply(cleaner).unique()
    configs['NEW_TIMETABLE'] = timetable[-1] == 'Sim' if len(timetable) > 0 else False
    
    success(f'configurações realizadas com sucesso!')
    return configs

def remove_by_filters(data:pd.DataFrame) -> pd.DataFrame:
    info('Removendo horários solicitados manualmente')
    data_filters_raw = pd.read_excel(filename, 'Remoção Horários Space', header=1)
    columns_to_check = ['Curso', 'Série', 'Turma', 'Nome da Disciplina']
    for i, row in data_filters_raw.iterrows():
        if row[columns_to_check].isnull().any():
            empty_columns = row[columns_to_check].index[row[columns_to_check].isnull()].tolist()
            warning(f'Linha [{i+3}] da aba [Remoção Horários Space] será ignorada devido a coluna [{empty_columns}] estar vazia')
    data_filters = data_filters_raw.dropna(subset=columns_to_check).fillna('')
    data_filters = data_filters.map(cleaner)
    data_filters.rename(columns={
        'Curso': 'curso',
        'Série': 'serie',
        'Turma': 'turma',
        'Nome da Disciplina': 'nome_disciplina',
        'Tipo Atividade': 'tipo_atividade',
        'Dia da Semana': 'dia_semana',
    }, inplace=True)
    
    if not data_filters.empty:
        data_filters['filtro'] = data_filters.apply(lambda row: f"{row['curso']}_{row['serie']}{row['turma']}_[{row['nome_disciplina']}]{'' if row['tipo_atividade'] == '' else '_'+row['tipo_atividade']}{'' if row['dia_semana'] == '' else '_'+row['dia_semana']}", axis=1)
        filters = data_filters['filtro'].unique().tolist()
        
        data_filtering = data.copy()
        data_filtering['filtro_base'] = data_filtering.apply(lambda row: f"{row['cod_turma']}_[{row['nome_disciplina']}]", axis=1)
        data_filtering['filtro_tipo'] = data_filtering.apply(lambda row: f"{row['cod_turma']}_[{row['nome_disciplina']}]_{row['tipo_atividade']}", axis=1)
        data_filtering['filtro_dia'] = data_filtering.apply(lambda row: f"{row['cod_turma']}_[{row['nome_disciplina']}]_{row['dia_semana']}", axis=1)
        data_filtering['filtro_completo'] = data_filtering.apply(lambda row: f"{row['cod_turma']}_[{row['nome_disciplina']}]_{row['tipo_atividade']}_{row['dia_semana']}", axis=1)
        
        removed = data_filtering[
            (data_filtering['filtro_base'].isin(filters))
            | (data_filtering['filtro_tipo'].isin(filters))
            | (data_filtering['filtro_dia'].isin(filters))
            | (data_filtering['filtro_completo'].isin(filters))
        ]
        
        for i, row in removed.iterrows():
            warning(f'a {row["tipo_atividade"]} da turma {row["cod_turma"]}, de {row["nome_disciplina"]} que ocorre as {row["dia_semana"]} foi removida com sucesso')
        
        data_filtering = data_filtering.drop(removed.index)
        success(f'Horários removidos com sucesso!')
        return data_filtering
    return data
    
    
    