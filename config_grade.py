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
    data_config = pd.read_excel(config_filepath, 'Dados Configuráveis', header=1)
    configs = {}
    configs['DISCIPLINES_2_SLOTS_ATTENDANCE'] = data_config[data_config['Tipo'] == 'Disciplina com 2 Atendimentos']['Dado'].apply(cleaner).to_list()
    configs['DISCIPLINES_4_SLOTS_CLASS'] = data_config[data_config['Tipo'] == 'Disciplina de 110 Horas']['Dado'].apply(cleaner).to_list()
    configs['DISCIPLINES_GRUPED_CLASS'] = data_config[data_config['Tipo'] == 'Disciplina com Turmas Unidas']['Dado'].apply(cleaner).to_list()
    configs['ASSISTANT_PROFESSORS'] = data_config[data_config['Tipo'] == 'Professor Auxiliar / Assistente de Ensino']['Dado'].apply(cleaner).to_list()
    configs['NINJA_MONITORIES'] = data_config[data_config['Tipo'] == 'Monitoria Ninja']['Dado'].apply(cleaner).to_list()
    configs['SPECIAL_CLASSES'] = data_config[data_config['Tipo'] == 'Aula Especial']['Dado'].apply(cleaner).to_list()
    configs['EXCLUSIVE_TIMETABLE'] = data_config[data_config['Tipo'] == 'Grade Separada']['Dado'].apply(cleaner).to_list()
    configs['DEVELOPER_LIFE_NAME'] = data_config[data_config['Tipo'] == 'Developer Life']['Dado'].apply(cleaner).unique()[0]
    configs['NEW_TIMETABLE'] = data_config[data_config['Tipo'] == 'Nova Grade']['Dado'].apply(cleaner).unique()[0] == 'SIM'
    return configs

def remove_by_filters(data:pd.DataFrame) -> pd.DataFrame:
    data_filters_raw = pd.read_excel(config_filepath, 'Remoção Horários Space', header=1)
    columns_to_check = ['Curso', 'Série', 'Turma', 'Código da Disciplina']
    for i, row in data_filters_raw.iterrows():
        if row[columns_to_check].isnull().any():
            empty_columns = row[columns_to_check].index[row[columns_to_check].isnull()].tolist()
            warning(f'Linha [{i+3}] da aba [Remoção Horários Space] será ignorada devido as colunas {empty_columns} estarem vazias')
    data_filters_raw = data_filters_raw.dropna(subset=columns_to_check)
    data_filters_raw.rename(columns={
        'Curso': 'curso',
        'Série': 'serie',
        'Turma': 'turma',
        'Código da Disciplina': 'cod_disciplina',
        'Tipo Atividade': 'tipo_atividade',
        'Dia da Semana': 'dia_semana',
    }, inplace=True)
    print(data_filters_raw)
    data_filters = data_filters_raw.copy()
    data_filters['cod_turma'] = data_filters.apply(lambda row: f"{row['curso']}_{row['serie']}{row['turma']}", axis=1)
    
    data_filters['filtro_base'] = data_filters.apply(lambda row: f"{row['cod_turma']}_{row['cod_disciplina']}", axis=1)
    data_filters.loc[(~data_filters['tipo_atividade'].isnull()) & (~data_filters['dia_semana'].isnull()), 'filtro_base'] = None
    print(data_filters['filtro_base'].dropna().to_list())
    
    data_filters['filtro_tipo'] = data_filters.apply(lambda row: f"{row['cod_turma']}_{row['cod_disciplina']}_{row['tipo_atividade']}", axis=1)
    data_filters.loc[(data_filters['dia_semana'].isnull()), 'filtro_tipo'] = None
    print(data_filters['filtro_tipo'].dropna().to_list())
    
    data_filters['filtro_dia'] = data_filters.apply(lambda row: f"{row['cod_turma']}_{row['cod_disciplina']}_{row['dia_semana']}", axis=1)
    data_filters.loc[(data_filters['tipo_atividade'].isnull()), 'filtro_dia'] = None
    print(data_filters['filtro_dia'].dropna().to_list())
    
    data_filters['filtro_completo'] = data_filters.apply(lambda row: f"{row['cod_turma']}_{row['cod_disciplina']}_{row['tipo_atividade']}_{row['dia_semana']}", axis=1)
    data_filters.loc[(data_filters['tipo_atividade'].isnull()) & (data_filters['dia_semana'].isnull()), 'filtro_completo'] = None
    print(data_filters['filtro_completo'].dropna().to_list())

    
    data_filtering = data.copy()
    data_filtering['filtro_base'] = data_filtering.apply(lambda row: f"{row['cod_turma']}_{row['cod_disciplina']}", axis=1)
    data_filtering['filtro_tipo'] = data_filtering.apply(lambda row: f"{row['cod_turma']}_{row['cod_disciplina']}_{row['tipo_atividade']}", axis=1)
    data_filtering['filtro_dia'] = data_filtering.apply(lambda row: f"{row['cod_turma']}_{row['cod_disciplina']}_{row['dia_semana']}", axis=1)
    data_filtering['filtro_completo'] = data_filtering.apply(lambda row: f"{row['cod_turma']}_{row['cod_disciplina']}_{row['tipo_atividade']}_{row['dia_semana']}", axis=1)
    
    data_filtering = data_filtering[
        (~data_filtering['filtro_base'].isin(data_filters['filtro_base']))
        & (~data_filtering['filtro_tipo'].isin(data_filters['filtro_tipo']))
        & (~data_filtering['filtro_dia'].isin(data_filters['filtro_dia']))
        & (~data_filtering['filtro_completo'].isin(data_filters['filtro_completo']))
    ]

    
    data_filtered = data_filtering.drop(['filtro_base', 'filtro_tipo', 'filtro_dia', 'filtro_completo'], axis=1)
    return data_filtered, data_filtering
    
    
    