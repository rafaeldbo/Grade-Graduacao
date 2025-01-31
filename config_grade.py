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
    data_manual_raw = pd.read_excel(config_filepath, 'Horários Manuais', header=1)
    data_manual_raw.rename(columns={
        'Curso': 'curso',
        'Série': 'serie',
        'Turma Pref': 'turma',
        'Disciplina': 'nome_disciplina',
        'Tipo': 'tipo_atividade',
        'Dia da Semana': 'dia_semana',
        'Hora início': 'hora_inicio',
        'Hora fim': 'hora_fim',
        'Docente': 'docentes',
    }, inplace=True)
    columns_to_check = ['curso', 'serie', 'turma', 'nome_disciplina', 'tipo_atividade', 'dia_semana', 'hora_inicio', 'hora_fim']
    for i, row in data_manual_raw.iterrows():
        if row[columns_to_check].isnull().any():
            empty_columns = row[columns_to_check].index[row[columns_to_check].isnull()].tolist()
            warning(f'Linha [ {i+3} ] será removida devido as colunas {empty_columns} estarem vazias')
    data_manual = data_manual_raw.dropna(subset=columns_to_check)

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