import pandas as pd
from datetime import datetime
import unicodedata

from settings import get_digit, error, success, warning, info

import warnings
warnings.filterwarnings('ignore')

# CURRENT_DATE = datetime.now()
CURRENT_DATE = '2024-08-05'

cleaner = lambda x: x.strip().upper() if isinstance(x, str) else x

# ================================================== Dados de entrada ==================================================

data_config = pd.read_excel('Config_Grade.xlsx', 'Dados Configuráveis', header=1)

DISCIPLINES_2_SLOTS_ATTENDANCE = data_config[data_config['Tipo'] == 'Disciplina com 2 Atendimentos']['Dado'].apply(cleaner).to_list()
DISCIPLINES_4_SLOTS_CLASS = data_config[data_config['Tipo'] == 'Disciplina de 110 Horas']['Dado'].apply(cleaner).to_list()
DISCIPLINES_GRUPED_CLASS = data_config[data_config['Tipo'] == 'Disciplina com Turmas Unidas']['Dado'].apply(cleaner).to_list()
ASSISTANT_PROFESSORS = data_config[data_config['Tipo'] == 'Professor Auxiliar / Assistente de Ensino']['Dado'].apply(cleaner).to_list()
NINJA_MONITORIES = data_config[data_config['Tipo'] == 'Monitoria Ninja']['Dado'].apply(cleaner).to_list()
SPECIAL_CLASSES = data_config[data_config['Tipo'] == 'Aula Especial']['Dado'].apply(cleaner).to_list()
EXCLUSIVE_TIMETABLE = data_config[data_config['Tipo'] == 'Grade Separada']['Dado'].apply(cleaner).to_list()
DEVELOPER_LIFE_NAME = data_config[data_config['Tipo'] == 'Developer Life']['Dado'].apply(cleaner).unique()[0]
NEW_TIMETABLE = data_config[data_config['Tipo'] == 'Nova Grade']['Dado'].apply(cleaner).unique()[0] == 'SIM'

COURSE_PARSER = {
    'GRENGCOMP': ['ENG', 'ENG', 'COMP', 'COMP', 'COMP', 'COMP', 'COMP', 'COMP', 'COMP', 'COMP'],
    'GRENGMECA': ['ENG', 'ENG', 'MEC/MECAT', 'MEC/MECAT', 'MEC', 'MEC', 'MEC', 'MEC', 'MEC', 'MEC'],
    'GRENGMECAT': ['ENG', 'ENG', 'MEC/MECAT', 'MEC/MECAT', 'MECAT', 'MECAT', 'MECAT', 'MECAT', 'MECAT', 'MECAT'],
    'GRADM': ['ADM/ECO', 'ADM/ECO', 'ADM/ECO', 'ADM', 'ADM', 'ADM', 'ADM', 'ADM'],
    'GRECO': ['ADM/ECO', 'ADM/ECO', 'ADM/ECO','ECO', 'ECO', 'ECO', 'ECO', 'ECO'],
    'GRDIR': ['DIR']*10,
    'GRCIECOMP': ['CCOMP']*8,
}
TYPE_PRIORITY = {
    'AULA': 3,
    'ATIVIDADE EXTRA CURRICULAR': 3,
    'ATENDIMENTO / PLANTÃO': 2,
    'MONITORIA': 1,
    'MONITORIA NINJA': 0,
    'DIA RESERVADO': -1,
}

# Definição de horários fixos para aulas
FIXED_START_HOURS = ['07:30', '09:45', '12:00', '14:15', '16:30', '19:00'] if NEW_TIMETABLE else  ['07:30', '09:45', '12:00', '13:30', '15:45', '18:00'] 
FIXED_END_HOURS = ['09:30', '11:45', '14:00', '16:15', '18:30', '21:00'] if NEW_TIMETABLE else ['09:30', '11:45', '13:15', '15:30', '17:45' ,'20:00']

info('Carregando dados adicionados manualmente')
data_manual_raw = pd.read_excel('Config_Grade.xlsx', 'Horários Manuais', header=1)
columns_to_check = ['Curso', 'Série', 'Turma Pref', 'Disciplina', 'Tipo', 'Dia da Semana', 'Hora início', 'Hora fim']
for i, row in data_manual_raw.iterrows():
    if row[columns_to_check].isnull().any():
        empty_columns = row[columns_to_check].index[row[columns_to_check].isnull()].tolist()
        warning(f'Linha [ {i+3} ] será removida devido as colunas {empty_columns} estarem vazias')
data_manual = data_manual_raw.dropna(subset=columns_to_check)

info('Carregando dados do Relatório')
data_raw = pd.read_excel('ReservasAcademicas.xlsx', header=1)
success('Dados carregados!')
info('Processando dados')

# ================================================== Limpeza dos Dados ==================================================

data_course_trated = data_raw.copy()
data_course_trated['Curso'] = data_course_trated.apply(lambda row: COURSE_PARSER.get(row['Curso'], ['']*10)[row['Série']-1], axis=1)

data_no_duplicates = data_course_trated.drop_duplicates([
    'Data aula', 'Dia da Semana', 'Hora início\n(com tempo pré)',
    'Hora fim\n(com tempo pós)', 'Prédio', 'Andar',
    'Turma Pref', 'Curso', 'Disciplina'
])


data = data_no_duplicates.drop(columns=[
    'ID', 'Família', 'Previsão de Alunos', 'Prédio', 'Andar', 'Capacidade',
    'Pré Matriculado', 'Matriculados', 'Total\n(Pré + Matr)', 'Qtde Disponíveis de Assentos',
    'Status', 'Aula Externa?'
])
data['Data aula'] = pd.to_datetime(data['Data aula'], format='%d/%m/%Y')
data = data.rename(columns={'Hora início\n(com tempo pré)': 'Hora início', 'Hora fim\n(com tempo pós)': 'Hora fim'})

data_cleaned = data[
    (data['Data aula'] >= CURRENT_DATE)
    & (data['Descrição do Lyceum'].isna())
    & (~data['Disciplina'].str.startswith('AC -'))
    & (~data['Turma Pref'].str.contains('DPFERIAS'))
    & (~data['Turma Pref'].str.contains('_OPTATIVA'))
    & (~data['Turma Pref'].str.contains('_OPT'))	
    & (~data['Turma'].str.contains('_ELET'))
    & (~data['Turma'].str.contains('ELET_'))
    & (~data['Turma'].str.contains('_ANPEC'))
    & (data['Curso'] != 'GLOACA')
].copy()



data_cleaned = data_cleaned.map(cleaner)
data_cleaned = data_cleaned[~data_cleaned['Disciplina'].isin(['PROJETO FINAL - CAPSTONE'])]
data_cleaned = data_cleaned[
    ((data_cleaned['Disciplina'] == DEVELOPER_LIFE_NAME) 
     & (data_cleaned['Tipo'] == 'ATENDIMENTO / PLANTÃO')) 
    | (data_cleaned['Disciplina'] != DEVELOPER_LIFE_NAME)
]

data_cleaned['Docente'] = data_cleaned['Docente'].map(
    lambda x: ''.join(
        [c for c in unicodedata.normalize('NFKD', x) if not unicodedata.combining(c)]
    ) if pd.notna(x) else x
)

# ================================================== Processamento dos dados ==================================================

counts = data_cleaned.groupby([
    'Curso', 'Série', 'Disciplina', 'Tipo', 'Turma Pref', 'Hora início', 'Hora fim', 'Docente',
])['Dia da Semana'].value_counts().reset_index(name='counts')

rooms = data_cleaned.groupby([
    'Curso', 'Série', 'Disciplina', 'Tipo', 'Turma Pref', 'Hora início', 'Hora fim', 'Docente',
])['Sala'].apply(lambda x: ', '.join(set(x.dropna()))).reset_index()

classes = data_cleaned.drop_duplicates([
    'Curso', 'Série', 'Disciplina', 'Tipo', 'Turma Pref', 'Hora início', 'Hora fim'
]).groupby([
    'Curso', 'Série', 'Disciplina', 'Tipo', 'Turma Pref'
]).size().reset_index(name='Ocorrências')

data_counted = pd.merge(rooms, counts, 'left', [
    'Curso', 'Série', 'Disciplina', 'Tipo', 'Turma Pref', 'Hora início', 'Hora fim', 'Docente',
]).drop_duplicates()
data_counted = data_counted.merge(classes, 'left', ['Curso', 'Série', 'Disciplina', 'Tipo', 'Turma Pref'])

def Treat_docente(row: pd.Series) -> pd.Series:
    if row['Tipo'] not in ['MONITORIA', 'MONITORIA NINJA']:
        docente = row['Docente'].split(' / ')
        true_docente = [professor.strip() for professor in docente if professor not in ASSISTANT_PROFESSORS]
        row['Titular'] = ' / '.join(true_docente)
    return row

data_docente_treated = data_counted.apply(Treat_docente, axis=1)

separated_docente = data_docente_treated.copy()
separated_docente['Titular'] = separated_docente['Docente'].str.split(' / ')
separated_docente = separated_docente.explode('Titular').reset_index(drop=True)
separated_docente = pd.concat([data_docente_treated, separated_docente]).drop_duplicates().reset_index(drop=True)

teacher_frequency = separated_docente.groupby(['Curso', 'Turma Pref', 'Disciplina', 'Titular']).size().reset_index(name='Frequency')

def select_titular(group):
    max_frequency = group['Frequency'].max()
    candidates = group[group['Frequency'] == max_frequency]
    if len(candidates) > 1:
        candidates_with_slash = candidates[candidates['Titular'].str.contains('/')]
        if not candidates_with_slash.empty:
            return candidates_with_slash.iloc[0]['Titular']
    return candidates.iloc[0]['Titular']

titular = teacher_frequency.groupby(['Curso', 'Turma Pref', 'Disciplina']).apply(select_titular).reset_index(name='Titular')

data_docente_treated['Titular'] = data_docente_treated.set_index(['Curso', 'Turma Pref', 'Disciplina']).index.map(titular.set_index(['Curso', 'Turma Pref', 'Disciplina'])['Titular'])

def update_tipo(row: pd.Series) -> pd.Series:
    if row['Tipo'] == 'ATIVIDADE EXTRA CURRICULAR' and (row['Titular'] not in row['Docente']):
        row['Tipo'] = 'MONITORIA NINJA'
    return row

data_docente_treated = data_docente_treated.apply(update_tipo, axis=1)

def gambirra_special_titular(row: pd.Series) -> pd.Series:
    if row['Disciplina'] in EXCLUSIVE_TIMETABLE:
        start_date = data_raw[data_raw['Disciplina'] == row['Disciplina']]['Data aula'].min()
        end_date = data_raw[data_raw['Disciplina'] == row['Disciplina']]['Data aula'].max()
        row['Titular'] = f"{start_date.strftime('%d/%m')} a {end_date.strftime('%d/%m')}"
    return row

# Aplicando a função a cada linha do DataFrame
data_docente_treated = data_docente_treated.apply(gambirra_special_titular, axis=1)

def treat_turma(row: pd.Series) -> pd.DataFrame:
    turma = ''.join(filter(str.isalpha, row['Turma Pref'].split('_')[-1]))
    if ('DP' not in turma) and len(turma) > 1:
        if row['Disciplina'] in DISCIPLINES_GRUPED_CLASS:
            new_rows = []
            for turma_part in turma:
                new_row = row.copy()
                new_row['Turma Pref'] = turma_part
                new_rows.append(new_row)
            return pd.DataFrame(new_rows)
        turma = turma[-1]
    if row['Disciplina'] in EXCLUSIVE_TIMETABLE:
        turma = 'Z@'+turma
    row['Turma Pref'] = turma
    return pd.DataFrame([row])
    
data_turma_treated = pd.concat(data_docente_treated.apply(treat_turma, axis=1).tolist(), ignore_index=True)\
                .drop_duplicates()\
                .sort_values(['Curso','Disciplina', 'Turma Pref', 'Tipo', 'Dia da Semana'])

def treat_serie(row):
    if row['Curso'] == 'ADM/ECO' and ('DP' in row['Turma Pref']):
        return str(row['Série']) + 'DP'
    elif row['Curso'] in ['ADM', 'ECO']:
        return str(row['Série']) + row['Curso']
    return row['Série']

data_serie_treated = data_turma_treated.copy()
data_serie_treated['Série'] = data_serie_treated.apply(treat_serie, axis=1)
            
# ================================================== Filtragem dos Dados ==================================================

max_counts = data_serie_treated.groupby(['Curso', 'Série', 'Turma Pref'])['counts'].max().reset_index(name='max')
data_slots_options = data_serie_treated.merge(max_counts, on=['Curso', 'Série', 'Turma Pref'], how='left')
data_slots_options['Mínimo de Aulas'] = data_slots_options['max'].fillna(0).apply(lambda x: min([max([x//4, 4]), x]))
data_slots_options.loc[data_slots_options['Disciplina'].isin(EXCLUSIVE_TIMETABLE), 'Mínimo de Aulas'] = data_slots_options.loc[data_slots_options['Disciplina'].isin(EXCLUSIVE_TIMETABLE), 'max']//2

one_slots = data_slots_options[
    (data_slots_options['Tipo'].isin(['ATENDIMENTO / PLANTÃO', 'MONITORIA NINJA', 'MONITORIA'])) 
    & (~data_slots_options['Disciplina'].isin(DISCIPLINES_2_SLOTS_ATTENDANCE))
    & ((data_slots_options['Ocorrências'] == 1) 
        | (data_slots_options['counts'] >= data_slots_options['Mínimo de Aulas']) 
        | (data_slots_options['Sala'] == 'AULA REMOTA'))
    & (~data_slots_options['Disciplina'].isin(EXCLUSIVE_TIMETABLE))
    & (data_slots_options['Disciplina'] != DEVELOPER_LIFE_NAME)
]
one_slots = one_slots\
                .groupby(['Curso', 'Série', 'Disciplina', 'Tipo', 'Turma Pref'], group_keys=False)\
                .apply(lambda group: group.nlargest(1, columns='counts'))
                
two_slots = data_slots_options[
    ((((~data_slots_options['Tipo'].isin(['ATENDIMENTO / PLANTÃO', 'MONITORIA NINJA', 'MONITORIA'])) 
        | (data_slots_options['Disciplina'].isin(DISCIPLINES_2_SLOTS_ATTENDANCE)))
            & (~data_slots_options['Disciplina'].isin(DISCIPLINES_4_SLOTS_CLASS)))
        | (data_slots_options['Tipo'] == 'ATIVIDADE EXTRA CURRICULAR'))
    & ((data_slots_options['Ocorrências'] == 2) 
        | (data_slots_options['counts'] >= data_slots_options['Mínimo de Aulas']) 
        | (data_slots_options['Sala'] == 'AULA REMOTA'))
    & (~data_slots_options['Disciplina'].isin(EXCLUSIVE_TIMETABLE))
    & (data_slots_options['Disciplina'] != DEVELOPER_LIFE_NAME)
]                
two_slots = two_slots\
                .groupby(['Curso', 'Série', 'Disciplina', 'Tipo', 'Turma Pref'], group_keys=False)\
                .apply(lambda group: group.nlargest(2, columns='counts'))
                
four_slots = data_slots_options[
    (data_slots_options['Tipo'] == 'AULA') 
    & (data_slots_options['Disciplina'].isin(DISCIPLINES_4_SLOTS_CLASS))
    & ((data_slots_options['Ocorrências'] == 4) 
        | (data_slots_options['counts'] >= data_slots_options['Mínimo de Aulas']) 
        | (data_slots_options['Sala'] == 'AULA REMOTA'))
    & (~data_slots_options['Disciplina'].isin(EXCLUSIVE_TIMETABLE))
]
four_slots = four_slots\
                .groupby(['Curso', 'Série', 'Disciplina', 'Tipo', 'Turma Pref'], group_keys=False)\
                .apply(lambda group: group.nlargest(4, columns='counts'))
                
special_slots = data_slots_options[
    (data_slots_options['Disciplina'].isin(EXCLUSIVE_TIMETABLE))
    & (data_slots_options['counts'] >= data_slots_options['Mínimo de Aulas'])
]

developer_life_slots = data_slots_options[
    (data_slots_options['Disciplina'] == DEVELOPER_LIFE_NAME)
    & (data_slots_options['Tipo'] == 'ATENDIMENTO / PLANTÃO')
    & (data_slots_options['counts'] >= data_slots_options['Mínimo de Aulas'])
]

data_slots = pd.concat([one_slots, two_slots, four_slots, special_slots, developer_life_slots])

                
# ================================================== Pós Processamento dos Dados ==================================================

def time_difference(start_time, end_time):
    start_time = datetime.strptime(start_time, '%H:%M')
    end_time = datetime.strptime(end_time, '%H:%M')
    return abs(start_time - end_time).total_seconds() / 3600

def closest_time(target_time:str, time_list:list[str]) -> str:
    time_differences = [(t, time_difference(target_time, t)) for t in time_list]
    closest = min(time_differences, key=lambda x: x[1])[0]
    return closest

def Treat_slots(row: pd.Series) -> pd.Series:
    if (row['Tipo'] in ['ATIVIDADE EXTRA CURRICULAR', 'AULA']) and (row['Disciplina'] not in DISCIPLINES_4_SLOTS_CLASS) and (row['Disciplina'] not in SPECIAL_CLASSES):
        time = time_difference(row['Hora início'], row['Hora fim'])
        if  time != 2:
            if row['Hora início'] not in FIXED_START_HOURS:
                row['Hora início'] = closest_time(row['Hora início'], FIXED_START_HOURS)
            if row['Hora fim'] not in FIXED_END_HOURS:
                row['Hora fim'] = closest_time(row['Hora fim'], FIXED_END_HOURS)
            if time < 1.75: 
                warning(f'A aula da disciplina {row["Disciplina"]} da turma {row["Curso"]}-{get_digit(row['Série'])}{row["Turma Pref"]} foi agendada na {row["Dia da Semana"]} com menos de 2 horas de duração')
                if time_difference(row['Hora início'], row['Hora fim']) < 1.75:
                    error(f'A aula da disciplina {row["Disciplina"]} da turma {row["Curso"]}-{get_digit(row['Série'])}{row["Turma Pref"]} foi removida por ter menos de 2h de duração mesmo após correção')
                    return pd.Series()  
    return row

data_slots_treated = data_slots.apply(Treat_slots, axis=1)
data_slots_treated['Origem'] = 'Reserva Acadêmica'
data_slots_treated.loc[data_slots_treated['Disciplina'] == DEVELOPER_LIFE_NAME, 'Disciplina'] = 'DEVELOPER LIFE'

data_manual = data_manual.map(lambda x: x.strip().upper() if isinstance(x, str) else x)
data_manual = data_manual.drop(['Observação'], axis=1)
data_manual['Hora início'] = pd.to_datetime(data_manual['Hora início'], format='%H:%M:%S').dt.strftime('%H:%M')
data_manual['Hora fim'] = pd.to_datetime(data_manual['Hora fim'], format='%H:%M:%S').dt.strftime('%H:%M')
data_manual['Curso'] = data_manual.apply(lambda row: COURSE_PARSER.get(row['Curso'], ['']*10)[row['Série']-1], axis=1)
data_manual['Docente'] = data_manual['Docente'].fillna('')
data_manual['Titular'] = data_manual['Docente']
data_manual['Origem'] = 'Manual'
data_manual['counts'] = 0
data_manual['Série'] = data_manual.apply(treat_serie, axis=1)
data_manual['Sala'] = ''

data_full = pd.concat([data_slots_treated, data_manual], ignore_index=True)

def position_class(data: pd.DataFrame) -> pd.DataFrame:
    data = data.sort_values(by='Hora início').reset_index(drop=True)
    data['Posição'] = 0
    for i in range(len(data) - 1):
        for j in range(i+1, len(data)):
            if ((data.loc[i, 'Hora fim'] > data.loc[j, 'Hora início'] or data.loc[j, 'Hora início'] < data.loc[i, 'Hora fim'])):
                if TYPE_PRIORITY[data.loc[i, 'Tipo']] != TYPE_PRIORITY[data.loc[j, 'Tipo']]:
                    if TYPE_PRIORITY[data.loc[i, 'Tipo']] > TYPE_PRIORITY[data.loc[j, 'Tipo']]:
                        data.loc[i, 'Posição'] = 0
                        data.loc[j, 'Posição'] = -1
                    else:
                        data.loc[i, 'Posição'] = -1
                        data.loc[j, 'Posição'] = 0
                else:
                    data.loc[i, 'Posição'] = 1
                    data.loc[j, 'Posição'] = 2
                
                    if (data.loc[i, 'Tipo'] not in ['AULA', 'ATIVIDADE EXTRA CURRICULAR']
                    and data.loc[j, 'Tipo'] in ['AULA', 'ATIVIDADE EXTRA CURRICULAR']):
                        data.loc[i, 'Posição'] = -1
                        data.loc[j, 'Posição'] = 0
                        
                    if (data.loc[i, 'Tipo'] in ['AULA', 'ATIVIDADE EXTRA CURRICULAR']
                    and data.loc[j, 'Tipo'] not in ['AULA', 'ATIVIDADE EXTRA CURRICULAR']):
                        data.loc[i, 'Posição'] = 0
                        data.loc[j, 'Posição'] = -1
    return data

data_slots_positioned = data_full.groupby(['Curso', 'Série', 'Turma Pref', 'Dia da Semana']).apply(position_class).reset_index(drop=True)
data_slots_positioned['Posição'] = data_slots_positioned.apply(
    lambda row: -2 if (row['Tipo'] == 'MONITORIA NINJA' and row['Disciplina'] not in NINJA_MONITORIES) else row['Posição'], 
    axis=1
)

DATA = data_slots_positioned.copy()
success('Processamento conluido!\n')
