import pandas as pd
import unicodedata

import urllib
from sqlalchemy import create_engine, Engine
from dotenv import load_dotenv
import warnings
from os import getenv, path

load_dotenv(override=True)
ABS_PATH = path.abspath(path.dirname(__file__))
warnings.filterwarnings('ignore')

from config_grade import load_manual_data
from utils import info, success, cleaner 
from settings import CONFIGS, TYPE_PRIORITY, CURRENT_DATE, CURRENT_YEAR, CURRENT_SEMESTER

# Criando a string de conexão manualmente
params = urllib.parse.quote_plus(
    f"DRIVER={getenv('driver')};"
    f"SERVER={getenv('server')};"
    f"DATABASE={getenv('database')};"
    f"UID={getenv('username')};"
    f"PWD={getenv('password')};"
)
if 'None' in params:
    raise ValueError('As Variáveis de Ambiente não foram configuradas corretamente!')

connection_url = f"mssql+pyodbc:///?odbc_connect={params}"

# Cria a engine do SQLAlchemy
engine:Engine = create_engine(connection_url)

query = f"""--sql
WITH temp_tratamento_1 AS (
    SELECT 
        *,
        REPLACE(TRANSLATE(RIGHT(turma_pref, LEN(turma_pref) - CHARINDEX('_', turma_pref)), '0123456789', '          '), ' ', '') AS turma_tratada, -- tratamento inicial da turma
        CONCAT(CAST(ano AS NVARCHAR), '.', CAST(semestre-60 AS NVARCHAR)) AS periodo, -- coluna de periodo para facilitar identificação
        CASE -- Correção do tempo pós-aula colocado automaticamente pela configuração antiga do Space
            WHEN Descricao != 'Aula' THEN HORA_INICIO
            WHEN HORA_INICIO = '15:30' AND HORA_FIM = '17:45' THEN '15:45'
            WHEN HORA_INICIO = '13:15' AND HORA_FIM = '15:30' THEN '13:30'
            WHEN HORA_INICIO = '09:30' AND HORA_FIM = '11:45' THEN '09:45'
            WHEN HORA_INICIO = '07:15' AND HORA_FIM = '09:30' THEN '07:30'
            ELSE HORA_INICIO 
        END AS hora_inicio_corrigida,
        CASE -- Correção do tempo pré-aula colocado automaticamente pela configuração antiga do Space
            WHEN Descricao != 'Aula' THEN HORA_FIM
            WHEN HORA_INICIO = '15:45' AND HORA_FIM = '18:00' THEN '17:45'
            WHEN HORA_INICIO = '13:30' AND HORA_FIM = '15:45' THEN '15:30'
            WHEN HORA_INICIO = '09:45' AND HORA_FIM = '12:00' THEN '11:45'
            WHEN HORA_INICIO = '07:30' AND HORA_FIM = '09:45' THEN '09:30'
            ELSE HORA_FIM 
        END AS hora_fim_corrigida,
        CASE
            WHEN curso IN ('GRADM', 'GRECO') AND serie <= 3 THEN 'ADM/ECO' -- ciclo em conjunto de ADM e ECO
            WHEN curso IN ('GRENGCOMP', 'GRENGMECA', 'GRENGMECAT') AND serie <= 2 THEN 'ENG' -- ciclo em conjunto de ENG
            WHEN curso IN ('GRENGMECA', 'GRENGMECAT') AND serie <= 4 THEN 'MECA/MECAT' -- ciclo em conjunto de MECA e MECAT
            ELSE REPLACE(REPLACE(curso, 'GRENG', ''), 'GR', '') --removendo GR e GRENG para facilitar a leitura
        END AS curso_tratado,
        CASE 
            WHEN turma LIKE '%ELET[_]%' OR turma LIKE '%[_]OPT%' OR turma LIKE '%ELETFERIAS%' THEN 'ELETIVA' -- identificando disciplinas eletivas
            WHEN turma LIKE 'AC[_]%' THEN 'ATIVIDADE COMPLEMENTAR' -- identificando atividades complementares
            WHEN turma LIKE '%[_]ANPEC%' THEN 'ANPEC' -- identificando disciplinas da ANPEC
            WHEN turma LIKE '%GLOACA%' THEN 'GLOBAL ACADEMY' -- identificando disciplinas da GLOBAL ACADEMY
            WHEN turma LIKE '%DPFERIAS%' OR turma LIKE '%DPFÉRIAS%' THEN 'DPFERIAS' -- identificando DPs de férias
            ELSE 'OBRIGATORIA' -- definindo demais disciplinas como obrigatórias
        END AS tipo_disciplina
    FROM tb_dta_reservas_academicas 
    WHERE 
        familia_curso = 'GRADUACAO' -- filtrando apenas pela graduação (esses tratamentos são específicos para a graduação)
        AND ano = {CURRENT_YEAR} -- filtrando apenas pelo ano atual para acelerar a consulta
),
temp_tratamento_2 AS (
    SELECT
        *,
        CASE
            WHEN turma_tratada IN ('OPT[_]A', 'OPTATIVA', 'ELET[_]A') THEN 'ELET_A' -- turma eleita A
            WHEN (turma_tratada = 'A' AND tipo_disciplina = 'ELETIVA') THEN 'ELET_A' -- turma eleita A
            WHEN (turma_tratada = 'B' AND tipo_disciplina = 'ELETIVA') THEN 'ELET_B' -- turma eleita B
            WHEN turma LIKE 'EXAME[_]QUALI%' THEN 'A' -- padronizando turma do exame de qualificação de ADM e ECO
            WHEN turma LIKE 'AC[_]%' THEN 'AC_A' -- distinção da turma das atividades complementares
            WHEN turma_tratada LIKE 'ECO%' THEN RIGHT(turma_tratada, 1) -- padronizando turma de ECO
            WHEN turma_tratada = 'DPFÉRIAS' THEN 'DPFERIAS' -- padronizando turma de DP de férias
            ELSE turma_tratada
        END AS turma_real
    FROM temp_tratamento_1
),
temp_duplicatas AS (
    SELECT 
        *, 
        ROW_NUMBER() OVER (PARTITION BY data_aula, HORA_INICIO, HORA_FIM, curso_tratado, disciplina, serie, turma_real, sala, Docente ORDER BY data_aula) AS n_duplicacao
        -- identificando duplicatas de aulas (aula de ciclo em conjunto de ADM/ECO, ENG, MECA/MECAT geram duplicatas, uam para cada curso)
    FROM temp_tratamento_2
),
tb_space AS (
    SELECT 
        -- selecionando apenas as colunas necessárias para a view e renomeando-as para facilitar identificação
        data_aula, periodo,         
        DIA_SEMANA AS dia_semana,
        data_aula + CAST(PARSE(hora_inicio_corrigida AS TIME) AS DATETIME) AS hora_inicio, -- convertendo a hora de início para datetime para facilitar a comparações    
        data_aula + CAST(PARSE(hora_fim_corrigida AS TIME) AS DATETIME) AS hora_fim, -- convertendo a hora de término para datetime para facilitar a comparações    
        Descricao AS tipo_atividade, 
        disciplina AS cod_disciplina,
        TRIM(nome_disciplina) AS nome_disciplina,
        tipo_disciplina,
        curso_tratado AS curso,
        turma AS cod_turma_disciplina,
        CONCAT(curso_tratado, '_', serie, turma_real) AS cod_turma,
        turma_real AS turma,
        serie,
        TRIM(Docente) AS docentes,
        sala,
        descricao_lyceum AS observacao,
        DT_ATUALIZACAO AS dt_atualizacao
    FROM temp_duplicatas
    WHERE n_duplicacao = 1 -- removendo duplicatas
)
SELECT * FROM tb_space
ORDER BY data_aula, hora_inicio, hora_fim, curso, serie, turma;
"""

def load_space_data() -> pd.DataFrame:


    # ================================================== Dados de entrada ==================================================

    info('Extraindo dados do Banco de Dados')
    df_space = pd.read_sql(query, engine)
    success('Dados carregados!')
    info('Processando dados')

    # ================================================== Limpeza dos Dados ==================================================

    data_cleaned = df_space[
        (df_space['periodo'] == f'{CURRENT_YEAR}.{CURRENT_SEMESTER}')
        & (df_space['data_aula'] >= CURRENT_DATE)
        & (df_space['observacao'].isna()) # removendo aulas com observações (reposições, provas, etc)
        & (df_space['tipo_disciplina'] == 'OBRIGATORIA') # filtrando disciplinas obrigatórias	
    ].copy()
    data_cleaned.drop(['observacao', 'dt_atualizacao'], axis=1)

    data_cleaned = data_cleaned.map(cleaner)
    data_cleaned['hora_inicio'] = pd.to_datetime(data_cleaned['hora_inicio']).dt.time
    data_cleaned['hora_fim'] = pd.to_datetime(data_cleaned['hora_fim']).dt.time

    data_cleaned = data_cleaned[~data_cleaned['nome_disciplina'].isin(['PROJETO FINAL - CAPSTONE'])] # CAPSTONE será tratado manualmente
    data_cleaned = data_cleaned[ # as aulas do DEVELOPER LIFE serão tratadas manualmente devido a variação dos nomes das disciplinas filhas
        ((data_cleaned['nome_disciplina'] == CONFIGS['DEVELOPER_LIFE_NAME']) 
        & (data_cleaned['tipo_atividade'] == 'ATENDIMENTO / PLANTÃO')) 
        | (data_cleaned['nome_disciplina'] != CONFIGS['DEVELOPER_LIFE_NAME'])
    ]

    # removendo acentos do nome dos professores
    # como os nomes são colocados manualmente, pode haver diferença entre os nomes, principalmente acentuação
    data_cleaned['docentes'] = data_cleaned['docentes'].map(
        lambda x: ''.join(
            [c for c in unicodedata.normalize('NFKD', x) if not unicodedata.combining(c)]
        ) if pd.notna(x) else x
    )

    # ================================================== Processamento dos dados ==================================================
    # OBS.: chamaremos de SLOT um horario do dia da semana em que uma atividade ocorre

    contagem = data_cleaned.groupby([ # contando quantidade de SLOTs de cada disciplina
        'cod_turma', 'curso', 'serie', 'turma', 'cod_disciplina', 'nome_disciplina', 'tipo_atividade', 'hora_inicio', 'hora_fim', 'docentes',
    ])['dia_semana'].value_counts().reset_index(name='contagem')

    rooms = data_cleaned.groupby([  # agrupando salas de cada SLOT
        'cod_turma', 'curso', 'serie', 'turma', 'cod_disciplina', 'nome_disciplina', 'tipo_atividade', 'hora_inicio', 'hora_fim', 'docentes',
    ])['sala'].apply(lambda x: ', '.join(set(x.dropna()))).reset_index()

    classes = data_cleaned.drop_duplicates([ # contando quantas vezes cada disciplina ocorre na semana
        'cod_turma', 'nome_disciplina', 'tipo_atividade', 'hora_inicio', 'hora_fim'
    ]).groupby([
        'cod_turma', 'nome_disciplina', 'tipo_atividade'
    ]).size().reset_index(name='n_ocorrencias')

    data_counted = pd.merge(rooms, contagem, 'left', [ # juntando as informações
        'cod_turma', 'curso', 'serie', 'turma', 'cod_disciplina', 'nome_disciplina', 'tipo_atividade', 'hora_inicio', 'hora_fim', 'docentes',
    ]).drop_duplicates()
    data_counted = data_counted.merge(classes, 'left', ['cod_turma', 'nome_disciplina', 'tipo_atividade'])

    # removendo professores assistentes
    def Treat_docente(row: pd.Series) -> pd.Series:
        if row['tipo_atividade'] in ['AULA', 'ATIVIDADE EXTRA CURRICULAR', 'ATENDIMENTO / PLANTÃO']:
            docente = row['docentes'].split(' / ')
            true_docente = [professor.strip() for professor in docente if professor not in CONFIGS['ASSISTANT_PROFESSORS']]
            row['titular'] = ' / '.join(true_docente)
        return row

    data_docente_treated = data_counted.apply(Treat_docente, axis=1)

    # calculando a frequencia de aparecimento de cada professor em cada disciplina para determinar o professor titular (com mais aulas daquela disciplina)
    separated_docente = data_docente_treated.copy()
    separated_docente['titular'] = separated_docente['docentes'].str.split(' / ')
    separated_docente = separated_docente.explode('titular').reset_index(drop=True)
    separated_docente = pd.concat([data_docente_treated, separated_docente]).drop_duplicates().reset_index(drop=True)

    teacher_frequency = separated_docente.groupby(['cod_turma', 'nome_disciplina', 'titular']).size().reset_index(name='freq_docente')

    # determinando o professor titular (com mais aulas daquela disciplina)
    def select_titular(group):
        max_frequency = group['freq_docente'].max()
        candidates = group[group['freq_docente'] == max_frequency]
        if len(candidates) > 1:
            candidates_with_slash = candidates[candidates['titular'].str.contains('/')]
            if not candidates_with_slash.empty:
                return candidates_with_slash.iloc[0]['titular']
        return candidates.iloc[0]['titular']

    titular = teacher_frequency.groupby(['cod_turma', 'nome_disciplina']).apply(select_titular).reset_index(name='titular')
    data_docente_treated['titular'] = data_docente_treated.set_index(['cod_turma', 'nome_disciplina']).index.map(titular.set_index(['cod_turma', 'nome_disciplina'])['titular'])

    # caso a atividade extracurricular seja ministrada apenas por um professor não titular, ela será considerada uma monitoria ninja
    def update_tipo(row: pd.Series) -> pd.Series:
        if row['tipo_atividade'] == 'ATIVIDADE EXTRA CURRICULAR' and (row['titular'] not in row['docentes']):
            row['tipo_atividade'] = 'MONITORIA NINJA'
        return row

    data_docente_treated = data_docente_treated.apply(update_tipo, axis=1)

    # gambiarra para disciplinas especiais (como CAPSTONE E REP) que mostram o horário no local do docente
    def gambirra_special_titular(row: pd.Series) -> pd.Series:
        if (row['nome_disciplina'] in CONFIGS['EXCLUSIVE_TIMETABLE']):
            start_date = df_space[df_space['nome_disciplina'] == row['nome_disciplina']]['data_aula'].min()
            end_date = df_space[df_space['nome_disciplina'] == row['nome_disciplina']]['data_aula'].max()
            row['titular'] = f"{start_date.strftime('%H%M')} a {end_date.strftime('%H%M')}"
        return row

    data_docente_treated = data_docente_treated.apply(gambirra_special_titular, axis=1)

    # Separando turmas unidas (disciplinas de Tópicos)
    def treat_turma(row: pd.Series) -> pd.DataFrame:
        if (row['nome_disciplina'] in CONFIGS['DISCIPLINES_GRUPED_CLASS']) and ('DP' not in row['turma']) and (len(row['turma']) > 1):
            new_rows = []
            for turma_part in row['turma']:
                new_row = row.copy()
                new_row['turma'] = turma_part
                new_row['cod_turma'] = f"{row['curso']}_{row['serie']}{turma_part}"
                new_rows.append(new_row)
            return pd.DataFrame(new_rows)
        if (row['nome_disciplina'] in CONFIGS['EXCLUSIVE_TIMETABLE']):
            row['turma'] = 'Z@'+row['turma']
        return pd.DataFrame([row])
        
    data_turma_treated = pd.concat(data_docente_treated.apply(treat_turma, axis=1).tolist(), ignore_index=True)\
                    .drop_duplicates()\
                    .sort_values(['cod_turma', 'nome_disciplina', 'tipo_atividade', 'dia_semana'])

    # Separando turmas de DP de ADM/ECO para facilitar vizualização na grade
    def treat_serie(row):
        if row['curso'] == 'ADM/ECO' and ('DP' in row['turma']):
            return str(row['serie']) + 'DP'
        elif row['curso'] in ['ADM', 'ECO']:
            return str(row['serie']) + row['curso']
        return row['serie']

    data_serie_treated = data_turma_treated.copy()
    data_serie_treated['serie'] = data_serie_treated.apply(treat_serie, axis=1)
                
    # ================================================== Filtragem dos Dados ==================================================

    # Encontrando a quantidade total de aulas de cada turma e definindo a quantidade mínima de aulas de cada disciplina
    max_contagem = data_serie_treated.groupby(['cod_turma'])['contagem'].max().reset_index(name='max')
    data_slots_options = data_serie_treated.merge(max_contagem, on=['cod_turma'], how='left')
    data_slots_options['min_aulas'] = data_slots_options['max'].fillna(0).apply(lambda x: min([max([x//4, 4]), x]))
    data_slots_options.loc[data_slots_options['nome_disciplina'].isin(CONFIGS['EXCLUSIVE_TIMETABLE']), 'min_aulas'] = data_slots_options.loc[data_slots_options['nome_disciplina'].isin(CONFIGS['EXCLUSIVE_TIMETABLE']), 'max']//2

    # OBS.: 

    # Filtrando SLOTs de disciplinas de apenas um SLOT
        # Atendimentos e monitorias normalmente possuem apenas um SLOT
    one_slots = data_slots_options[
        (data_slots_options['tipo_atividade'].isin(['ATENDIMENTO / PLANTÃO', 'MONITORIA NINJA', 'MONITORIA'])) 
        & (~data_slots_options['nome_disciplina'].isin(CONFIGS['DISCIPLINES_2_SLOTS_ATTENDANCE'])) # existrem disciplinas especificas com 2 atendimentos
        & ((data_slots_options['n_ocorrencias'] == 1) 
            | (data_slots_options['contagem'] >= data_slots_options['min_aulas']) 
            | (data_slots_options['sala'] == 'AULA REMOTA'))
        & (~data_slots_options['nome_disciplina'].isin(CONFIGS['EXCLUSIVE_TIMETABLE']))
        & (data_slots_options['nome_disciplina'] != CONFIGS['DEVELOPER_LIFE_NAME'])
    ]
    one_slots = one_slots\
                    .groupby(['cod_turma', 'nome_disciplina', 'tipo_atividade'], group_keys=False)\
                    .apply(lambda group: group.nlargest(1, columns='contagem'))
                    
    two_slots = data_slots_options[
        ((((~data_slots_options['tipo_atividade'].isin(['ATENDIMENTO / PLANTÃO', 'MONITORIA NINJA', 'MONITORIA'])) 
            | (data_slots_options['nome_disciplina'].isin(CONFIGS['DISCIPLINES_2_SLOTS_ATTENDANCE'])))
                & (~data_slots_options['nome_disciplina'].isin(CONFIGS['DISCIPLINES_4_SLOTS_CLASS'])))
            | (data_slots_options['tipo_atividade'] == 'ATIVIDADE EXTRA CURRICULAR'))
        & ((data_slots_options['contagem'] >= data_slots_options['min_aulas']) 
            | (data_slots_options['sala'] == 'AULA REMOTA'))
        & (~data_slots_options['nome_disciplina'].isin(CONFIGS['EXCLUSIVE_TIMETABLE']))
        & (data_slots_options['nome_disciplina'] != CONFIGS['DEVELOPER_LIFE_NAME'])
    ]                
    two_slots = two_slots\
                    .groupby(['cod_turma', 'nome_disciplina', 'tipo_atividade'], group_keys=False)\
                    .apply(lambda group: group.nlargest(2, columns='contagem'))
                    
    four_slots = data_slots_options[
        (data_slots_options['tipo_atividade'] == 'AULA') 
        & (data_slots_options['nome_disciplina'].isin(CONFIGS['DISCIPLINES_4_SLOTS_CLASS']))
        & ((data_slots_options['n_ocorrencias'] == 4) 
            | (data_slots_options['contagem'] >= data_slots_options['min_aulas']) 
            | (data_slots_options['sala'] == 'AULA REMOTA'))
        & (~data_slots_options['nome_disciplina'].isin(CONFIGS['EXCLUSIVE_TIMETABLE']))
    ]
    four_slots = four_slots\
                    .groupby(['cod_turma', 'nome_disciplina', 'tipo_atividade'], group_keys=False)\
                    .apply(lambda group: group.nlargest(4, columns='contagem'))
                    
    special_slots = data_slots_options[
        (data_slots_options['nome_disciplina'].isin(CONFIGS['EXCLUSIVE_TIMETABLE']))
        & (data_slots_options['contagem'] >= data_slots_options['min_aulas'])
    ]

    developer_life_slots = data_slots_options[
        (data_slots_options['nome_disciplina'] == CONFIGS['DEVELOPER_LIFE_NAME'])
        & (data_slots_options['tipo_atividade'] == 'ATENDIMENTO / PLANTÃO')
        & (data_slots_options['contagem'] >= data_slots_options['min_aulas'])
    ]

    data_slots = pd.concat([one_slots, two_slots, four_slots, special_slots, developer_life_slots])
    data_slots.loc[data_slots['nome_disciplina'] == CONFIGS['DEVELOPER_LIFE_NAME'], 'nome_disciplina'] = 'DEVELOPER LIFE'
    data_slots['Origem'] = 'Reserva Acadêmica'

    # ================================================== Pós Processamento dos Dados ==================================================

    data_manual = load_manual_data()
    data_manual['serie'] = data_manual.apply(treat_serie, axis=1)
    data_full = pd.concat([data_slots, data_manual], ignore_index=True)

    # precisa de refatoração 
    def position_class(data: pd.DataFrame) -> pd.DataFrame:
        data = data.sort_values(by='hora_inicio').reset_index(drop=True)
        data['posicao'] = 0
        for i in range(len(data) - 1):
            for j in range(i+1, len(data)):
                if ((data.loc[i, 'hora_fim'] > data.loc[j, 'hora_inicio'] or data.loc[j, 'hora_inicio'] < data.loc[i, 'hora_fim'])):
                    if TYPE_PRIORITY[data.loc[i, 'tipo_atividade']] != TYPE_PRIORITY[data.loc[j, 'tipo_atividade']]:
                        if TYPE_PRIORITY[data.loc[i, 'tipo_atividade']] > TYPE_PRIORITY[data.loc[j, 'tipo_atividade']]:
                            data.loc[i, 'posicao'] = 0
                            data.loc[j, 'posicao'] = -1
                        else:
                            data.loc[i, 'posicao'] = -1
                            data.loc[j, 'posicao'] = 0
                    else:
                        data.loc[i, 'posicao'] = 1
                        data.loc[j, 'posicao'] = 2
                    
                        if (data.loc[i, 'tipo_atividade'] not in ['AULA', 'ATIVIDADE EXTRA CURRICULAR']
                        and data.loc[j, 'tipo_atividade'] in ['AULA', 'ATIVIDADE EXTRA CURRICULAR']):
                            data.loc[i, 'posicao'] = -1
                            data.loc[j, 'posicao'] = 0
                            
                        if (data.loc[i, 'tipo_atividade'] in ['AULA', 'ATIVIDADE EXTRA CURRICULAR']
                        and data.loc[j, 'tipo_atividade'] not in ['AULA', 'ATIVIDADE EXTRA CURRICULAR']):
                            data.loc[i, 'posicao'] = 0
                            data.loc[j, 'posicao'] = -1
        return data

    data_slots_positioned = data_full.groupby(['cod_turma', 'dia_semana']).apply(position_class).reset_index(drop=True)
    data_slots_positioned['posicao'] = data_slots_positioned.apply(
        lambda row: -2 if (row['tipo_atividade'] == 'MONITORIA NINJA' and row['nome_disciplina'] not in CONFIGS['NINJA_MONITORIES']) else row['posicao'], 
        axis=1
    )

    disciplines_colors:dict[str, int] = {}
    colors_used_by_class:dict[str, set] = {}
    def assign_color(row):
        turma = row['cod_turma']
        disciplina = row['nome_disciplina']

        if disciplina not in disciplines_colors:
            available_colors = [color for color in range(6) if color not in colors_used_by_class.get(turma, set())]
            disciplines_colors[disciplina] = available_colors[0]
        colors_used_by_class.get(turma, set()).add(disciplines_colors[disciplina])
        return disciplines_colors[disciplina]

    # Aplicar a função para atribuir cores
    data_slots_colored = data_slots_positioned.copy()
    data_slots_colored['cor'] = data_slots_colored.apply(assign_color, axis=1)

    success('Processamento conluido!\n')
    return data_slots_colored.copy()
