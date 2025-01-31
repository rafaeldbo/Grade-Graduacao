from datetime import datetime, time

TERMINAL_COLORS = {'red': '\033[91m', 'green': '\033[92m', 'dark green': '\033[38;2;0;100;0m', 'yellow': '\033[93m', 'reset': '\033[0m', 'cyan': '\033[96m'}
def colorize(text:str, color:str) -> str:
    return f'{TERMINAL_COLORS[color.lower()]}{text}{TERMINAL_COLORS["reset"]}'

def error(text:str) -> str:
    print(colorize('[ERRO] ', 'red')+text)
def success(text:str) -> str:
    print(colorize('[SUCESSO] ', 'green')+text)
def warning(text:str) -> str:
    print(colorize('[AVISO] ', 'yellow')+text)
def info(text:str) -> str:
    print(colorize('[INFO] ', 'cyan')+text)

def cleaner(text:str) -> str:
    if isinstance(text, str): 
        return text.strip().upper()
    return text

def time_to_integer(time:str|datetime, start_value:int=7) -> int:
    time_str = time.strftime('%H:%M') if not isinstance(time, str) else time    
    hours, minutes = map(int, time_str.split(':'))
    return (hours - start_value) * 4 + (minutes - 30) // 15

def get_digit(value:str|int) -> int:
    if isinstance(value, int): return value
    return int(''.join(filter(str.isdigit, value)))

def time_differece(start_time:time, end_time:time) -> int:
    return (end_time.hour - start_time.hour) + (end_time.minute - start_time.minute)/60

def float_to_time(time_float:float) -> str:
    minutes = time_float*60
    if minutes%60 == 0: 
        return f"{int(minutes//60)}h"
    elif minutes//60 == 0:
        return f"{int(minutes%60)}min"
    return f"{int(minutes//60)}h{int(minutes%60)}"