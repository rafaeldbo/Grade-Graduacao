import warnings
from dotenv import load_dotenv
from os import getenv, path


from code.grade import construct_calendar

load_dotenv(override=True)
ABS_PATH = path.abspath(path.dirname(__file__))
warnings.filterwarnings('ignore')

if __name__ == '__main__':
    construct_calendar(ABS_PATH)