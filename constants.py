'''프로그램에서 이용되는 상수들이 있는 모듈
'''
import os

PROGRAM_PATH = os.getcwd()

BASE_DATA_PATH ='C:\\Hwpauto\data'
RULE_FILE_NAME = '\\rule.json'

EMPTY_RULE_STR = '''{
    "eqn_edit":{
    }
}'''

EMPTY_RULE_JSON = {
    "eqn_edit":{}
}