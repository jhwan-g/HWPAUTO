HwpAuto (ver3.1)
=======

한글 파일 작업 자동화를 위한 프로그램

Pyinstaller compile 명령어 :
pyinstaller -F -w main.py --additional-hooks-dir=.

기본 데이터 저장 위치 :
BASE_DATA_PATH ='C:\\Hwpauto\data'

기본 규칙 저장 위치 :
RULE_FILE_NAME = '\\rule.json'

rule.json 파일 형식 :
{
    "eqn_edit":{ # 수식에서 문자열 변경
        "rule1"{
            "init_str":"string",
            "fin_str":"string",
            "apply":true
        }
    }
}
