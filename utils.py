import os
import re

def clear_console():
    # 플랫폼에 따라 다른 명령어 실행
    os.system('cls' if os.name == 'nt' else 'clear')
    
def clean_value(value):
    # 제거할 문자 정의
    forbidden_chars = r'[\\/\?\*\[\]]'  # \ / ? * [ ]는 정규식으로 표현
    # 정규식으로 forbidden_chars에 해당하는 문자 제거
    cleaned_value = re.sub(forbidden_chars, '', value)
    return cleaned_value