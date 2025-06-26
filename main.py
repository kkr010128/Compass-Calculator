from flask import Flask, request, render_template;
import pandas as pd
import msoffcrypto
from io import BytesIO

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload():
    file = request.files['file']
    password = '0000'

    # 암호 해제
    office_file = msoffcrypto.OfficeFile(file)
    office_file.load_key(password=password)

    decrypted = BytesIO()
    office_file.decrypt(decrypted)

    # 엑셀 읽기
    df = pd.read_excel(decrypted, engine='openpyxl')

    # 원본 엑셀 파일에서 필요 없는 컬럼 리스트
    drop_col = ['Unnamed: 0','No','년도','학기','분반','담당교원','성적삭제명','공학인증','공학요소','공학세부요소','원어강의종류','인정구분','성적인정대학명','교과목영문명']

    replace_target = ['제4영역:예술/문화','제2영역:정치/경제/사회/심리','제3영역:과학/기술/생명','제1영역:문학/역사/철학/종교']

    # 평점 평균 계산할때 사용
    grade_points = {
        'A+': 4.5,
        'A0': 4.0,
        'B+': 3.5,
        'B0': 3.0,
        'C+': 2.5,
        'C0': 2.0,
        'D+': 1.5,
        'D0': 1.0,
        'F': 0.0,
        'P': None  # 평점에 미반영
    }

    rules = {
        'Minimum Total': 130, # 종합 학점
        'Minimum Major': 63, # 전공 학점
        'Major Professional': 32, # 전문 학점
        'Major Basic': 31, # 기초 학점
        'General Education': 12, # 교양
        'Departmental Education': 6, # 계열교육
        'Self-Development': 2, # 자기계발
        'Personality Education': 4, # 인성
        'Common Education': 8, # 공통 교육
        'Minimum GPA': 2.8,
    }

    df.drop(drop_col,axis=1, inplace=True)
    df.replace(replace_target, '교양',inplace=True)
    df.replace(grade_points, inplace=True)

    df['이수구분'].unique()

    df['영역'].unique()

    # 졸업 사정 결과 저장 딕셔너리
    result = {}

    # 'Minimum Total': 130
    total = int(df['학점'].sum()) # 112
    # 전공 학점
    major = int(df[df['이수구분'] == '전공']['학점'].sum()) #65
    # 전문 학점
    major_professional = int(df[df['영역'] == '전문']['학점'].sum()) #28
    # 기초 학점
    major_basic = int(df[df['영역'] == '기초']['학점'].sum()) #37
    # 공통 교육
    general_education = int(df[df['영역'] == '교양']['학점'].sum()) #8
    # 계열 교육
    departmental_education = int(df[df['영역'] == '계열교육']['학점'].sum()) #12
    # 자기 계발
    self_development = int(df[df['영역'] == '자기계발']['학점'].sum()) #2
    # 인성
    personality_education = int(df[df['영역'] == '인성']['학점'].sum()) #4
    # 공통 교육
    common_education = int(df[df['영역'] == '공통교육']['학점'].sum()) #8
    # 평점 평균
    GPA = float(str(df['등급'].dropna().mean())[:4]) #4.38

    user = {
        'Minimum Total': total, # 종합 학점
        'Minimum Major': major, # 전공 학점
        'Major Professional': major_professional, # 전문 학점
        'Major Basic': major_basic, # 기초 학점
        'General Education': general_education, # 교양
        'Departmental Education': departmental_education, # 계열교육
        'Self-Development': self_development, # 자기계발
        'Personality Education': personality_education, # 인성
        'Common Education': common_education, # 공통 교육
        'Minimum GPA': GPA, 
    }

    for key in rules:
        v1, v2 = rules[key], user[key]
        result[key] = v1 - v2

    print(result)

    return render_template('result.html', result=result)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=1241)