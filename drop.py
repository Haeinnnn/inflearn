import os
import xlwings as xw
import pandas as pd
import numpy as np
import pickle
import re

# 엑셀폴더 경로
XLSX_PATH = r'C:\Users\
# 패스워드
XLSX_PASS = "5471"
# 구문시작컬럼
SMST_START_IDX = 13
# 결과폴더 경로
TARGET_PATH = r'C:\USERS

def processXlsxToCsv(xlsx_file) : 
    # 경로설정
    xlsx_path = os.path.join(XLSX_PATH, xlsx_file)
    # 엑셀파일가져오기
    xlsx_workbook = xw.Book(xlsx_path, password=XLSX_PASS)
    print("1. 엑셀파일 가져오기 완료")
    # 원본데이터
    orgn_data = xlsx_workbook.sheets(1).used_range.options(pd.DataFrame).value
    # 컬럼명 변경
    orgn_data = xlsx_workbook.sheets(1).used_range.options(pd.DataFrame).value
    #  nan 제거
    orgn_data = orgn_data.fillna('')
    print(" 2. nan 제거 완료")

    #### 구문 결합
    # 결합컬럼 생성
    orgn_data['CONCAT'] = ""
    # 구문 결합
    for idx in range(STMT_START_IDX, len(orgn_data.columns) -1): orgn_data["CONCAT'] += orgn_data['COL' + str(idx)]
    print("3.  구문 결합 완료")

    #### 결과 데이터셋 생성
    # 필요컬럼만 추출, 원본컬럼을 남겨두기 위해 하나 복제
    work_df = orgn_data[['COL10', 'COL11', 'CONCAT']]
    work_df['CONCAT1'] = work_df['CONCAT']
    print(" 4. 작업 DataFrame 생성")

    # 무료를 기준으로 뒤 글자를 자름
    work_df['CONCAT'] = work_df['CONCAT'].apply(lambda x:x.split('(무료)')[0]
    print(" 5. 무료를 기준으로 글자 자르기)
    work_df['CONCAT'] = work_df['CONCAT'].apply(lambda x:re.sub(r'[^가-힣a-zA-Z\)\(\]\[]', '', X))
    print("6. 문자 기준으로 삭제")
    work_df['CONCAT'] = work_df['CONCAT'].apply(lambda x:x.rstrip("()"))
    work_df = work_df.drop_duplicates(['CONCAT'])
    work_df = work_df.sort_values("CONCAT", ascending=True)
    work_df = work_df.drop_duplicates(['CONCAT'])
    print("9. 정렬을 해줌 / 7번은 빈괄호 삭제 8번은 정렬 - 안적음)

    ### 엑셀파일출력
    work_df.columns = ("판정결과", "전송기본ID", "수정문구", "원문구")
    target_path = os.path.join(TARGET_PATH, xlsx_file.replace(".", "_result."))
    work_df.to_excel(target_path)
    print("10. 엑셀파일출력")


# 하위 오브젝트 명칭
file_list = os.listdir(XLSX_PATH)
# 확장자가 XLSX 인것만 추출
xlsx_list = [file for file in file_list if file.endswith(".xlsx")]

for idx, xlsx_file in enumerate(xlsx_list, start=1):
    print("[작업시작 %s/%s] %s" % (str(idx), str(len(file_list)), xlsx_file))
    processXlsxToCsv(xlsx_file)
    print("[작업종료 %s/%s] %s" % (str(idx), str(len(file_list)), slxs_file))


