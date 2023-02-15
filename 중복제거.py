import xlwings as xw
import pandas as pd
import numpy as np
import pickle
import re

book = xw.BOok("경로.csv")
df = book.sheets(1).used_range.options(pd.DataFrame).value

# 데이터프레임은 불러온뒤 인덱스 꼭 초기화 시켜주기! 안그러면 인덱스값으로 값이 들어감
df.reset_index(inplace=True)

# 컬럼명 변경
df.columns = list(map(lambda x: "COL" + str(x), range(len(df.columns))))

# 널값 확인하기
df.isnull().sum()

# 널값제거하기
df = df.fillna('')

# 구문 결합용 빈 열
df['CONCAT'] = ""

# 형식을 문자열로 변환 - 나중에 오류가 생겨서
df = df.astype('str')

# 따로 떨어져있는 구문 결합하기
cols = ['col10', 'col11', ~~마지막까지 입력]
df['CONCAT'] = df[cols].apply(lambda row: " ".join(row.values.astype(str)), axis=1)

# 원 결합 구문은
# STMT_START_IDX = 12 #구문시작
# for idx in range(STMT_START_IDX, len(df.columns) - 1): df['CONCAT'] += D

# 필요한 열만 추출
df_concat = df_drop[['COL10', 'COL11', 'CONCAT']]

# for문 돌릴값만 추출
# 필요한 열만 for문 돌리는걸 못해서 새 데이터프레임을 만듦
# ? 어떻게 나누지 않고 for문에서 데이터프레임 값만 빼서 돌릴까 ?
df_concat = df_drop[['CONCAT']]

# 중복된 문구를 제거하는 for 문
# messages가 자꾸 오류가 떠서 차장님이 뒤에 []를 붙여주심
# 콘솔창에 message.tostring 도 치셨던거 같음!

message = list()
text = list()
for messages in df_concat.values : 
    msg_free = "".join(messages[0].split('(무료)')[:-1]) # 무료 기준 앞그룹
    only_text = re.sub(r'[^가-힣a-zA-Z)(\]\[]', '', str(msg_free)) # 숫자제거
    name_delete = re.sub("]...고객님", "]고객님", str(only_text)) # 고객명제거
    breacket_delete = re.sub("\)\(", "", name_delete) #  지워지고남은 빈괄호 제거 ()
    text.append(bracket_delete)

# 데이터프레임으로 변환
df_text = pd.DataFrame(text)
# 형식을 문자열로 변환
df_text = df_text.astype('str')


# concat이 잘 안되어서.... index로 결합을 사용
# 원래썼던 코드
# findresult = pd.merge(file1, file2, on="MID") # MID 기준으로 오른쪽을 왼쪽에 붙임

df_text.reset_index(inplace=True)

# 컬럼명 수정
df_text.columns = ["index", "전송문구_수정"]

# index 생성 (출력순번으로 떠서 두번해버림)
df_drop.reset_index(inplace=True)
df_drop.reset_index(inplace=True)

# index를 기준으로 원본문구와 중복된 문구가 제거된 데이터프레임을 결합
df_result = pd.merge(df_drop, df_text, on="index")

# 필요한 행만 다시 추출
df_result = df_result[['COL10', 'COL11', 'CONCAT', '전송문구_수정']]

# 일반통지성만 추출
# df_result = df_result[df_result["COL10"] == "일반통지성", ["COL11", "CONCAT", "전송문구_수정"]]

# 컬럼명 수정
df.rename(columns={"어쩌고" : "저쩌고"}, inplace=True)

# 전송문구_수정을 기준으로 중복제거
df_result = df_result.drop_duplicates(["전송문구_수정"])

# null값 확인
df_result.isnull().sum()

# 전송문구를 기준으로 솔트
df_result = df_result.sort_values("전송문구", ascending=True)

# csv 파일로 저장, dat 추가
df_result.to_csv(r"C:\USER\저장.csv", index=Flase, encoding="euc-kr")
df_result[["전송기본ID", "전송문구"]].to_csv(r"C:\이건왜넣냐면/를안쓰고 그냥 복붙했을때 읽히라고")

끝

