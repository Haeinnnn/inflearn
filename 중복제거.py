import xlwings as xw
import pandas as pd
import numpy as np
import pickle
import re

book = xw.BOok("경로.csv")
df = book.sheets(1).used_range.options(pd.DataFrame).value

# 데이터프레임은 불러온뒤 인덱스 꼭 초기화 시켜주기! 안그러면 인덱스값으로 값이 들어감
df.reset_index(inplace=True)

# 널값 확인하기
df.isnull().sum()

