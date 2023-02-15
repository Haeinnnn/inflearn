import random
import xlwings as xw
import pandas as pd
import numpy as np
import pickle
import re

filebook = xw.Book('경로.csv')
file1 = filebook.sheet(1).used_range.options(pd.DataFrame).value
file1.reset_index(inplace = True)

# 열 합치기
cols = ['TYPE', 'PRED']
file1['Combine'] = file1[cols].apply(lambda row: "/".join(row.values.astype(str)), axis=1)

# 열 이름 바꾸기
file1.rename(columns={"Combine" : "결과값합침"}, inplace=True)
