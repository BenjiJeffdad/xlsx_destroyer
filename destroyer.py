"""
Author: Jeonghoon Lee
Last Modification: 2022.01.09
https://github.com/Benjijeffdad/xlsx_destroyer
"""

import sys
import os
import pyexcel as px
import random
import time

print("Process Start")
start_time=time.time()

# 파괴하려는 xlsx 파일 저장 폴더 이름
directory = sys.argv[1]

# 몇퍼센트의 데이터를 파괴할까요?
percent = float(sys.argv[2])/100

# 폴더안에 있는 파일 목록을 받아오기
files = os.listdir(directory)

# 원래 있떤 자료 대신 집어 넣을 가짜 단어를 모아줍니다.

TERROR = ["고양이", "야옹", "냥코", "네롱", "고양이 사랑해요"]

# for 문을 돌면서 파일을 하나씩 읽어옵니다.
for filename in files:
    # xlsx 파일이 아닌 경우 건너뜁니다.
    if not filename.endswith(".xlsx"):
        continue

    file_array = px.get_array(file_name=directory + "/" + filename)

    # 엑셀 파일을 위에서부터 한 줄씩 불러옵니다.
    for i in range(len(file_array)):
        # 엑셀 파일을 왼쪽에서부터 한 개씩 불러옵니다.
        for j in range(len(file_array[0])):
            if random.random() < percent:
                #엑셀 파일 내용을 바꿔치기
                file_array[i][j] = random.choice(TERROR)

    # 수정이 끝난 파일을 저장합니다.
    px.save_as(array=file_array, dest_file_name = directory + "/ " + filename)

print("Process Done")
end_time=time.time()
print("The job took " + str(end_time - start_time) + " seconds.")
