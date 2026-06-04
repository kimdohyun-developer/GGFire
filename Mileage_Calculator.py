from collections import Counter
import pandas as pd
from openpyxl import load_workbook

# 1. 원본 엑셀 파일 로드
load_wb = load_workbook("/content/xsell/2026년 소집내역(수당연계, 활동실적 계).xlsx", data_only=True)
load_ws = load_wb['소집활동 내역']

col_size = 348 
target_col1 = 4   # 이름 열 (D열)
target_col2 = 16  # 활동 내용 열 (P열)
start_row = 6 

# 2. 원본 데이터에서 [이름, 활동] 이차원 배열(combined_data) 생성
combined_data = []
for row_num in range(start_row, col_size + start_row):
    name_value = load_ws.cell(row=row_num, column=target_col1).value
    work_value = load_ws.cell(row=row_num, column=target_col2).value
    
    if name_value is not None and work_value is not None:
        combined_data.append([str(name_value).strip(), str(work_value).strip()])

# --------------------------------------------------
# [새로운 시퀀스 시작] 표를 먼저 만들고 하나씩 계산하기
# --------------------------------------------------

# 요구하신 25개 항목 정확하게 리스트로 정의 (가로축 헤더)
activity_columns = [
    "화재진압", "구조구급", "생활안전", "지원근무", "예방순찰", 
    "행사안전", "점검지원", "경계근무", "용수통로", "행사참여", 
    "급수지원", "안전교육", "캠페인활동", "정기교육", "소방학교", 
    "소방훈련", "소방기술경연대회", "불우이웃돕기", "재해복구", "기타", 
    "자격취득", "표창", "혁신제안", "안전사고", "부정사례"
]

# 중복 없는 사람 이름 명단 추출 (세로축 인덱스)
unique_names = sorted(list(set([row[0] for row in combined_data])))

# 단계 1: 가로는 25개 항목, 세로는 이름으로 구성된 빈 표(모든 칸이 0)를 먼저 생성
summary_df = pd.DataFrame(0, index=unique_names, columns=activity_columns)

# 단계 2: combined_data를 하나씩 불러와서 행/열에 맞는 칸의 빈도 계산(누적)
for name, work in combined_data:
    # 혹시 25개 항목 리스트에 없는 텍스트가 들어올 경우를 대비한 예외 처리
    if work in summary_df.columns:
        summary_df.loc[name, work] += 1

summary_df = summary_df.replace(0, None)

#print("=== 활동이 없는 칸을 빈칸 처리한 요약 표 ===")
#print(summary_df.head(10))

# 4. 최종 엑셀 파일로 저장
summary_df.to_excel("/content/xsell/소집내역_최종요약표.xlsx", index_label="이름")
print("\n[성공] 활동이 없는 칸이 모두 Null 처리되어 '소집내역_최종요약표.xlsx'로 저장되었습니다.")

#파일명이랑 전체 행길이만 입력해주면됨
