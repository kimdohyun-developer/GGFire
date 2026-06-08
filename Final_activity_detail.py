import os
import pandas as pd
from openpyxl.utils import get_column_letter

# 1. 파일 경로 및 시트 설정
input_file = "/content/xsell/2026년 소집내역(수당연계, 활동실적 계).xlsx"
output_file = "/content/xsell/항목별_소집내역_분류.xlsx"
target_sheet = 'temp'

# 원본 데이터 로드 (6번째 행부터 데이터 본문이므로 header=5)
df_raw = pd.read_excel(input_file, sheet_name=target_sheet, header=4)

# -------------------------------------------------------------------------
# [단계 1] 14번째 열(인덱스 13) 항목 추출 후 1차원 배열에 저장
# -------------------------------------------------------------------------
df_raw.iloc[:, 13] = df_raw.iloc[:, 13].astype(str).str.strip()
raw_unique = df_raw.iloc[:, 13].unique()

# 고유 항목들을 담을 1차원 배열(리스트) 생성 (결측치 제외)
unique_items_array = [item for item in raw_unique if item and item != 'nan']

print(f"📦 [1차원 배열 생성 완료] 총 {len(unique_items_array)}개의 항목이 순서대로 배열에 저장되었습니다.\n")


# -------------------------------------------------------------------------
# [데이터 포맷 정제] F열 뒤 8자리 슬라이싱 및 H, I, J열 시간 복원
# -------------------------------------------------------------------------
# F열 (인덱스 5) 단일 열만 정확히 지정하여 뒤에서 8자리만 남기기
if len(df_raw.columns) > 5:
    f_col_name = df_raw.columns[5]  # 정확히 6번째(F열)의 이름을 가져옴
    df_raw[f_col_name] = df_raw[f_col_name].astype(str).str.strip()
    df_raw[f_col_name] = df_raw[f_col_name].apply(
        lambda x: x[-8:] if (x != 'nan' and len(x) >= 8) else x
    )
    df_raw[f_col_name] = df_raw[f_col_name].replace({'nan': '', 'None': ''})

# H, I, J열 (인덱스 7, 8, 9) 시간 초단위 자르기
time_cols = [7, 8, 9]
for idx in time_cols:
    if idx < len(df_raw.columns):
        col_name = df_raw.columns[idx]
        df_raw[col_name] = df_raw[col_name].astype(str).str.strip()
        df_raw[col_name] = df_raw[col_name].apply(
            lambda x: x[:5] if (':' in x and len(x) >= 8 and x != 'nan') else x
        )
        df_raw[col_name] = df_raw[col_name].replace({'nan': '', 'None': ''})


# -------------------------------------------------------------------------
# [단계 2] 1차원 배열 순서대로 시트를 하나씩 만들고 분류 데이터 주입
# -------------------------------------------------------------------------
# 맨 첫 번째 행(1행)에 무조건 먼저 채워 넣을 고정 대제목 명단 (총 14개 열)
fixed_headers = [
    "연번", "소속", "직위", "성명", "장소", "활동일자", "월", 
    "시작시간","종료시간", "시간", "합계", "인정시간", "금액", "비고", "활동내역"
]

# 저장할 폴더 생성
os.makedirs(os.path.dirname(output_file), exist_ok=True)

print("🚀 [시트 순차 생성 시작] 1행에 항목을 먼저 작성한 후 2행부터 데이터를 주입합니다...")

with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    for index, item in enumerate(unique_items_array):
        
        # 현재 항목에 해당하는 데이터만 필터링
        filtered_df = df_raw[df_raw.iloc[:, 13] == item].copy()
        
        # 고정 헤더 구조를 가진 데이터프레임 틀 생성
        final_df = pd.DataFrame(columns=fixed_headers)
        
        # 필터링된 데이터를 14개 고정 헤더 위치에 정확히 1:1 주입
        for i, header_name in enumerate(fixed_headers):
            if i < len(filtered_df.columns):
                final_df[header_name] = filtered_df.iloc[:, i]
        
        # 시트 이름 명명 규칙 적용 (최대 31자 제한 반영)
        sheet_name = f"{index + 1}. {item}"[:31]
        
        # 빈 시트를 활성화하거나 새로 만듭니다.
        if sheet_name not in writer.book.sheetnames:
            ws = writer.book.create_sheet(title=sheet_name)
        else:
            ws = writer.book[sheet_name]
            
        # [순서 유지] 데이터를 넣기 전, 무조건 1번째 행에 항목 제목을 수동으로 먼저 입력합니다.
        for col_idx, header_text in enumerate(fixed_headers):
            ws.cell(row=1, column=col_idx + 1).value = header_text
        
        # 1행 선입력이 끝난 직후, 그 아래 공간인 2번째 행(startrow=1)부터 실제 데이터 본문을 주입합니다.
        final_df.to_excel(writer, sheet_name=sheet_name, index=False, header=False, startrow=1)
        
        # ---------------------------------------------------------------------
        # [버그 박멸 🛠️] 빈 도화지 상태에서도 안전하게 구동되는 열 너비 맞춤 로직
        # ---------------------------------------------------------------------
        for col_cells in ws.columns:
            max_len = 0
            # 튜플 형태의 열 데이터에서 첫 번째 셀(인덱스 0)을 콕 찝어 정확한 열 번호를 구합니다.
            col_letter = get_column_letter(col_cells[0].column)
            
            for cell in col_cells:
                if cell.value:
                    val_str = str(cell.value)
                    # 한글(전각문자) 길이를 고려하여 너비 계산 (한글은 2글자 취급)
                    cell_len = sum([2 if ord(char) > 128 else 1 for char in val_str])
                    if cell_len > max_len:
                        max_len = cell_len
            
            # 최소 너비 14 확보 및 글자 길이에 여유 마진(+4)을 부여합니다.
            ws.column_dimensions[col_letter].width = max(max_len + 4, 14)
        # ---------------------------------------------------------------------
        
        print(f"   ▶ [{index + 1}번 시트] {sheet_name} 완료 ({len(final_df)}건)")

print(f"\n🎉 모든 분류 작업이 완료되었습니다! 최종 파일 확인: {output_file}")
