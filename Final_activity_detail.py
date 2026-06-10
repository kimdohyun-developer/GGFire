import pandas as pd
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# 1. 파일 이름 설정
input_file = "/content/2026년 소집내역(수당연계, 활동실적 계).xlsx"
output_file = "/content/대별활동실적.xlsx"

# 분류할 6개 소속(대) 목록
sheet_names = ["산악전문대", "대학생전문대", "장안남성대", "장안여성대", "영통남성대", "영통여성대"]

# 공백을 제거한 비교용 딕셔너리 생성 (띄어쓰기 오류 방지)
sheet_mapping = {name.replace(" ", ""): name for name in sheet_names}

# 순번 없이 원본에서 그대로 가져올 실제 열 번호 목록 (2열: 소속, 6열: 날짜, 14열: 대분류, 15열: 활동명칭)
target_columns = [2, 6, 14, 15]

# 2. 원본 파일 읽기 및 새 워크북 생성
wb_source = load_workbook(input_file, data_only=True)
ws_source = wb_source.active

wb_target = Workbook()
# 첫 번째 시트 이름 지정
wb_target.active.title = sheet_names[0]

# 두 번째 시트부터 순서대로 생성
for name in sheet_names[1:]:
    wb_target.create_sheet(title=name)

# 6개 시트의 행 포인터 초기화 (각 시트의 1행부터 데이터 입력 시작)
row_pointers = {name: 1 for name in sheet_names}

# 3. 원본 데이터를 행별로 돌며 분류 및 복사
for row_idx in range(1, ws_source.max_row + 1):
    # 2열(소속) 값 가져오기
    dept_val = ws_source.cell(row=row_idx, column=2).value

    if dept_val is not None:
        clean_dept = str(dept_val).strip().replace(" ", "")

        # 일치하는 소속 시트가 있다면 데이터 복사 실행
        if clean_dept in sheet_mapping:
            target_sheet_name = sheet_mapping[clean_dept]
            ws_target = wb_target[target_sheet_name]

            # 현재 시트에 채워야 할 행 번호
            current_row = row_pointers[target_sheet_name]

            # 1열부터 빈칸 없이 추출한 4개의 열을 촘촘하게 채웁니다.
            # t_idx는 새 시트의 열 번호 (1열부터 시작)
            for t_idx, col_num in enumerate(target_columns, start=1):
                cell_value = ws_source.cell(row=row_idx, column=col_num).value

                # 원본 6열(날짜) 데이터일 경우 문자열로 변환 후 뒤에서 8자리를 지움
                if col_num == 6 and cell_value is not None:
                    val_str = str(cell_value).strip()
                    if len(val_str) > 8:
                        cell_value = val_str[:-8]

                # 새 시트에 값 입력
                ws_target.cell(row=current_row, column=t_idx).value = cell_value

            # 해당 시트의 다음 입력 위치를 1 증가
            row_pointers[target_sheet_name] += 1

# 4. 분류된 새 엑셀 파일 저장
wb_target.save(output_file)

print(f"🎉 번호 없이 [소속, 날짜, 14열, 15열] 데이터만 1행 1열부터 분류를 완료했습니다! -> {output_file}")


# 1. 파일 경로 설정
input_file = "/content/대별활동실적.xlsx"
output_file = "/content/★대별 활동실적.xlsx"

# 처리할 시트 목록
sheet_names = ["산악전문대", "대학생전문대", "장안남성대", "장안여성대", "영통남성대", "영통여성대"]

# 2. 결과 생성을 위한 엑셀 파일 열기 (원본 데이터 읽기용)
wb_in = load_workbook(input_file, data_only=True)

# 완전히 새로 채워 넣기 위해 깨끗한 새 Workbook 생성 (원본 데이터 제외)
wb_out = Workbook()

# 리스트의 첫 번째 값인 "산악전문대" 문자열을 첫 시트 이름으로 지정
wb_out.active.title = sheet_names[0]

# 두 번째 시트부터 순서대로 생성
for name in sheet_names[1:]:
    wb_out.create_sheet(title=name)

# 가독성을 위한 서식/스타일 정의 (감청색 톤)
font_title = Font(name="맑은 고딕", size=12, bold=True, color="1B365D")
font_header = Font(name="맑은 고딕", size=10, bold=True, color="FFFFFF") # 흰색 글씨 헤더
font_data = Font(name="맑은 고딕", size=10)
fill_header = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid") # 감청색 배경
align_center = Alignment(horizontal="center", vertical="center")
align_left = Alignment(horizontal="left", vertical="center")
thin_border = Border(
    left=Side(style='thin', color='D9D9D9'),
    right=Side(style='thin', color='D9D9D9'),
    top=Side(style='thin', color='D9D9D9'),
    bottom=Side(style='thin', color='D9D9D9')
)

for sheet_name in sheet_names:
    if sheet_name not in wb_in.sheetnames:
        continue

    ws_in = wb_in[sheet_name]
    if ws_in.max_row == 0:
        continue

    # 데이터 추출 (1열: 소속, 2열: 날짜, 3열: 대분류, 4열: 활동명칭)
    data = []
    for r in range(1, ws_in.max_row + 1):
        row_vals = [ws_in.cell(row=r, column=c).value for c in range(1, 5)]
        if row_vals[1] is not None and row_vals[3] is not None:
            data.append(row_vals)

    if not data:
        continue

    # 데이터프레임 변환
    df = pd.DataFrame(data, columns=['소속', '날짜', '대분류', '활동명칭'])

    # [1단계 집계] 날짜와 활동명이 같은 중복 행을 카운트 (+1씩 누적하여 참여인원 계산)
    group_date_name = df.groupby(['날짜', '대분류', '활동명칭']).size().reset_index(name='참여인원')

    # [2단계 집계] 날짜는 다르지만 활동명이 같은 데이터를 그룹화
    # - 제거할 때마다 +1 카운트 (=회차)
    # - 해당 활동명칭의 최종 누적 인원수 계산 (=총참여인원)
    final_summary = group_date_name.groupby(['대분류', '활동명칭']).agg(
        회차=('참여인원', 'count'),
        총참여인원=('참여인원', 'sum')
    ).reset_index()

    ws_out = wb_out[sheet_name]

    # 3. 요약 표 작성 (1행 타이틀)
    ws_out.cell(row=1, column=1, value=f"📊 [{sheet_name}] 대별 활동").font = font_title

    # 2행: 요약 표 헤더 설정
    headers = ["항목(대분류)", "활동명(명칭)", "회차(횟수)", "참여 인원수"]
    for c_idx, header in enumerate(headers, start=1):
        cell = ws_out.cell(row=2, column=c_idx, value=header)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = align_center
        cell.border = thin_border

    # 3행부터 데이터 삽입
    current_row = 3
    for index, row in final_summary.iterrows():
        vals = [
            row['대분류'],
            row['활동명칭'],
            f"{row['회차']}",      # 맨 마지막에서 한 칸 앞 (회차)
            f"{row['총참여인원']}"  # 맨 마지막 열 (참여 인원수)
        ]

        for c_idx, val in enumerate(vals, start=1):
            cell = ws_out.cell(row=current_row, column=c_idx, value=val)
            cell.font = font_data
            cell.border = thin_border

            if c_idx >= 3:
                cell.alignment = align_center
            else:
                cell.alignment = align_left

        current_row += 1

    # 📌 [완벽 수정 부] 튜플을 사용하지 않고 열 번호 숫자(1~4)로 안전하게 너비를 조절합니다.
    for col_idx in range(1, 5):
        # 해당 열의 모든 셀을 돌며 가장 긴 글자수 찾기
        max_len = 0
        for row_idx in range(1, current_row):
            val = ws_out.cell(row=row_idx, column=col_idx).value
            if val is not None:
                max_len = max(max_len, len(str(val)))

        col_letter = get_column_letter(col_idx)
        ws_out.column_dimensions[col_letter].width = max(max_len * 2 + 4, 16)

# 5. 최종 결과 저장
wb_out.save(output_file)

print(f"🎉 모든 에러를 방지한 최종 요약 표 생성을 완료했습니다! -> {output_file}")
