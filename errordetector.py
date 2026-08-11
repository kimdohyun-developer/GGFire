#---------------------------------------
#특근매식비 엑셀파일 오류 검출기 입니다. 
#---------------------------------------


from openpyxl import load_workbook
import os

def to_total_minutes(time_str):
    # 빈 값이나 문자열이 아닌 경우 기본 커트
    if not time_str or not isinstance(time_str, str) or ':' not in time_str:
        return "이상치"
    
    # 익일 여부 확인 (앞이나 뒤에 '익일'이 포함되어 있는지 검사)
    is_next_day = "익일" in time_str
    
    try:
        # 숫자와 콜론(:)만 남기고 한글 및 공백 제거
        clean_str = "".join(c for c in time_str if c.isdigit() or c == ':')
        hour, minute = map(int, clean_str.split(':'))
        
        # 시간 형식 자체의 오류 검증 (시는 0~23, 분은 0~59)
        if not (0 <= hour < 24 and 0 <= minute < 60):
            return "이상치"
            
        # 기본 분 계산
        total_min = hour * 60 + minute
        
        # 만약 '익일' 데이터라면 하루치 분(24시간 = 1440분)을 더해줌
        if is_next_day:
            total_min += 24 * 60  # + 1440분
            
        return total_min
        
    except ValueError:
        return "이상치"

# 경로 문자열 앞 r 추가 (안전망)
folder_path = r"C:\Users\User\Desktop\매식비 검사기"
files = os.listdir(folder_path)

print("\n-----------------------------[FILE LIST]-------------------------------\n")
for i in range(0, len(files)):
    print(f"{i+1}. {files[i]}")
print("\n----------------------------[SELECT STEP]------------------------------\n")

num = int(input("몇번으로 진행하시겠습니까?\n--> "))
obj_file = files[num-1]
print(f"{obj_file}으로 파일을 실행합니다.")

# 여기서 이미 풀 경로를 조화롭게 잘 만들어 두셨습니다!
full_path = os.path.join(folder_path, obj_file)

name, ext = os.path.splitext(obj_file)
size = os.path.getsize(full_path)

print("\n-------------------------------[INFO]----------------------------------\n")
print(f"[ 파일명: {name}{ext or '(확장자 없음)'} / 크기: {size} bytes ]")
print("\n------------------------------[RUNNING]--------------------------------\n")

# 1. 엑셀 파일 로드 (문제가 있던 부분을 올바른 풀 경로 변수로 대체)
wb = load_workbook(full_path, data_only=True)
sheet = wb['재난대응과']

# 2. 전체 행의 길이(마지막 행 번호) 알아내기
max_row = sheet.max_row

# 데이터를 저장할 2차원 배열
matrix_2d = []

# 3. 4행부터 마지막 행까지 한 행씩 반복 처리
for row_idx in range(4, max_row + 1):
    current_row = []
    
    # [수정 포인트] 원하는 열 번호(2열, 8열, 9열)만 특정하여 순회
    target_columns = [2, 7, 8, 9]
    for col_idx in target_columns:
        cell_value = sheet.cell(row=row_idx, column=col_idx).value
        current_row.append(cell_value)
        
    # 모든 원소가 None이 아닌 경우에만(데이터가 있는 경우에만) 배열에 추가
    if any(cell_value is not None for cell_value in current_row):
        matrix_2d.append(current_row)

# 결과 확인 (전체 행 개수 및 데이터 출력)
print(f"총 저장된 행의 개수: {len(matrix_2d)}개")
#print(matrix_2d)

# 결과를 저장할 리스트
invalid_records = []

for i in range(len(matrix_2d)):
    name = matrix_2d[i][0]
    day_of_week = matrix_2d[i][1]
    start_time_str = matrix_2d[i][2]
    end_time_str = matrix_2d[i][3]
    
    # 1. 숫자로 시간 환산
    start_min = to_total_minutes(start_time_str)
    end_min = to_total_minutes(end_time_str)
    
    # [이상치 탐지] 데이터 자체에 오류가 있는 경우 탈락 리스트에 바로 포함
    if start_min == "이상치" or end_min == "이상치":
        invalid_records.append({
            "행": i + 4, "이름": name, "요일": day_of_week, 
            "출근": start_time_str, "퇴근": end_time_str, "사유": "시간 데이터 오류(이상치)"
        })
        continue

    # 근무 시간(분) 계산
    work_duration = end_min - start_min
    
    # 통과 여부를 체크할 플래그 (True면 정상, False면 탈락)
    is_valid = False

    # 2. 주말 (토요일, 일요일) 조건 검사
    if day_of_week in ["토", "일"]:
        if work_duration >= 60:  # 1시간(60분) 이상 근무 시 통과
            is_valid = True
        else:
            fail_reason = f"주말 근무 시간 미달 ({work_duration}분 근무)"

    # 3. 평일 조건 검사
    else:
        # 조건 A: 8시 이전 출근 (<= 08:00)
        condition_A = start_min <= 8 * 60
        
        # 조건 B: 9시 출근(<= 09:00) 이고 21시 이후 퇴근(>= 21:00)
        condition_B = (start_min <= 9 * 60) and (end_min >= 21 * 60)
        
        if condition_A or condition_B:  # 둘 중 하나라도 만족하면 통과
            is_valid = True
        else:
            fail_reason = "평일 출퇴근 조건 미달 (8시 이전 출근 불가 또는 21시 이전 퇴근)"

    # 4. 조건에 해당하지 않는(탈락한) 데이터만 따로 추출하여 저장
    if not is_valid:
        invalid_records.append({
            "행": i + 4, "이름": name, "요일": day_of_week, 
            "출근": start_time_str, "퇴근": end_time_str, "사유": fail_reason
        })
print("\n-------------------------------[RESULT]--------------------------------\n")
print(f"조건 추출 결과: 오류 총 {len(invalid_records)}건")
print("\n-------------------------------[ERROR]---------------------------------\n")
for record in invalid_records:
    print(f"[엑셀 {record['행']}행] {record['이름']} ({record['요일']}) | 출근: {record['출근']}, 퇴근: {record['퇴근']} ->  {record['사유']}")

print("\n---------------------------------[END]----------------------------------\n")
input("\n프로그램을 종료하려면 엔터(Enter) 키를 누르세요...")
