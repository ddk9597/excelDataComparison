import pandas as pd
import re
import unicodedata

def clean_name(name: str, remove_english: bool = False) -> str:
    """
    통합 전처리 함수.
    1. 유니코드 정규화 (NFKC)
    2. non-breaking space 등 특수 공백 제거
    3. 개행문자가 있으면, 줄 단위로 분리 후 한자가 등장하기 전까지의 줄을 합침
    4. 괄호와 쌍따옴표 제거
    5. 모든 공백(띄어쓰기, 탭, 개행 등)을 제거
    6. remove_english가 True이면 영어 알파벳([A-Za-z])를 제거
    7. 모든 특수문자 제거 (한글과 숫자만 남김)
    """
    if pd.isnull(name):
        return ''
    text = unicodedata.normalize('NFKC', str(name))
    text = text.replace('\u00A0', ' ').replace('\u200B', '')
    if '\n' in text:
        lines = text.split('\n')
        kept_lines = []
        for line in lines:
            line = line.strip()
            if re.search(r'[\u4E00-\u9FFF]', line):
                break
            kept_lines.append(line)
        text = ''.join(kept_lines)
    text = re.sub(r'[()]', '', text)
    text = text.replace('"', '')
    text = re.sub(r'\s+', '', text)
    if remove_english:
        text = re.sub(r'[A-Za-z]', '', text)
    text = re.sub(r'[^0-9가-힣]', '', text)
    return text.strip()

# 파일 경로 설정
file_a = "문화재정보A.xlsx"  # A 파일: 단일 시트
file_b = "문화재정보B.xlsx"  # B 파일: 시트별 자료

# A 파일 읽기 및 원본 보존
df_a = pd.read_excel(file_a)
if "문화재명(한글)" not in df_a.columns:
    raise Exception("A 파일에 '문화재명(한글)' 컬럼이 없습니다.")
# 원본 값 보존
df_a["원본문화재명"] = df_a["문화재명(한글)"]
# 전처리된 값(비교용)은 별도 컬럼에 저장 (영어 제거 적용)
df_a["문화재명(한글)_가공"] = df_a["원본문화재명"].apply(lambda x: clean_name(x, remove_english=True))

# B 파일 읽기 (모든 시트를 dict로)
df_b = pd.read_excel(file_b, sheet_name=None)

# A 파일의 '지정구분'과 B 파일 시트 이름 매핑
mapping = {
    '국가지정문화재': '국가 지정문화재',
    '국가등록문화재': '국가 등록문화재',
    '시도지정문화재': '서울시 지정문화재',
    '시등록문화재': '서울시 등록문화재'
}

missing_in_b_list = []  # A 파일에는 있으나 B 파일에는 없는 항목 (지정구분, 연번, 원본문화재명)
missing_in_a_list = []  # B 파일에는 있으나 A 파일에는 없는 항목 (시트명, 지정번호, 원본문화재명)

for a_value, sheet_name in mapping.items():
    # A 파일에서 해당 지정구분 필터링 (가공된 값 비교)
    df_a_filtered = df_a[df_a['지정구분'] == a_value].copy()
    # 이미 "문화재명(한글)_가공" 컬럼에 전처리 값 저장되어 있음
    
    # B 파일에서 해당 시트 가져오기
    df_b_sheet = df_b.get(sheet_name)
    if df_b_sheet is None:
        print(f"시트 '{sheet_name}'가 B 파일에 없습니다.")
        continue
    df_b_sheet.columns = df_b_sheet.columns.str.strip()
    if "문화재명" not in df_b_sheet.columns:
        print(f"'{sheet_name}' 시트에 '문화재명' 컬럼이 없습니다.")
        continue
    # B 파일도 원본 보존 및 전처리
    df_b_sheet["원본문화재명"] = df_b_sheet["문화재명"]
    df_b_sheet["문화재명_가공"] = df_b_sheet["원본문화재명"].apply(lambda x: clean_name(x, remove_english=False))
    
    # 추가 필터링: 서울시 지정문화재의 경우 "문화유산" 컬럼이 있으면 무형문화유산 항목 제외,
    # 국가 지정문화재의 경우 "종목" 컬럼이 있으면 국가무형유산 항목 제외
    if sheet_name == "서울시 지정문화재":
        if "문화유산" in df_b_sheet.columns:
            df_b_sheet = df_b_sheet[df_b_sheet["문화유산"] != "무형문화유산"]
    elif sheet_name == "국가 지정문화재":
        if "종목" in df_b_sheet.columns:
            df_b_sheet = df_b_sheet[df_b_sheet["종목"] != "국가무형유산"]
    
    # 집합 비교: 비교는 전처리된(가공) 값 기준으로 수행
    names_a = set(df_a_filtered["문화재명(한글)_가공"])
    names_b = set(df_b_sheet["문화재명_가공"])
    
    diff_a = df_a_filtered[~df_a_filtered["문화재명(한글)_가공"].isin(names_b)]
    if not diff_a.empty:
        # 출력 시 원본 컬럼 사용: 지정구분, 연번, 원본문화재명
        missing_in_b_list.append(diff_a[["지정구분", "연번", "원본문화재명"]])
    
    diff_b = df_b_sheet[~df_b_sheet["문화재명_가공"].isin(names_a)]
    # 서울시 지정문화재 시트: 공백 제거
    if sheet_name == "서울시 지정문화재":
        diff_b = diff_b[diff_b["문화재명_가공"] != ""]
    if not diff_b.empty:
        diff_b = diff_b.copy()
        diff_b["시트명"] = sheet_name
        # 출력 시 원본 값: 시트명, 지정번호, 원본문화재명
        missing_in_a_list.append(diff_b[["시트명", "지정번호", "원본문화재명"]])

# 결과 문자열 구성
output_str = ""

if missing_in_b_list:
    result_a_missing = pd.concat(missing_in_b_list, ignore_index=True)
    output_str += "A 파일에는 있으나 B 파일에는 없는 항목 (지정구분, 연번, 원본문화재명):\n"
    output_str += result_a_missing.to_string() + "\n"
else:
    output_str += "A 파일의 모든 항목이 B 파일에 존재합니다.\n"

if missing_in_a_list:
    result_b_missing = pd.concat(missing_in_a_list, ignore_index=True)
    output_str += "\nB 파일에는 있으나 A 파일에는 없는 항목 (시트명, 지정번호, 원본문화재명):\n"
    output_str += result_b_missing.to_string() + "\n"
else:
    output_str += "B 파일의 모든 항목이 A 파일에 존재합니다.\n"

print(output_str)

with open("result.txt", "w", encoding="utf-8") as f:
    f.write(output_str)
