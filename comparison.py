import pandas as pd
import re
import unicodedata

def clean_name(name: str, remove_english: bool = False) -> str:
    """
    A, B 모두에 사용할 통합 전처리 함수.
    1. 유니코드 정규화 (NFKC)
    2. non-breaking space 등 특수 공백 제거
    3. 개행문자가 있으면, 줄 단위로 분리 후 한자가 등장하기 전까지의 줄을 합침
    4. 괄호와 쌍따옴표 제거
    5. 모든 공백(띄어쓰기, 탭, 개행 등)을 제거
    6. remove_english가 True이면 영어 알파벳([A-Za-z])를 제거
    """
    if pd.isnull(name):
        return ''
    # 1. 유니코드 정규화
    text = unicodedata.normalize('NFKC', str(name))
    # 2. 특수 공백 제거: non-breaking space, zero-width space 등
    text = text.replace('\u00A0', ' ').replace('\u200B', '')
    
    # 3. 개행문자가 있으면, 줄 단위로 분리하여 한자가 등장하기 전까지의 줄을 사용
    if '\n' in text:
        lines = text.split('\n')
        kept_lines = []
        for line in lines:
            line = line.strip()
            # 한자(유니코드 \u4E00-\u9FFF)가 나오면 그 줄부터는 무시
            if re.search(r'[\u4E00-\u9FFF]', line):
                break
            kept_lines.append(line)
        text = ''.join(kept_lines)
    
    # 4. 괄호 및 쌍따옴표 제거
    text = re.sub(r'[()]', '', text)
    text = text.replace('"', '')
    
    # 5. 모든 공백 제거 (문자 사이 공백 포함)
    text = re.sub(r'\s+', '', text)
    
    # 6. 영어 알파벳 제거 (옵션)
    if remove_english:
        text = re.sub(r'[A-Za-z]', '', text)
    
    return text.strip()

def keep_only_hangul(text: str) -> str:
    """
    문자열에서 오직 한글([가-힣])만 추출하여 반환.
    """
    if pd.isnull(text):
        return ''
    return ''.join(re.findall(r'[가-힣]', text))

# === 1차 검증 (첫 번째 비교) ===

# 파일 경로 설정
file_a = "문화재정보A.xlsx"  # A 파일: 단일 시트
file_b = "문화재정보B.xlsx"  # B 파일: 시트별 자료

# A 파일 읽기
df_a = pd.read_excel(file_a)
# B 파일은 모든 시트를 딕셔너리 형태로 읽음
df_b = pd.read_excel(file_b, sheet_name=None)

# A 파일의 문화재명(한글) 컬럼 전처리: 영어 제거 옵션(True) 적용
df_a['문화재명(한글)'] = df_a['문화재명(한글)'].apply(lambda x: clean_name(x, remove_english=True))

# A 파일의 '지정구분' 값과 B 파일 시트 이름 매핑
mapping = {
    '국가지정문화재': '국가 지정문화재',
    '국가등록문화재': '국가 등록문화재',
    '시도지정문화재': '서울시 지정문화재',
    '시등록문화재': '서울시 등록문화재'
}

# 결과 저장용 리스트 (1차 검증)
missing_in_b_list = []  # A에는 있으나 B에는 없는 항목 (지정구분, 연번, 문화재명(한글))
missing_in_a_list = []  # B에는 있으나 A에는 없는 항목 (시트명, 문화재명)

for a_value, sheet_name in mapping.items():
    # A 파일에서 해당 지정구분 필터링 및 전처리(영어 제거 포함)
    df_a_filtered = df_a[df_a['지정구분'] == a_value].copy()
    df_a_filtered['문화재명(한글)'] = df_a_filtered['문화재명(한글)'].apply(lambda x: clean_name(x, remove_english=True))
    
    # B 파일에서 해당 시트 가져오기
    df_b_sheet = df_b.get(sheet_name)
    if df_b_sheet is None:
        print(f"시트 '{sheet_name}'가 B 파일에 없습니다.")
        continue

    # B 파일 시트의 컬럼명 정리
    df_b_sheet.columns = df_b_sheet.columns.str.strip()
    
    # 비교에 필요한 컬럼 확인
    if '문화재명' not in df_b_sheet.columns:
        print(f"'{sheet_name}' 시트에 '문화재명' 컬럼이 없습니다.")
        continue

    # ----- 추가 필터링: 해당 시트의 불필요 항목 제외 -----
    if sheet_name == "서울시 지정문화재":
        if "문화유산" in df_b_sheet.columns:
            df_b_sheet = df_b_sheet[df_b_sheet["문화유산"] != "무형문화유산"]
    elif sheet_name == "국가 지정문화재":
        if "종목" in df_b_sheet.columns:
            df_b_sheet = df_b_sheet[df_b_sheet["종목"] != "국가무형유산"]
    # --------------------------------------------------------

    # B 파일의 '문화재명' 컬럼에도 전처리 적용 (영어 제거 없이)
    df_b_sheet['문화재명'] = df_b_sheet['문화재명'].apply(clean_name)

    # 집합 비교 (1차 검증)
    names_a = set(df_a_filtered['문화재명(한글)'])
    names_b = set(df_b_sheet['문화재명'])

    diff_a = df_a_filtered[~df_a_filtered['문화재명(한글)'].isin(names_b)]
    if not diff_a.empty:
        missing_in_b_list.append(diff_a[['지정구분', '연번', '문화재명(한글)']])
    
    diff_b = df_b_sheet[~df_b_sheet['문화재명'].isin(names_a)]
    if sheet_name == "서울시 지정문화재":
        diff_b = diff_b[diff_b['문화재명'] != ""]
    if not diff_b.empty:
        diff_b = diff_b.copy()
        diff_b['시트명'] = sheet_name
        missing_in_a_list.append(diff_b[['시트명', '문화재명']])

# 결과 문자열 구성 (1차 검증)
output_str = ""
if missing_in_b_list:
    result_a_missing = pd.concat(missing_in_b_list, ignore_index=True)
    output_str += "1차 검증 - A 파일에는 있으나 B 파일에는 없는 항목 (지정구분, 연번, 문화재명(한글)):\n"
    output_str += result_a_missing.to_string() + "\n"
else:
    output_str += "1차 검증 - A 파일의 모든 항목이 B 파일에 존재합니다.\n"

if missing_in_a_list:
    result_b_missing = pd.concat(missing_in_a_list, ignore_index=True)
    output_str += "\n1차 검증 - B 파일에는 있으나 A 파일에는 없는 항목 (시트명, 문화재명):\n"
    output_str += result_b_missing.to_string() + "\n"
else:
    output_str += "1차 검증 - B 파일의 모든 항목이 A 파일에 존재합니다.\n"

# 1차 검증 결과 출력 및 저장
with open("result_first_pass.txt", "w", encoding="utf-8") as f:
    f.write(output_str)
print(output_str)

# === 2차 검증: 1차 검증된 자료에서 오직 한글만 남긴 자료로 재비교 ===

# 1차 검증된 A, B 데이터를 대상으로 각 문화재명에서 오직 한글만 추출하여 비교합니다.
missing_in_b_list_hangul = []  # A에는 있으나 B에는 없는 항목 (지정구분, 연번, 문화재명(한글)_hangul)
missing_in_a_list_hangul = []  # B에는 있으나 A에는 없는 항목 (시트명, 문화재명_hangul)

for a_value, sheet_name in mapping.items():
    df_a_filtered = df_a[df_a['지정구분'] == a_value].copy()
    # 새 컬럼: 1차 검증 결과에서 오직 한글만 남김
    df_a_filtered['문화재명(한글)_hangul'] = df_a_filtered['문화재명(한글)'].apply(keep_only_hangul)
    
    df_b_sheet = df_b.get(sheet_name)
    if df_b_sheet is None:
        continue
    df_b_sheet.columns = df_b_sheet.columns.str.strip()
    if '문화재명' not in df_b_sheet.columns:
        continue
    # 기존 전처리된 B 파일 데이터에서 오직 한글만 남김
    df_b_sheet['문화재명_hangul'] = df_b_sheet['문화재명'].apply(keep_only_hangul)
    
    names_a_hangul = set(df_a_filtered['문화재명(한글)_hangul'])
    names_b_hangul = set(df_b_sheet['문화재명_hangul'])
    
    diff_a_hangul = df_a_filtered[~df_a_filtered['문화재명(한글)_hangul'].isin(names_b_hangul)]
    if not diff_a_hangul.empty:
        missing_in_b_list_hangul.append(diff_a_hangul[['지정구분', '연번', '문화재명(한글)_hangul']])
    
    diff_b_hangul = df_b_sheet[~df_b_sheet['문화재명_hangul'].isin(names_a_hangul)]
    if sheet_name == "서울시 지정문화재":
        diff_b_hangul = diff_b_hangul[diff_b_hangul['문화재명_hangul'] != ""]
    if not diff_b_hangul.empty:
        diff_b_hangul = diff_b_hangul.copy()
        diff_b_hangul['시트명'] = sheet_name
        missing_in_a_list_hangul.append(diff_b_hangul[['시트명', '문화재명_hangul']])


## 결과 
# B가 83개 더 많음
# AO BX : 81개 
# AX BO : 72개

# 진짜 열받는네
