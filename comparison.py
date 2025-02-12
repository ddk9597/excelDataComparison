import pandas as pd

# 파일의 경로 지정
file_a = "문화재정보A.xlsx"
file_b = "문화재정보B.xlsx"

# 기준 파일 
df_a = pd.read_excel(file_a)

# 파일 B 읽기 
# 각 시트를 딕셔너리 형태로 저장. K : 시트이름, V : 해당 시트의 DataFrame
df_b = pd.read_excel(file_b, sheet_name=None)

# 데이터 구조 확인 
# 각 데이터의 기본 구조와 컬럼을 확인하여 이후 비교 작업을 위한 전처리 방향을 정한다.

# 파일 A데이터 미리 보기
print("파일 A 미리보기 : ")
print(df_a.head())
print("파일 A 컬럼 :", df_a.columns.to_list())


# 파일 B 각 시트의 데이터 미리보기
print("\n파일 B 각 시트 미리보기:")
for sheet_name, df in df_b.items():
    print(f"--- 시트: {sheet_name} ---")
    print(df.head())
    print("컬럼:", df.columns.tolist(), "\n")
    
## -- 02.12. 여기까지.. 이제 할 일 : 
# A 지정구분 : B시트 이름 통일하기
# A의 head()중 어떤 녀석들만 사용할지 자르기. -> 동일여부만 확인하는거니까

