import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Alignment
import io

# # 데이터 프레임으로 읽는 함수
def read_df(file) :
    df = pd.read_excel(file, engine='pyxlsb')

    # 엑셀 파일 불러오기 전, 필요한 탭만 추출하기 위해 탭의 리스트를 작성
    sheets= ["SUN", "MON", "TUE", "WED", "THUR", "FRI", "SAT"]

    # 엑셀 파일의 필요한 탭만 불러오기
    file_path = 'CornDog.xlsb'
    df_dict = pd.read_excel(file_path, engine='pyxlsb', sheet_name=sheets)

    # 불러온 파일은 딕셔너리 타입이므로 각각의 변수로 매핑
    dfs = []
    for df in df_dict.values() :
        dfs.append(df)
    
    return dfs


# # 데이터 프레임 처리 함수
def process_data(df) :

    # Columns 숫자로 변경
    rename_col = range(len(df.columns))
    df.columns = rename_col

    # # 필요 없는 데이터들 제거
    # Tempreture 제거
    df = df[df.iloc[:, 0] != 'Tempreture']

    # NAME 제거
    df = df[df.iloc[:, 0] != 'NAME']

    # 표 이름 제거
    df = df[df.iloc[:, 0] != 'ROLLER GRILL - SUNDAY']
    df = df[df.iloc[:, 0] != 'BURRITOS - SUNDAY']
    df = df[df.iloc[:, 0] != 'HOT TO GO - SUNDAY']
    df = df[df.iloc[:, 0] != 'DELI EXPRESS - SUNDAY']
    df = df[df.iloc[:, 0] != '0x7']

    # NaN값 제거
    df = df.dropna(subset=[df.columns[0]])

    # 표의 아래 통계 계산해주는 셀 제거
    df = df[df.iloc[:, 0] != '0x7']
    df = df[df.iloc[:, 0] != 'ROLLER GRIL\nHOURS WASTE %']
    df = df[df.iloc[:, 0] != 'BURRITOS\nHOURS WASTE %']
    df = df[df.iloc[:, 0] != 'PAPA PRIMOS\nHOURS WASTE %']
    df = df[df.iloc[:, 0] != 'DELI EXPRESS\nHOURS WASTE %']
    df = df[df.iloc[:, 0] != 'TOTAL SUNDAY\nHOURS WASTE %']

    # ITEM # 제거
    df = df.drop(columns = [1, 50, 51, 52], axis=1)

    # Columns 변동되었으므로 한번 더 변경
    rename_col = range(len(df.columns))
    df.columns = rename_col

    # 셀에 값이 입력되면 생기는 값이 숫자인 행을 제거
    df = df[~df.iloc[:, 0].apply(lambda x: isinstance(x, (int, float)) or str(x).isdigit())]

    # 인덱스 재설정
    df.reset_index(drop=True, inplace=True)

    return df


# # TOTAL WEEK BY DAY 계산 함수 (같은 시간으로 계산)
# def cal_total_week_by_day(df) :
#     # 시간별 DISPOSAL, PUT, WASTE 계산
#     time_put = [0 for i in range(24)]
#     time_disposal = [0 for i in range(24)]
#     time_waste = [0 for i in range(24)]

#     for i in range(len(df)):
#         for j in range(1, len(df.columns)):
#             if pd.notna(df.iloc[i, j]):
#                 hour = (j - 1) // 2
#                 if (j % 2) == 1:  # 홀수 열 처리 (즉, put 열)
#                     time_put[hour] += int(df.iloc[i, j])
#                 else:  # 짝수 열 처리 (즉, disposal 열)
#                     time_disposal[hour] += int(df.iloc[i, j])

#     # time_waste 계산 (백분율로, 소수점 첫째 자리까지)
#     for hour in range(24):
#         if time_put[hour] != 0:
#             time_waste[hour] = round((time_disposal[hour] / time_put[hour]) * 100, 1)
#         else:
#             time_waste[hour] = 0  # put이 0인 경우 waste도 0으로 설정

#     # total_waste 계산 (백분율로, 소수점 첫째 자리까지)
#     if sum(time_put) != 0:
#         total_waste = round((sum(time_disposal) / sum(time_put)) * 100, 1)
#     else:
#         total_waste = 0  # time_put의 합이 0인 경우 waste도 0으로 설정

#     return [time_disposal, time_put, time_waste, total_waste]


# # TOTAL WEEK BY DAY 계산 함수 (4시간 전 시간으로 계산)
def cal_total_week_by_day(df):
    # 시간별 DISPOSAL, PUT, WASTE 계산
    time_put = [0 for i in range(24)]
    time_disposal = [0 for i in range(24)]
    time_waste = ['0%' for i in range(24)]  # 문자열로 초기화

    for i in range(len(df)):
        for j in range(1, len(df.columns)):
            if pd.notna(df.iloc[i, j]):
                hour = (j - 1) // 2
                if (j % 2) == 1:  # 홀수 열 처리 (즉, put 열)
                    time_put[hour] += int(df.iloc[i, j])
                else:  # 짝수 열 처리 (즉, disposal 열)
                    time_disposal[hour] += int(df.iloc[i, j])

    # time_waste 계산 (백분율로, 소수점 첫째 자리까지, '%' 기호 추가)
    for hour in range(24):
        if hour >= 4:
            if time_put[hour - 4] != 0:
                waste_value = round((time_disposal[hour] / time_put[hour - 4]) * 100, 1)
                time_waste[hour] = f"{waste_value}%"
            else:
                time_waste[hour] = "0%"  # 4시간 전 put이 0인 경우 waste도 0%로 설정
        else:
            time_waste[hour] = "0%"  # 0시부터 3시까지는 이전 날의 데이터가 필요하므로 0%로 설정

    # total_waste 계산 (백분율로, 소수점 첫째 자리까지, '%' 기호 추가)
    total_disposal = sum(time_disposal[4:])  # 4시부터 23시까지의 disposal 합계
    total_put = sum(time_put[:20])  # 0시부터 19시까지의 put 합계
    if total_put != 0:
        total_waste = f"{round((total_disposal / total_put) * 100, 1)}%"
    else:
        total_waste = "0%"  # put의 합이 0인 경우 waste도 0%로 설정

    return [time_disposal, time_put, time_waste, total_waste]


# # TOTAL WEEK BY DAY 엑셀 파일로 변환하는 함수
def create_excel_file(waste_by_day):
    wb = Workbook()
    ws = wb.active
    ws.title = "Weekly Waste Log"

    days = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday']
    
    for i, day in enumerate(days):
        start_col = 1 + i * 5  # 각 요일마다 5열씩 사용 (Hour, Disposal, Put, Waste (%), 빈열)

        # Write the day name
        ws.cell(row=1, column=start_col, value=day)
        ws.merge_cells(start_row=1, start_column=start_col, end_row=1, end_column=start_col + 3)
        ws.cell(row=1, column=start_col).alignment = Alignment(horizontal="center")

        time_disposal, time_put, time_waste, total_waste = waste_by_day[i]
        
        df = pd.DataFrame({
            'Hour': range(1, 25),
            'Disposal': time_disposal,
            'Put': time_put,
            'Waste (%)': time_waste
        })

        for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), start=2):
            for c_idx, value in enumerate(row, start=start_col):
                ws.cell(row=r_idx, column=c_idx, value=value)
                ws.cell(row=r_idx, column=c_idx).alignment = Alignment(horizontal="center")
        
        # Write the total waste for the day
        ws.cell(row=r_idx + 1, column=start_col, value='Total Waste (%)')
        ws.cell(row=r_idx + 1, column=start_col + 3, value=total_waste)
        ws.cell(row=r_idx + 1, column=start_col).alignment = Alignment(horizontal="center")
        ws.cell(row=r_idx + 1, column=start_col + 3).alignment = Alignment(horizontal="center")
    
    # Save to a BytesIO object
    excel_file = io.BytesIO()
    wb.save(excel_file)
    excel_file.seek(0)
    
    return excel_file



## # 메인 화면

st.title("A detailed summary of WASTE LOG")
st.subheader("For Troop Mini Mall", divider=True)

file = st.file_uploader("Upload excel file", type=["xlsb"])

# 파일이 업로드된 경우 데이터 프레임으로 읽기
if file :
    # 데이터 프레임 읽기
    dfs = read_df(file)

    # 각 데이터 프레임 처리
    for i in range(7) :
        dfs[i] = process_data(dfs[i])
    
    # TOTAL WEEK BY DAY 계산
    waste_by_day = []
    for i in range(7):
        waste_by_day.append(cal_total_week_by_day(dfs[i]))
    
    # TOTAL WEEK BY DAY 엑셀 파일로 변환
    excel_file = create_excel_file(waste_by_day)

    # 다운로드 버튼 생성
    st.download_button(
        label="Download detailed TOTAL WEEK BY DAY Excel file",
        data=excel_file,
        file_name="waste_log_summary.xlsx",  # .xlsx로 변경
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"  # MIME 타입 변경
    )

    # TOTAL WEEK BY DAY 데이터 미리보기
    st.subheader("Preview")
    sunday_data = waste_by_day[0]
    sunday_df = pd.DataFrame({
        'Hour': range(1, 25),
        'Disposal': sunday_data[0],
        'Put': sunday_data[1],
        'Waste (%)': sunday_data[2]
    })
    st.table(sunday_df)
