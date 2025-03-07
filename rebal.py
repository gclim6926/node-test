import tkinter as tk
from tkinter import ttk
from datetime import datetime
import yfinance as yf
import time
import pandas as pd

# 오늘 날짜를 기본값으로 설정
today_date = datetime.today().strftime('%Y-%m-%d')

# 기본 대기 시간 (초)
sleep_time = 0.3

def fetch_data():
    end_date = entry_date.get()
    years = int(entry_years.get())
    tickers = [entry_ticker_a.get(), entry_ticker_b.get(), entry_ticker_c.get()]
    balancing = int(entry_balancing.get())

    start_date = datetime.strptime(end_date, '%Y-%m-%d').replace(year=datetime.strptime(end_date, '%Y-%m-%d').year - years)
    start_date = start_date.strftime('%Y-%m-%d')
    
    # Text 위젯 초기화
    text_output.delete(1.0, tk.END)
    
    # 각 티커에 대한 데이터를 변수에 저장
    text_output.insert(tk.END, f"loading period: from{start_date} to {end_date}\n\n")
    text_output.insert(tk.END, f"loading.... ticker {tickers[0]} \n")
    stock_data_a = yf.download(tickers[0], start=start_date, end=end_date)
    time.sleep(sleep_time)

    text_output.insert(tk.END, f"loading.... ticker {tickers[1]} \n")
    stock_data_b = yf.download(tickers[1], start=start_date, end=end_date)
    time.sleep(sleep_time)
    
    text_output.insert(tk.END, f"loading.... ticker {tickers[2]} \n")
    stock_data_c = yf.download(tickers[2], start=start_date, end=end_date)
    time.sleep(sleep_time)

    # stock_df 데이터프레임 생성
    stock_df = pd.DataFrame({
        'date': stock_data_a.index,
        tickers[0]: stock_data_a['Close'].values,
        tickers[1]: stock_data_b['Close'].values,
        tickers[2]: stock_data_c['Close'].values
    })

    # 초기 투자금 컬럼 생성
    stock_df['초기투자금_A'] = 50000
    stock_df['초기투자금_B'] = 25000
    stock_df['초기투자금_C'] = 25000
    stock_df['A+B+C합계'] = stock_df['초기투자금_A'] + stock_df['초기투자금_B'] + stock_df['초기투자금_C']

    for i in range(1, len(stock_df)):
        stock_df.at[i, '초기투자금_A'] = stock_df.at[i-1, '초기투자금_A'] * (stock_df.at[i, tickers[0]] / stock_df.at[i-1, tickers[0]])
        stock_df.at[i, '초기투자금_B'] = stock_df.at[i-1, '초기투자금_B'] * (stock_df.at[i, tickers[1]] / stock_df.at[i-1, tickers[1]])
        stock_df.at[i, '초기투자금_C'] = stock_df.at[i-1, '초기투자금_C'] * (stock_df.at[i, tickers[2]] / stock_df.at[i-1, tickers[2]])
        stock_df.at[i, 'A+B+C합계'] = stock_df.at[i, '초기투자금_A'] + stock_df.at[i, '초기투자금_B'] + stock_df.at[i, '초기투자금_C']

    # 리밸런싱 컬럼 생성
    stock_df['리발_A'] = 50000
    stock_df['리발_B'] = 25000
    stock_df['리발_C'] = 25000
    stock_df['비율'] = 0
    stock_df['리발유무'] = ''
    stock_df['리발A+B+C합계'] = 100000

    for i in range(1, len(stock_df)):
        stock_df.at[i, '리발_A'] = stock_df.at[i-1, '리발_A'] * (stock_df.at[i, tickers[0]] / stock_df.at[i-1, tickers[0]])
        stock_df.at[i, '리발_B'] = stock_df.at[i-1, '리발_B'] * (stock_df.at[i, tickers[1]] / stock_df.at[i-1, tickers[1]])
        stock_df.at[i, '리발_C'] = stock_df.at[i-1, '리발_C'] * (stock_df.at[i, tickers[2]] / stock_df.at[i-1, tickers[2]])
        stock_df.at[i, '리발A+B+C합계'] = stock_df.at[i, '리발_A'] + stock_df.at[i, '리발_B'] + stock_df.at[i, '리발_C']

        stock_df.at[i, '비율'] = (stock_df.at[i, '리발_A'] / (stock_df.at[i, '리발_B'] + stock_df.at[i, '리발_C'])) * 100
        
        if stock_df.at[i, '비율'] > 100 + balancing:
            stock_df.at[i, '리발유무'] = '상승리발'
        elif stock_df.at[i, '비율'] < 100 - balancing:
            stock_df.at[i, '리발유무'] = '하락리발'
        else:
            stock_df.at[i, '리발유무'] = ''

        if stock_df.at[i, '리발유무'] in ['상승리발', '하락리발'] and i+1 < len(stock_df):
            total_value = stock_df.at[i, '리발A+B+C합계']
            stock_df.at[i, '리발_A'] = total_value * 0.50
            stock_df.at[i, '리발_B'] = total_value * 0.25
            stock_df.at[i, '리발_C'] = total_value * 0.25

    # 결과를 창에 출력
    text_output.insert(tk.END, "Stock DataFrame:\n")
    text_output.insert(tk.END, stock_df.head().to_string())
    text_output.insert(tk.END, "\n\n")

    # 엑셀 파일로 저장
    now = datetime.now().strftime('%Y%m%d_%H%M%S')
    filename = f'리발_결과_{now}.xlsx'
    stock_df.to_excel(filename, index=False)
    text_output.insert(tk.END, f"Excel 파일로 저장되었습니다: {filename}\n\n")

    # 최종 값 출력
    final_sum_A_B_C = stock_df['A+B+C합계'].iloc[-1]
    final_rebal_sum_A_B_C = stock_df['리발A+B+C합계'].iloc[-1]
    rate_of_return = (final_rebal_sum_A_B_C / final_sum_A_B_C) * 100

    text_output.insert(tk.END, f"A+B+C합계 최종 값: {final_sum_A_B_C}\n")
    text_output.insert(tk.END, f"리발A+B+C합계 최종 값: {final_rebal_sum_A_B_C}\n")
    text_output.insert(tk.END, f"상승률: {rate_of_return:.2f}%\n")

# Tkinter 윈도우 생성
root = tk.Tk()
root.title('교잡종 프로그램')

# 라벨과 입력 필드 추가
label_date = ttk.Label(root, text='결산일:')
label_date.grid(row=0, column=0, padx=5, pady=5)
entry_date = ttk.Entry(root)
entry_date.grid(row=0, column=1, padx=5, pady=5)
entry_date.insert(0, today_date)

label_years = ttk.Label(root, text='몇 년치?:')
label_years.grid(row=1, column=0, padx=5, pady=5)
entry_years = ttk.Entry(root)
entry_years.grid(row=1, column=1, padx=5, pady=5)
entry_years.insert(0, '1')

label_ticker_a = ttk.Label(root, text='ticker_a(50%):')
label_ticker_a.grid(row=2, column=0, padx=5, pady=5)
entry_ticker_a = ttk.Entry(root)
entry_ticker_a.grid(row=2, column=1, padx=5, pady=5)
entry_ticker_a.insert(0, 'TQQQ')

label_ticker_b = ttk.Label(root, text='ticker_b(25%):')
label_ticker_b.grid(row=3, column=0, padx=5, pady=5)
entry_ticker_b = ttk.Entry(root)
entry_ticker_b.grid(row=3, column=1, padx=5, pady=5)
entry_ticker_b.insert(0, 'QQQ')

label_ticker_c = ttk.Label(root, text='ticker_c(25%):')
label_ticker_c.grid(row=4, column=0, padx=5, pady=5)
entry_ticker_c = ttk.Entry(root)
entry_ticker_c.grid(row=4, column=1, padx=5, pady=5)
entry_ticker_c.insert(0, 'SCHD')

label_balancing = ttk.Label(root, text='Balancing(%):')
label_balancing.grid(row=5, column=0, padx=5, pady=5)
entry_balancing = ttk.Entry(root)
entry_balancing.grid(row=5, column=1, padx=5, pady=5)
entry_balancing.insert(0, '20')

# 버튼 추가
button_fetch = ttk.Button(root, text='데이터 가져오기', command=fetch_data)
button_fetch.grid(row=6, column=0, columnspan=2, pady=10)

# Text 위젯 추가
text_output = tk.Text(root, wrap='word', width=80, height=20)
text_output.grid(row=7, column=0, columnspan=2, padx=5, pady=5)

# Tkinter 이벤트 루프 시작
root.mainloop()
