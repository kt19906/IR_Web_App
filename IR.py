import streamlit as st
import pandas as pd
from bokeh.plotting import figure
import xlsxwriter
import base64
from io import BytesIO


####様々な関数を定義####
#エクセルにインポートする関数
def to_excel(df):
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df.to_excel(writer, sheet_name='sheet1',index = False)
    workbook  = writer.book
    worksheet = writer.sheets['sheet1']
    chart = workbook.add_chart({'type': 'scatter', 'subtype': 'smooth'})#散布図・スムージング

    for index in range(1,len(files)+1):
        chart.add_series({'categories' : ["sheet1", 1, 0, 3551, 0], #線の設定A2からA3552を選択
                          'values' : ['sheet1', 1, index, 3551, index], 
                          'line' : {'color': 'black', 'width': 1},
                          'name' :["sheet1", 0, index],
                          })

    chart.set_size({'width': width, 'height': height}) #外枠のサイズ設定

    chart.set_x_axis({#x軸の設定
        'name': "Wave Number / cm  ̄¹",
        'name_font': {'size': 11},
        'min': min, 'max': max,
        'minor_unit': 100, 'major_unit': 500,
        'major_tick_mark': 'inside',
        'minor_tick_mark': 'inside',
        'line': {'color': 'black','width': 1},
        'reverse': True,
        })

    chart.set_y_axis({'visible': False,#y軸の設定
                      'major_gridlines': {'visible': False},
                      'min': y_min, 'max': y_max,
                      'crossing': y_min,
                      })

    chart.set_legend({'none': True})#凡例の設定

    if frame_border:
        chart.set_plotarea({#内枠の設定
            'border': {'color': 'black', 'width': 1},
            })

    chart.set_chartarea({#外枠の設定
        'border': {'none': True},
        })

    worksheet.insert_chart('H1', chart)#グラフのおく場所
    writer.save()
    processed_data = output.getvalue()
    return processed_data
#ダウンロードリンクを出す
def get_table_download_link(df):
    """Generates a link allowing the data in a given panda dataframe to be downloaded
    in:  dataframe
    out: href string
    """
    val = to_excel(df)
    b64 = base64.b64encode(val)  # val looks like b'...'
    return f'<a href="data:application/octet-stream;base64,{b64.decode()}" download="IR.xlsx">ここをクリック</a>' # decode b'abc' => abc
#ファイルから波長を取り出す
def Extraction(file):
  df = pd.read_csv(file,skiprows=1,nrows=3551)
  df = df["%T"]
  return df
#引数に波数とし，規格化係数を算出する
def Standardization(wave_max,wave_min):
  standards = ys.iloc[4000-wave_max:4001-wave_min]
  standards = pd.DataFrame.min(standards)
  standards = (baselines - standards)*0.01
  return standards


files = st.sidebar.file_uploader("ファイルを選択",accept_multiple_files = True)
#####生データ読み込み####
#吸光度をDataFrame化
ys = pd.DataFrame()
file_names = []
for file in files:
  y = Extraction(file)
  ys = pd.concat([ys, y], axis=1)
  file_names.append(file.name.split('.', 1)[0])
ys.columns = file_names
#波数をDataFrame化
x = pd.DataFrame(range(4000,449,-1), columns=["cm-1"])
#波数と吸光度を一つにしたDataFrameを作成(生データをdf0に保存)
df0 = pd.concat([x,ys],axis=1)

####エクスパンダ―#####
expander_stand = st.sidebar.beta_expander("規格化")
standardization = expander_stand.radio(
                            "規格化条件を選択",
                            ('規格化なし', 'ピンポイント規格化', '範囲規格化'))
if standardization == 'ピンポイント規格化':
    wave_number = expander_stand.number_input("波数",value = 3500,step=1) 
elif standardization == '範囲規格化':
    wave_number_max = expander_stand.number_input("波数(最大値)",value = 4000,step=1)
    wave_number_min = expander_stand.number_input("波数(最小値)",value = 500,step=1)

expander_spc = st.sidebar.beta_expander("スペクトル間隔")
expander_spc.markdown("## 全体")
interval = expander_spc.slider('スペクトル間隔', 0, 200, 40)

expander_spc.markdown("## 個別")
individuals = []
for file_name in file_names:
    individual = expander_spc.slider(file_name, -100, 100, 0)
    individuals.append(individual)

expander_prop = st.sidebar.beta_expander("プロパティー")
expander_prop.markdown("## 軸範囲")
max = expander_prop.number_input("最大値", value = 4000,step = 100)
min = expander_prop.number_input("最小値", value = 500,step = 100)
expander_prop.markdown("## サイズ")
height = expander_prop.number_input("高さ", value = 500,step = 25)
width = expander_prop.number_input("幅", value = 450,step = 25)
frame_border = expander_prop.checkbox('枠線(枠線付きで保存される)')

####ファイルが選択されたら実行される####
if len(files) > 0:
    ####加工####
    #ベースライン
    baselines = ys.iloc[0]

    #全体
    if interval != 0 :
        intervals = list(range(0,len(files)*interval,interval))
    else:
        intervals = 0

    #個別(すでに##エクスパンダ―でリスト化済み)

    #規格化
    if standardization == '規格化なし':
        standards = 1
    elif standardization == 'ピンポイント規格化':
        standards = Standardization(wave_number,wave_number)
    elif standardization == '範囲規格化':
        standards = Standardization(wave_number_max,wave_number_min)

    #計算
    ys = (ys -baselines) / standards - intervals + individuals
    df = pd.concat([x,ys],axis=1)

    

    ####グラフ描写####
    y_min = round(ys.min()[len(files)-1], -1)-10 #yの最小値を計算
    y_max = round(ys.max()[0], -1)+20
    p = figure(
            title= "Stacked IR spectrum",
            x_axis_label = 'Wave Number / cm  ̄¹',
            y_axis_label = '',
            x_range = [max, min],
            y_range = [y_min,y_max],
            plot_width = width , 
            plot_height = height)


    for file_name in file_names: 
        p.line(x=x["cm-1"],y=ys[file_name],
               line_width = 1,
               line_color = "black",
               legend = file_name)

    p.legend.click_policy = "hide"

    if frame_border:
        p.outline_line_color = "black"

    st.bokeh_chart(p)

    #エクセルファイルをダウンロード
    st.sidebar.markdown("**エクセルファイルをダウンロード**")
    st.sidebar.markdown(get_table_download_link(df), unsafe_allow_html=True)

####使い方####
'''
# 使い方
1. "ファイルを選択"をクリックし，ドラックしてCSVファイルを読み込む   
    (一気にアップロードすると，グラフの順がおかしくなる場合がある．ドラックして一個ずつアップロードした方が確実．
    その場合，ファイルを閉じるとき，「開く」でなく，右上の「×」をクリック)

2. 規格化する場合  
    **ピンポイント規格化**   
    記入した波数の位置で規格化される   
    **範囲規格化**   
    記入した波数の範囲から最も吸光度の低い位置を見つけて規格化する 

3. 保存   
    エクセルファイルを開いた状態で保存することはできない
'''





