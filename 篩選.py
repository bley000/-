import pandas as pd

file_path = '/workspaces/-/CQ861標-門樘烤漆顔色區分.xlsx'
sheet_name = '工作表1'
new_file_path = '/workspaces/-/CQ861標-門樘烤漆顔色區分_更新.xlsx'

# 讀取指定的工作表
df = pd.read_excel(file_path, sheet_name=sheet_name)

# 檢查是否已經存在「已查詢」欄位，如果不存在則新增該欄位
if '已查詢' not in df.columns:
    df.insert(0, '已查詢', '')

# 刪除第一欄空白欄位
df = df.iloc[:, 1:]

# 設定欄位名稱
df.columns = ['已查詢', '門號', '門尺寸 (mm)', '門型', '開向', '防火時效', '五金組別', '備註']

# 新增「已查詢」欄位
if '已查詢' not in df.columns:
    df.insert(0, '已查詢', '')

while True:
    # 使用者輸入門號
    door_number = input("請輸入門號 (或輸入 '結束' 來退出): ")
    
    if door_number == '結束':
        break

    # 根據門號篩選資料
    filtered_data = df[df['門號'] == door_number]

    # 列印出對應的資料
    if not filtered_data.empty:
        for index, row in filtered_data.iterrows():
            print(f"門尺寸: {row['門尺寸 (mm)']}")
            print(f"門型: {row['門型']}")
            print(f"開向: {row['開向']}")
            print(f"防火時效: {row['防火時效']}")
            print(f"五金組別: {row['五金組別']}")
            print(f"備註: {row['備註']}")
            # 標記已查詢
            df.at[index, '已查詢'] = '✔'
    else:
        print("找不到對應的門號。")

# 將更新後的資料另存為新的 Excel 檔案
df.to_excel(new_file_path, sheet_name=sheet_name, index=False)