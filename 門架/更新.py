import pandas as pd

# 讀取 Excel 檔案
file_path = '/workspaces/-/CQ861標-門樘烤漆顔色區分_更新.xlsx'
sheet_name = '工作表1'

# 讀取指定的工作表
df = pd.read_excel(file_path, sheet_name=sheet_name)

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
            # 標記已查詢
            df.at[index, '已查詢'] = '✔'
    else:
        print("找不到對應的門號。")

# 將更新後的資料寫回 Excel 檔案
df.to_excel(file_path, sheet_name=sheet_name, index=False)