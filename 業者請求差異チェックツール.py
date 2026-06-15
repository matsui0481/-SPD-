import pandas as pd
import tkinter as tk
from tkinter import ttk
import os

def calculate():
        # 集計中メッセージを表示
    result_label.config(text="現在集計中...", foreground="blue")
    root.update_idletasks()  # GUIを更新してメッセージを表示

    # 選択されたファイル名を取得
    file_name = file_var.get()
    if  file_name == "ムトウ四国":
        excel_path = "ムトウ四国.xlsx"
        # 'ムトウ四国.xlsx'を読み込む
        df = pd.read_excel("ムトウ四国.xlsx", skiprows=11)  # 最初の11行をスキップ

        # 'お客様商品ｺｰﾄﾞ'でグループ化し、'数量'と'売上金額'の合計を計算
        result = df.groupby('お客様商品ｺｰﾄﾞ').agg({'数量': 'sum', '金額': 'sum'}).reset_index()

        # 'お客様商品ｺｰﾄﾞ'の値から '.0' を削除
        result['お客様商品ｺｰﾄﾞ'] = result['お客様商品ｺｰﾄﾞ'].astype(str).replace(r'\.0', '', regex=True)

        # 売上金額のデータ型を浮動小数点数に変更し、売上金額を1.1倍に調整
        result['金額'] = (result['金額'].astype(float) * 1.1).astype(int)

        # 結果を新しいExcelファイルに保存
        result.to_excel("集計結果_ムトウ四国.xlsx", index=False)

        # 仕入れデータを読み込む
        purchase_df = pd.read_excel("仕入れデータ.xlsx")
        # '業者名'が'株式会社ﾒﾃﾞｨｶﾙｻｰﾋﾞｽ'のデータに絞り込む
        filtered_purchase = purchase_df[purchase_df['業者名'] == 'ﾑﾄｳ四国']
        # '原価管理CD'でグループ化し、'購入数'と'購入合価'の合計を計算
        purchase_summary = filtered_purchase.groupby('AptageCD').agg({'購入数': 'sum', '購入合価': 'sum', '材料名': 'first', '規格': 'first'}).reset_index()

        # '購入合価'のデータ型を浮動小数点数に変更
        purchase_summary['購入合価'] = purchase_summary['購入合価']

        # データ型を文字列に変更
        result['お客様商品ｺｰﾄﾞ'] = result['お客様商品ｺｰﾄﾞ'].astype(str)
        purchase_summary['AptageCD'] = purchase_summary['AptageCD'].astype(str)

        # 読み込んだ仕入れデータをExcelファイルとして出力
        purchase_summary.to_excel("集計結果_仕入れデータ_ムトウ四国.xlsx", index=False)

        # データをマージ
        merged_data = pd.merge(result, purchase_summary, left_on='お客様商品ｺｰﾄﾞ', right_on='AptageCD', how='outer')

        # merged_data が空でない場合のみ処理を実行
        if not merged_data.empty:
            # 条件に一致する行を削除する処理
            condition = (merged_data['お客様商品ｺｰﾄﾞ'] == merged_data['AptageCD'])& \
                         (merged_data['数量'] == merged_data['購入数'])& \
            (merged_data['金額'] == merged_data['購入合価'])
            filtered_data = merged_data[~condition].copy()
            filtered_data.to_excel(f"{file_var.get()}差異データ.xlsx", index=False)

        # 条件に一致する行を削除
        condition = (merged_data['お客様商品ｺｰﾄﾞ'] == merged_data['AptageCD']) & \
                    (merged_data['金額'] == merged_data['購入合価'])
        filtered_data = merged_data[~condition].copy()

        # 金額と購入合価の差を計算し、新しい列として追加
        filtered_data['差異金額'] = filtered_data['金額'] - filtered_data['購入合価']

        # 列を指定された順番に並び替える
        filtered_data = filtered_data[['お客様商品ｺｰﾄﾞ', '数量', '金額', 'AptageCD', '購入数', '購入合価', '差異金額', '材料名', '規格']]

        # 結果をExcelファイルに保存
        filtered_data.to_excel(f"{file_var.get()}差異データ.xlsx", index=False)

        os.remove("集計結果_ムトウ四国.xlsx")
        os.remove("集計結果_仕入れデータ_ムトウ四国.xlsx")

    elif file_name == "四国医療器":
        excel_path = "四国医療器.xlsx"
        # Excelファイルを読み込む
        df = pd.read_excel(excel_path)
        # '相手先品番'でグループ化し、'売上数量'と'売上金額'の合計を計算
        result = df.groupby('相手先品番').agg({'売上数量': 'sum', '売上金額': 'sum'}).reset_index()

        # 仕入れデータを読み込む
        purchase_df = pd.read_excel("仕入れデータ.xlsx")
        # '業者名'が'四国医療器'のデータに絞り込む
        filtered_purchase = purchase_df[purchase_df['業者名'] == '四国医療器']
        # '原価管理CD'でグループ化し、'購入数'と'購入合価'の合計を計算
        purchase_summary = filtered_purchase.groupby('原価管理CD').agg({'購入数': 'sum', '購入合価': 'sum', '材料名': 'first', '規格': 'first'}).reset_index()
    
        # データ型を文字列に変更し、'.0'を削除
        result['相手先品番'] = result['相手先品番'].astype(str).replace(r'\.0$', '', regex=True)
        purchase_summary['原価管理CD'] = purchase_summary['原価管理CD'].astype(str).replace(r'\.0$', '', regex=True)

        # 読み込んだ仕入れデータをExcelファイルとして出力
        purchase_summary.to_excel("仕入れデータ_四国医療器元データ.xlsx", index=False)

        # 条件分岐前に merged_data を空のDataFrameとして初期化
        merged_data = pd.DataFrame()

        # 四国医療器のデータと仕入れデータをマージ（全てのデータを保持）
        merged_data = pd.merge(result, purchase_summary, left_on='相手先品番', right_on='原価管理CD', how='outer')

        # merged_data が空でない場合のみ処理を実行
        if not merged_data.empty:
            merged_data['売上金額'] = merged_data['売上金額'].astype(float)  # 売上金額は浮動小数点数として扱う
            # 条件に一致する行を削除する処理
            condition = (merged_data['相手先品番'] == merged_data['原価管理CD']) & \
                        (merged_data['売上数量'] == merged_data['購入数']) & \
                        (merged_data['売上金額'] == merged_data['購入合価'])
            filtered_data = merged_data[~condition]

            # 金額と購入合価の差を算し、新しい列として追加
            filtered_data['差異金額'] = filtered_data['売上金額'] - filtered_data['購入合価']

            # 列を指定された順番に並び替える
            filtered_data = filtered_data[['相手先品番', '売上数量', '売上金額', '原価管理CD', '購入数', '購入合価', '差異金額', '材料名', '規格']]

            # 結果をExcelファイルに保存
            filtered_data.to_excel("四国医療器差異データ.xlsx", index=False)

            os.remove("仕入れデータ_四国医療器元データ.xlsx")

        
    elif file_name == "メディカルサービス":
        excel_path = "メディカルサービス.xlsx"
        # 'メディカルサービス.xlsx'を読み込む
        df = pd.read_excel("メディカルサービス.xlsx")
        # '◆労災コード\n(原価管理CD)'でグループ化し、'数量'と'売上金額'の合計を計算
        result = df.groupby('◆労災コード').agg({'数量': 'sum', '売上金額': 'sum'}).reset_index()

        # '◆労災コード\n(原価管理CD)'の値から '.0' を削除
        result['◆労災コード'] = result['◆労災コード'].astype(str).replace(r'\.0', '', regex=True)

        # 売上金額のデータ型を浮動小数点数に変更し、売上金額を1.1倍に調整
        result['売上金額'] = (result['売上金額'].astype(float) * 1.1).astype(int)

        # 結果を新しいExcelファイルに保存
        result.to_excel("メディカルサービス集計結果.xlsx", index=False)

        # 仕入れデータを読み込む
        purchase_df = pd.read_excel("仕入れデータ.xlsx")
        # '業者名'が'株式会社ﾒﾃﾞｨｶﾙｻｰﾋﾞｽ'のデータに絞り込む
        filtered_purchase = purchase_df[purchase_df['業者名'] == '株式会社ﾒﾃﾞｨｶﾙｻｰﾋﾞｽ']
        # '原価管理CD'でグループ化し、'購入数'と'購入合価'の合計を計算
        purchase_summary = filtered_purchase.groupby('原価管理CD').agg({'購入数': 'sum', '購入合価': 'sum' , '材料名': 'first', '規格': 'first'}).reset_index()

        # '購入合価'のデータ型を浮動小数点数に変更
        purchase_summary['購入合価'] = purchase_summary['購入合価']

        # データ型を文字列に変更
        result['◆労災コード'] = result['◆労災コード'].astype(str)
        purchase_summary['原価管理CD'] = purchase_summary['原価管理CD'].astype(str)

        # 読み込んだ仕入れデータをExcelァイルとして出力
        purchase_summary.to_excel("仕入れデータ_メリット元データ.xlsx", index=False)

        # データをマージ
        merged_data = pd.merge(result, purchase_summary, left_on='◆労災コード', right_on='原価管理CD', how='outer')

        # merged_data が空でない場合のみ処理を実行
        if not merged_data.empty:
            # 条件に一致する行を削除する処理
            condition = (merged_data['◆労災コード'] == merged_data['原価管理CD']) & \
                        (merged_data['数量'] == merged_data['購入数']) & \
                        (merged_data['売上金額'] == merged_data['購入合価'])
            filtered_data = merged_data[~condition].copy()  # コピーを作成

            # 金額と購入合価の差を計算し、新しい列として追加
            filtered_data.loc[:, '差異金額'] = filtered_data['売上金額'] - filtered_data['購入合価']

            # 列を指定された順番に並び替える
            filtered_data = filtered_data[['◆労災コード', '数量', '売上金額', '原価管理CD', '購入数', '購入合価', '差異金額', '材料名', '規格']]

            # 結果をExcelファイルに保存
            filtered_data.to_excel("メディカルサービス差異データ.xlsx", index=False)

        # 仕入れデータ_メリット元データ.xlsxとメディカルサービス集計結果.xlsxを削除
        os.remove("仕入れデータ_メリット元データ.xlsx")
        os.remove("メディカルサービス集計結果.xlsx")

    else:
        return
 
    # 完了メッセージを表示
    # result_label.config(text=f"集計結果が保存されました: 集計結果_{file_name}.xlsx")
    result_label.config(text=f"集計結果が保存されました: 集計結果_{file_name}.xlsx", foreground="green")

# GUIの設定
root = tk.Tk()
root.title("Excelファイル集計ツール")

# ウィンドウサイズの設定
root.geometry("400x200")

# フレームの設定
frame = ttk.Frame(root, padding="10")
frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

# ドロップダウンメニューの設定
file_var = tk.StringVar()
file_choices = ['ムトウ四国', '四国医療器', 'メディカルサービス']
file_var.set(file_choices[0])  # デフォルト値
file_label = ttk.Label(frame, text="メーカーを選択:")
file_label.grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
file_dropdown = ttk.Combobox(frame, textvariable=file_var, values=file_choices)
file_dropdown.grid(row=0, column=1, padx=5, pady=5, sticky=(tk.W, tk.E))

# 集計ボタンの設定
calculate_button = ttk.Button(frame, text="集計開始", command=calculate)
calculate_button.grid(row=1, column=0, columnspan=2, pady=10)

# 結果表示用ラベル
result_label = ttk.Label(frame, text="", foreground="green")
result_label.grid(row=2, column=0, columnspan=2, pady=5)

# レスポンシブデザインの設定
for child in frame.winfo_children():
    child.grid_configure(padx=5, pady=5)

root.mainloop()