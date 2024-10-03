import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import os

def process_csv(file_path):
    """電圧が0（0～+100）になるすべての時間を取得する関数"""
    try:
        # CSVファイルの読み込み
        data = pd.read_csv(file_path, encoding='shift_jis')
        
        # D列（時間）とE列（電圧）の取得
        if data.shape[1] < 5:
            return file_path, "エラー: ファイルに必要な列が不足しています"

        time_column = data.iloc[:, 3]  # D列 (4列目)
        voltage_column = data.iloc[:, 4]  # E列 (5列目)

        # 電圧が 0 から +100 の範囲にある時間を取得
        near_zero_voltage_index = voltage_column[(voltage_column >= 0) & (voltage_column <= 100)].index

        if near_zero_voltage_index.empty:
            return file_path, "該当する電圧の範囲（0～100）が存在しません。"
        
        times_when_voltage_is_near_zero = time_column[near_zero_voltage_index].tolist()
        return file_path, ', '.join(map(str, times_when_voltage_is_near_zero))

    except pd.errors.EmptyDataError:
        return file_path, "エラー: ファイルが空です"
    except pd.errors.ParserError:
        return file_path, "エラー: CSVファイルの構文エラーが発生しました"
    except FileNotFoundError:
        return file_path, "エラー: ファイルが見つかりません"
    except IndexError:
        return file_path, "エラー: 指定された列が存在しません。ファイルの形式を確認してください"
    except Exception as e:
        return file_path, f"エラー: 予期しないエラーが発生しました\n詳細: {str(e)}"


def select_files():
    """ユーザーが複数のCSVファイルを選択して処理する関数"""
    try:
        file_paths = filedialog.askopenfilenames(filetypes=[("CSV Files", "*.csv")])
        if not file_paths:
            messagebox.showwarning("警告", "ファイルが選択されていません。")
            return

        results = [process_csv(file_path) for file_path in file_paths]
        save_results_to_file(results)

    except Exception as e:
        messagebox.showerror("エラー", f"ファイル選択中にエラーが発生しました: {str(e)}")


def save_results_to_file(results):
    """処理結果を指定したファイル形式で保存する関数"""
    try:
        file_type = filedialog.asksaveasfilename(defaultextension=".txt", 
                                                 filetypes=[("Text files", "*.txt"), 
                                                            ("CSV files", "*.csv"), 
                                                            ("Excel files", "*.xlsx")])
        if not file_type:
            messagebox.showwarning("警告", "保存先が指定されていません。")
            return

        ext = os.path.splitext(file_type)[1]
        if ext == ".txt":
            save_as_txt(file_type, results)
        elif ext == ".csv":
            save_as_csv(file_type, results)
        elif ext == ".xlsx":
            save_as_xlsx(file_type, results)
        else:
            messagebox.showerror("エラー", "対応していないファイル形式です。")
    
    except Exception as e:
        messagebox.showerror("エラー", f"結果の保存中にエラーが発生しました: {str(e)}")


def save_as_txt(file_path, results):
    """結果をテキストファイルに保存"""
    try:
        with open(file_path, 'w', encoding='utf-8') as f:
            for file_name, result in results:
                f.write(f"{file_name}: {result}\n")
        messagebox.showinfo("成功", f"結果が {file_path} に保存されました。")
    except Exception as e:
        messagebox.showerror("エラー", f"テキストファイルの保存中にエラーが発生しました: {str(e)}")


def save_as_csv(file_path, results):
    """結果をCSVファイルに保存"""
    try:
        df = pd.DataFrame(results, columns=["ファイル名", "結果"])
        df.to_csv(file_path, index=False, encoding='utf-8-sig')
        messagebox.showinfo("成功", f"結果が {file_path} に保存されました。")
    except Exception as e:
        messagebox.showerror("エラー", f"CSVファイルの保存中にエラーが発生しました: {str(e)}")


def save_as_xlsx(file_path, results):
    """結果をExcelファイルに保存"""
    try:
        df = pd.DataFrame(results, columns=["ファイル名", "結果"])
        df.to_excel(file_path, index=False)
        messagebox.showinfo("成功", f"結果が {file_path} に保存されました。")
    except Exception as e:
        messagebox.showerror("エラー", f"Excelファイルの保存中にエラーが発生しました: {str(e)}")


# GUIを作成
root = tk.Tk()
root.title("CSV処理アプリ")
root.geometry("400x200")

# ラベル
label = tk.Label(root, text="複数のCSVファイルを選択してください")
label.pack(pady=20)

# ファイル選択ボタン
button = tk.Button(root, text="ファイル選択", command=select_files)
button.pack(pady=10)

# GUI起動
root.mainloop()
