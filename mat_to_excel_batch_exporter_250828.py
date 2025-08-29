import scipy.io
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, MULTIPLE
import os

def load_mat_variables(mat_path):
    data = scipy.io.loadmat(mat_path)
    return {k: v for k, v in data.items() if not k.startswith("__")}

def export_selected_variables_batch(folder_path, selected_vars, save_folder):
    mat_files = [f for f in os.listdir(folder_path) if f.lower().endswith(".mat")]
    if not mat_files:
        messagebox.showerror("エラー", "指定フォルダに.matファイルが存在しません。")
        return

    for mat_file in mat_files:
        mat_path = os.path.join(folder_path, mat_file)
        variables = load_mat_variables(mat_path)
        combined_data = {}
        lengths = []

        for var_display in selected_vars:
            var = var_display.split(" (")[0]  # 元の変数名を抽出
            if var in variables:
                array = variables[var]
                try:
                    flat_array = array.flatten()
                    combined_data[var] = flat_array
                    lengths.append(len(flat_array))
                except Exception as e:
                    print(f"{mat_file} の変数 {var} の処理中にエラー: {e}")

        if not combined_data:
            print(f"{mat_file} にエクスポート可能な変数がありません。")
            continue

        if len(set(lengths)) != 1:
            print(f"{mat_file} の変数の長さが一致しません。スキップされます。")
            continue

        df = pd.DataFrame(combined_data)
        excel_name = os.path.splitext(mat_file)[0] + "_export.xlsx"
        save_path = os.path.join(save_folder, excel_name)
        df.to_excel(save_path, sheet_name="ExportedVariables", index=False)
        print(f"保存しました: {save_path}")

    messagebox.showinfo("完了", f"{len(mat_files)} 個のファイルを処理しました。")

def batch_process():
    folder_path = filedialog.askdirectory(title="MATファイルが入っているフォルダを選択")
    if not folder_path:
        return

    example_mat = None
    for f in os.listdir(folder_path):
        if f.lower().endswith(".mat"):
            example_mat = os.path.join(folder_path, f)
            break
    if not example_mat:
        messagebox.showerror("エラー", "フォルダ内に.matファイルが見つかりません。")
        return

    variables = load_mat_variables(example_mat)
    var_names = [f"{k} {v.shape}" for k, v in variables.items()]

    select_win = tk.Toplevel(root)
    select_win.title("エクスポートする変数を選択")
    listbox = tk.Listbox(select_win, selectmode=MULTIPLE, width=60, height=20)
    listbox.pack(padx=10, pady=10)
    for name in var_names:
        listbox.insert(tk.END, name)

    def on_export():
        selected_indices = listbox.curselection()
        selected_vars = [var_names[i] for i in selected_indices]
        if not selected_vars:
            messagebox.showwarning("警告", "少なくとも1つの変数を選択してください。")
            return
        save_folder = filedialog.askdirectory(title="エクスポート先フォルダを選択")
        if not save_folder:
            return
        export_selected_variables_batch(folder_path, selected_vars, save_folder)
        select_win.destroy()

    export_btn = tk.Button(select_win, text="バッチエクスポート", command=on_export)
    export_btn.pack(pady=(0, 10))

root = tk.Tk()
root.title("MAT Batch Exporter")
root.geometry("300x150")

btn = tk.Button(root, text="MATフォルダを選んで一括エクスポート", command=batch_process)
btn.pack(expand=True)

root.mainloop()
