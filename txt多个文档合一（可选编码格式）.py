import os
import tkinter as tk
from tkinter import filedialog, messagebox


def read_txt_auto_encoding(file_path):
    """自动尝试 UTF-8 → GBK 读取文本"""
    try:
        with open(file_path, "r", encoding="utf-8") as f:
            return f.read()
    except UnicodeDecodeError:
        with open(file_path, "r", encoding="gbk") as f:
            return f.read()


def merge_txt_files(input_files, output_file, output_encoding):
    """合并多个 txt 文件"""
    first_file = True

    with open(output_file, "w", encoding=output_encoding) as outfile:
        for file_path in input_files:
            if not os.path.isfile(file_path):
                continue

            content = read_txt_auto_encoding(file_path).rstrip("\n")

            if not first_file:
                outfile.write("\n")  # 正常换行，不插空行

            outfile.write(content)
            first_file = False


# ===== GUI 逻辑 =====

def choose_input_files():
    paths = filedialog.askopenfilenames(
        title="选择 TXT 文件（可多选 Ctrl / Shift）",
        filetypes=[("TXT 文件", "*.txt")]
    )
    if paths:
        input_files_text.delete("1.0", tk.END)
        input_files_text.insert(tk.END, "\n".join(paths))


def choose_output_dir():
    path = filedialog.askdirectory()
    if path:
        output_dir_var.set(path)


def start_merge():
    input_text = input_files_text.get("1.0", tk.END).strip()
    output_dir = output_dir_var.get()
    output_name = output_name_var.get().strip()
    encoding_choice = encoding_var.get()

    if not input_text or not output_dir or not output_name:
        messagebox.showerror("错误", "请完整选择输入 TXT、输出路径和文件名")
        return

    input_files = [p.strip() for p in input_text.splitlines() if p.strip()]

    if not input_files:
        messagebox.showerror("错误", "未选择任何 TXT 文件")
        return

    if not output_name.lower().endswith(".txt"):
        output_name += ".txt"

    output_encoding = "utf-8" if encoding_choice == "UTF-8" else "gbk"
    output_file = os.path.join(output_dir, output_name)

    try:
        merge_txt_files(input_files, output_file, output_encoding)
        messagebox.showinfo("完成", "TXT 合并完成！")
    except Exception as e:
        messagebox.showerror("异常", str(e))


# ===== 窗口 =====

window = tk.Tk()
window.title("多个 TXT 合并为一个 TXT 工具（支持多选） by LIV")
window.geometry("650x360")
window.resizable(False, False)

output_dir_var = tk.StringVar()
output_name_var = tk.StringVar()
encoding_var = tk.StringVar(value="ANSI")

# ===== 布局 =====

tk.Label(
    window,
    text="输入 TXT 文件\n（可多选 Ctrl / Shift）："
).grid(row=0, column=0, padx=15, pady=15, sticky="ne")

input_files_text = tk.Text(window, width=45, height=4)
input_files_text.grid(row=0, column=1)

tk.Button(
    window,
    text="选择",
    width=10,
    command=choose_input_files
).grid(row=0, column=2, padx=10)

tk.Label(window, text="输出 TXT 文件夹路径：").grid(
    row=1, column=0, padx=15, pady=15, sticky="e"
)
tk.Entry(window, textvariable=output_dir_var, width=45).grid(row=1, column=1)
tk.Button(window, text="选择", width=10, command=choose_output_dir)\
    .grid(row=1, column=2, padx=10)

tk.Label(window, text="输出 TXT 文件名：").grid(
    row=2, column=0, padx=15, pady=15, sticky="e"
)
tk.Entry(window, textvariable=output_name_var, width=45).grid(row=2, column=1)

tk.Label(window, text="输出编码格式：").grid(
    row=3, column=0, padx=15, pady=15, sticky="e"
)
tk.Radiobutton(
    window, text="ANSI (GBK)",
    variable=encoding_var, value="ANSI"
).grid(row=3, column=1, sticky="w")

tk.Radiobutton(
    window, text="UTF-8",
    variable=encoding_var, value="UTF-8"
).grid(row=3, column=1, padx=120, sticky="w")

tk.Button(
    window,
    text="开始合并",
    width=20,
    height=2,
    command=start_merge
).grid(row=4, column=1, pady=25)

window.mainloop()
