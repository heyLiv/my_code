import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import openpyxl
import os


# ================= 规则加载 =================

def load_replace_rules(rule_file):
    rules = {}
    with open(rule_file, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            if "=" in line:
                old, new = line.split("=", 1)
                rules[old] = new
    return rules


# ================= 核心处理逻辑 =================

def process_excel(input_path, output_path, rule_file, gui_prefix):

    # ---------- 加载替换规则 ----------
    replace_rules = load_replace_rules(rule_file)

    # ---------- 1. 用 openpyxl 找「名称列标红的行」 ----------
    wb = openpyxl.load_workbook(input_path)
    ws = wb.active

    red_rows = []
    for row in ws.iter_rows(min_row=2):  # 跳过表头
        cell = row[3]  # 第4列：名称
        font_color = cell.font.color
        if font_color and font_color.type == "rgb":
            rgb = font_color.rgb
            if rgb:
                rgb_str = rgb if isinstance(rgb, str) else rgb.rgb
                if rgb_str[-6:].upper() == "FF0000":
                    red_rows.append(cell.row)

    if not red_rows:
        raise ValueError("未找到任何「名称列为红色」的行")

    # ---------- 2. pandas 只保留这些行 ----------
    df = pd.read_excel(input_path)

    rows_idx = sorted(set(r - 2 for r in red_rows))  # Excel 行号 → df 行号
    df = df.iloc[rows_idx].reset_index(drop=True)

    # ---------- 必须列检查 ----------
    required_cols = {"序号", "前缀", "描述", "名称"}
    if not required_cols.issubset(df.columns):
        raise ValueError("Excel 必须包含列：序号、前缀、描述、名称")

    # ---------- 3. 拼接完整原始文本（先拼接，再替换） ----------
    df["原始文本"] = (
        gui_prefix +
        df["前缀"].astype(str) +
        df["名称"].astype(str)
    )

    # ---------- 4. 对整段文本做规则替换 ----------
    df["最终文本"] = (
        df["原始文本"]
        .replace(replace_rules, regex=True)
    )

    # ---------- 5. 描述加 .wav ----------
    df["描述_处理后"] = df["描述"].astype(str) + ".wav"

    # ---------- 6. 导出 txt ----------
    out_df = df[["序号", "最终文本", "描述_处理后"]]

    out_df.to_csv(
        output_path,
        sep=" ",
        index=False,
        header=False,
        encoding="utf-8"
    )


# ================= GUI =================

def choose_input_file():
    path = filedialog.askopenfilename(
        title="选择输入 Excel 文件",
        filetypes=[("Excel 文件", "*.xlsx")]
    )
    if path:
        entry_input.delete(0, tk.END)
        entry_input.insert(0, path)


def choose_output_dir():
    path = filedialog.askdirectory(title="选择输出文件夹")
    if path:
        entry_output_dir.delete(0, tk.END)
        entry_output_dir.insert(0, path)


def choose_rule_file():
    path = filedialog.askopenfilename(
        title="选择规则文件",
        filetypes=[("规则文件", "*.txt")]
    )
    if path:
        entry_rule.delete(0, tk.END)
        entry_rule.insert(0, path)


def run():
    input_path = entry_input.get().strip()
    output_dir = entry_output_dir.get().strip()
    output_name = entry_output_name.get().strip()
    rule_file = entry_rule.get().strip()
    gui_prefix = entry_prefix.get().strip()

    if not input_path:
        messagebox.showerror("错误", "请选择输入 Excel 文件")
        return
    if not output_dir:
        messagebox.showerror("错误", "请选择输出文件夹")
        return
    if not output_name:
        messagebox.showerror("错误", "请输入输出文件名称")
        return
    if not rule_file:
        messagebox.showerror("错误", "请选择规则文件")
        return

    if not output_name.lower().endswith(".txt"):
        output_name += ".txt"

    output_path = os.path.join(output_dir, output_name)

    try:
        process_excel(
            input_path=input_path,
            output_path=output_path,
            rule_file=rule_file,
            gui_prefix=gui_prefix
        )
        messagebox.showinfo(
            "完成",
            f"处理完成！\n\n生成文件：\n{output_path}"
        )
    except Exception as e:
        messagebox.showerror("失败", str(e))


# ================= 主窗口 =================

root = tk.Tk()
root.title("Excel → TXT 点表处理工具")
root.geometry("740x390")
root.resizable(False, False)

# 输入 Excel
tk.Label(root, text="输入 Excel 文件：").grid(row=0, column=0, padx=10, pady=12, sticky="e")
entry_input = tk.Entry(root, width=55)
entry_input.grid(row=0, column=1)
tk.Button(root, text="选择文件", command=choose_input_file).grid(row=0, column=2)

# 输出文件夹
tk.Label(root, text="输出文件夹：").grid(row=1, column=0, padx=10, pady=12, sticky="e")
entry_output_dir = tk.Entry(root, width=55)
entry_output_dir.grid(row=1, column=1)
tk.Button(root, text="选择文件夹", command=choose_output_dir).grid(row=1, column=2)

# 输出文件名
tk.Label(root, text="输出文件名称：").grid(row=2, column=0, padx=10, pady=12, sticky="e")
entry_output_name = tk.Entry(root, width=30)
entry_output_name.grid(row=2, column=1, sticky="w")
tk.Label(root, text="（可不写 .txt）").grid(row=2, column=2, sticky="w")

# 规则文件
tk.Label(root, text="规则文件：").grid(row=3, column=0, padx=10, pady=12, sticky="e")
entry_rule = tk.Entry(root, width=55)
entry_rule.grid(row=3, column=1)
tk.Button(root, text="选择规则", command=choose_rule_file).grid(row=3, column=2)

# GUI 前缀
tk.Label(root, text="额外前缀（参与替换）：").grid(row=4, column=0, padx=10, pady=12, sticky="e")
entry_prefix = tk.Entry(root, width=30)
entry_prefix.grid(row=4, column=1, sticky="w")

# 运行按钮
tk.Button(
    root,
    text="开始处理",
    width=22,
    height=2,
    command=run
).grid(row=5, column=1, pady=25)

root.mainloop()
