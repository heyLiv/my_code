import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import threading

# --- 核心逻辑配置 ---
TYPE_CONFIG = {
    'DI':  {'prefix': 'KGZ_DIN'},
    'SOE': {'prefix': 'KGZ_SOE'},
    'DO':  {'prefix': 'KGZ_DON'},
    'AI':  {'prefix': 'KGZ_AIN'},
    'RTD': {'prefix': 'KGZ_RTD'},
}
STATION_PREFIX = "1.1.11.3"

class PointListApp:
    def __init__(self, root):
        self.root = root
        self.root.title("点表自动化整理工具 v1.0")
        self.root.geometry("600x500")

        # --- 界面布局 ---
        # 1. 输入文件选择
        tk.Label(root, text="原始 Excel 文件:").pack(pady=(10, 0), anchor="w", padx=20)
        self.entry_input = tk.Entry(root, width=70)
        self.entry_input.pack(side="top", padx=20, pady=5)
        tk.Button(root, text="选择文件", command=self.select_file).pack(anchor="e", padx=20)

        # 2. 输出文件夹选择
        tk.Label(root, text="结果保存目录:").pack(pady=(10, 0), anchor="w", padx=20)
        self.entry_output = tk.Entry(root, width=70)
        self.entry_output.pack(side="top", padx=20, pady=5)
        tk.Button(root, text="选择文件夹", command=self.select_dir).pack(anchor="e", padx=20)

        # 3. 运行按钮
        self.btn_run = tk.Button(root, text="开始执行优化", bg="#4CAF50", fg="white", 
                                 height=2, width=20, command=self.start_task)
        self.btn_run.pack(pady=20)

        # 4. 日志显示区
        tk.Label(root, text="运行日志:").pack(anchor="w", padx=20)
        self.log_area = scrolledtext.ScrolledText(root, height=12, width=75)
        self.log_area.pack(padx=20, pady=10)

    def select_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if path:
            self.entry_input.delete(0, tk.END)
            self.entry_input.insert(0, path)

    def select_dir(self):
        path = filedialog.askdirectory()
        if path:
            self.entry_output.delete(0, tk.END)
            self.entry_output.insert(0, path)

    def log(self, message):
        self.log_area.insert(tk.END, message + "\n")
        self.log_area.see(tk.END)

    def start_task(self):
        # 使用线程运行逻辑，防止界面卡死
        thread = threading.Thread(target=self.run_logic)
        thread.start()

    def run_logic(self):
        input_path = self.entry_input.get()
        output_dir = self.entry_output.get()

        if not input_path or not output_dir:
            messagebox.showwarning("提示", "请先选择输入文件和输出目录！")
            return

        self.btn_run.config(state="disabled")
        self.log_area.delete(1.0, tk.END)
        self.log(">>> 开始任务...")

        try:
            excel_file = pd.ExcelFile(input_path, engine='openpyxl')
            for sheet_name in excel_file.sheet_names:
                if any(kw in sheet_name for kw in ['说明', '时钟', '网络', '串口', '目录']):
                    self.log(f"[排除] 系统 Sheet: {sheet_name}")
                    continue

                df_raw = pd.read_excel(excel_file, sheet_name=sheet_name)
                result_df = self.process_sheet_logic(sheet_name, df_raw)

                if result_df is not None:
                    out_path = os.path.join(output_dir, f"优化_{sheet_name}.csv")
                    result_df.to_csv(out_path, index=False, encoding='utf-8-sig')
                    self.log(f"[完成] 已导出: {sheet_name}")

            self.log("-" * 30)
            self.log("所有任务已完成！")
            messagebox.showinfo("成功", "点表处理完毕！")
        except Exception as e:
            self.log(f"[错误] 发生异常: {str(e)}")
            messagebox.showerror("错误", f"程序运行出错: {e}")
        finally:
            self.btn_run.config(state="normal")

    def process_sheet_logic(self, sheet_name, df_source):
        # 这里保留你之前的核心清洗逻辑
        cols_to_fill = [c for c in ['本侧盘柜', '模块号'] if c in df_source.columns]
        if cols_to_fill:
            df_source[cols_to_fill] = df_source[cols_to_fill].ffill()

        if '点名' not in df_source.columns:
            return None
        
        df_valid = df_source.dropna(subset=['点名']).copy()
        df_valid['点名'] = df_valid['点名'].astype(str)
        df_valid = df_valid[~df_valid['点名'].str.contains('备用|预留|/|#NAME', na=False)]
        
        if df_valid.empty: return None

        s_type = 'DI'
        for key in TYPE_CONFIG.keys():
            if key in sheet_name.upper():
                s_type = key
                break
        
        # 48列标准模板构造
        target_columns = [
            '实际测点地址', '描述', '点号', '节点别名', '取反', '事件启动', '远动', '双节点', '双节点名',
            '测值', '输入', '虚拟点', '检修', '手动', '品质', '入历史库', '语音报警', '生值', '反值',
            '开入实测值', '事故追忆启动源', '远动投退', '远动测值', '事件处理方式', 'ACC测点名',
            '内部点号', '驱动内部点号', '序号', '板号', '驱动名称', '设备类型', '一览表',
            '1->0描述', '0->1描述', '0->1报警', '0->1登录', '0->1寻呼', '1->0报警',
            '1->0登录', '1->0寻呼', '上位机报警', '0->1语音号', '1->0语音号', '镜头号',
            '报警', '电话语音号', '电话报警', '屏蔽'
        ]
        
        target_df = pd.DataFrame(columns=target_columns)
        count = len(df_valid)
        target_df['描述'] = df_valid['点名'].values
        target_df['序号'] = range(count)
        target_df['点号'] = target_df['序号'].astype(float)
        target_df['ACC测点名'] = [f"{TYPE_CONFIG[s_type]['prefix']}{i:03d}" for i in range(count)]
        target_df['实际测点地址'] = [f"{STATION_PREFIX}.{i}" for i in range(count)]
        
        # 填充默认值
        target_df['节点别名'] = 'COM1'
        for col in ['入历史库', '报警', '上位机报警', '0->1报警', '1->0报警']: target_df[col] = 'true'
        if s_type in ['DI', 'SOE', 'DO']:
            target_df['1->0描述'], target_df['0->1描述'] = '复归', '动作'
            
        for col in target_columns:
            if col not in target_df.columns or target_df[col].isnull().all():
                target_df[col] = 0.0
        return target_df

if __name__ == "__main__":
    root = tk.Tk()
    app = PointListApp(root)
    root.mainloop()