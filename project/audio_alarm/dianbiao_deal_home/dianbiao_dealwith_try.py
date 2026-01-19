import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import threading

# --- 全局配置 ---
STATION_PREFIX = "1.1.11.3"
TARGET_COLUMNS = [
    '实际测点地址', '描述', '点号', '节点别名', '取反', '事件启动', '远动', '双节点', '双节点名',
    '测值', '输入', '虚拟点', '检修', '手动', '品质', '入历史库', '语音报警', '生值', '反值',
    '开入实测值', '事故追忆启动源', '远动投退', '远动测值', '事件处理方式', 'ACC测点名',
    '内部点号', '驱动内部点号', '序号', '板号', '驱动名称', '设备类型', '一览表',
    '1->0描述', '0->1描述', '0->1报警', '0->1登录', '0->1寻呼', '1->0报警',
    '1->0登录', '1->0寻呼', '上位机报警', '0->1语音号', '1->0语音号', '镜头号',
    '报警', '电话语音号', '电话报警', '屏蔽'
]

class SmartPointListApp:
    def __init__(self, root):
        self.root = root
        self.root.title("电力全兼容点表优化工具 v2.0")
        self.root.geometry("700x550")
        
        # 路径选择 UI
        tk.Label(root, text="原始文件:").pack(pady=(10, 0), anchor="w", padx=20)
        self.entry_input = tk.Entry(root, width=80); self.entry_input.pack(padx=20, pady=5)
        tk.Button(root, text="选择文件", command=self.select_file).pack(anchor="e", padx=20)

        tk.Label(root, text="保存目录:").pack(pady=(10, 0), anchor="w", padx=20)
        self.entry_output = tk.Entry(root, width=80); self.entry_output.pack(padx=20, pady=5)
        tk.Button(root, text="选择目录", command=self.select_dir).pack(anchor="e", padx=20)

        self.btn_run = tk.Button(root, text="识别并转换", bg="#2196F3", fg="white", height=2, width=20, command=self.start)
        self.btn_run.pack(pady=20)

        self.log_area = scrolledtext.ScrolledText(root, height=15, width=90); self.log_area.pack(padx=20)

    def select_file(self):
        p = filedialog.askopenfilename(); self.entry_input.delete(0, tk.END); self.entry_input.insert(0, p)
    def select_dir(self):
        p = filedialog.askdirectory(); self.entry_output.delete(0, tk.END); self.entry_output.insert(0, p)
    def log(self, m): self.log_area.insert(tk.END, m + "\n"); self.log_area.see(tk.END)

    def start(self): threading.Thread(target=self.main_logic).start()

    def main_logic(self):
        in_p = self.entry_input.get(); out_d = self.entry_output.get()
        if not in_p or not out_d: return messagebox.showwarning("警告", "请选好路径！")
        
        self.btn_run.config(state="disabled"); self.log_area.delete(1.0, tk.END)
        try:
            excel = pd.ExcelFile(in_p, engine='openpyxl' if '.xlsx' in in_p else 'xlrd')
            for sn in excel.sheet_names:
                df = pd.read_excel(excel, sheet_name=sn)
                self.log(f"正在分析: {sn}...")
                
                # 识别模式并处理
                res = self.adapter(df, sn)
                if res is not None:
                    res.to_csv(os.path.join(out_d, f"数据库模板_{sn}.csv"), index=False, encoding='utf-8-sig')
                    self.log(f"  [成功] 转换完成")
            messagebox.showinfo("完成", "处理成功！")
        except Exception as e: self.log(f"  [报错] {e}")
        finally: self.btn_run.config(state="normal")

    def adapter(self, df, sheet_name):
        # 1. 识别 PLC 模式 (针对“赣江尾闾”这种左右并列的表)
        if '寄存器地址' in str(df.columns.tolist()):
            self.log("  模式识别：PLC通讯地址表 (进行并列拆分)")
            # 拆分逻辑：取左边5列，再取右边5列，垂直合并
            part1 = df.iloc[:, 0:5].copy()
            part2 = df.iloc[:, 6:11].copy() # 假设中间有空列
            part2.columns = part1.columns
            df = pd.concat([part1, part2]).dropna(subset=[part1.columns[1]])
            name_col = part1.columns[1] # "Name"
            addr_col = part1.columns[4] # "modbus上送地址"
            return self.build_template(df, name_col, addr_col, "PLC")

        # 2. 识别硬件 LCU 模式
        elif '模块号' in df.columns:
            self.log("  模式识别：LCU硬接线表")
            df[['本侧盘柜', '模块号']] = df[['本侧盘柜', '模块号']].ffill()
            df = df.dropna(subset=['点名'])
            df = df[~df['点名'].astype(str).str.contains('备用|/')]
            return self.build_template(df, '点名', None, "LCU")

        # 3. 识别逻辑报警模式
        elif '描述' in df.columns:
            self.log("  模式识别：逻辑映射表")
            df = df.dropna(subset=['描述'])
            return self.build_template(df, '描述', None, "LOGIC")
        
        return None

    def build_template(self, df, name_col, addr_col, mode):
        count = len(df)
        res = pd.DataFrame(columns=TARGET_COLUMNS)
        res['描述'] = df[name_col].astype(str).values
        res['序号'] = range(count)
        
        # 地址生成算法
        if mode == "PLC":
            # 直接取 PLC 表里的 Modbus 地址
            res['实际测点地址'] = df[addr_col].fillna(0).astype(str).values
            res['ACC测点名'] = [f"PLC_PT_{i:03d}" for i in range(count)]
        elif mode == "LCU":
            res['实际测点地址'] = [f"{STATION_PREFIX}.{i}" for i in range(count)]
            res['ACC测点名'] = [f"KGZ_DIN{i:03d}" for i in range(count)]
        else: # LOGIC 模式
            res['实际测点地址'] = "0"
            res['虚拟点'] = "true"
            res['ACC测点名'] = [f"VIRT_PT_{i:03d}" for i in range(count)]

        # 统一步骤：填充默认值
        res['节点别名'] = 'COM1'
        for col in ['入历史库', '报警', '上位机报警']: res[col] = 'true'
        res = res.fillna(0.0)
        return res

if __name__ == "__main__":
    root = tk.Tk(); app = SmartPointListApp(root); root.mainloop()