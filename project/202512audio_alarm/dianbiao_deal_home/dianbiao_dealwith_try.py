import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import threading

# --- 全局配置区 ---
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

# 不同信号类型的 ID 前缀定义
TYPE_CONFIG = {
    'DI':  {'prefix': 'KGZ_DIN'},
    'SOE': {'prefix': 'KGZ_SOE'},
    'DO':  {'prefix': 'KGZ_DON'},
    'AI':  {'prefix': 'KGZ_AIN'},
    'RTD': {'prefix': 'KGZ_RTD'},
}

class SmartPointListApp:
    def __init__(self, root):
        self.root = root
        self.root.title("电力全兼容点表自动化优化工具 v2.5 (支持固卷并列格式)")
        self.root.geometry("750x600")
        
        # UI 组件
        tk.Label(root, text="原始文件 (支持 .xlsx / .xls):").pack(pady=(10, 0), anchor="w", padx=20)
        self.entry_input = tk.Entry(root, width=85); self.entry_input.pack(padx=20, pady=5)
        tk.Button(root, text="选择文件", command=self.select_file).pack(anchor="e", padx=20)

        tk.Label(root, text="保存目录:").pack(pady=(10, 0), anchor="w", padx=20)
        self.entry_output = tk.Entry(root, width=85); self.entry_output.pack(padx=20, pady=5)
        tk.Button(root, text="选择目录", command=self.select_dir).pack(anchor="e", padx=20)

        self.btn_run = tk.Button(root, text="识别并转换", bg="#2196F3", fg="white", height=2, width=25, command=self.start)
        self.btn_run.pack(pady=20)

        self.log_area = scrolledtext.ScrolledText(root, height=18, width=95); self.log_area.pack(padx=20)

    def select_file(self):
        p = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
        self.entry_input.delete(0, tk.END); self.entry_input.insert(0, p)
    def select_dir(self):
        p = filedialog.askdirectory(); self.entry_output.delete(0, tk.END); self.entry_output.insert(0, p)
    def log(self, m): self.log_area.insert(tk.END, m + "\n"); self.log_area.see(tk.END)

    def start(self): threading.Thread(target=self.main_logic).start()

    def main_logic(self):
        in_p = self.entry_input.get(); out_d = self.entry_output.get()
        if not in_p or not out_d: return messagebox.showwarning("警告", "请先选好路径！")
        
        self.btn_run.config(state="disabled"); self.log_area.delete(1.0, tk.END)
        self.log(">>> 启动处理任务...")
        
        try:
            # 根据文件后缀名选择读取引擎
            engine = 'xlrd' if in_p.endswith('.xls') else 'openpyxl'
            excel = pd.ExcelFile(in_p, engine=engine)
            
            for sn in excel.sheet_names:
                # 排除系统级 Sheet
                if any(kw in sn for kw in ['说明', '时钟', '网络', '串口', '目录']): continue
                
                df = pd.read_excel(excel, sheet_name=sn)
                self.log(f"--- 正在分析 Sheet: {sn} ---")
                
                # 模式适配与转换
                res = self.adapter(df, sn)
                
                if res is not None:
                    # 导出结果，使用 utf-8-sig 防止中文乱码
                    out_path = os.path.join(out_d, f"数据库模板_{sn}.csv")
                    res.to_csv(out_path, index=False, encoding='utf-8-sig')
                    self.log(f"  [完成] 成功导出: {out_path}")
            
            messagebox.showinfo("完成", "所有 Sheet 已处理完毕！")
        except Exception as e:
            self.log(f"  [严重错误] {str(e)}")
            messagebox.showerror("错误", f"程序异常: {e}")
        finally:
            self.btn_run.config(state="normal")

    def adapter(self, df, sheet_name):
        """核心适配器：增加模糊匹配，解决空格导致的无法识别问题"""
        # 1. 获取所有列名，并进行极度清洗：转字符串、去两端空格、去中间空格
        raw_cols = [str(c) for c in df.columns.tolist()]
        clean_cols = [c.replace(" ", "").replace("\n", "") for c in raw_cols]
        
        # --- 模式 1：并列 PLC 结构识别 (针对固卷表) ---
        # 模糊检查：只要列名里包含关键字，且重复出现，就判定为固卷表
        if clean_cols.count('序号') >= 2 or '寄存器地址' in "".join(clean_cols):
            self.log(f"  [匹配成功] 发现并列 PLC 结构: {sheet_name}")
            try:
                # 固卷表手术：根据位置切分左右两部分
                # 这里使用坐标切分最为稳妥
                part1 = df.iloc[:, 0:5].copy()  # 左侧 5 列
                part2 = df.iloc[:, 7:12].copy() # 右侧 5 列
                
                # 强制对齐列名，防止因为空格导致的拼接失败
                part2.columns = part1.columns
                df_final = pd.concat([part1, part2], ignore_index=True)
                
                # 识别 Name 和地址列的具体索引位置（防止列名微变）
                name_idx = 1 # 第二列通常是名字
                addr_idx = 4 # 第五列通常是 Modbus 地址
                
                # 清洗空行
                df_final = df_final.dropna(subset=[df_final.columns[name_idx]])
                
                return self.build_template(df_final, df_final.columns[name_idx], df_final.columns[addr_idx], "PLC")
            except Exception as e:
                self.log(f"  [手术失败] 结构解析异常: {e}")

        # --- 模式 2：硬件 LCU 表识别 ---
        elif '模块号' in clean_cols or '本侧盘柜' in clean_cols:
            self.log(f"  [匹配成功] 识别为 LCU 硬接线表: {sheet_name}")
            # 找到对应的原始列名进行填充
            fill_cols = [c for c in df.columns if '盘柜' in str(c) or '模块' in str(c)]
            df[fill_cols] = df[fill_cols].ffill()
            df_valid = df.dropna(subset=[c for c in df.columns if '点名' in str(c)][0])
            return self.build_template(df_valid, [c for c in df.columns if '点名' in str(c)][0], None, "LCU")

        # --- 模式 3：逻辑报警映射表识别 ---
        elif '描述' in clean_cols:
            self.log(f"  [匹配成功] 识别为逻辑报警表: {sheet_name}")
            name_col = [c for c in df.columns if '描述' in str(c)][0]
            df_valid = df.dropna(subset=[name_col])
            return self.build_template(df_valid, name_col, None, "LOGIC")
        
        self.log(f"  [跳过] 无法匹配该 Sheet 结构。检测到的清洗后列名为: {clean_cols[:5]}")
        return None

    def build_template(self, df, name_col, addr_col, mode):
        """填充 48 列数据库模板"""
        count = len(df)
        res = pd.DataFrame(columns=TARGET_COLUMNS)
        
        # 基础点名映射
        res['描述'] = df[name_col].astype(str).str.strip().values
        res['序号'] = range(count)
        res['点号'] = res['序号'].astype(float)
        res['内部点号'] = res['序号'].astype(float)
        
        # 核心逻辑：地址与 ID 翻译
        if mode == "PLC":
            # 处理 PLC 地址中的小数点 (如 40011.0 -> 40011)
            raw_addr = df[addr_col].astype(str)
            res['实际测点地址'] = raw_addr.apply(lambda x: x.split('.')[0] if '.' in x else x).values
            res['ACC测点名'] = [f"PLC_PT_{i:03d}" for i in range(count)]
            res['虚拟点'] = 'false'
        elif mode == "LCU":
            res['实际测点地址'] = [f"{STATION_PREFIX}.{i}" for i in range(count)]
            res['ACC测点名'] = [f"KGZ_DIN{i:03d}" for i in range(count)]
            res['虚拟点'] = 'false'
        else: # LOGIC
            res['实际测点地址'] = "0"
            res['虚拟点'] = 'true'
            res['ACC测点名'] = [f"VIRT_PT_{i:03d}" for i in range(count)]

        # 填充系统要求的默认参数
        res['节点别名'] = 'COM1'
        for col in ['入历史库', '报警', '上位机报警', '0->1报警', '1->0报警']:
            res[col] = 'true'
        res['1->0描述'], res['0->1描述'] = '复归', '动作'
        res['取反'], res['事件启动'], res['远动'] = 'false', 'false', 'false'
        
        # 补全缺失列为 0.0
        res = res.fillna(0.0)
        # 如果列是空的（pd.concat生成的全Nan），再次强制补 0
        for col in TARGET_COLUMNS:
            if res[col].isnull().all(): res[col] = 0.0
            
        return res

if __name__ == "__main__":
    root = tk.Tk(); app = SmartPointListApp(root); root.mainloop()