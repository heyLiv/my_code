import json
from datetime import datetime
import tkinter as tk
from tkinter import messagebox

CONFIG_FILE = "travel_records_v7.json"

class TravelCalculator:
    def __init__(self, root):
        self.root = root
        self.root.title("出差补贴计算器 v7.4 [多项目管理+原版逻辑]")
        self.root.geometry("600x750")  # 调整为更适合两列布局的宽度
        
        self.all_data = self.load_all_config()
        self.current_project_name = None

        self.setup_ui()
        self.refresh_project_list()

    def load_all_config(self):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                return data if data else {"默认项目": self.get_default_values()}
        except:
            return {"默认项目": self.get_default_values()}

    def get_default_values(self):
        today_str = datetime.today().strftime("%Y-%m-%d")
        return {
            "start_date": today_str, "end_date": today_str,
            "traffic": "80", "other": "0", "house_fixed": "0",
            "house_invoice": "0", "cost_day": "0", "room_money_350": "0",
            "special_money": "0", "special_days": "0", "house_type": 1
        }

    def save_all_config(self):
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(self.all_data, f, ensure_ascii=False, indent=4)

    # —— 工具函数：生成你原有的“双列一行”输入 —— 
    def add_row(self, parent, label1, label2):
        row = tk.Frame(parent)
        row.pack(fill="x", pady=4)

        tk.Label(row, text=label1, width=12, anchor="e").pack(side="left")
        e1 = tk.Entry(row, width=12)
        e1.pack(side="left", padx=5)

        tk.Label(row, text=label2, width=15, anchor="e").pack(side="left", padx=5)
        e2 = tk.Entry(row, width=12)
        e2.pack(side="left")
        return e1, e2

    def setup_ui(self):
        self.paned = tk.PanedWindow(self.root, orient=tk.HORIZONTAL, sashwidth=4)
        self.paned.pack(fill="both", expand=True)

        # --- 左侧：项目列表 ---
        left_frame = tk.Frame(self.paned, bg="#f0f0f0", width=150)
        self.paned.add(left_frame)

        tk.Label(left_frame, text="项目列表", bg="#f0f0f0", font=("微软雅黑", 9, "bold")).pack(pady=5)
        self.listbox = tk.Listbox(left_frame, font=("微软雅黑", 9))
        self.listbox.pack(fill="both", expand=True, padx=5, pady=5)
        self.listbox.bind("<<ListboxSelect>>", self.on_project_select)

        btn_f = tk.Frame(left_frame, bg="#f0f0f0")
        btn_f.pack(fill="x", pady=5)
        tk.Button(btn_f, text="新增", command=self.add_new_project, width=6).pack(side="left", padx=2)
        tk.Button(btn_f, text="删除", command=self.delete_project, width=6).pack(side="right", padx=2)

        # --- 右侧：原版显示格式 ---
        self.right_scroll = tk.Frame(self.paned, padx=10, pady=10)
        self.paned.add(self.right_scroll)

        # 项目名称 (对齐布局)
        row_proj = tk.Frame(self.right_scroll)
        row_proj.pack(fill="x", pady=4)
        tk.Label(row_proj, text="当前项目名称", width=12, anchor="e").pack(side="left")
        self.entry_proj_name = tk.Entry(row_proj, width=12, fg="blue")
        self.entry_proj_name.pack(side="left", padx=5)

        # 行1：日期
        self.entry_start, self.entry_end = self.add_row(self.right_scroll, "出差开始日期", "出差结束日期")
        
        # 行2：补贴
        self.entry_traffic, self.entry_other = self.add_row(self.right_scroll, "交通补贴金额", "其他补贴金额")

        # 行3：包干/节支
        self.entry_house_fixed, self.entry_house_invoice = self.add_row(self.right_scroll, "住宿【包干】\n(包干金额)", "住宿【节支】\n(节支剩余)")

        # 行4：支出 + 酒店档位 (350/450/550)
        self.entry_cost, self.entry_room_350 = self.add_row(self.right_scroll, "每日支出金额", "住宿补贴金额\n350/450/550")

        # 行5：特殊房补
        self.entry_special_money, self.entry_special_days = self.add_row(self.right_scroll, "特殊房补金额", "特殊房补天数")

        # 房补方式选择
        row_choice = tk.Frame(self.right_scroll)
        row_choice.pack(fill="x", pady=6)
        tk.Label(row_choice, text="房补方式", width=12, anchor="e").pack(side="left")
        self.house_choice = tk.IntVar(value=1)
        tk.Radiobutton(row_choice, text="包干方式", variable=self.house_choice, value=1).pack(side="left", padx=10)
        tk.Radiobutton(row_choice, text="节支补贴", variable=self.house_choice, value=2).pack(side="left")

        # 按钮
        btn_row = tk.Frame(self.right_scroll)
        btn_row.pack(fill="x", pady=10)
        tk.Button(btn_row, text="保存当前项目", command=self.save_current_settings).pack(side="left", padx=20, expand=True)
        tk.Button(btn_row, text="开始计算", command=self.calculate, bg="#e1f5fe").pack(side="left", padx=20, expand=True)

        # 输出框
        self.text_result = tk.Text(self.right_scroll, height=15, font=("微软雅黑", 9))
        self.text_result.pack(fill="both", expand=True)

    # ========= 逻辑部分 =========
    def refresh_project_list(self):
        self.listbox.delete(0, tk.END)
        for name in self.all_data.keys():
            self.listbox.insert(tk.END, name)

    def on_project_select(self, event):
        selection = self.listbox.curselection()
        if selection:
            name = self.listbox.get(selection[0])
            self.load_project_data(name)

    def load_project_data(self, name):
        self.current_project_name = name
        cfg = self.all_data[name]
        self.entry_proj_name.delete(0, tk.END)
        self.entry_proj_name.insert(0, name)
        
        fields = {
            self.entry_start: "start_date", self.entry_end: "end_date",
            self.entry_traffic: "traffic", self.entry_other: "other",
            self.entry_house_fixed: "house_fixed", self.entry_house_invoice: "house_invoice",
            self.entry_cost: "cost_day", self.entry_room_350: "room_money_350",
            self.entry_special_money: "special_money", self.entry_special_days: "special_days"
        }
        for widget, key in fields.items():
            widget.delete(0, tk.END)
            widget.insert(0, cfg.get(key, "0"))
        self.house_choice.set(cfg.get("house_type", 1))

    def add_new_project(self):
        new_name = f"项目_{datetime.now().strftime('%m%d%H%M')}"
        self.all_data[new_name] = self.get_default_values()
        self.refresh_project_list()
        self.load_project_data(new_name)

    def delete_project(self):
        if not self.current_project_name: return
        if messagebox.askyesno("提示", f"确定删除项目：{self.current_project_name}？"):
            del self.all_data[self.current_project_name]
            self.current_project_name = None
            self.refresh_project_list()
            self.save_all_config()

    def save_current_settings(self):
        name = self.entry_proj_name.get().strip()
        if not name: return
        if self.current_project_name and self.current_project_name != name:
            if self.current_project_name in self.all_data: del self.all_data[self.current_project_name]

        self.all_data[name] = {
            "start_date": self.entry_start.get().strip(),
            "end_date": self.entry_end.get().strip(),
            "traffic": self.entry_traffic.get().strip(),
            "house_fixed": self.entry_house_fixed.get().strip(),
            "house_invoice": self.entry_house_invoice.get().strip(),
            "other": self.entry_other.get().strip(),
            "cost_day": self.entry_cost.get().strip(),
            "room_money_350": self.entry_room_350.get().strip(),
            "special_days": self.entry_special_days.get().strip(),
            "special_money": self.entry_special_money.get().strip(),
            "house_type": self.house_choice.get()
        }
        self.current_project_name = name
        self.save_all_config()
        self.refresh_project_list()
        messagebox.showinfo("提示", "设置保存成功！")

    def safe_float(self, entry):
        val = entry.get().strip()
        try: return float(val) if val else 0.0
        except: return 0.0

    def calculate(self):
        try:
            # 日期逻辑容错
            raw_start = self.entry_start.get().replace('.', '-').replace('/', '-').strip()
            raw_end = self.entry_end.get().replace('.', '-').replace('/', '-').strip()
            start_date = datetime.strptime(raw_start, "%Y-%m-%d")
            end_date = datetime.strptime(raw_end, "%Y-%m-%d")
            today = datetime.today()

            # 获取输入值
            traffic = self.safe_float(self.entry_traffic)
            house_fixed = self.safe_float(self.entry_house_fixed)
            house_invoice = self.safe_float(self.entry_house_invoice)
            other = self.safe_float(self.entry_other)
            cost_day = self.safe_float(self.entry_cost)
            # 酒店补贴档位 (350/450/550)
            room_money = self.safe_float(self.entry_room_350)
            special_money = self.safe_float(self.entry_special_money)
            special_days = int(self.safe_float(self.entry_special_days))

            # 互斥逻辑
            house = house_fixed if self.house_choice.get() == 1 else house_invoice

            # 日期运算
            days_diff = (today - start_date).days
            remain_days = (end_date - today).days
            all_days = (end_date - start_date).days
            progress = 100 * (days_diff / all_days) if all_days > 0 else 0

            # 总补贴计算 (整合了 room_money 的影响，如果你的公式不涉及它，可以自行微调这里)
            total_money = (traffic + house + other) * (days_diff - special_days) + special_days * (special_money+traffic+other)
            total_money_cost = (traffic + house + other - cost_day) * (days_diff - special_days) + special_days * (special_money+traffic+other-cost_day)
            total_cost = days_diff * cost_day
            remain_money = remain_days * (traffic + house + other)

            msg = (
                f"起始日：{start_date.strftime('%Y-%m-%d')} --> 结束日：{end_date.strftime('%Y-%m-%d')}\n"
                f"出差总天数：{all_days} 天 \n"
                f"已出差：{days_diff} 天 | 剩余：{remain_days} 天 | 目前出差进度：{progress:.2f}%\n"
                f"——————————————————————————\n"
                f"累计总补贴(不扣支出)：【{total_money:.2f}】 元\n"
                f"—————————————————\n"
                f"累计总支出：{total_cost:.2f} 元\n"
                f"累计净补贴(扣除每日支出)：{total_money_cost:.2f} 元\n"
                f"剩余未完成补贴：{remain_money:.2f} 元\n"
                f"——————————————————————————\n"
                f"住宿标准参考：{room_money} 元档位\n"
                f"注：累计总补贴计算方式为:(交通+住宿[包干/节支]2选1+其他)\n"
                f"注：特殊房补为出差住宿补贴金额不同的情况，\n    将从总天数里扣除单独进行计算并相加至总额。\n"
            )
            self.text_result.delete("1.0", tk.END)
            self.text_result.insert(tk.END, msg)

        except Exception as e:
            messagebox.showerror("计算失败", f"请检查日期格式是否为 YYYY-MM-DD\n错误信息：{e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = TravelCalculator(root)
    root.mainloop()