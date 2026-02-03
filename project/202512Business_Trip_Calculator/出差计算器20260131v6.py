import json
import os
from datetime import datetime
import tkinter as tk
from tkinter import messagebox, ttk

# ========= 方案 C：系统用户数据目录配置 =========
def get_config_path():
    # 获取系统为当前用户分配的 AppData\Local 目录
    app_data_dir = os.path.join(os.environ.get('LOCALAPPDATA', os.path.expanduser('~')), "TravelCalculator")
    # 如果该目录不存在，则自动创建
    if not os.path.exists(app_data_dir):
        os.makedirs(app_data_dir)
    # 返回完整的文件存储路径
    return os.path.join(app_data_dir, "travel_records_v7.json")

CONFIG_FILE = get_config_path()
SUMMARY_KEY = "📋 [累计补贴汇总 - 置顶]" 

class TravelCalculator:
    def __init__(self, root):
        self.root = root
        self.root.title("出差补贴计算器 v6 [2026.01版 by Liv]")
        self.root.geometry("620x600") 
        
        self.all_data = self.load_all_config()
        self.current_project_name = None

        self.setup_ui()
        self.refresh_project_list()

    def load_all_config(self):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                if SUMMARY_KEY in data: del data[SUMMARY_KEY]
                return data if data else {"默认项目": self.get_default_values()}
        except:
            return {"默认项目": self.get_default_values()}

    def get_default_values(self):
        today_str = datetime.today().strftime("%Y-%m-%d")
        return {
            "start_date": today_str,
            "end_date": today_str,
            "traffic": "180",
            "other": "0",
            "house_fixed": "0",
            "house_invoice": "0",
            "cost_day": "0",
            "room_money_350": "350",
            "special_money": "0",
            "special_days": "0",
            "house_type": 1
        }

    def save_all_config(self):
        save_data = {k: v for k, v in self.all_data.items() if k != SUMMARY_KEY}
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(save_data, f, ensure_ascii=False, indent=4)

    def add_row(self, parent, label1, label2):
        row = tk.Frame(parent)
        row.pack(fill="x", pady=4)
        tk.Label(row, text=label1, width=14, anchor="e").pack(side="left")
        e1 = tk.Entry(row, width=12)
        e1.pack(side="left", padx=5)
        tk.Label(row, text=label2, width=16, anchor="e").pack(side="left", padx=5)
        e2 = tk.Entry(row, width=12)
        e2.pack(side="left")
        return e1, e2

    def setup_ui(self):
        self.paned = tk.PanedWindow(self.root, orient=tk.HORIZONTAL, sashwidth=4)
        self.paned.pack(fill="both", expand=True)

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

        self.right_container = tk.Frame(self.paned, padx=10, pady=10)
        self.paned.add(self.right_container)

        self.edit_frame = tk.Frame(self.right_container)
        
        row_proj = tk.Frame(self.edit_frame)
        row_proj.pack(fill="x", pady=4)
        tk.Label(row_proj, text="当前项目名称", width=14, anchor="e").pack(side="left")
        self.entry_proj_name = tk.Entry(row_proj, width=25, fg="blue") 
        self.entry_proj_name.pack(side="left", padx=5)

        self.entry_start, self.entry_end = self.add_row(self.edit_frame, "出差开始日期", "出差结束日期")
        self.entry_traffic, self.entry_other = self.add_row(self.edit_frame, "交通+吃饭补贴", "其他补贴金额")
        self.entry_house_fixed, self.entry_house_invoice = self.add_row(self.edit_frame, "住宿【包干】\n(包干金额)", "住宿【节支】\n(携程价格)")
        self.entry_cost, self.entry_room_350 = self.add_row(self.edit_frame, "每日支出金额", "住宿补贴金额\n350/450/550")
        self.entry_special_money, self.entry_special_days = self.add_row(self.edit_frame, "特殊房补金额", "特殊房补天数")

        row_choice = tk.Frame(self.edit_frame)
        row_choice.pack(fill="x", pady=6)
        tk.Label(row_choice, text="房补方式", width=14, anchor="e").pack(side="left")
        self.house_choice = tk.IntVar(value=1)
        tk.Radiobutton(row_choice, text="包干方式", variable=self.house_choice, value=1).pack(side="left", padx=10)
        tk.Radiobutton(row_choice, text="节支补贴", variable=self.house_choice, value=2).pack(side="left")

        btn_row = tk.Frame(self.edit_frame)
        btn_row.pack(fill="x", pady=10)
        tk.Button(btn_row, text="💾 保存当前设置", command=self.save_current_settings).pack(side="left", padx=20, expand=True)
        tk.Button(btn_row, text="🧮 开始计算", command=self.calculate, bg="#e1f5fe").pack(side="left", padx=20, expand=True)

        self.summary_frame = tk.Frame(self.right_container)
        tk.Label(self.summary_frame, text="选择需要汇总的项目", font=("微软雅黑", 9, "bold")).pack(pady=5)
        
        self.check_canvas = tk.Canvas(self.summary_frame, height=120)
        self.check_scroll = ttk.Scrollbar(self.summary_frame, orient="vertical", command=self.check_canvas.yview)
        self.check_inner = tk.Frame(self.check_canvas)
        
        self.check_canvas.create_window((0, 0), window=self.check_inner, anchor="nw")
        self.check_canvas.configure(yscrollcommand=self.check_scroll.set)
        self.check_canvas.pack(side="left", fill="both", expand=True)
        self.check_scroll.pack(side="right", fill="y")
        
        tk.Button(self.summary_frame, text="确定", command=self.calculate_all_selected, bg="#d1ffcf", height=2).pack(fill="x", pady=10)

        self.text_result = tk.Text(self.right_container, height=10, font=("微软雅黑", 9), bg="#fafafa")
        self.text_result.pack(fill="both", expand=True)

    def refresh_project_list(self):
        self.listbox.delete(0, tk.END)
        self.listbox.insert(0, SUMMARY_KEY)
        self.listbox.itemconfig(0, fg="green")
        for name in sorted(self.all_data.keys()):
            self.listbox.insert(tk.END, name)

    def on_project_select(self, event):
        selection = self.listbox.curselection()
        if not selection: return
        name = self.listbox.get(selection[0])
        if name == SUMMARY_KEY:
            self.show_summary_view()
        else:
            self.show_edit_view(name)

    def show_edit_view(self, name):
        self.summary_frame.pack_forget()
        self.edit_frame.pack(fill="x")
        self.load_project_data(name)

    def show_summary_view(self):
        self.edit_frame.pack_forget()
        self.summary_frame.pack(fill="x")
        self.current_project_name = SUMMARY_KEY
        for widget in self.check_inner.winfo_children(): widget.destroy()
        self.check_vars = {}
        for name in sorted(self.all_data.keys()):
            var = tk.BooleanVar(value=True)
            cb = tk.Checkbutton(self.check_inner, text=name, variable=var, font=("微软雅黑", 9))
            cb.pack(anchor="w")
            self.check_vars[name] = var
        self.check_inner.update_idletasks()
        self.check_canvas.config(scrollregion=self.check_canvas.bbox("all"))

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
        new_name = f"新项目_{datetime.now().strftime('%m%d%H%M')}"
        self.all_data[new_name] = self.get_default_values()
        self.refresh_project_list()
        self.load_project_data(new_name)

    def delete_project(self):
        if not self.current_project_name or self.current_project_name == SUMMARY_KEY: 
            messagebox.showwarning("警告", "汇总项不可删除！")
            return
        if messagebox.askyesno("提示", f"确定删除项目：{self.current_project_name}？"):
            del self.all_data[self.current_project_name]
            self.current_project_name = None
            self.refresh_project_list()
            self.save_all_config()

    def save_current_settings(self):
        name = self.entry_proj_name.get().strip()
        if not name or name == SUMMARY_KEY: return
        if self.current_project_name and self.current_project_name != name and self.current_project_name != SUMMARY_KEY:
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

    def safe_float(self, val):
        try: return float(val) if val else 0.0
        except: return 0.0

    def calculate_logic(self, cfg):
        try:
            raw_start = cfg['start_date'].replace('.', '-').replace('/', '-').strip()
            raw_end = cfg['end_date'].replace('.', '-').replace('/', '-').strip()
            start_date = datetime.strptime(raw_start, "%Y-%m-%d")
            end_date = datetime.strptime(raw_end, "%Y-%m-%d")
            today = datetime.today()

            traffic_food = self.safe_float(cfg['traffic'])
            house_fixed = self.safe_float(cfg['house_fixed'])
            house_invoice = self.safe_float(cfg['house_invoice'])
            other = self.safe_float(cfg['other'])
            cost_day = self.safe_float(cfg['cost_day'])
            special_money = self.safe_float(cfg['special_money'])
            special_days = int(self.safe_float(cfg['special_days']))
            house_type = int(cfg['house_type'])

            house = house_fixed if house_type == 1 else house_invoice

            all_days = (end_date - start_date).days + 1
            if today > end_date:
                days_diff = all_days
                remain_days = 0
            else:
                days_diff = (today - start_date).days + 1
                remain_days = max(0, (end_date - today).days) 

            progress = 100 * (days_diff / all_days) if all_days > 0 else 0

            total_money = (traffic_food + house + other) * (days_diff - special_days) + special_days * (special_money+traffic_food+other)
            total_money_cost = (traffic_food + house + other - cost_day) * (days_diff - special_days) + special_days * (special_money+traffic_food+other-cost_day)
            total_cost = days_diff * cost_day
            remain_money = remain_days * (traffic_food + house + other)

            return {
                "start": start_date, "end": end_date, "all": all_days, "diff": days_diff, 
                "remain": remain_days, "prog": progress, "total": total_money, 
                "total_cost_net": total_money_cost, "cost": total_cost, "rem_mon": remain_money
            }
        except:
            return None

    def calculate(self):
        cfg = {
            "start_date": self.entry_start.get(), "end_date": self.entry_end.get(),
            "traffic": self.entry_traffic.get(), "house_fixed": self.entry_house_fixed.get(),
            "house_invoice": self.entry_house_invoice.get(), "other": self.entry_other.get(),
            "cost_day": self.entry_cost.get(), "special_money": self.entry_special_money.get(),
            "special_days": self.entry_special_days.get(), "house_type": self.house_choice.get()
        }
        res = self.calculate_logic(cfg)
        if res:
            msg = (
                f"起始日：{res['start'].strftime('%Y-%m-%d')} --> 结束日：{res['end'].strftime('%Y-%m-%d')}\n"
                f"出差总天数：{res['all']} 天 \n"
                f"已出差：{res['diff']} 天 | 剩余：{res['remain']} 天 | 目前出差进度：{res['prog']:.2f}%\n"
                f"——————————————————————————\n"
                f"累计总补贴(不扣支出)：【{res['total']:.2f}】 元\n"
                f"—————————————————\n"
                f"累计总支出：{res['cost']:.2f} 元\n"
                f"累计净补贴(扣除每日支出)：{res['total_cost_net']:.2f} 元\n"
                f"剩余未完成补贴：{res['rem_mon']:.2f} 元\n"
                f"——————————————————————————\n"
                f"注：累计总补贴计算方式为:(交通+吃饭+住宿[包干/节支]2选1+其他)\n"
                f"注：特殊房补为出差住宿补贴金额不同的情况，\n    节支补贴计算方式=（酒店额度-实际消费）*0.7\n"
            )
            self.text_result.delete("1.0", tk.END)
            self.text_result.insert(tk.END, msg)
        else:
            messagebox.showerror("错误", "日期格式不正确")

    def calculate_all_selected(self):
        total_sum = 0
        cost_sum = 0
        remain_sum = 0
        count = 0
        detail_msg = "" 

        for name, var in self.check_vars.items():
            if var.get():
                res = self.calculate_logic(self.all_data[name])
                if res:
                    total_sum += res['total']
                    cost_sum += res['cost']
                    remain_sum += res['rem_mon']
                    detail_msg += f"· {name[:12]}... : 已计 {res['diff']}天 | 净补贴 {res['total_cost_net']:.2f}元\n"
                    count += 1
        
        if count == 0:
            messagebox.showwarning("提示", "请先勾选要汇总的项目")
            return
            
        msg = (
            f"📊 【累计补贴汇总统计】\n"
            f"汇总项目总数：{count} 个\n"
            f"——————————————————————————\n"
            f"{detail_msg}"
            f"—————————————————\n"
            f"汇总累计总补贴(总额)：【{total_sum:.2f}】 元\n"
            f"汇总累计总支出(合计)：{cost_sum:.2f} 元\n"
            f"汇总累计净补贴(到手)：{(total_sum - cost_sum):.2f} 元\n"
            f"汇总剩余未完成补贴：{remain_sum:.2f} 元\n"
            f"——————————————————————————\n"
            f"注：该结果包含勾选项目截止今日的所有累计数值总和。\n"
        )
        self.text_result.delete("1.0", tk.END)
        self.text_result.insert(tk.END, msg)

if __name__ == "__main__":
    root = tk.Tk()
    app = TravelCalculator(root)
    root.mainloop()