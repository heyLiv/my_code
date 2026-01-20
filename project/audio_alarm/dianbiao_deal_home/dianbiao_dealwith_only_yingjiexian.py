import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext

# --- 配置区 ---
# 测点地址前缀，可根据实际站号修改
STATION_PREFIX = "1.1.11.3"

# 数据库 48 列标准模板
TARGET_COLUMNS = [
    '实际测点地址', '描述', '点号', '节点别名', '取反', '事件启动', '远动', '双节点', '双节点名',
    '测值', '输入', '虚拟点', '检修', '手动', '品质', '入历史库', '语音报警', '生值', '反值',
    '开入实测值', '事故追忆启动源', '远动投退', '远动测值', '事件处理方式', 'ACC测点名',
    '内部点号', '驱动内部点号', '序号', '板号', '驱动名称', '设备类型', '一览表',
    '1->0描述', '0->1描述', '0->1报警', '0->1登录', '0->1寻呼', '1->0报警',
    '1->0登录', '1->0寻呼', '上位机报警', '0->1语音号', '1->0语音号', '镜头号',
    '报警', '电话语音号', '电话报警', '屏蔽'
]

def process_hardwired_logic(input_path, output_dir):
    try:
        # 1. 加载 Excel
        excel = pd.ExcelFile(input_path, engine='openpyxl')
        
        for sheet_name in excel.sheet_names:
            # 只处理包含 DI/DO/AI/SOE/RTD 关键字的硬接线表
            if not any(kw in sheet_name.upper() for kw in ['DI', 'DO', 'AI', 'SOE', 'RTD']):
                continue
            
            print(f">>> 正在专项处理硬接线表: {sheet_name}")
            df = pd.read_excel(excel, sheet_name=sheet_name)
            
            # 2. 核心清洗：补齐合并单元格（向下填充）
            # 针对硬接线表的“本侧盘柜”和“模块号”进行填充
            cols_to_fill = [c for c in ['本侧盘柜', '模块号'] if c in df.columns]
            if cols_to_fill:
                df[cols_to_fill] = df[cols_to_fill].ffill()
            
            # 3. 过滤：去掉没有点名的行，并剔除备用点
            if '点名' not in df.columns:
                continue
            df_valid = df.dropna(subset=['点名']).copy()
            df_valid['点名'] = df_valid['点名'].astype(str)
            df_valid = df_valid[~df_valid['点名'].str.contains('备用|预留|/|#NAME', na=False)]
            
            if df_valid.empty:
                continue

            # 4. 构建目标模板
            count = len(df_valid)
            res = pd.DataFrame(columns=TARGET_COLUMNS)
            
            res['描述'] = df_valid['点名'].values
            res['序号'] = range(count)
            res['点号'] = res['序号'].astype(float)
            res['内部点号'] = res['序号'].astype(float)
            
            # 根据 Sheet 名称决定 ID 前缀
            prefix = "KGZ_PT"
            if 'DI' in sheet_name.upper(): prefix = "KGZ_DIN"
            elif 'DO' in sheet_name.upper(): prefix = "KGZ_DON"
            elif 'AI' in sheet_name.upper(): prefix = "KGZ_AIN"
            elif 'SOE' in sheet_name.upper(): prefix = "KGZ_SOE"
            elif 'RTD' in sheet_name.upper(): prefix = "KGZ_RTD"
            
            res['ACC测点名'] = [f"{prefix}{i:03d}" for i in range(count)]
            
            # 地址生成：前缀 + 序号（最稳妥的物理偏移量生成方式）
            res['实际测点地址'] = [f"{STATION_PREFIX}.{i}" for i in range(count)]
            
            # 5. 填充硬接线表专用的默认业务参数
            res['节点别名'] = 'COM1'
            res['虚拟点'] = 'false'  # 硬接线表全是物理点
            for col in ['入历史库', '报警', '上位机报警', '0->1报警', '1->0报警']:
                res[col] = 'true'
            
            # 设置开关量动作描述
            if any(kw in sheet_name.upper() for kw in ['DI', 'DO', 'SOE']):
                res['1->0描述'], res['0->1描述'] = '复归', '动作'
            
            # 6. 最终清理：所有空列补 0.0
            res = res.fillna(0.0)
            for col in TARGET_COLUMNS:
                if res[col].isnull().all(): res[col] = 0.0
            
            # 7. 保存导出
            output_filename = os.path.join(output_dir, f"硬接线数据库_{sheet_name}.csv")
            res.to_csv(output_filename, index=False, encoding='utf-8-sig')
            print(f"    [成功] 已生成: {output_filename}")

        return True
    except Exception as e:
        print(f"处理出错: {e}")
        return False

# --- 极简调用区 ---
if __name__ == "__main__":
    # 你可以手动修改这里的路径，或者结合之前的 GUI 使用
    input_file = r"D:\B5_项目库\20260119_点表导入优化尝试\raw_input\4、开关站LCU（2025.11.29）.xlsx"
    output_folder = r"D:\B5_项目库\20260119_点表导入优化尝试\raw_input\output"
    
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
        
    if process_hardwired_logic(input_file, output_folder):
        print("\n" + "="*30)
        print("硬接线点表专项优化完成！")
        print("="*30)