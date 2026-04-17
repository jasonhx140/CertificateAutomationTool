import pandas as pd
import numpy as np
import os
import json
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
import datetime

# ===================== 配置 =====================
CONFIG_FILE = "tool_config.json"
MERGE_COL_NAME = "销售凭证"

# 必须删除的物料编码
REMOVE_MATERIALS = {
    "590000013", "590000014", "590000015", "590000027", "590000028",
    "590000029", "590000030", "590000003", "590000004", "590000005",
    "590000006", "590000007", "590000008", "590000009", "590000010",
    "590000011", "590000012", "590000032", "590000033", "590000031",
    "590000000"
}

ORG_CONFIG = {
    "维通利华北京销售组织": {
        "test_sheet": "北京",
        "license": "SCXK（京）2021001"
    },
    "维通利华湖北销售组织": {
        "test_sheet": "湖北",
        "license": "SCXK（鄂）2022030"
    }
}

# ===================== 配置记忆 =====================
def load_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except:
            if os.path.exists(CONFIG_FILE):
                os.remove(CONFIG_FILE)
    return {}

def save_config(export_path, cert_path, test_path):
    cfg = {"export": export_path, "cert": cert_path, "test": test_path}
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)

# ===================== 文件夹命名规则 =====================
def get_output_folder():
    today = datetime.datetime.now().strftime("%Y%m%d")
    if not os.path.exists(today):
        os.makedirs(today)
        return today
    idx = 2
    while True:
        folder_name = f"{today}-{idx}"
        if not os.path.exists(folder_name):
            os.makedirs(folder_name)
            return folder_name
        idx += 1

# ===================== 【美化加强】自动列宽 + 表头筛选 =====================
def save_pretty_excel(df, filename, text_col_name):
    wb = Workbook()
    ws = wb.active
    ws.title = "合格证"

    # 写入数据
    for col_idx, col_name in enumerate(df.columns, 1):
        ws.cell(row=1, column=col_idx, value=col_name)
    for r_idx, row in enumerate(df.values, 2):
        for c_idx, value in enumerate(row, 1):
            ws.cell(row=r_idx, column=c_idx, value=value)

    # 样式
    font_normal = Font(name="微软雅黑", size=9)
    font_header = Font(name="微软雅黑", size=9, bold=True)
    fill_header = PatternFill(start_color="4472C9", end_color="4472C9", fill_type="solid")
    align_center = Alignment(horizontal='center', vertical='center')

    # 表头样式
    for cell in ws[1]:
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = align_center

    # 内容样式
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.font = font_normal
            cell.alignment = align_center

    # ===================== 【新增】自动适配列宽 =====================
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = min(max_length + 2, 60)  # 最大60，防止过宽
        ws.column_dimensions[column].width = adjusted_width

    # ===================== 【新增】表头添加筛选 =====================
    ws.auto_filter.ref = ws.dimensions

    ws.freeze_panes = "A2"
    wb.save(filename)

# ===================== 核心处理 =====================
def process_all(export_path, cert_path, test_path, log_widget):
    def log(msg):
        now = datetime.datetime.now().strftime("%H:%M:%S")
        log_widget.insert(tk.END, f"[{now}] {msg}\n")
        log_widget.see(tk.END)
        log_widget.update()

    try:
        log("开始读取数据源...")
        export_df = pd.read_excel(export_path, dtype=str)
        cert_template = pd.read_excel(cert_path, dtype=str)
        log(f"原始数据：{len(export_df)} 行")

        # 按表头【物料】过滤，彻底删除指定编码
        if "物料" in export_df.columns:
            export_df["物料"] = export_df["物料"].astype(str).str.strip()
            before = len(export_df)
            export_df = export_df[~export_df["物料"].isin(REMOVE_MATERIALS)]
            log(f"✅ 已按表头【物料】删除无效数据：{before - len(export_df)} 行")

        # 剔除销售凭证2开头
        if MERGE_COL_NAME in export_df.columns:
            export_df = export_df[~export_df[MERGE_COL_NAME].str.startswith("2", na=False)]

        # 原有过滤
        condition_exclude = (export_df["单位"] == "EA") | export_df["拒绝原因描述"].notna()
        export_df = export_df[~condition_exclude]
        export_df = export_df[export_df["销售组织描述"].isin(ORG_CONFIG.keys())]

        log(f"最终有效数据：{len(export_df)} 行")
        if len(export_df) == 0:
            messagebox.showwarning("提示", "无有效数据")
            return

        # 合并凭证
        export_df["合并凭证"] = export_df["销售凭证"].str.strip() + export_df["销售订单行号"].str.strip()

        # 日期格式化
        date_cols = [c for c in export_df.columns if "日期" in c or "时间" in c]
        for col in date_cols:
            export_df[col] = pd.to_datetime(export_df[col], errors="coerce").dt.strftime("%Y-%m-%d")

        output_dir = get_output_folder()
        log(f"输出文件夹：{output_dir}")

        for org_name, cfg in ORG_CONFIG.items():
            org_data = export_df[export_df["销售组织描述"] == org_name].copy()
            if org_data.empty:
                log(f"{org_name} 无数据，跳过")
                continue

            log(f"处理 {org_name}...")
            test_df = pd.read_excel(test_path, sheet_name=cfg["test_sheet"], dtype=str)

            if "品系" in test_df.columns:
                key_col = "品系"
            elif "SAP系统品系名称" in test_df.columns:
                key_col = "SAP系统品系名称"
            elif "SAP对应名称" in test_df.columns:
                key_col = "SAP对应名称"
            else:
                raise Exception("未找到品系相关列")

            test_df["检测日期"] = pd.to_datetime(test_df["检测日期"], errors="coerce").dt.strftime("%Y-%m-%d")
            test_df = test_df.drop_duplicates(subset=[key_col], keep="last")
            date_map = dict(zip(test_df[key_col], test_df["检测日期"]))

            final_df = pd.DataFrame(columns=cert_template.columns)
            for col in final_df.columns:
                if col in org_data.columns and col != "备注":
                    final_df[col] = org_data[col]

            final_df[MERGE_COL_NAME] = org_data["合并凭证"]
            if "备注" in final_df.columns:
                final_df["备注"] = ""

            if "品系" in final_df.columns and "最后一次检测日期" in final_df.columns:
                final_df["最后一次检测日期"] = final_df["品系"].map(date_map)

            fixed_fields = {
                "生产许可证号": cfg["license"],
                "用途": "科学研究",
                "质检单位": "北京维通利华实验动物技术有限公司",
                "质量负责人": "韩雪",
                "质量等级": "SPF级"
            }
            for k, v in fixed_fields.items():
                if k in final_df.columns:
                    final_df[k] = v

            out_file = os.path.join(output_dir, f"合格证_{cfg['test_sheet']}.xlsx")
            save_pretty_excel(final_df, out_file, MERGE_COL_NAME)
            log(f"已生成：{os.path.basename(out_file)}")

        log("==================================================")
        log("✅ 全部处理完成！格式已优化：自动列宽 + 表头筛选")
        if messagebox.askyesno("完成", f"已保存至：{output_dir}\n是否打开文件夹？"):
            try:
                import subprocess
                import platform
                if platform.system() == 'Windows':
                    os.startfile(output_dir)
                elif platform.system() == 'Darwin':  # macOS
                    subprocess.run(['open', output_dir])
                else:  # Linux
                    subprocess.run(['xdg-open', output_dir])
            except Exception as e:
                log(f"打开文件夹失败: {str(e)}")

    except Exception as e:
        log(f"❌ 错误：{str(e)}")
        messagebox.showerror("失败", str(e))

# ===================== GUI界面 =====================
def main_gui():
    root = tk.Tk()
    root.title("实验动物合格证自动生成工具")
    root.geometry("760x750")
    root.resizable(False, False)

    PRIMARY = "#2C70C9"
    SECONDARY = "#F5F7FA"
    ACCENT = "#3E8D66"
    FG = "#FFFFFF"

    top_frame = tk.Frame(root, bg=PRIMARY, height=90)
    top_frame.pack(fill=tk.X)
    tk.Label(top_frame, text="实验动物合格证自动生成工具", font=("微软雅黑", 20, "bold"),
             bg=PRIMARY, fg=FG).place(x=30, y=20)
    tk.Label(top_frame, text="文件记忆｜一键生成", font=("微软雅黑", 11),
             bg=PRIMARY, fg="#D0E0F0").place(x=32, y=58)

    content_frame = tk.Frame(root, bg=SECONDARY, padx=20, pady=12)
    content_frame.pack(fill=tk.BOTH, expand=True)
    content_frame.pack_propagate(False)

    cfg = load_config()
    paths = {
        "export": cfg.get("export", ""),
        "cert": cfg.get("cert", ""),
        "test": cfg.get("test", "")
    }

    PANEL_WIDTH = 700
    FILE_H = 160
    LOG_H = 210
    BTN_FRAME_H = 120
    GAP = 10

    # 文件选择区域
    file_frame = tk.LabelFrame(content_frame, text="文件选择", font=("微软雅黑", 11, "bold"), bg=SECONDARY)
    file_frame.place(x=0, y=GAP, width=PANEL_WIDTH, height=FILE_H)

    rows = [
        ("1. export 源数据文件", "export", 10),
        ("2. 合格证模板", "cert", 40),
        ("3. 检测日期报告文件", "test", 70)
    ]

    entrys = {}
    for label, key, y in rows:
        tk.Label(file_frame, text=label, font=("微软雅黑", 10), bg=SECONDARY).place(x=15, y=y)
        ent = tk.Entry(file_frame, font=("微软雅黑", 10), state="readonly", bg="white")
        ent.place(x=180, y=y-2, width=380, height=24)
        entrys[key] = ent
        if paths[key]:
            ent.config(state=tk.NORMAL)
            ent.insert(0, paths[key])
            ent.config(state="readonly")

        def choose(k=key, e=ent):
            f = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
            if f:
                e.config(state=tk.NORMAL)
                e.delete(0, tk.END)
                e.insert(0, f)
                e.config(state="readonly")
                paths[k] = f
                save_config(paths["export"], paths["cert"], paths["test"])

        tk.Button(file_frame, text="选择文件", width=12, command=choose).place(x=570, y=y-3)

    # 运行日志
    log_frame = tk.LabelFrame(content_frame, text="运行日志", font=("微软雅黑", 11, "bold"), bg=SECONDARY)
    log_frame.place(x=0, y=GAP + FILE_H + 10, width=PANEL_WIDTH, height=LOG_H)

    log_txt = tk.Text(log_frame, font=("Consolas", 10), bg="white", relief=tk.FLAT)
    log_txt.place(x=8, y=6, width=PANEL_WIDTH-25, height=LOG_H-35)
    scr = ttk.Scrollbar(log_frame, command=log_txt.yview)
    scr.place(x=PANEL_WIDTH-15, y=6, width=12, height=LOG_H-35)
    log_txt.config(yscrollcommand=scr.set)

    # 操作按钮
    btn_frame = tk.LabelFrame(content_frame, text="操作", font=("微软雅黑", 11, "bold"), bg=SECONDARY)
    btn_frame.place(x=0, y=GAP + FILE_H + LOG_H + 40, width=PANEL_WIDTH, height=BTN_FRAME_H)

    def run():
        if not all(paths.values()):
            messagebox.showwarning("提示", "请选择全部3个文件")
            return
        process_all(paths["export"], paths["cert"], paths["test"], log_txt)

    tk.Button(btn_frame, text="🚀 一键生成合格证", font=("微软雅黑", 14, "bold"),
              bg=ACCENT, fg=FG, relief=tk.FLAT, command=run)\
        .place(relx=0.5, y=25, anchor="n", width=PANEL_WIDTH-40, height=40)

    # 底部
    bottom = tk.Frame(root, bg=PRIMARY, height=28)
    bottom.pack(side=tk.BOTTOM, fill=tk.X)
    tk.Label(bottom, text="© 内部专用工具", font=("微软雅黑", 9),
             bg=PRIMARY, fg="#D0E0F0").pack(anchor=tk.CENTER)

    root.mainloop()

if __name__ == "__main__":
    main_gui()