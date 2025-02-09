import tkinter as tk
from tkinter import messagebox, scrolledtext
from tkcalendar import DateEntry
import pandas as pd

file_path = "/Users/yimengduan/Desktop/2024实习/南京银行/压降/Data.xlsx"
all_sheets = {}
cutoff_date_str = "未选择"  

def import_branch_data():
    global all_sheets  
    sheet_names = pd.ExcelFile(file_path).sheet_names
    all_sheets = {sheet: pd.read_excel(file_path, sheet_name=sheet) for sheet in sheet_names}
    result_text.delete("1.0", tk.END)
    summary_text.delete("1.0", tk.END)

    for sheet in sheet_names[1:]: 
        df = all_sheets[sheet]
        if "时间" in df.columns and not pd.api.types.is_datetime64_any_dtype(df["时间"]):
            df["时间"] = pd.to_datetime("1899-12-30") + pd.to_timedelta(df["时间"], unit="D")
        result_text.insert(tk.END, f"\n=== {sheet} ===\n")
        result_text.insert(tk.END, df.to_string(index=False))  
        result_text.insert(tk.END, "\n\n")

    update_summary_display()
    messagebox.showinfo("成功", f"已成功导入 {len(sheet_names) - 1} 个工作表")

def calculate_risk_balance():
    global cutoff_date_str
    try:
        cutoff_date = pd.to_datetime(date_picker.get_date())  
        cutoff_date_str = cutoff_date.strftime("%Y-%m-%d")  
        risk_levels = ["观察类", "持续关注", "重点关注"]
        result_text.delete("1.0", tk.END)
        sheet1_name = list(all_sheets.keys())[0]
        df_sheet1 = all_sheets[sheet1_name].copy()

        if "当前风险客户敞口余额" not in df_sheet1.columns:
            df_sheet1["当前风险客户敞口余额"] = None  

        branch_sums = []
        for sheet in list(all_sheets.keys())[1:]:  
            df = all_sheets[sheet].copy()
            if "时间" in df.columns and "客户名称" in df.columns and "客户敞口余额" in df.columns and "客户风险级别" in df.columns:
                if not pd.api.types.is_datetime64_any_dtype(df["时间"]):
                    df["时间"] = pd.to_datetime("1899-12-30") + pd.to_timedelta(df["时间"], unit="D")
                closest_dates = (
                    df[df["时间"] < cutoff_date]
                    .sort_values(by=["客户名称", "时间"])
                    .groupby("客户名称")["时间"]
                    .last()
                    .reset_index()
                )
                df_filtered = df.merge(closest_dates, on=["客户名称", "时间"], how="inner")
                df_filtered = df_filtered[df_filtered["客户风险级别"].isin(risk_levels)]
                branch_risk_balance = df_filtered["客户敞口余额"].sum()
                branch_sums.append(branch_risk_balance)
                result_text.insert(tk.END, f"\n=== {sheet} (已筛选) ===\n")
                result_text.insert(tk.END, df_filtered.to_string(index=False))  
                result_text.insert(tk.END, f"\n📊 {sheet} 当前风险客户敞口余额总和: {branch_risk_balance:.2f} 元\n")
                result_text.insert(tk.END, "\n\n")

        min_rows = min(len(df_sheet1), len(branch_sums)) 
        df_sheet1.loc[:min_rows - 1, "当前风险客户敞口余额"] = branch_sums[:min_rows]
        all_sheets[sheet1_name] = df_sheet1
        update_summary_display()
        messagebox.showinfo("计算完成", "所有中支的当前风险客户敞口余额计算完成，并更新至汇总展示")
    except Exception as e:
        messagebox.showerror("错误", f"发生错误: {e}")

def calculate_risk_reduction_rate():
    global cutoff_date_str
    try:
        cutoff_date = pd.to_datetime(date_picker.get_date())
        cutoff_date_str = cutoff_date.strftime("%Y-%m-%d")
        result_text.delete("1.0", tk.END)
        sheet1_name = list(all_sheets.keys())[0]
        df_sheet1 = all_sheets[sheet1_name].copy()
        branch_reduction_rates = []
        for i, sheet in enumerate(list(all_sheets.keys())[1:]):  
            df = all_sheets[sheet].copy()
            if "时间" in df.columns:
                if not pd.api.types.is_datetime64_any_dtype(df["时间"]):
                    df["时间"] = pd.to_datetime("1899-12-30") + pd.to_timedelta(df["时间"], unit="D")
            closest_dates = (
                df[df["时间"] < cutoff_date]
                .sort_values(by=["客户名称", "时间"])
                .groupby("客户名称")["时间"]
                .last()
                .reset_index()
            )
            df_filtered = df.merge(closest_dates, on=["客户名称", "时间"], how="inner")
            initial_risk_balance = df_sheet1.loc[i, "年初存量风险客户敞口余额"]
            if initial_risk_balance is None or pd.isna(initial_risk_balance):
                initial_risk_balance = 0  
            current_risk_labels = df_filtered.set_index("客户名称")["客户风险级别"]
            if current_risk_labels.isin(["观察类", "持续关注", "重点关注"]).any():
                branch_risk_balance = df_filtered["客户敞口余额"].sum()
                risk_exit_amount = initial_risk_balance - branch_risk_balance
            elif current_risk_labels.isin(["正常监测"]).any():
                risk_exit_amount = initial_risk_balance
            elif current_risk_labels.isin(["不良客户"]).any():
                bad_customer_dates = (
                    df[df["客户风险级别"] == "不良客户"]
                    .groupby("客户名称")["时间"]
                    .min()
                    .reset_index()
                )
                bad_customers = df.merge(bad_customer_dates, on=["客户名称", "时间"], how="inner")
                bad_customer_balance = bad_customers["客户敞口余额"].sum()
                risk_exit_amount = initial_risk_balance - bad_customer_balance
            else:
                risk_exit_amount = 0  

            reduction_rate = (risk_exit_amount / initial_risk_balance) * 100 if initial_risk_balance > 0 else 0
            branch_reduction_rates.append(reduction_rate)

            result_text.insert(tk.END, f"\n=== {sheet} (压降计算) ===\n")
            result_text.insert(
                tk.END,
                f" 年初存量风险客户敞口余额: {initial_risk_balance:.2f} 元\n"
                f" 年初风险客户压降退出敞口金额: {risk_exit_amount:.2f} 元\n"
                f" {sheet} 当前压降率: {reduction_rate:.2f}%\n"
            )
            result_text.insert(tk.END, "\n\n")

        min_rows = min(len(df_sheet1), len(branch_reduction_rates))
        df_sheet1.loc[:min_rows - 1, "当前压降率"] = branch_reduction_rates[:min_rows]
        all_sheets[sheet1_name] = df_sheet1
        update_summary_display()
        messagebox.showinfo("计算完成", "当前压降率计算完成")
    except Exception as e:
        messagebox.showerror("错误", f"发生错误: {e}")

def update_summary_display():
    sheet1_name = list(all_sheets.keys())[0] 
    df_sheet1 = all_sheets[sheet1_name].copy()
    if "当前风险客户敞口余额" not in df_sheet1.columns:
        df_sheet1["当前风险客户敞口余额"] = None
    if "当前压降率" not in df_sheet1.columns:
        df_sheet1["当前压降率"] = None
    all_sheets[sheet1_name] = df_sheet1
    summary_text.delete("1.0", tk.END)
    summary_text.insert(tk.END, f"\n=== {sheet1_name} (汇总展示) ===\n")
    summary_text.insert(tk.END, f" 截止到 {cutoff_date_str}\n")
    summary_text.insert(tk.END, df_sheet1.to_string(index=False))
    summary_text.insert(tk.END, "\n\n")
    messagebox.showinfo("更新完成", "汇总展示已更新")

def export_summary_to_excel():
    try:
        with pd.ExcelWriter(file_path, engine="openpyxl", mode="w") as writer:
            for sheet_name, df in all_sheets.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)
        messagebox.showinfo("导出成功", f"汇总展示已成功导出至 {file_path}")
    except Exception as e:
        messagebox.showerror("导出失败", f"导出时发生错误: {e}")

root = tk.Tk()
root.title("分行风险部操作界面")
root.geometry("900x600")
title_label = tk.Label(root, text="分行风险部操作界面", font=("Helvetica", 16, "bold"))
title_label.pack(pady=10)

import_button = tk.Button(root, text="导入各中支底表", command=import_branch_data, font=("Helvetica", 12))
import_button.pack(pady=10)

date_label = tk.Label(root, text="选择截止日期:", font=("Helvetica", 12))
date_label.pack(pady=5)
date_picker = DateEntry(root, width=12, background='darkblue', foreground='white', borderwidth=2, date_pattern='yyyy-mm-dd')
date_picker.pack(pady=5)

# 计算当前风险客户敞口余额
calculate_button = tk.Button(root, text="计算当前风险客户敞口余额", command=calculate_risk_balance, font=("Helvetica", 12))
calculate_button.pack(pady=10)

# 当前压降率 
reduction_button = tk.Button(root, text="计算当前压降率", command=calculate_risk_reduction_rate, font=("Helvetica", 12))
reduction_button.pack(pady=10)

# 导出按钮
export_button = tk.Button(root, text="导出", command=export_summary_to_excel, font=("Helvetica", 12))
export_button.pack(pady=10)

# 处理反馈
feedback_label = tk.Label(root, text="处理反馈:", font=("Helvetica", 14, "bold"))
feedback_label.pack(pady=5)
result_text = scrolledtext.ScrolledText(root, width=100, height=15, font=("Courier", 10))
result_text.pack(pady=10)

# 汇总展示
summary_label = tk.Label(root, text="汇总展示:", font=("Helvetica", 14, "bold"))
summary_label.pack(pady=5)
summary_text = scrolledtext.ScrolledText(root, width=100, height=10, font=("Courier", 10))
summary_text.pack(pady=10)
root.mainloop()