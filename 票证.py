import tkinter as tk
from tkinter import filedialog, messagebox
from tkcalendar import DateEntry
import pandas as pd

# 主界面
interface = tk.Tk()
interface.title("票证之星")
interface.geometry("800x800")

# 主界面header
label = tk.Label(interface, text="操作界面")
label.pack(pady=10)

# 存储导入的文件和处理结果
imported_volume_df = None
imported_income_df = None
processed_volume_df = None
processed_income_df = None
merged_df = None

# 点击导入业务量表
def on_import_volume_click():
    global imported_volume_df
    # 根据Path打开文件
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
    if file_path:
        try:
            # 读取Excel文件
            imported_volume_df = pd.read_excel(file_path)
            label.config(text=f"{file_path}已导入业务量表")
            text_widget1.delete(1.0, tk.END)
            text_widget1.insert(tk.END, imported_volume_df.to_string(index=False))
        except Exception as e:
            messagebox.showerror("Error", f"无法导入文件: {str(e)}")

# 点击导入业务收入表
def on_import_revenue_click():
    global imported_income_df
    # 根据Path打开文件
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
    if file_path:
        try:
            # 读取Excel文件
            imported_income_df = pd.read_excel(file_path)
            # 尝试解析日期列
            if '日期' in imported_income_df.columns:
                # 仅当日期列为浮点数时才进行转换
                if pd.api.types.is_numeric_dtype(imported_income_df['日期']):
                    imported_income_df['日期'] = pd.to_datetime(imported_income_df['日期'], origin='1899-12-30', unit='D')
            label.config(text=f"{file_path}已导入业务收入表")
            text_widget2.delete(1.0, tk.END)
            text_widget2.insert(tk.END, imported_income_df.to_string(index=False))
        except Exception as e:
            messagebox.showerror("Error", f"无法导入文件: {str(e)}")

# 获取并筛选数据
def filter_data(df):
    try:
        # 获取用户选择的日期范围
        start_date = start_date_entry.get_date()
        end_date = end_date_entry.get_date()

        # 确保日期列为datetime格式
        if pd.api.types.is_numeric_dtype(df['日期']):
            df['日期'] = pd.to_datetime(df['日期'], origin='1899-12-30', unit='D')

        # 筛选日期范围
        mask = (pd.to_datetime(df['日期'])>= pd.to_datetime(start_date)) & (pd.to_datetime(df['日期']) <= pd.to_datetime(end_date))
        filtered_df = df[mask]

        # 筛选收费代码
        valid_chargetypes = ["福费廷", "福费廷分层收入", "福费廷转卖投资收益", "国内信用证（承兑）", "国内信用证（付款）", "国内信用证（开立）", "国内信用证（审单）", "国内信用证（通知）", "国内信用证（修改/取证）"]
        filtered_df = filtered_df[filtered_df['收费代码名称'].isin(valid_chargetypes)]
        return filtered_df
    
    except Exception as e:
        messagebox.showerror("Error", f"筛选条件有误 {str(e)}")
        return None

# 处理并显示结果
def process_and_display():
    global imported_volume_df, imported_income_df, processed_volume_df, processed_income_df, merged_df
    try:
        # 业务量表处理逻辑
        if imported_volume_df is not None:
            filtered_volume_df = imported_volume_df.copy()

            # 筛选产品名称并集合每个三级机构名称对应的折人民币放款金额
            filtered_volume_df = filtered_volume_df[(filtered_volume_df['产品名称'] == '国内信用证') | (filtered_volume_df['产品名称'] == '国内信用证福费廷')]
            processed_volume_df = filtered_volume_df.groupby('三级机构名称')['折人民币放款金额'].sum().reset_index()
            
            # 重命名-统一命名中支和三级机构名称为中支/三级机构
            if '三级机构名称' in processed_volume_df.columns:
                processed_volume_df.rename(columns={'三级机构名称': '中支/三级机构'}, inplace=True)
        
         # 业务收入表处理逻辑
        if imported_income_df is not None:
            filtered_income_df = filter_data(imported_income_df)
            if filtered_income_df is not None:
                processed_income_df = filtered_income_df.groupby('中支')['折合人民币金额'].sum().reset_index()
                
                # 重命名-统一命名中支和三级机构名称为中支/三级机构
                if '中支' in processed_income_df.columns:
                    processed_income_df.rename(columns={'中支': '中支/三级机构'}, inplace=True)
        
        # 未导入表单处理
        else:
            messagebox.showwarning("Warning", "请先导入业务量表或业务收入表！")
        
        # 合并两个处理后的表格
        if processed_volume_df is not None and processed_income_df is not None:
            # 合并两个表格[根据中支/三级机构列进行outer join]
            merged_df = pd.merge(processed_income_df, processed_volume_df, on='中支/三级机构', how='outer')
            
            # 更新显示最终处理结果
            text_widget_final.delete(1.0, tk.END)
            text_widget_final.insert(tk.END, merged_df.to_string(index=False))
    
    except Exception as e:
        messagebox.showerror("Error", f"无法处理表单: {str(e)}")

# 导出处理结果函数
def export_to_excel():
    global merged_df
    try:
        if merged_df is not None:
            # 导出合并后的最终表格
            final_file_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                           filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
            if final_file_path:
                merged_df.to_excel(final_file_path, index=False)
                messagebox.showinfo("Success", "最终处理结果已成功导出！")
        else:
            messagebox.showwarning("Warning", "没有可导出的最终处理结果！")

    except Exception as e:
        messagebox.showerror("Error", f"无法导出表单: {str(e)}")

# 点击删除
def on_delete_click():
    global imported_volume_df, imported_income_df, processed_volume_df, processed_income_df, merged_df
    label.config(text="已删除")
    text_widget1.delete(1.0, tk.END)
    text_widget2.delete(1.0, tk.END)
    text_widget_final.delete(1.0, tk.END)
    imported_volume_df = None
    imported_income_df = None
    processed_volume_df = None
    processed_income_df = None
    merged_df = None

# 按钮设计
button_frame = tk.Frame(interface)
button_frame.pack(side=tk.LEFT, padx=10, pady=10)

# 通用按钮样式
button_style = {"width": 15, "height": 3, "relief": "raised", "bd": 4}

# 导入业务量表按钮
import_volume_button = tk.Button(button_frame, text="导入业务量表", command=on_import_volume_click, **button_style)
import_volume_button.grid(row=0, column=0, padx=5, pady=5)

# 导入业务收入表按钮
import_button = tk.Button(button_frame, text="导入业务收入表", command=on_import_revenue_click, **button_style)
import_button.grid(row=1, column=0, padx=5, pady=5)

# 日期选择标签和日期选择器
start_date_label = tk.Label(button_frame, text="开始日期:")
start_date_label.grid(row=2, column=0, padx=5, pady=5)

start_date_entry = DateEntry(button_frame, width=12, background='darkgray', foreground='black', borderwidth=2)
start_date_entry.grid(row=3, column=0, padx=5, pady=5)

end_date_label = tk.Label(button_frame, text="结束日期:")
end_date_label.grid(row=4, column=0, padx=5, pady=5)

end_date_entry = DateEntry(button_frame, width=12, background='darkgray', foreground='black', borderwidth=2)
end_date_entry.grid(row=5, column=0, padx=5, pady=5)

# 处理按钮
process_button = tk.Button(button_frame, text="处理", command=process_and_display, **button_style)
process_button.grid(row=6, column=0, padx=5, pady=5)

# 导出处理结果按钮
export_button = tk.Button(button_frame, text="导出处理结果", command=export_to_excel, **button_style)
export_button.grid(row=7, column=0, padx=5, pady=5)

# 删除按钮
delete_button = tk.Button(button_frame, text="删除", command=on_delete_click, **button_style)
delete_button.grid(row=8, column=0, padx=5, pady=5)

# 文字框设计
text_frame = tk.Frame(interface)


text_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=10, pady=10)

# 添加预览导入文字标签 (业务量表)
label_preview_import_volume = tk.Label(text_frame, text="预览导入 (业务量表)")
label_preview_import_volume.pack(side=tk.TOP, padx=10, pady=5)

# 添加一个文字框以显示导入的业务量表
text_widget1 = tk.Text(text_frame, wrap='word', width=20, height=10) 
text_widget1.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

# 添加预览导入文字标签 (业务收入表)
label_preview_import_income = tk.Label(text_frame, text="预览导入 (业务收入表)")
label_preview_import_income.pack(side=tk.TOP, padx=10, pady=5)

# 添加一个文字框以显示导入的业务收入表
text_widget2 = tk.Text(text_frame, wrap='word', width=20, height=10)  
text_widget2.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

# 添加预览最终处理结果标签
label_preview_final = tk.Label(text_frame, text="预览导出")
label_preview_final.pack(side=tk.TOP, padx=10, pady=5)

# 添加一个文字框以显示处理后的最终结果
text_widget_final = tk.Text(text_frame, wrap='word', width=50, height=10) 
text_widget_final.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

# 主程序运行
interface.mainloop()






