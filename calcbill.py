import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
from tkinterdnd2 import DND_FILES, TkinterDnD
import pandas as pd
import openpyxl
import os
import re 

class BillCalculator:
    def __init__(self, master):
        self.master = master
        master.title("批量账单计算器")
        master.geometry("600x400")  # 增加窗口大小
        master.configure(bg="#f0f0f0")  # 设置背景颜色

        self.files = []
        self.service_fees = {}  # 存储每个文件的服务费
        
        # 创建一个框架用于按钮
        self.button_frame = tk.Frame(master, bg="#f0f0f0")
        self.button_frame.pack(pady=10)

        # 设置按钮样式
        self.select_button = tk.Button(self.button_frame, text="选择Excel文件", command=self.select_files, 
                                        bg="#4CAF50", fg="white", font=("Arial", 12), padx=10, pady=5)
        self.select_button.pack(side=tk.LEFT, padx=5)  # 横向排列

        self.process_button = tk.Button(self.button_frame, text="处理所选文件", command=self.process_files, 
                                         bg="#2196F3", fg="white", font=("Arial", 12), padx=10, pady=5)
        self.process_button.pack(side=tk.LEFT, padx=5)  # 横向排列

        self.remove_button = tk.Button(self.button_frame, text="移除Excel文件", command=self.remove_files, 
                                        bg="#f44336", fg="white", font=("Arial", 12), padx=10, pady=5)
        self.remove_button.pack(side=tk.LEFT, padx=5)  # 横向排列

        self.file_frame = tk.Frame(master)
        self.file_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # 设置列表框样式
        self.file_listbox = tk.Listbox(self.file_frame, width=50, height=10, 
                                        bg="#ffffff", font=("Arial", 10), selectbackground="#cce5ff")
        self.file_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.scrollbar = tk.Scrollbar(self.file_frame, orient=tk.VERTICAL)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.file_listbox.config(yscrollcommand=self.scrollbar.set)
        self.scrollbar.config(command=self.file_listbox.yview)

        # 配置文件列表框以接受拖放
        self.file_listbox.drop_target_register(DND_FILES)
        self.file_listbox.dnd_bind('<<Drop>>', self.drop_files)

        self.result_label = tk.Label(master, text="", bg="#f0f0f0", font=("Arial", 10))
        self.result_label.pack(pady=10)

        # 绑定双击事件
        self.file_listbox.bind('<Double-1>', self.edit_service_fee)  # 双击编辑服务费

    def select_files(self):
        new_files = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx *.xls")])
        self.add_files(new_files)

    def drop_files(self, event):
        files = self.file_listbox.tk.splitlist(event.data)
        excel_files = [f for f in files if f.lower().endswith(('.xlsx', '.xls'))]
        self.add_files(excel_files)

    def add_files(self, new_files):
        for file in new_files:
            if file not in self.files:
                self.files.append(file)
                self.service_fees[file] = 6.5  # 默认服务费
                self.update_file_list()

    def update_file_list(self):
        self.file_listbox.delete(0, tk.END)
        for file in self.files:
            self.file_listbox.insert(tk.END, f"{os.path.basename(file)} - 服务费: {self.service_fees[file]}")

    def edit_service_fee(self, event):
        selection = self.file_listbox.curselection()
        if selection:
            index = selection[0]
            file = self.files[index]
            new_fee = simpledialog.askfloat("修改服务费", f"请输入 {os.path.basename(file)} 的新服务费:", 
                                            initialvalue=self.service_fees[file])
            if new_fee is not None:
                self.service_fees[file] = new_fee
                self.update_file_list()

    def remove_files(self):
        selection = self.file_listbox.curselection()
        if selection:
            index = selection[0]
            file = self.files[index]
            self.files.pop(index)  # 移除选中的文件
            del self.service_fees[file]  # 同时移除服务费
            self.update_file_list()  # 更新文件列表

    def process_files(self):
        if not self.files:
            messagebox.showwarning("警告", "请先选择Excel文件")
            return

        for file in self.files:
            self.process_single_file(file, self.service_fees[file])

        messagebox.showinfo("成功", "所有文件处理完成")

    def process_single_file(self, file_path, per_order_service_fee):
        try:
            df = pd.read_excel(file_path,dtype={"订单编号":str,"快递单号":str})
            # 确保 "商家" 列存在
            if '商家' not in df.columns:
                df['商家'] = None

            # 检查是否存在“备注”列，如果不存在则插入
            if '备注' not in df.columns:
                df['备注'] = ""  # 插入空的"备注"列

            # 增加Q列,根据条件填充内容
            df['Q'] = df.apply(lambda row: row['卖家备注'] if pd.notna(row['卖家备注']) and row['卖家备注'].strip() != "" else row['商品编码'], axis=1)
            
            # 从Q列提取单价
            def extract_price(s):
                s = s.lower()
                match = re.search(r'-p(\d+)', s)  # 匹配最后一个'-p'后面的数字
                price= match.group(1) if match else None  # 返回匹配的数字
                print(s,price)
                return price  # 返回匹配的数字

            df['单价'] = df['Q'].apply(extract_price)
            
            # 将单价转换为数值类型
            df['单价'] = pd.to_numeric(df['单价'], errors='coerce')
            
            # 计算订单金额
            df['订单金额'] = df['单价'] * df['商品数量']
            
            # 过滤掉"备注"列中包含"次日发货"的行
            df_filtered = df[~df['备注'].str.contains("次日发货", na=False)]
            df_next_day = df[df['备注'].str.contains("次日发货", na=False)]  # 单独提取"次日发货"的行
            
            # 计算汇总数据
            total_order_amount = df_filtered['订单金额'].sum()
            express_count = df_filtered['快递单号'].nunique()
            total_service_fee = per_order_service_fee * express_count
            total_payable = total_order_amount + total_service_fee

            # 创建汇总表格
            summary_data = {
                '项目': ['货值总额', '快递数量', '服务费', '需支付总额'],
                '金额': [total_order_amount, express_count, total_service_fee, total_payable]
            }
            summary_df = pd.DataFrame(summary_data)

            # 将汇总表格和原始数据写入同一个Excel文件的不同工作表，调整顺序
            output_path = file_path.rsplit('.', 1)[0] + '_processed.xlsx'
            with pd.ExcelWriter(output_path) as writer:
                summary_df.to_excel(writer, sheet_name='对账金额', index=False)
                df_filtered.to_excel(writer, sheet_name='订单数据', index=False)  # 写入过滤后的数据
                df_next_day.to_excel(writer, sheet_name='次日发货数据', index=False)  # 写入“次日发货”的数据

            self.result_label.config(text=f"处理完成: {os.path.basename(file_path)}\n"
                                          f"新文件已保存为: {os.path.basename(output_path)}")
            self.master.update()  # 更新GUI，显示最新处理的文件信息
            
        except Exception as e:
            messagebox.showerror("错误", f"处理文件 {os.path.basename(file_path)} 时出错: {str(e)}")

root = TkinterDnD.Tk()
app = BillCalculator(root)
app.file_listbox.bind('<Double-1>', app.edit_service_fee)  # 双击编辑服务费
root.mainloop()
