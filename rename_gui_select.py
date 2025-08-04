import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk


class PhotoRenamerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("照片批量重命名工具")
        self.root.geometry("650x480")  # 窗口大小
        self.root.resizable(False, False)  # 禁止调整窗口大小

        # 设置中文字体
        self.style = ttk.Style()
        self.style.configure("TLabel", font=("SimHei", 10))
        self.style.configure("TButton", font=("SimHei", 10))
        self.style.configure("TEntry", font=("SimHei", 10))
        self.style.configure("TCombobox", font=("SimHei", 10))

        # 存储选择的路径和列名
        self.excel_path = ""
        self.photos_dir = ""
        self.source_column = ""
        self.target_column = ""
        self.excel_columns = []  # 存储Excel表头

        # 创建界面组件
        self.create_widgets()

    def create_widgets(self):
        # 标题
        title_label = ttk.Label(
            self.root,
            text="照片批量重命名工具",
            font=("SimHei", 16, "bold")
        )
        title_label.pack(pady=20)

        # Excel文件选择
        excel_frame = ttk.Frame(self.root)
        excel_frame.pack(fill=tk.X, padx=50, pady=10)

        ttk.Label(excel_frame, text="Excel文件:").pack(side=tk.LEFT, padx=5)

        self.excel_entry = ttk.Entry(excel_frame, width=40)
        self.excel_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)

        excel_btn = ttk.Button(
            excel_frame,
            text="浏览...",
            command=self.select_excel
        )
        excel_btn.pack(side=tk.LEFT, padx=5)

        # 照片文件夹选择
        photos_frame = ttk.Frame(self.root)
        photos_frame.pack(fill=tk.X, padx=50, pady=10)

        ttk.Label(photos_frame, text="照片文件夹:").pack(side=tk.LEFT, padx=5)

        self.photos_entry = ttk.Entry(photos_frame, width=40)
        self.photos_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)

        photos_btn = ttk.Button(
            photos_frame,
            text="浏览...",
            command=self.select_photos_dir
        )
        photos_btn.pack(side=tk.LEFT, padx=5)

        # 字段选择
        columns_frame = ttk.Frame(self.root)
        columns_frame.pack(fill=tk.X, padx=50, pady=10)

        ttk.Label(columns_frame, text="从重命名:").pack(side=tk.LEFT, padx=5)

        self.source_combobox = ttk.Combobox(columns_frame, state="disabled", width=15)
        self.source_combobox.pack(side=tk.LEFT, padx=5)

        ttk.Label(columns_frame, text="到:").pack(side=tk.LEFT, padx=5)

        self.target_combobox = ttk.Combobox(columns_frame, state="disabled", width=15)
        self.target_combobox.pack(side=tk.LEFT, padx=5)

        # 处理按钮
        process_btn = ttk.Button(
            self.root,
            text="开始重命名",
            command=self.process_rename
        )
        process_btn.pack(pady=20)

        # 进度条
        self.progress = ttk.Progressbar(self.root, orient=tk.HORIZONTAL, length=550, mode='determinate')
        self.progress.pack(pady=10)

        # 结果显示区域
        result_frame = ttk.LabelFrame(self.root, text="处理结果")
        result_frame.pack(fill=tk.BOTH, expand=True, padx=50, pady=10)

        self.result_text = tk.Text(result_frame, height=6, width=65, font=("SimHei", 10))
        self.result_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.result_text.config(state=tk.DISABLED)  # 初始设为只读

    def select_excel(self):
        """选择Excel文件并加载表头"""
        file_path = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[("Excel files", "*.xlsx;*.xls")]
        )
        if file_path:
            self.excel_path = file_path
            self.excel_entry.delete(0, tk.END)
            self.excel_entry.insert(0, file_path)

            # 尝试读取Excel表头
            try:
                df = pd.read_excel(self.excel_path, sheet_name=0, nrows=0)  # 只读取表头
                self.excel_columns = df.columns.tolist()

                # 更新下拉菜单
                self.source_combobox['values'] = self.excel_columns
                self.target_combobox['values'] = self.excel_columns
                self.source_combobox['state'] = 'readonly'
                self.target_combobox['state'] = 'readonly'

                # 如果有默认的姓名和学号列，自动选中
                if '姓名' in self.excel_columns:
                    self.source_combobox.current(self.excel_columns.index('姓名'))
                if '学号' in self.excel_columns:
                    self.target_combobox.current(self.excel_columns.index('学号'))

                self.log(f"已加载Excel表头：{', '.join(self.excel_columns)}")

            except Exception as e:
                messagebox.showerror("错误", f"读取Excel表头失败：{str(e)}")
                self.excel_columns = []
                self.source_combobox['state'] = 'disabled'
                self.target_combobox['state'] = 'disabled'

    def select_photos_dir(self):
        """选择照片文件夹"""
        dir_path = filedialog.askdirectory(title="选择照片文件夹")
        if dir_path:
            self.photos_dir = dir_path
            self.photos_entry.delete(0, tk.END)
            self.photos_entry.insert(0, dir_path)

    def log(self, message):
        """在结果区域显示信息"""
        self.result_text.config(state=tk.NORMAL)
        self.result_text.insert(tk.END, message + "\n")
        self.result_text.see(tk.END)  # 滚动到最新内容
        self.result_text.config(state=tk.DISABLED)
        self.root.update_idletasks()  # 刷新界面

    def process_rename(self):
        """执行重命名逻辑"""
        # 检查路径是否已选择
        if not self.excel_path:
            messagebox.showerror("错误", "请选择Excel文件")
            return
        if not self.photos_dir:
            messagebox.showerror("错误", "请选择照片文件夹")
            return
        # 检查列是否已选择
        if not self.excel_columns:
            messagebox.showerror("错误", "未加载Excel表头，请重新选择Excel文件")
            return
        try:
            self.source_column = self.source_combobox.get()
            self.target_column = self.target_combobox.get()
            if not self.source_column or not self.target_column:
                raise Exception("请选择重命名的源字段和目标字段")
            if self.source_column == self.target_column:
                raise Exception("源字段和目标字段不能相同")
        except Exception as e:
            messagebox.showerror("错误", str(e))
            return

        # 重置进度和结果
        self.progress["value"] = 0
        self.result_text.config(state=tk.NORMAL)
        self.result_text.delete(1.0, tk.END)
        self.result_text.config(state=tk.DISABLED)

        try:
            # 读取Excel并清洗数据
            self.log(f"正在读取Excel文件，使用 {self.source_column} 重命名为 {self.target_column}...")
            df = pd.read_excel(self.excel_path, sheet_name=0)

            # 检查必要的列是否存在
            if self.source_column not in df.columns or self.target_column not in df.columns:
                messagebox.showerror("错误", f"Excel中必须包含'{self.source_column}'和'{self.target_column}'列")
                return

            # 清洗数据
            df[self.source_column] = df[self.source_column].astype(str).str.replace("\u200b", "",
                                                                                    regex=False).str.strip()
            df[self.target_column] = df[self.target_column].astype(str).str.replace("\u200b", "",
                                                                                    regex=False).str.strip()

            # 创建源字段到目标字段的映射
            result_dict = dict(zip(df[self.source_column], df[self.target_column]))
            self.log(f"成功加载 {len(result_dict)} 条{self.source_column}-{self.target_column}对应关系")

            # 获取照片文件列表
            image_extensions = {'.jpg', '.jpeg', '.png', '.gif', '.bmp', '.webp'}
            files = os.listdir(self.photos_dir)
            total_files = len(files)
            renamed_count = 0
            not_matched_count = 0

            # 遍历文件并处理
            self.log("开始处理文件...")
            for i, file in enumerate(files):
                # 更新进度条
                self.progress["value"] = (i + 1) / total_files * 100
                self.root.update_idletasks()

                old_path = os.path.join(self.photos_dir, file)
                # 跳过文件夹
                if not os.path.isfile(old_path):
                    continue

                # 处理文件名
                file_name, file_ext = os.path.splitext(file)
                file_name = file_name.replace("\u200b", "").strip()

                # 检查是否为图片
                if file_ext.lower() not in image_extensions:
                    continue

                # 匹配并重命名
                if file_name in result_dict:
                    target_value = result_dict[file_name]
                    new_file = f"{target_value}{file_ext}"
                    new_path = os.path.join(self.photos_dir, new_file)

                    # 检查新文件是否存在
                    if os.path.exists(new_path):
                        self.log(f"跳过：{file} -> {new_file}（文件已存在）")
                        continue

                    os.rename(old_path, new_path)
                    self.log(f"已重命名：{file} -> {new_file}")
                    renamed_count += 1
                else:
                    not_matched_count += 1

            # 处理完成
            self.log("\n处理完成！")
            self.log(f"成功重命名：{renamed_count} 个文件")
            self.log(f"未找到匹配：{not_matched_count} 个文件")
            messagebox.showinfo("完成", f"成功重命名 {renamed_count} 个文件\n未找到匹配 {not_matched_count} 个文件")

        except Exception as e:
            self.log(f"发生错误：{str(e)}")
            messagebox.showerror("错误", f"处理失败：{str(e)}")


if __name__ == "__main__":
    root = tk.Tk()
    app = PhotoRenamerApp(root)
    root.mainloop()
