from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox

from openpyxl import load_workbook
import ttkbootstrap as ttk

from Method import Method


EXCEL_EXTENSIONS = {".xlsx", ".xlsm", ".xltx", ".xltm"}


class ProcessFrame(ttk.Frame):
    def __init__(self, root):
        super().__init__(root)

        self.config = Method.readData()

        self.start_path = tk.StringVar(value=self.config.get("start_path", ""))
        self.target_path = tk.StringVar(value=self.config.get("target_path", ""))
        self.excel = tk.StringVar(value=self.config.get("excel", ""))
        self.sheet = tk.StringVar(value=self.config.get("sheetname", "Sheet1"))
        self.file_column = tk.StringVar(value=self.config.get("file_column", "1"))
        self.original_column = tk.StringVar(value=self.config.get("original_column", "2"))
        self.new_column = tk.StringVar(value=self.config.get("new_column", "3"))
        self.output_column = tk.StringVar(value=self.config.get("output_column", "4"))
        self.check_var = tk.BooleanVar(value=False)
        self.folder_only_var = tk.BooleanVar(value=False)
        self.action_var = tk.StringVar(value="文件读取")

        self.process_input = None
        self.process_show = None
        self.sheet_combo = None

        self.processInputPage()
        self.processShowPage()

    def _sync_config_fields(self):
        self.start_path_data = self.config.get("start_path", "")
        self.target_path_data = self.config.get("target_path", "")
        self.excel_data = self.config.get("excel", "")
        self.sheet_data = self.config.get("sheetname", "Sheet1")
        self.file_column_data = self.config.get("file_column", self.config.get("start", "1"))
        self.original_column_data = self.config.get("original_column", "2")
        self.new_column_data = self.config.get("new_column", self.config.get("target", "3"))
        self.output_column_data = self.config.get("output_column", "4")

    def _add_row(self, parent, label_text, value_text):
        row = ttk.Frame(parent)
        row.pack(side=tk.TOP, anchor=tk.W, fill=tk.X, pady=8)
        ttk.Label(row, text=label_text, width=14, anchor=tk.E).pack(side=tk.LEFT)
        ttk.Label(row, text=value_text or "", wraplength=760, anchor=tk.W).pack(side=tk.LEFT, padx=8)
        return row

    def _add_input_row(self, parent, label_text, variable, width=80):
        row = ttk.Frame(parent)
        row.pack(side=tk.TOP, anchor=tk.W, fill=tk.X, pady=8)
        ttk.Label(row, text=label_text, width=14, anchor=tk.E).pack(side=tk.LEFT)
        ttk.Entry(row, textvariable=variable, width=width).pack(side=tk.LEFT, padx=8)
        return row

    def _add_folder_input_row(self, parent, label_text, variable, command):
        row = ttk.Frame(parent)
        row.pack(side=tk.TOP, anchor=tk.W, fill=tk.X, pady=8)
        ttk.Label(row, text=label_text, width=14, anchor=tk.E).pack(side=tk.LEFT)
        ttk.Entry(row, textvariable=variable, width=70, state="readonly").pack(side=tk.LEFT, padx=8)
        ttk.Button(row, text="选择", command=command).pack(side=tk.LEFT)
        return row

    def _add_excel_input_row(self, parent):
        row = ttk.Frame(parent)
        row.pack(side=tk.TOP, anchor=tk.W, fill=tk.X, pady=8)
        ttk.Label(row, text="表格文件地址:", width=14, anchor=tk.E).pack(side=tk.LEFT)
        ttk.Entry(row, textvariable=self.excel, width=70, state="readonly").pack(side=tk.LEFT, padx=8)
        ttk.Button(row, text="选择", command=self.chooseExcelFile).pack(side=tk.LEFT)
        return row

    def _add_sheet_input_row(self, parent):
        row = ttk.Frame(parent)
        row.pack(side=tk.TOP, anchor=tk.W, fill=tk.X, pady=8)
        ttk.Label(row, text="工作表名称:", width=14, anchor=tk.E).pack(side=tk.LEFT)
        self.sheet_combo = ttk.Combobox(row, textvariable=self.sheet, state="readonly", width=70)
        self.sheet_combo.pack(side=tk.LEFT, padx=8)
        self.refreshSheetOptions()
        return row

    def _desktop_path(self):
        desktop = Path.home() / "Desktop"
        if desktop.is_dir():
            return desktop

        return Path.home()

    def chooseExcelFile(self):
        file_path = filedialog.askopenfilename(
            title="选择表格文件",
            initialdir=self._desktop_path(),
            filetypes=[
                ("Excel 文件", "*.xlsx *.xlsm *.xltx *.xltm"),
                ("所有文件", "*.*"),
            ],
        )

        if file_path:
            self.excel.set(file_path)
            self.refreshSheetOptions(show_warning=True)

    def chooseStartFolder(self):
        folder_path = filedialog.askdirectory(
            title="选择文件起始地址",
            initialdir=self._desktop_path(),
            mustexist=True,
        )

        if folder_path:
            self.start_path.set(folder_path)

    def chooseTargetFolder(self):
        folder_path = filedialog.askdirectory(
            title="选择文件目标地址",
            initialdir=self._desktop_path(),
            mustexist=True,
        )

        if folder_path:
            self.target_path.set(folder_path)

    def _load_sheet_names(self, excel_path):
        workbook = load_workbook(excel_path, read_only=True)
        try:
            return workbook.sheetnames
        finally:
            workbook.close()

    def refreshSheetOptions(self, show_warning=False):
        if self.sheet_combo is None:
            return

        excel_path = Path(self.excel.get().strip())
        if not excel_path.is_file() or excel_path.suffix.lower() not in EXCEL_EXTENSIONS:
            self.sheet_combo.configure(values=())
            return

        try:
            sheet_names = self._load_sheet_names(excel_path)
        except Exception as exc:
            self.sheet_combo.configure(values=())
            if show_warning:
                messagebox.showwarning(title="警告", message=f"无法读取工作表：{exc}")
            return

        self.sheet_combo.configure(values=sheet_names)
        if not sheet_names:
            self.sheet.set("")
            return

        if self.sheet.get() not in sheet_names:
            self.sheet.set(sheet_names[0])

    def _file_names_from_excel(self):
        workbook = load_workbook(self.excel_data, read_only=True, data_only=True)
        try:
            sheet = workbook[self.sheet_data]
            file_column = int(self.file_column_data)

            for line in range(2, sheet.max_row + 1):
                value = sheet.cell(line, file_column).value
                if value is None:
                    continue

                file_name = str(value).strip()
                if file_name:
                    yield file_name
        finally:
            workbook.close()

    def _run_action(self, action_name, action):
        try:
            count = action()
        except Exception as exc:
            messagebox.showerror(title="错误", message=f"{action_name}失败：{exc}")
            return

        messagebox.showinfo(title="完成", message=f"{action_name}完成，共处理 {count} 个文件")

    # 数据及操作页面显示
    def processShowPage(self):
        self.config = Method.readData()
        self._sync_config_fields()

        if self.process_show is not None:
            self.process_show.destroy()

        self.process_show = ttk.Frame(self)

        self._add_row(self.process_show, "文件起始地址:", self.start_path_data)
        self._add_row(self.process_show, "文件目标地址:", self.target_path_data)
        self._add_row(self.process_show, "表格文件地址:", self.excel_data)
        self._add_row(self.process_show, "工作表名称:", self.sheet_data)
        self._add_row(self.process_show, "文件名列:", self.file_column_data)
        self._add_row(self.process_show, "原名称列:", self.original_column_data)
        self._add_row(self.process_show, "新名称列:", self.new_column_data)
        self._add_row(self.process_show, "写入列:", self.output_column_data)

        init_show = ttk.Frame(self.process_show)
        init_show.pack(side=tk.TOP, pady=10)
        ttk.Button(init_show, text="初始化数据表格", command=self.initializeSheetButton, width=16).pack()

        check_show = ttk.Frame(self.process_show)
        check_show.pack(side=tk.BOTTOM, anchor=tk.N, pady=40)
        ttk.Checkbutton(check_show, text="包含子目录", variable=self.check_var).pack(side=tk.LEFT, padx=10)
        ttk.Checkbutton(check_show, text="仅修改文件夹", variable=self.folder_only_var).pack(side=tk.LEFT, padx=10)

        action_show = ttk.Frame(self.process_show)
        action_show.pack(side=tk.BOTTOM, anchor=tk.S, pady=24)

        ttk.Button(action_show, text="修改数据", command=self.inputDataButton, width=12).pack(side=tk.LEFT, padx=6, pady=(28, 0))

        execute_show = ttk.Frame(action_show)
        execute_show.pack(side=tk.LEFT, padx=6)
        ttk.Combobox(
            execute_show,
            textvariable=self.action_var,
            values=("文件读取", "文件复制", "文件移动", "文件改名", "文件删除"),
            state="readonly",
            width=12,
        ).pack(side=tk.TOP, pady=(0, 6))
        ttk.Button(execute_show, text="执行", command=self.executeAction, width=12).pack(side=tk.TOP)

        self.process_show.pack(fill=tk.BOTH, expand=True, padx=24, pady=18)

    def executeAction(self):
        action_map = {
            "文件读取": self.readButton,
            "文件复制": self.copyButton,
            "文件移动": self.moveButton,
            "文件改名": self.renamingButton,
            "文件删除": self.removeButton,
        }
        action = action_map.get(self.action_var.get(), self.readButton)
        action()

    def initializeSheetButton(self):
        try:
            Method.initializeSheetMethod(
                self.excel_data,
                self.sheet_data,
                self.file_column_data,
                self.original_column_data,
                self.new_column_data,
                self.output_column_data,
            )
        except Exception as exc:
            messagebox.showerror(title="错误", message=f"初始化表格数据失败：{exc}")
            return

        messagebox.showinfo(title="完成", message="初始化表格数据完成")

    def inputDataButton(self):
        self.start_path.set(self.start_path_data)
        self.target_path.set(self.target_path_data)
        self.excel.set(self.excel_data)
        self.sheet.set(self.sheet_data)
        self.file_column.set(self.file_column_data)
        self.original_column.set(self.original_column_data)
        self.new_column.set(self.new_column_data)
        self.output_column.set(self.output_column_data)
        self.refreshSheetOptions()

        self.process_show.pack_forget()
        self.process_input.pack(fill=tk.BOTH, expand=True, padx=24, pady=18)

    def readButton(self):
        self._run_action(
            "文件读取",
            lambda: Method.readMethod(
                self.start_path_data,
                self.excel_data,
                self.file_column_data,
                self.original_column_data,
                self.new_column_data,
                self.output_column_data,
                self.sheet_data,
                self.check_var.get(),
            ),
        )

    # 文件复制方法
    def copyButton(self):
        def action():
            copied_files = []
            for file_name in self._file_names_from_excel():
                copied_files.extend(
                    Method.copyMethod(
                        self.start_path_data,
                        self.target_path_data,
                        file_name,
                        self.check_var.get(),
                    )
                )
            return len(copied_files)

        self._run_action("文件复制", action)

    # 文件移动方法调用
    def moveButton(self):
        def action():
            moved_files = []
            for file_name in self._file_names_from_excel():
                moved_files.extend(
                    Method.moveMethod(
                        self.start_path_data,
                        self.target_path_data,
                        file_name,
                        self.check_var.get(),
                    )
                )
            return len(moved_files)

        self._run_action("文件移动", action)

    # 文件改名方法调用
    def renamingButton(self):
        if self.folder_only_var.get():
            action_name = "文件夹改名"
            rename_method = Method.renamingFolderMethod
        else:
            action_name = "文件改名"
            rename_method = Method.renamingMethod

        self._run_action(
            action_name,
            lambda: len(
                rename_method(
                    self.start_path_data,
                    self.excel_data,
                    self.original_column_data,
                    self.new_column_data,
                    self.sheet_data,
                    self.check_var.get(),
                )
            ),
        )

    # 文件删除方法调用
    def removeButton(self):
        def action():
            removed_files = []
            for file_name in self._file_names_from_excel():
                removed_files.extend(Method.removeMethod(self.start_path_data, file_name, self.check_var.get()))
            return len(removed_files)

        self._run_action("文件删除", action)

    def processInputPage(self):
        self.process_input = ttk.Frame(self)

        self._add_folder_input_row(self.process_input, "文件起始地址:", self.start_path, self.chooseStartFolder)
        self._add_folder_input_row(self.process_input, "文件目标地址:", self.target_path, self.chooseTargetFolder)
        self._add_excel_input_row(self.process_input)
        self._add_sheet_input_row(self.process_input)
        self._add_input_row(self.process_input, "文件名列:", self.file_column, width=20)
        self._add_input_row(self.process_input, "原名称列:", self.original_column, width=20)
        self._add_input_row(self.process_input, "新名称列:", self.new_column, width=20)
        self._add_input_row(self.process_input, "写入列:", self.output_column, width=20)

        button_show = ttk.Frame(self.process_input)
        button_show.pack(side=tk.TOP, anchor=tk.S, pady=10)
        ttk.Button(button_show, text="确认修改", command=self.defineDataButton).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_show, text="返回", command=self.backButton).pack(side=tk.LEFT, padx=5)

    def backButton(self):
        self.process_input.pack_forget()
        self.process_show.pack(fill=tk.BOTH, expand=True, padx=24, pady=18)

    def defineDataButton(self):
        start_path = Path(self.start_path.get().strip())
        target_path = Path(self.target_path.get().strip())
        excel_path = Path(self.excel.get().strip())
        sheet_name = self.sheet.get().strip()
        file_column = self.file_column.get().strip()
        original_column = self.original_column.get().strip()
        new_column = self.new_column.get().strip()
        output_column = self.output_column.get().strip()

        if not start_path.is_dir() or not target_path.is_dir():
            messagebox.showwarning(title="警告", message="请输入正确的文件夹地址")
            return

        if not excel_path.is_file() or excel_path.suffix.lower() not in EXCEL_EXTENSIONS:
            messagebox.showwarning(title="警告", message="请使用可读取的 Excel 表格文件")
            return

        try:
            sheet_names = self._load_sheet_names(excel_path)
        except Exception as exc:
            messagebox.showwarning(title="警告", message=f"无法打开表格文件：{exc}")
            return

        if sheet_name not in sheet_names:
            messagebox.showwarning(title="警告", message="没有找到工作表")
            return

        columns = {
            "文件名列": file_column,
            "原名称列": original_column,
            "新名称列": new_column,
            "写入列": output_column,
        }
        if any(not column.isdigit() or int(column) < 1 for column in columns.values()):
            messagebox.showwarning(title="警告", message="请输入正确的数据列")
            return

        if len({int(column) for column in columns.values()}) != len(columns):
            messagebox.showwarning(title="警告", message="文件名列、原名称列、新名称列、写入列不能重复")
            return

        self.config.update(
            {
                "start_path": str(start_path),
                "target_path": str(target_path),
                "excel": str(excel_path),
                "sheetname": sheet_name,
                "file_column": file_column,
                "original_column": original_column,
                "new_column": new_column,
                "output_column": output_column,
                "start": file_column,
                "target": new_column,
            }
        )
        Method.storeData(self.config)

        self.process_input.pack_forget()
        self.processShowPage()
