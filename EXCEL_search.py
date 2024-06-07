import os
import webbrowser
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from threading import Thread
import gettext

# 设置多语言支持
langs = {
    "en": {
        "title": "Excel Content Search Tool",
        "folder_path": "Folder Path:",
        "browse": "Browse",
        "search_content": "Search Content:",
        "search": "Search",
        "stop": "Stop",
        "open_selected_file": "Open Selected File",
        "show_log": "Show Log",
        "input_error": "Input Error",
        "input_warning": "Please fill in the folder path and search content",
        "search_stopped": "Search Stopped",
        "no_result": "No files containing the search content were found.",
        "searched_files": "Searched Files: {}/{}"
    },
    "zh": {
        "title": "Excel内容查询工具",
        "folder_path": "文件夹路径:",
        "browse": "浏览",
        "search_content": "查询内容:",
        "search": "查询",
        "stop": "停止",
        "open_selected_file": "打开选定文件",
        "show_log": "显示日志",
        "input_error": "输入错误",
        "input_warning": "请填写文件夹路径和查询内容",
        "search_stopped": "查询已停止",
        "no_result": "未找到包含查询内容的文件。",
        "searched_files": "已查询文件数: {}/{}"
    }
}

current_lang = "en"


def set_language(lang):
    global current_lang
    current_lang = lang
    app.update_language()


class ExcelSearchApp:
    def __init__(self, root):
        self.root = root
        self.root.title(langs[current_lang]["title"])

        self.folder_path = tk.StringVar()
        self.search_content = tk.StringVar()
        self.stop_search = False
        self.errors = []
        self.result_files = []

        self.create_widgets()

    def create_widgets(self):
        # 语言选择框
        tk.Label(self.root, text="Language:").grid(row=0, column=0, padx=10, pady=5)
        lang_menu = tk.OptionMenu(self.root, tk.StringVar(value=current_lang), "en", "zh", command=set_language)
        lang_menu.grid(row=0, column=1, padx=10, pady=5)

        # 文件夹路径输入框
        tk.Label(self.root, text=langs[current_lang]["folder_path"]).grid(row=1, column=0, padx=10, pady=5)
        tk.Entry(self.root, textvariable=self.folder_path, width=50).grid(row=1, column=1, padx=10, pady=5)
        tk.Button(self.root, text=langs[current_lang]["browse"], command=self.browse_folder).grid(row=1, column=2,
                                                                                                  padx=10, pady=5)

        # 查询内容输入框
        tk.Label(self.root, text=langs[current_lang]["search_content"]).grid(row=2, column=0, padx=10, pady=5)
        tk.Entry(self.root, textvariable=self.search_content, width=50).grid(row=2, column=1, padx=10, pady=5)

        # 查询按钮
        tk.Button(self.root, text=langs[current_lang]["search"], command=self.start_search).grid(row=3, column=1,
                                                                                                 padx=10, pady=5)

        # 停止按钮
        tk.Button(self.root, text=langs[current_lang]["stop"], command=self.stop_searching).grid(row=3, column=2,
                                                                                                 padx=10, pady=5)

        # 结果显示区域
        self.result_text = scrolledtext.ScrolledText(self.root, width=60, height=15)
        self.result_text.grid(row=4, column=0, columnspan=3, padx=10, pady=10)
        self.result_text.bind('<Button-1>', self.select_result)

        # 进度显示标签
        self.progress_label = tk.Label(self.root, text="", justify=tk.LEFT)
        self.progress_label.grid(row=5, column=0, columnspan=3, padx=10, pady=10)

        # 打开选定文件按钮
        tk.Button(self.root, text=langs[current_lang]["open_selected_file"], command=self.open_selected_file).grid(
            row=6, column=0, columnspan=3, padx=10, pady=5)

        # 日志按钮
        tk.Button(self.root, text=langs[current_lang]["show_log"], command=self.show_log).grid(row=7, column=0,
                                                                                               columnspan=3, padx=10,
                                                                                               pady=5)

    def update_language(self):
        self.root.title(langs[current_lang]["title"])
        for widget in self.root.winfo_children():
            widget.destroy()
        self.create_widgets()

    def browse_folder(self):
        folder_selected = filedialog.askdirectory()
        self.folder_path.set(folder_selected)

    def start_search(self):
        self.stop_search = False
        self.errors = []
        self.result_files = []
        folder = self.folder_path.get()
        content = self.search_content.get()
        if not folder or not content:
            messagebox.showwarning(langs[current_lang]["input_error"], langs[current_lang]["input_warning"])
            return

        self.result_text.delete(1.0, tk.END)
        self.progress_label.config(text="")

        # 使用线程进行查询，避免阻塞GUI
        search_thread = Thread(target=self.search, args=(folder, content))
        search_thread.start()

    def stop_searching(self):
        self.stop_search = True

    def search(self, folder, content):
        found_files = []
        total_files = len([f for f in os.listdir(folder) if f.endswith(".xlsx") or f.endswith(".xls")])
        searched_files = 0

        for filename in os.listdir(folder):
            if self.stop_search:
                self.update_result(langs[current_lang]["search_stopped"])
                return

            if filename.endswith(".xlsx") or filename.endswith(".xls"):
                file_path = os.path.join(folder, filename)
                try:
                    df = pd.read_excel(file_path, sheet_name=None)
                    for sheet_name, sheet_df in df.items():
                        sheet_df = sheet_df.applymap(self.clean_data)
                        if sheet_df.isin([content]).any().any():
                            found_files.append((filename, sheet_name, file_path))
                            self.update_result(f"文件: {filename}, 工作表: {sheet_name}\n", file_path)
                except Exception as e:
                    error_message = f"无法读取文件 {filename}：{e}"
                    self.errors.append(error_message)
                    print(error_message)

                searched_files += 1
                self.update_progress(searched_files, total_files)

        if not found_files:
            self.update_result(langs[current_lang]["no_result"])

    def clean_data(self, data):
        if isinstance(data, str):
            return data.strip()
        return str(data).strip()

    def update_progress(self, searched_files, total_files):
        progress_text = langs[current_lang]["searched_files"].format(searched_files, total_files)
        self.progress_label.config(text=progress_text)

    def update_result(self, result_text, file_path=None):
        if file_path:
            self.result_files.append(file_path)
        self.result_text.insert(tk.END, result_text)
        self.result_text.yview(tk.END)

    def select_result(self, event):
        try:
            index = self.result_text.index("@%s,%s" % (event.x, event.y))
            line = int(index.split('.')[0])
            self.selected_file = self.result_files[line - 1]
        except IndexError:
            self.selected_file = None

    def open_selected_file(self):
        if hasattr(self, 'selected_file') and self.selected_file:
            webbrowser.open(f'file://{self.selected_file}')
        else:
            messagebox.showwarning(langs[current_lang]["input_error"], langs[current_lang]["input_warning"])

    def show_log(self):
        log_window = tk.Toplevel(self.root)
        log_window.title(langs[current_lang]["show_log"])
        log_text = scrolledtext.ScrolledText(log_window, width=80, height=20)
        log_text.pack(padx=10, pady=10)
        for error in self.errors:
            log_text.insert(tk.END, error + "\n")


# 创建主窗口
root = tk.Tk()
app = ExcelSearchApp(root)
# 运行主循环
root.mainloop()