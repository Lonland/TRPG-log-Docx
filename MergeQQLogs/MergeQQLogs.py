import tkinter as tk
from tkinter import filedialog, messagebox

import reQQLogs
import toDocx

class Step1(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller

        label = tk.Label(self, text="第 1 步：选择要处理的 TXT 文件", font=("Helvetica", 14))
        label.pack(pady=20)

        self.file_listbox = tk.Listbox(self, width=60, height=6)
        self.file_listbox.pack(pady=10)

        choose_btn = tk.Button(self, text="选择文件", command=self.choose_files)
        choose_btn.pack(pady=5)

        next_btn = tk.Button(self, text="下一步", command=self.go_next)
        next_btn.pack(pady=20)

    def choose_files(self):
        file_paths = filedialog.askopenfilenames(
            title="选择 TXT 文件",
            filetypes=[("Text Files", "*.txt")]
        )
        if file_paths:
            self.controller.selected_files = list(file_paths)
            self.file_listbox.delete(0, tk.END)
            for path in file_paths:
                self.file_listbox.insert(tk.END, path)

    def go_next(self):
        if not hasattr(self.controller, 'selected_files') or not self.controller.selected_files:
            messagebox.showwarning("未选择文件", "请至少选择一个 TXT 文件。")
            return
        
        self.controller.all_messages = reQQLogs.all_file_parse(self.controller.selected_files)
        self.controller.name_list = unique_names = list({msg["name"] for msg in self.controller.all_messages})
        print(self.controller.name_list)

        step2_frame = Step2(self.master, self.controller)
        self.controller.frames[Step2] = step2_frame
        step2_frame.grid(row=0, column=0, sticky="nsew")

        # ✅ 切换到 Step2
        self.controller.show_frame(Step2)
        
class Step2(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller
        self.entries = {}

        label = tk.Label(self, text="第 2 步：为每个名字指定颜色", font=("Helvetica", 14))
        label.pack(pady=10)

        frame_inner = tk.Frame(self)
        frame_inner.pack(pady=10)

        # 遍历名字列表，创建输入框
        for name in controller.name_list:
            row = tk.Frame(frame_inner)
            row.pack(anchor='w', pady=2)

            label_name = tk.Label(row, text=name, width=15, anchor='w')
            label_name.pack(side='left')

            entry_color = tk.Entry(row, width=10)
            entry_color.pack(side='left')
            entry_color.insert(0, "#000000")  # 默认黑色，可自定义

            self.entries[name] = entry_color

        next_btn = tk.Button(self, text="下一步", command=self.go_next)
        next_btn.pack(pady=20)

        back_btn = tk.Button(self, text="返回", command=lambda: controller.show_frame(Step1))
        back_btn.pack()

    def go_next(self):
        self.controller.color_map_hex = {}

        for name, entry in self.entries.items():
            value = entry.get().strip()
            if not value.startswith("#") or len(value) != 7:
                messagebox.showwarning("格式错误", f"名字 {name} 的颜色必须是 #RRGGBB 格式")
                return
            self.controller.color_map_hex[name] = value
        
        toDocx.export_to_docx(self.controller.all_messages, self.controller.color_map_hex, "output.docx")

        self.controller.show_frame(Step3)

# 主窗口运行包装
if __name__ == "__main__":
    root = tk.Tk()
    root.title("Step1 Demo")
    root.geometry("600x400")

    # 创建一个伪 controller，模拟 MultiStepApp
    class DummyController:
        def __init__(self, master):
            self.master = master
            self.selected_files = []
            self.name_list = []
            self.color_map_hex = {}
            self.frames = {}

        def show_frame(self, cls):
            frame = self.frames[cls]
            frame.tkraise()

    controller = DummyController(root)

    # 创建 Step1 页面
    step1 = Step1(root, controller)
    step1.grid(row=0, column=0, sticky="nsew")
    controller.frames[Step1] = step1

    # 创建空 Step2 作为跳转目标
    step2 = tk.Frame(root)
    tk.Label(step2, text="Step 2 Placeholder").pack()
    step2.grid(row=0, column=0, sticky="nsew")
    controller.frames[Step2] = step2

    controller.show_frame(Step1)
    root.mainloop()

