import tkinter as tk
from tkinter import filedialog, messagebox
from PyPDF2 import PdfReader, PdfWriter

def add_file():
    file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    if file_path:
        file_listbox.insert(tk.END, file_path)

def remove_selected_file():
    selected_indices = file_listbox.curselection()
    for index in selected_indices[::-1]:  # Reverse the list to avoid index shifting issues
        file_listbox.delete(index)

def clear_list():
    file_listbox.delete(0, tk.END)

def set_output_file():
    output_file_path = filedialog.asksaveasfilename(defaultextension=".pdf",
                                                    filetypes=[("PDF files", "*.pdf")])
    if output_file_path:
        output_file_label.config(text=f"Output File: {output_file_path}")
        root.output_file_path = output_file_path

def move_up():
    selected_indices = file_listbox.curselection()
    if not selected_indices:
        return
    for index in selected_indices:
        if index == 0:
            continue
        file_path = file_listbox.get(index)
        file_listbox.delete(index)
        file_listbox.insert(index - 1, file_path)
        file_listbox.select_set(index - 1)

def move_down():
    selected_indices = file_listbox.curselection()
    if not selected_indices:
        return
    for index in selected_indices[::-1]:
        if index == file_listbox.size() - 1:
            continue
        file_path = file_listbox.get(index)
        file_listbox.delete(index)
        file_listbox.insert(index + 1, file_path)
        file_listbox.select_set(index + 1)

def merge_pdfs():
    if not file_listbox.size():
        messagebox.showwarning("No Files", "Please add PDF files to merge.")
        return

    if not root.output_file_path:
        messagebox.showwarning("No Output File", "Please set the output file location.")
        return

    pdf_writer = PdfWriter()

    for idx in range(file_listbox.size()):
        pdf_path = file_listbox.get(idx)
        pdf_reader = PdfReader(pdf_path)
        for page_num in range(len(pdf_reader.pages)):
            pdf_writer.add_page(pdf_reader.pages[page_num])

    with open(root.output_file_path, 'wb') as out_file:
        pdf_writer.write(out_file)

    messagebox.showinfo("Success", "PDF files merged successfully!")

def split_pdfs():
    if not file_listbox.size():
        messagebox.showwarning("No Files", "Please add PDF files to split.")
        return

    if not root.output_file_path:
        messagebox.showwarning("No Output File", "Please set the output file location.")
        return

    for idx in range(file_listbox.size()):
        pdf_path = file_listbox.get(idx)
        pdf_reader = PdfReader(pdf_path)
        for page_num in range(len(pdf_reader.pages)):
            pdf_writer = PdfWriter()
            pdf_writer.add_page(pdf_reader.pages[page_num])
            output_filename = f"{root.output_file_path}_page_{page_num + 1}.pdf"
            with open(output_filename, 'wb') as out_file:
                pdf_writer.write(out_file)

    messagebox.showinfo("Success", "PDF files split successfully!")

# 创建主窗口
root = tk.Tk()
root.title("PDF 拆分合并器")
root.output_file_path = None  # Initialize the output file path

# 创建文件列表框
file_listbox = tk.Listbox(root, selectmode=tk.SINGLE)
file_listbox.pack(expand=1, fill='both', padx=10, pady=10)

# 创建按钮框架
button_frame = tk.Frame(root)
button_frame.pack(fill='x')

# 创建添加文件按钮
add_button = tk.Button(button_frame, text="添加文件", command=add_file)
add_button.pack(side='left', padx=5, pady=5)

# 创建移除选中文件按钮
remove_button = tk.Button(button_frame, text="移除选择文件", command=remove_selected_file)
remove_button.pack(side='left', padx=5, pady=5)

# 创建清空列表按钮
clear_button = tk.Button(button_frame, text="清除列表", command=clear_list)
clear_button.pack(side='left', padx=5, pady=5)

# 创建设置输出文件按钮
set_output_button = tk.Button(button_frame, text="设置输出文件名字", command=set_output_file)
set_output_button.pack(side='left', padx=5, pady=5)

# 创建移动文件顺序按钮
move_up_button = tk.Button(button_frame, text="向上移动", command=move_up)
move_up_button.pack(side='left', padx=5, pady=5)

move_down_button = tk.Button(button_frame, text="向下移动", command=move_down)
move_down_button.pack(side='left', padx=5, pady=5)

# 创建输出文件标签
output_file_label = tk.Label(root, text="输出文件名: 空")
output_file_label.pack(pady=10)

# 创建确认按钮框架
confirm_frame = tk.Frame(root)
confirm_frame.pack(fill='x')

# 创建合并PDF按钮
merge_button = tk.Button(confirm_frame, text="合并 PDFs", command=merge_pdfs)
merge_button.pack(side='left', padx=5, pady=5)

# 创建拆分PDF按钮
split_button = tk.Button(confirm_frame, text="拆分 PDFs", command=split_pdfs)
split_button.pack(side='left', padx=5, pady=5)

# 运行主循环
root.mainloop()