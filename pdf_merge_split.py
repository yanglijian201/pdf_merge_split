import tkinter as tk
from tkinter import filedialog, messagebox
from tkinterdnd2 import DND_FILES, TkinterDnD
from PyPDF2 import PdfReader, PdfWriter

def add_file(file_path=None):
    if not file_path:
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

def drop(event):
    files = root.tk.splitlist(event.data)
    for file in files:
        if file.lower().endswith('.pdf'):
            add_file(file)

def on_drag_start(event):
    widget = event.widget
    widget._drag_data = {"x": event.x, "y": event.y, "item": widget.nearest(event.y)}

def on_drag_motion(event):
    widget = event.widget
    x, y = event.x, event.y
    widget._drag_data["x"] = x
    widget._drag_data["y"] = y

def on_drag_end(event):
    widget = event.widget
    item = widget._drag_data["item"]
    new_index = widget.nearest(event.y)
    if new_index != item:
        file_path = widget.get(item)
        widget.delete(item)
        widget.insert(new_index, file_path)
        widget.select_set(new_index)

# 创建主窗口
root = TkinterDnD.Tk()
root.title("PDF Merger and Splitter")
root.output_file_path = None  # Initialize the output file path

# 创建文件列表框
file_listbox = tk.Listbox(root, selectmode=tk.SINGLE)
file_listbox.pack(expand=1, fill='both', padx=10, pady=10)

# Enable drag-and-drop functionality
file_listbox.drop_target_register(DND_FILES)
file_listbox.dnd_bind('<<Drop>>', drop)

# Enable drag-and-drop reordering within the listbox
file_listbox.bind("<ButtonPress-1>", on_drag_start)
file_listbox.bind("<B1-Motion>", on_drag_motion)
file_listbox.bind("<ButtonRelease-1>", on_drag_end)

# 创建按钮框架
button_frame = tk.Frame(root)
button_frame.pack(fill='x')

# 创建添加文件按钮
add_button = tk.Button(button_frame, text="Add File", command=add_file)
add_button.pack(side='left', padx=5, pady=5)

# 创建移除选中文件按钮
remove_button = tk.Button(button_frame, text="Remove Selected", command=remove_selected_file)
remove_button.pack(side='left', padx=5, pady=5)

# 创建清空列表按钮
clear_button = tk.Button(button_frame, text="Clear List", command=clear_list)
clear_button.pack(side='left', padx=5, pady=5)

# 创建设置输出文件按钮
set_output_button = tk.Button(button_frame, text="Set Output File", command=set_output_file)
set_output_button.pack(side='left', padx=5, pady=5)

# 创建输出文件标签
output_file_label = tk.Label(root, text="Output File: None")
output_file_label.pack(pady=10)

# 创建确认按钮框架
confirm_frame = tk.Frame(root)
confirm_frame.pack(fill='x')

# 创建合并PDF按钮
merge_button = tk.Button(confirm_frame, text="Merge PDFs", command=merge_pdfs)
merge_button.pack(side='left', padx=5, pady=5)

# 创建拆分PDF按钮
split_button = tk.Button(confirm_frame, text="Split PDFs", command=split_pdfs)
split_button.pack(side='left', padx=5, pady=5)

# 运行主循环
root.mainloop()