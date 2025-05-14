import os
import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image
import threading
from tkinter import ttk

def select_folder():
    folder_selected = filedialog.askdirectory()
    if folder_selected:
        entry_folder.delete(0, tk.END)
        entry_folder.insert(0, folder_selected)
        update_file_list(folder_selected)

def update_file_list(folder):
    listbox_files.delete(0, tk.END)
    images = [f for f in os.listdir(folder) if f.lower().endswith(('png', 'jpg', 'jpeg'))]
    for img in sorted(images):
        listbox_files.insert(tk.END, img)

def select_output():
    output_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
    if output_path:
        entry_output.delete(0, tk.END)
        entry_output.insert(0, output_path)

def merge_images_to_pdf():
    thread = threading.Thread(target=merge_images_worker)
    thread.start()

def merge_images_worker():
    folder = entry_folder.get()
    output_path = entry_output.get()
    
    if not folder or not output_path:
        messagebox.showerror("错误", "请选择文件夹和输出文件名")
        return
    
    images = [os.path.join(folder, f) for f in sorted(os.listdir(folder)) if f.lower().endswith(('png', 'jpg', 'jpeg'))]
    
    if not images:
        messagebox.showerror("错误", "指定文件夹内没有图像文件")
        return
    
    image_list = []
    progress_bar['maximum'] = len(images)
    for i, img_path in enumerate(images):
        img = Image.open(img_path).convert("RGB")
        image_list.append(img)
        progress_bar['value'] = i + 1
        root.update_idletasks()
    
    try:
        image_list[0].save(output_path, save_all=True, append_images=image_list[1:])
        messagebox.showinfo("成功", f"PDF 已保存至 {output_path}")
    except Exception as e:
        messagebox.showerror("错误", f"无法保存 PDF: {e}")
    finally:
        progress_bar['value'] = 0

# 创建GUI窗口
root = tk.Tk()
root.title("图像合并为PDF")
root.geometry("500x450")

frame = tk.Frame(root)
frame.pack(pady=10)

# 文件夹选择
btn_folder = tk.Button(frame, text="选择文件夹", command=select_folder)
btn_folder.grid(row=0, column=0, padx=5)
entry_folder = tk.Entry(frame, width=50)
entry_folder.grid(row=0, column=1)

# 文件列表
listbox_files = tk.Listbox(root, width=60, height=10)
listbox_files.pack(pady=10)

# 输出路径选择
frame_output = tk.Frame(root)
frame_output.pack(pady=10)
btn_output = tk.Button(frame_output, text="选择输出文件", command=select_output)
btn_output.grid(row=0, column=0, padx=5)
entry_output = tk.Entry(frame_output, width=50)
entry_output.grid(row=0, column=1)

# 进度条
progress_bar = ttk.Progressbar(root, length=400, mode='determinate')
progress_bar.pack(pady=10)

# 合并按钮
btn_merge = tk.Button(root, text="合并为PDF", command=merge_images_to_pdf, bg="green", fg="white")
btn_merge.pack(pady=10)

root.mainloop()
