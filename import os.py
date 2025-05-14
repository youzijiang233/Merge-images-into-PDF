import os
import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image
import threading
from tkinter import ttk
import time

def select_folder():
    folder_selected = filedialog.askdirectory()
    if folder_selected:
        entry_folder.delete(0, tk.END)
        entry_folder.insert(0, folder_selected)
        update_file_list(folder_selected)
        # 自动设置默认输出文件名
        if not entry_output.get():
            default_output = os.path.join(os.path.dirname(folder_selected), 
                                        os.path.basename(folder_selected) + ".pdf")
            entry_output.delete(0, tk.END)
            entry_output.insert(0, default_output)

def update_file_list(folder):
    listbox_files.delete(0, tk.END)
    images = [f for f in os.listdir(folder) if f.lower().endswith(('png', 'jpg', 'jpeg','webp'))]
    
    # 获取排序方式和顺序
    sort_by = sort_var.get()
    reverse_order = order_var.get()
    
    # 根据选择的排序方式排序
    if sort_by == "filename":
        images.sort(key=lambda x: x.lower(), reverse=reverse_order)
    elif sort_by == "time":
        images.sort(key=lambda x: os.path.getmtime(os.path.join(folder, x)), reverse=reverse_order)
    
    for img in images:
        listbox_files.insert(tk.END, img)

def select_output():
    initial_file = entry_output.get() or os.path.join(os.path.dirname(entry_folder.get()), 
                                                    os.path.basename(entry_folder.get()) + ".pdf")
    output_path = filedialog.asksaveasfilename(
        defaultextension=".pdf", 
        filetypes=[("PDF files", "*.pdf")],
        initialfile=os.path.basename(initial_file)
    )
    if output_path:
        entry_output.delete(0, tk.END)
        entry_output.insert(0, output_path)

def merge_images_to_pdf():
    thread = threading.Thread(target=merge_images_worker)
    thread.start()

def merge_images_worker():
    folder = entry_folder.get()
    output_path = entry_output.get()
    
    if not folder:
        messagebox.showerror("错误", "请选择文件夹")
        return
    
    # 如果没有指定输出路径，使用默认路径
    if not output_path:
        output_path = os.path.join(os.path.dirname(folder), os.path.basename(folder) + ".pdf")
        entry_output.delete(0, tk.END)
        entry_output.insert(0, output_path)
    
    images = [f for f in os.listdir(folder) if f.lower().endswith(('png', 'jpg', 'jpeg','webp'))]
    
    if not images:
        messagebox.showerror("错误", "指定文件夹内没有图像文件")
        return
    
    # 获取排序方式和顺序
    sort_by = sort_var.get()
    reverse_order = order_var.get()
    
    # 根据选择的排序方式排序
    if sort_by == "filename":
        images.sort(key=lambda x: x.lower(), reverse=reverse_order)
    elif sort_by == "time":
        images.sort(key=lambda x: os.path.getmtime(os.path.join(folder, x)), reverse=reverse_order)
    
    image_paths = [os.path.join(folder, f) for f in images]
    
    image_list = []
    progress_bar['maximum'] = len(image_paths)
    for i, img_path in enumerate(image_paths):
        try:
            img = Image.open(img_path).convert("RGB")
            image_list.append(img)
            progress_bar['value'] = i + 1
            root.update_idletasks()
        except Exception as e:
            print(f"无法加载图像 {img_path}: {e}")
    
    if not image_list:
        messagebox.showerror("错误", "没有可用的图像文件")
        progress_bar['value'] = 0
        return
    
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
root.geometry("500x550")

frame = tk.Frame(root)
frame.pack(pady=10)

# 文件夹选择
btn_folder = tk.Button(frame, text="选择文件夹", command=select_folder)
btn_folder.grid(row=0, column=0, padx=5)
entry_folder = tk.Entry(frame, width=50)
entry_folder.grid(row=0, column=1)

# 排序选项
sort_frame = tk.Frame(root)
sort_frame.pack(pady=5)

tk.Label(sort_frame, text="排序方式:").pack(side=tk.LEFT)
sort_var = tk.StringVar(value="filename")
tk.Radiobutton(sort_frame, text="文件名", variable=sort_var, value="filename", 
              command=lambda: update_file_list(entry_folder.get()) if entry_folder.get() else None).pack(side=tk.LEFT)
tk.Radiobutton(sort_frame, text="修改时间", variable=sort_var, value="time", 
              command=lambda: update_file_list(entry_folder.get()) if entry_folder.get() else None).pack(side=tk.LEFT)

order_var = tk.BooleanVar(value=False)
tk.Checkbutton(sort_frame, text="倒序", variable=order_var, 
              command=lambda: update_file_list(entry_folder.get()) if entry_folder.get() else None).pack(side=tk.LEFT)

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