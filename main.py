import os
import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image
import threading
from tkinter import ttk
import re
from datetime import datetime
from tkinterdnd2 import *

def extract_p_number(filename):
    match = re.search(r'p(\d+)', filename.lower())
    if match:
        return int(match.group(1))
    return 0

def extract_date_from_filename(filename):
    date_match = re.search(r'(\d{4}-\d{2}-\d{2})', filename)
    if date_match:
        try:
            return datetime.strptime(date_match.group(1), "%Y-%m-%d")
        except ValueError:
            pass

    date_match = re.search(r'(\d{4})(\d{2})(\d{2})', filename)
    if date_match:
        try:
            return datetime.strptime(f"{date_match.group(1)}-{date_match.group(2)}-{date_match.group(3)}", "%Y-%m-%d")
        except ValueError:
            pass
    return None

def get_filename_without_ext(filename):
    return os.path.splitext(filename)[0]


# ============= 排序方法 =============
def sort_images(images, folder):
    sort_by = sort_var.get()
    reverse_order = order_var.get()

    if sort_by == "filename":
        images.sort(key=lambda x: x.lower(), reverse=reverse_order)

    elif sort_by == "time":
        images.sort(key=lambda x: os.path.getmtime(os.path.join(folder, x)), reverse=reverse_order)

    elif sort_by == "prefix":
        try:
            n = int(prefix_n.get())
            images.sort(key=lambda x: (
                x[:n].lower() if len(x) >= n else x.lower(),
                -extract_p_number(x) if reverse_order else extract_p_number(x)
            ), reverse=reverse_order)
        except:
            messagebox.showerror("错误", "前N位不是有效的数字")

    elif sort_by == "suffix":
        try:
            n = int(suffix_n.get())
            images.sort(key=lambda x: (
                get_filename_without_ext(x)[-n:].lower()
                if len(get_filename_without_ext(x)) >= n else
                get_filename_without_ext(x).lower(),
                extract_date_from_filename(x) or "",
                -extract_p_number(x) if reverse_order else extract_p_number(x)
            ), reverse=reverse_order)
        except:
            messagebox.showerror("错误", "后N位不是有效的数字")

    return images


# ============= 添加任务 =============
task_folders = []

def add_folder():
    folder = filedialog.askdirectory()
    if folder and folder not in task_folders:
        task_folders.append(folder)
        task_list.insert(tk.END, folder)

        task_list.selection_clear(0, tk.END)
        task_list.selection_set(0)
        update_preview(task_list.get(0))


def drop_inside(event):
    paths = root.tk.splitlist(event.data)
    for p in paths:
        if os.path.isdir(p) and p not in task_folders:
            task_folders.append(p)
            task_list.insert(tk.END, p)

    if task_list.size() > 0:
        task_list.selection_clear(0, tk.END)
        task_list.selection_set(0)
        update_preview(task_list.get(0))


# ============= 预览区域 =============
def update_preview(folder):
    preview_list.delete(0, tk.END)
    if not folder or not os.path.exists(folder):
        return

    images = [f for f in os.listdir(folder) if f.lower().endswith(('png','jpg','jpeg','webp'))]
    images = sort_images(images, folder)

    for img in images:
        preview_list.insert(tk.END, img)


def refresh_current_preview(*args):
    if task_list.curselection():
        folder = task_list.get(task_list.curselection()[0])
        update_preview(folder)


# ============= 删除任务 =============
def delete_selected_tasks():
    selected = list(task_list.curselection())

    if not selected:
        messagebox.showwarning("提示", "请先选择要删除的任务文件夹")
        return

    for i in reversed(selected):
        folder = task_list.get(i)
        task_list.delete(i)
        if folder in task_folders:
            task_folders.remove(folder)

    preview_list.delete(0, tk.END)


# ============= 批量处理 =============
def start_batch():
    threading.Thread(target=batch_worker).start()


def batch_worker():
    if not task_folders:
        messagebox.showerror("错误", "没有任务文件夹")
        return

    for folder in task_folders:

        images = [f for f in os.listdir(folder) if f.lower().endswith(('png', 'jpg', 'jpeg', 'webp'))]
        if not images:
            continue

        images = sort_images(images, folder)
        image_paths = [os.path.join(folder, f) for f in images]
        image_list = []

        template = output_name_var.get().strip()
        if not template:
            template = "{folder_name}"

        file_name = template.replace("{folder_name}", os.path.basename(folder))
        if not file_name.lower().endswith(".pdf"):
            file_name += ".pdf"

        output_path = os.path.join(os.path.dirname(folder), file_name)

        progress_bar["maximum"] = len(image_paths)
        progress_bar["value"] = 0

        for i, img_path in enumerate(image_paths):
            try:
                img = Image.open(img_path).convert("RGB")
                image_list.append(img)
                progress_bar['value'] = i + 1
            except:
                pass

        if image_list:
            image_list[0].save(output_path, save_all=True, append_images=image_list[1:])
            
            #新增：释放内存
            # 1. 关闭所有Image对象
            for img in image_list:
                img.close()
            
            # 2. 清空列表
            image_list.clear()
            
            # 3. 强制垃圾回收
            import gc
            gc.collect()

        # 重置进度条
        progress_bar['value'] = 0

    progress_bar['value'] = 0
    messagebox.showinfo("完成", "所有任务已处理完成！")


# ================== GUI ==================

root = TkinterDnD.Tk()
root.title("多任务图像合成PDF工具")

# 获取屏幕尺寸
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()

# 设置窗口尺寸
window_width = 800
window_height = 720

# 计算右下角坐标 
margin_x = 100
margin_y = 200
x_position = screen_width - window_width - margin_x
y_position = screen_height - window_height - margin_y

# 设置窗口位置和大小
root.geometry(f"{window_width}x{window_height}+{x_position}+{y_position}")
root.resizable(False, False)

style = ttk.Style(root)
style.theme_use("clam")

# ========== 拖放说明 ==========
ttk.Label(root, text="⬇ 拖动文件夹到此窗口 ⬇", font=("微软雅黑", 12)).pack(pady=8)

root.drop_target_register(DND_FILES)
root.dnd_bind('<<Drop>>', drop_inside)

# ========== 主体左右分区 ==========
main_frame = ttk.Frame(root)
main_frame.pack(fill="both", expand=True, padx=10, pady=5)

left_frame = ttk.Frame(main_frame)
left_frame.pack(side="left", fill="y", padx=5)

right_frame = ttk.Frame(main_frame)
right_frame.pack(side="right", fill="both", expand=True, padx=5)


# ========== 任务区 ==========
frame_tasks = ttk.LabelFrame(left_frame, text="任务文件夹", width=360, height=250)
frame_tasks.pack(fill="x", pady=5)

task_list = tk.Listbox(frame_tasks, height=8, selectmode=tk.EXTENDED)
task_list.pack(fill="both", expand=True, padx=5, pady=5)

btn_frame = ttk.Frame(frame_tasks)
btn_frame.pack(fill="x")

ttk.Button(btn_frame, text="添加文件夹", command=add_folder).pack(side="left", expand=True, padx=5, pady=5)
ttk.Button(btn_frame, text="删除选中", command=delete_selected_tasks).pack(side="right", expand=True, padx=5, pady=5)


# ========== 排序区 ==========
frame_sort = ttk.LabelFrame(left_frame, text="排序选项")
frame_sort.pack(fill="x", pady=5)

sort_var = tk.StringVar(value="filename")
order_var = tk.BooleanVar(value=False)
prefix_n = tk.StringVar(value="9")
suffix_n = tk.StringVar(value="10")

ttk.Radiobutton(frame_sort, text="文件名", variable=sort_var, value="filename").grid(row=0, column=0, padx=4, pady=4)
ttk.Radiobutton(frame_sort, text="时间", variable=sort_var, value="time").grid(row=0, column=1, padx=4)

ttk.Radiobutton(frame_sort, text="前N位", variable=sort_var, value="prefix").grid(row=1, column=0)
ttk.Entry(frame_sort, width=5, textvariable=prefix_n).grid(row=1, column=1)

ttk.Radiobutton(frame_sort, text="后N位", variable=sort_var, value="suffix").grid(row=1, column=2)
ttk.Entry(frame_sort, width=5, textvariable=suffix_n).grid(row=1, column=3)

ttk.Checkbutton(frame_sort, text="倒序", variable=order_var).grid(row=0, column=2, padx=6)


# ========== 输出工具 ==========
frame_output = ttk.LabelFrame(left_frame, text="输出文件名")
frame_output.pack(fill="x", pady=5)

output_name_var = tk.StringVar(value="{folder_name}")
ttk.Entry(frame_output, textvariable=output_name_var).pack(fill="x", padx=6, pady=6)
ttk.Label(frame_output, text="变量：{folder_name}").pack(pady=2)


# ========== 右侧预览 ==========
frame_preview = ttk.LabelFrame(right_frame, text="排序预览")
frame_preview.pack(fill="both", expand=True)

preview_list = tk.Listbox(frame_preview)
preview_list.pack(fill="both", expand=True, padx=5, pady=5)

task_list.bind("<<ListboxSelect>>", lambda e: refresh_current_preview())


sort_var.trace_add("write", refresh_current_preview)
order_var.trace_add("write", refresh_current_preview)
prefix_n.trace_add("write", refresh_current_preview)
suffix_n.trace_add("write", refresh_current_preview)


# ========== 进度 & 执行 ==========
bottom_frame = ttk.Frame(root)
bottom_frame.pack(fill="x", pady=8)

progress_bar = ttk.Progressbar(bottom_frame)
progress_bar.pack(fill="x", padx=10)

ttk.Button(root, text="开始批量处理", command=start_batch, width=30).pack(pady=6)

root.mainloop()