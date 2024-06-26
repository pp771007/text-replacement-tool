import tkinter as tk
from tkinter import messagebox
import configparser
import os
import pyperclip
import base64
import win32clipboard

# 創建主視窗
window = tk.Tk()
window.title("文本替換工具")

# 定義icon.ico文件的路徑
icon_path = "icon.ico"

# 檢查圖標文件是否存在
if os.path.exists(icon_path):
    # 如果文件存在，設置應用程式窗口的圖標
    window.iconbitmap(icon_path)

# 創建配置文件
config = configparser.ConfigParser()
config_file = 'config.ini'

# 如果配置文件不存在，自動創建一個默認的配置文件
if not os.path.exists(config_file):
    config['Window'] = {
        'width': 800,
        'height': 600,
        'x': 100,
        'y': 100
    }

    config['Textarea'] = {
        'original_text_area_content': '',
        'template_text_area_content': ''
    }

    with open(config_file, 'w') as configfile:
        config.write(configfile)

def encode_to_base64(text):
    text_bytes = text.encode('utf-8')
    encoded_bytes = base64.b64encode(text_bytes)
    encoded_text = encoded_bytes.decode('utf-8')
    return encoded_text

def decode_from_base64(encoded_text):
    encoded_bytes = encoded_text.encode('utf-8')
    decoded_bytes = base64.b64decode(encoded_bytes)
    decoded_text = decoded_bytes.decode('utf-8')
    return decoded_text

def save_settings():
    config['Window'] = {
        'width': window.winfo_width(),
        'height': window.winfo_height(),
        'x': window.winfo_x(),
        'y': window.winfo_y()
    }

    config['Textarea'] = {
        'original_text_area_content': encode_to_base64(original_text_area.get("1.0", "end-1c")),
        'template_text_area_content': encode_to_base64(template_text_area.get("1.0", "end-1c"))
    }

    with open(config_file, 'w') as configfile:
        config.write(configfile)

    window.destroy()  # 關閉視窗

def load_settings():
    if os.path.exists(config_file):
        config.read(config_file)
        window.geometry(f"{config['Window']['width']}x{config['Window']['height']}+{config['Window']['x']}+{config['Window']['y']}")
        original_text_area.insert("1.0", decode_from_base64(config['Textarea']['original_text_area_content']))
        template_text_area.insert("1.0", decode_from_base64(config['Textarea']['template_text_area_content']))

def recursive_replace(template, source_items, target_items):
    # 基本情況：如果source_items為空，則返回空字串
    if not source_items:
        return template

    # 從source_items和target_items中取出第一個元素
    source_item = source_items[0]
    target_item = target_items[0]

    # 從列表中移除第一個元素，以便於下次遞迴調用
    source_items = source_items[1:]
    target_items = target_items[1:]

    # 分割template，並對分割後的每個部分進行遞迴處理
    parts = template.split(source_item)
    result = recursive_replace(parts[0], source_items, target_items)
    for part in parts[1:]:
        result += target_item + recursive_replace(part, source_items, target_items)

    return result
    
def replace_and_copy_to_clipboard():
    # 從原始文本框中獲取要替換的源項目
    source_items = original_text_area.get("1.0", "1.end").split("\t")
    # 從原始文本框中獲取要替換的目標行
    target_lines_text = original_text_area.get("2.0", "end-1c").split("\n")
    # 從範本文本框中獲取要替換成的範本文字
    replacement_template = template_text_area.get("1.0", "end-1c")
    
    processed_text = ""
    for index, line in enumerate(target_lines_text):
        if line == "":
            continue
        
        target_items = line.split("\t")
        
        # 如果項目數量不同，跳過這一行
        if len(source_items) != len(target_items):
            messagebox.showinfo("", f"第{index + 2}行 \"{line}\"數量不符，跳過此行", parent=window)
            continue

        processed_text += recursive_replace(replacement_template, source_items, target_items)

    # 將處理後的文本複製到剪貼簿
    pyperclip.copy(processed_text)
    
    # 更新按鈕的狀態為"已複製"，紅色
    replace_button.config(text="已複製", bg="red")

    # 使用after方法在500毫秒後恢復按鈕的狀態
    window.after(500, reset_button_state)

def reset_button_state():
    # 恢復按鈕的文本和顏色
    replace_button.config(text="複製到剪貼簿", bg="#444654")

def paste_from_clipboard_and_process(event):
    # 清空template_text_area
    template_text_area.delete("1.0", tk.END)
    
    win32clipboard.OpenClipboard()
    clipboard_data = win32clipboard.GetClipboardData()
    win32clipboard.CloseClipboard()
    
    # 將剪貼簿的值貼上到template_text_area
    template_text_area.insert(tk.END, clipboard_data)
    
    # 執行複製到剪貼簿的功能
    replace_and_copy_to_clipboard()

# 創建標籤
label1 = tk.Label(window, text="第一行為要被替換掉的字，第二行以後放要替換成的字。可以用Tab隔開，一次替換多個文字。")
label1.pack()

# 創建第一個 textarea，用於輸入原始文本
original_text_area = tk.Text(window, height=5, width=50, undo=True)
original_text_area.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)  # 使用fill和expand使original_text_area跟隨視窗大小變化

# 創建標籤
label2 = tk.Label(window, text="要替換的內容，右鍵可直接貼上並執行複製到剪貼簿的功能。")
label2.pack()

# 創建第二個 textarea，用於輸入替換模板
template_text_area = tk.Text(window, height=10, width=50, undo=True)
template_text_area.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)  # 使用fill和expand使original_text_area跟隨視窗大小變化

# 創建一個Frame小部件
button_frame = tk.Frame(window)
button_frame.pack(padx=10, pady=10, fill=tk.BOTH)

# 創建替換按鈕，放置在button_frame中，點擊後將處理結果複製到剪貼簿
replace_button = tk.Button(button_frame, text="複製到剪貼簿", command=replace_and_copy_to_clipboard, height=3, bg="#444654", fg="white", cursor="hand2")
replace_button.pack(fill=tk.BOTH)

# 載入設置
load_settings()

# 綁定事件: 右鍵貼上並執行替換邏輯
window.bind("<Button-3>", paste_from_clipboard_and_process)

# 儲存設置的函數綁定到視窗的關閉事件
window.protocol("WM_DELETE_WINDOW", save_settings)

# 啟動 Tkinter 主迴圈
window.mainloop()
