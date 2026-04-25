import pandas as pd
import win32com.client as win32
import os
import time
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import traceback
import re

def select_file(entry, title="選擇檔案", filetypes=[("Excel files", "*.xlsx *.xls")]):
    file_path = filedialog.askopenfilename(title=title, filetypes=filetypes)
    if file_path:
        entry.delete(0, tk.END)
        entry.insert(0, file_path)

def mask_email(email):
    """保護隱私：將郵件地址部分屏蔽 (e.g., ed***@example.com)"""
    if not isinstance(email, str) or "@" not in email:
        return str(email)
    prefix, domain = email.split("@")
    if len(prefix) <= 2:
        return f"{prefix}***@{domain}"
    return f"{prefix[:2]}***@{domain}"

def send_emails_from_excel(excel_path, log_box):
    try:
        # 1. 讀取 Excel 並檢查必要欄位
        df = pd.read_excel(excel_path)
        required_cols = ['To', 'Subject', 'Body']
        missing_cols = [col for col in required_cols if col not in df.columns]
        
        if missing_cols:
            messagebox.showerror("格式錯誤", f"Excel 缺少必要欄位：{', '.join(missing_cols)}")
            return

    except Exception as e:
        messagebox.showerror("讀取錯誤", f"無法讀取檔案，請檢查 Excel 是否已關閉。\n{e}")
        return

    outlook = win32.Dispatch('outlook.application')
    sent_count = 0
    skipped_rows = []

    for idx, row in df.iterrows():
        excel_row_num = idx + 2 
        try:
            # 2. 附件檢查邏輯
            attachments_list = []
            if 'Attachment' in df.columns and pd.notna(row['Attachment']):
                attachments = str(row['Attachment']).split(";")
                for path in attachments:
                    path = path.strip()
                    if path and os.path.exists(path):
                        attachments_list.append(path)
                    elif path:
                        log_box.insert(tk.END, f"⚠️ 第 {excel_row_num} 行附件不存在：{os.path.basename(path)}\n")
            
            # 3. 建立郵件物件
            mail = outlook.CreateItem(0)
            target_to = str(row['To']).strip()
            mail.To = target_to
            mail.CC = str(row['CC']).strip() if 'CC' in df.columns and pd.notna(row['CC']) else ""
            mail.Subject = str(row['Subject']).strip()

            # 4. 處理簽名檔 (預先顯示以獲取 HTMLBody)
            mail.Display()
            time.sleep(0.5) 
            signature = mail.HTMLBody

            # 5. 組合內文 (支援換行轉 HTML)
            raw_body = str(row['Body']).strip()
            formatted_body = raw_body.replace("\r\n", "<br>").replace("\n", "<br>")
            
            # 使用通用字體，避免特定公司格式限制
            mail.HTMLBody = f"""
            <html>
                <body style="font-family: 'Segoe UI', Calibri, sans-serif; font-size: 11pt;">
                    {formatted_body}
                    <br><br>
                </body>
            </html>
            """ + signature

            # 6. 加入附件
            for file_path in attachments_list:
                mail.Attachments.Add(file_path)

            # 7. 寄送並記錄 (隱私保護版 Log)
            mail.Send()
            sent_count += 1
            log_box.insert(tk.END, f"✅ 已寄出第 {excel_row_num} 行：{mask_email(target_to)}\n")
            log_box.see(tk.END)
            
            # 避免觸發郵件伺服器頻率限制
            time.sleep(1)

        except Exception as e:
            log_box.insert(tk.END, f"❌ 第 {excel_row_num} 行發生錯誤：{str(e)}\n")
            skipped_rows.append(excel_row_num)

    messagebox.showinfo("執行完畢", f"總計寄出 {sent_count} 封信件。\n失敗行數：{skipped_rows if skipped_rows else '無'}")

def main():
    root = tk.Tk()
    root.title("Python 辦公自動化：批次寄信工具")
    root.geometry("620x450")
    root.configure(padx=20, pady=20)

    # UI 佈局
    tk.Label(root, text="Step 1: 選擇包含名單的 Excel 檔案", font=("Microsoft JhengHei", 10, "bold")).pack(anchor="w")
    
    frame_excel = tk.Frame(root)
    frame_excel.pack(fill="x", pady=5)
    
    excel_entry = tk.Entry(frame_excel)
    excel_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
    
    tk.Button(frame_excel, text="瀏覽檔案", command=lambda: select_file(excel_entry)).pack(side="right")

    tk.Label(root, text="執行進度 Log:", font=("Microsoft JhengHei", 9)).pack(anchor="w", pady=(10, 0))
    log_box = scrolledtext.ScrolledText(root, width=70, height=12, font=("Consolas", 9))
    log_box.pack(pady=5)

    def run_process():
        path = excel_entry.get()
        if not path:
            messagebox.showwarning("提示", "請先選擇 Excel 檔案")
            return
        send_emails_from_excel(path, log_box)

    tk.Button(root, text="🚀 開始執行批次寄送", command=run_process, 
              bg="#28a745", fg="white", font=("Microsoft JhengHei", 10, "bold"), height=2).pack(fill="x", pady=10)

    root.mainloop()

if __name__ == "__main__":
    main()
