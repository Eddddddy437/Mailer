import pandas as pd
import win32com.client as win32
import os
import time
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import traceback

def select_file(entry, title="選擇檔案", filetypes=[("All files", "*.*")]):
    file_path = filedialog.askopenfilename(title=title, filetypes=filetypes)
    if file_path:
        entry.delete(0, tk.END)
        entry.insert(0, file_path)

def send_emails_from_excel(excel_path, log_box):
    # 讀取 Excel
    try:
        df = pd.read_excel(excel_path)
    except Exception as e:
        messagebox.showerror("讀取錯誤", f"無法讀取 Excel 檔案\n{e}")
        return

    outlook = win32.Dispatch('outlook.application')
    sent_count = 0
    skipped_rows = []

    for idx, row in df.iterrows():
        excel_row_num = idx + 2  # Excel 的行號（加上標題列）
        try:
            # 檢查附件是否存在
            attachments_missing = False
            attachments_list = []
            if pd.notna(row['Attachment']):
                attachments = str(row['Attachment']).split(";")
                for attachment_path in attachments:
                    attachment_path = attachment_path.strip()
                    if attachment_path and os.path.exists(attachment_path):
                        attachments_list.append(attachment_path)
                    elif attachment_path:
                        log_box.insert(tk.END, f"❌ Excel 第 {excel_row_num} 行找不到附件，該封信不寄出：{attachment_path}\n")
                        log_box.see(tk.END)
                        attachments_missing = True

            if attachments_missing:
                skipped_rows.append(excel_row_num)
                continue  # 跳過寄送這封信

            # 創建郵件
            mail = outlook.CreateItem(0)
            mail.To = str(row['To']).strip() if pd.notna(row['To']) else ""
            mail.CC = str(row['CC']).strip() if pd.notna(row['CC']) else ""
            mail.Subject = str(row['Subject']).strip() if pd.notna(row['Subject']) else "（無主旨）"

            # 先顯示，讓 Outlook 套用簽名
            mail.Display()
            time.sleep(0.5)  # 等 Outlook 寫入簽名
            signature = mail.HTMLBody

            # Excel 內容轉 HTML
            body_text = str(row['Body']).strip() if pd.notna(row['Body']) else ""
            body_text = body_text.replace("\r\n", "<br>").replace("\n", "<br>")
            body_html = f'<p style="font-family:Calibri; font-size:11pt; color:black;">{body_text}</p>'
            mail.HTMLBody = body_html + "<br><br>" + signature

            # 加入附件
            for file_path in attachments_list:
                mail.Attachments.Add(file_path)

            # 寄送郵件
            mail.Save()
            mail.Send()
            del mail

            sent_count += 1
            log_box.insert(tk.END, f"✅ 已寄出 Excel 第 {excel_row_num} 行 {row['Subject']} → {row['To']}\n")
            log_box.see(tk.END)
            time.sleep(1)

        except Exception as e:
            log_box.insert(tk.END, f"❌ Excel 第 {excel_row_num} 行錯誤：{e}\n{traceback.format_exc()}\n")
            log_box.see(tk.END)

    summary_msg = f"所有郵件處理完畢！\n已寄出 {sent_count} 封信件"
    if skipped_rows:
        summary_msg += f"\n以下 Excel 行數因附件缺失未寄出：{', '.join(map(str, skipped_rows))}"
    messagebox.showinfo("完成", summary_msg)

def main():
    root = tk.Tk()
    root.title("批次寄信工具 (Outlook)")
    root.geometry("600x400")

    frame_excel = tk.Frame(root)
    frame_excel.pack(anchor="w", padx=10, pady=5)
    tk.Label(frame_excel, text="Excel 檔案").pack(side="left")
    excel_entry = tk.Entry(frame_excel, width=50)
    excel_entry.pack(side="left", padx=10)
    tk.Button(frame_excel, text="瀏覽",
              command=lambda: select_file(excel_entry, "選擇 Excel 檔案", [("Excel files", "*.xlsx *.xls")])
              ).pack(side="left")

    log_box = scrolledtext.ScrolledText(root, width=70, height=15)
    log_box.pack(padx=10, pady=10)

    def run_send():
        excel_file = excel_entry.get()
        if not excel_file:
            messagebox.showwarning("缺少檔案", "請先選擇 Excel 檔案")
            return
        send_emails_from_excel(excel_file, log_box)

    tk.Button(root, text="開始寄送", command=run_send, bg="lightblue").pack(pady=5)

    root.mainloop()

if __name__ == "__main__":
    main()
