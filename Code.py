import os
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from docxtpl import DocxTemplate
import win32com.client as win32
from docx import Document
import re

# ฟังก์ชันหลัก
def run_program():
    root = tk.Tk()
    root.withdraw()  # ซ่อนหน้าต่างหลัก
    
    try:
        # เลือกไฟล์ Excel (Database)
        messagebox.showinfo("Information", "Please select the Excel file (Database).")
        excel_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel Files", "*.xlsx;*.xls")]
        )
        if not excel_path:
            messagebox.showwarning("Warning", "Excel file not selected")
            return
        
        df = pd.read_excel(excel_path, header=None)
        df = df.dropna(axis=0, how='all').reset_index(drop=True)
        df.columns = df.iloc[0]  # ใช้แถวแรกเป็น Header
        df = df[1:].reset_index(drop=True)
        
        # เลือกไฟล์ Word Template
        messagebox.showinfo("Information", "Please select the Word template file.")
        word_template_path = filedialog.askopenfilename(
            title="Select Word Template File",
            filetypes=[("Word Files", "*.docx")]
        )
        if not word_template_path:
            messagebox.showwarning("Warning", "Word template file not selected")
            return

        # เลือกไฟล์ Word ที่ใช้เป็นเนื้อหาอีเมล
        messagebox.showinfo("Information", "Please select the Word file for email content.")
        email_content_path = filedialog.askopenfilename(
            title="Select Email Content File",
            filetypes=[("Word Files", "*.docx")]
        )
        if not email_content_path:
            messagebox.showwarning("Warning", "Email content file not selected")
            return

        email_subject = os.path.splitext(os.path.basename(email_content_path))[0]  # ใช้ชื่อไฟล์เป็นหัวข้ออีเมล

        # เลือกโฟลเดอร์สำหรับบันทึกไฟล์
        messagebox.showinfo("Information", "Select a folder to save generated files.")
        save_folder = filedialog.askdirectory(title="Select Save Folder")
        if not save_folder:
            messagebox.showwarning("Warning", "No folder selected. Files will not be saved.")
            return

        doc = DocxTemplate(word_template_path)
        template_name = os.path.splitext(os.path.basename(word_template_path))[0]

        for index, row in df.iterrows():
            context = {col: row[col] for col in df.columns}
            doc.render(context)

            # บันทึกไฟล์ Word
            word_filename = os.path.join(save_folder, f"{template_name}_{row.get('name', 'Unknown')}.docx")
            doc.save(word_filename)
            print(f"Saved Word file: {word_filename}")

            # แปลงเป็น PDF
            pdf_filename = convert_to_pdf(word_filename, save_folder)

            if 'email' in df.columns and pd.notna(row['email']):
                recipient_name = row.get('name', 'Valued Customer')  # ดึงชื่อจาก Excel
                body = extract_and_replace_placeholders(email_content_path, row)  # ดึงและแทนที่ข้อความอีเมล
                attach_pdf_to_outlook(pdf_filename, row['email'], email_subject, body, recipient_name, row)

        messagebox.showinfo("Complete", "All files have been saved successfully")
        root.quit()
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")
        root.quit()

# แปลงไฟล์ Word เป็น PDF และบันทึกในโฟลเดอร์ที่เลือก
def convert_to_pdf(doc_filename, save_folder):
    try:
        word = win32.Dispatch("Word.Application")
        word.Visible = False
        
        doc = word.Documents.Open(os.path.abspath(doc_filename))
        pdf_filename = os.path.join(save_folder, os.path.splitext(os.path.basename(doc_filename))[0] + ".pdf")
        
        doc.SaveAs(os.path.abspath(pdf_filename), FileFormat=17)
        doc.Close()
        word.Quit()
        
        print(f"Saved PDF file: {pdf_filename}")
        return pdf_filename
    except Exception as e:
        print(f"Failed to convert to PDF: {e}")
        return None

# ดึงและแทนที่ placeholders จากไฟล์ Word สำหรับเนื้อหาอีเมล
def extract_and_replace_placeholders(word_path, row):
    try:
        doc = Document(word_path)

        # สร้าง body โดยการแทนที่ placeholders จากไฟล์ Word
        body = ""
        for para in doc.paragraphs:
            para_text = para.text

            # หา placeholders ที่อยู่ในรูปแบบ {{ placeholder }}
            placeholders = re.findall(r'{{\s*(\w+)\s*}}', para_text)
            for placeholder in placeholders:
                if placeholder in row:
                    para_text = para_text.replace(f"{{{{ {placeholder} }}}}", str(row[placeholder]))
                else:
                    para_text = para_text.replace(f"{{{{ {placeholder} }}}}", "")  # ถ้าไม่มีค่าจะลบ placeholder ออก

            body += para_text + "<br>"

        return body  # คืนค่าเนื้อหาที่แทนที่ placeholder แล้ว

    except Exception as e:
        print(f"Failed to extract and replace placeholders from Word: {e}")
        return ""

# แนบไฟล์ PDF ไปยังอีเมล Outlook และใช้ HTML รักษารูปแบบ
def attach_pdf_to_outlook(pdf_filename, recipient_email, subject, body, recipient_name, row):
    try:
        outlook = win32.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)

        # แทนที่ placeholder ในเนื้อหาอีเมลที่อ่านจาก Word
        mail.Subject = subject
        mail.HTMLBody = body  # ใช้ HTML เพื่อรักษารูปแบบ
        mail.To = recipient_email
        mail.Attachments.Add(os.path.abspath(pdf_filename))
        mail.Display()
        print(f"Sent email to {recipient_name} ({recipient_email}) with {pdf_filename}.")
    except Exception as e:
        print(f"Failed to attach PDF to Outlook: {e}")

if __name__ == "__main__":
    run_program()
