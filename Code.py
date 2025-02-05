import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from docxtpl import DocxTemplate
import win32com.client as win32
from docx import Document

# ฟังก์ชันหลักที่ใช้เลือกไฟล์ Excel และ Word Template
def run_program():
    root = tk.Tk()
    root.withdraw()  # ซ่อนหน้าต่างหลัก
    
    try:
        # แจ้งให้ผู้ใช้เลือกไฟล์ Excel ก่อน
        messagebox.showinfo("Information", "Please select the Excel file (Database).")
        
        # เปิดหน้าต่างเลือกไฟล์ Excel
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
        
        # แจ้งให้ผู้ใช้เลือกไฟล์ Word Template
        messagebox.showinfo("Information", "Please select the Word template file.")
        
        word_template_path = filedialog.askopenfilename(
            title="Select Word Template File",
            filetypes=[("Word Files", "*.docx")]
        )
        
        if not word_template_path:
            messagebox.showwarning("Warning", "Word template file not selected")
            return
        
        # แจ้งให้ผู้ใช้เลือกไฟล์ Word ที่ใช้เป็นเนื้อหาอีเมล
        messagebox.showinfo("Information", "Please select the Word file for email content.")
        email_content_path = filedialog.askopenfilename(
            title="Select Email Content File",
            filetypes=[("Word Files", "*.docx")]
        )
        
        if not email_content_path:
            messagebox.showwarning("Warning", "Email content file not selected")
            return
        
        email_subject = os.path.splitext(os.path.basename(email_content_path))[0]  # ใช้ชื่อไฟล์เป็นหัวข้ออีเมล
        email_body = extract_text_from_word(email_content_path)  # ดึงเนื้อหาจากไฟล์ Word
        
        doc = DocxTemplate(word_template_path)
        template_name = os.path.splitext(os.path.basename(word_template_path))[0]
        
        for index, row in df.iterrows():
            context = {col: row[col] for col in df.columns}
            doc.render(context)
            
            word_filename = f"{template_name}_{row.get('name', 'Unknown')}.docx"
            doc.save(word_filename)
            print(f"Saved Word file: {word_filename}")
            
            pdf_filename = convert_to_pdf(word_filename)
            
            if 'email' in df.columns and pd.notna(row['email']):
                attach_pdf_to_outlook(pdf_filename, row['email'], email_subject, email_body)
        
        messagebox.showinfo("Complete", "All files have been saved successfully")
        root.quit()  # ปิดหน้าต่าง Tkinter เพื่อหลีกเลี่ยงการค้าง
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")
        root.quit()

# ฟังก์ชันแปลงไฟล์ Word เป็น PDF
def convert_to_pdf(doc_filename):
    try:
        word = win32.Dispatch("Word.Application")
        word.Visible = False
        
        doc = word.Documents.Open(os.path.abspath(doc_filename))
        pdf_filename = os.path.splitext(doc_filename)[0] + ".pdf"
        
        doc.SaveAs(os.path.abspath(pdf_filename), FileFormat=17)
        doc.Close()
        word.Quit()
        
        print(f"Saved PDF file: {pdf_filename}")
        return pdf_filename
    except Exception as e:
        print(f"Failed to convert to PDF: {e}")
        return None

# ฟังก์ชันแนบไฟล์ PDF ไปยังอีเมล Outlook
def attach_pdf_to_outlook(pdf_filename, recipient_email, subject, body):
    try:
        outlook = win32.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)
        
        mail.Subject = subject
        mail.Body = body
        mail.To = recipient_email
        mail.Attachments.Add(os.path.abspath(pdf_filename))
        mail.Display()
        print(f"Attached {pdf_filename} to Outlook email for {recipient_email}.")
    except Exception as e:
        print(f"Failed to attach PDF to Outlook: {e}")

# ฟังก์ชันดึงข้อความจากไฟล์ Word
def extract_text_from_word(word_path):
    try:
        doc = Document(word_path)
        return "\n".join([para.text for para in doc.paragraphs])
    except Exception as e:
        print(f"Failed to extract text from Word: {e}")
        return ""

if __name__ == "__main__":
    run_program()
