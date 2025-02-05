# Word-Template-Automation
 Uses Excel data to automatically populate a Word template and send emails.


stateDiagram-v2
    Start_Program --> Select_Excel_DATABASE(.xlsx)
    Select_Excel_DATABASE(.xlsx) --> ProcessData : Success
    Select_Excel_DATABASE(.xlsx) --> End_program : Error Occurred
    
    ProcessData --> Select_Word_Template(.docx)
    Select_Word_Template(.docx) --> Generate_Word : Success
    Select_Word_Template(.docx) --> End_program : Error Occurred
    
    Generate_Word --> Convert_To_PDF
    Convert_To_PDF --> Attach_PDF_file_to_Outlook
    Attach_PDF_file_to_Outlook --> Select_Email_Content(.docx)
    
    Select_Email_Content(.docx) --> Copy_Email_Content_To_Outlook : Success
    Select_Email_Content(.docx) --> End_program : Error Occurred

    Copy_Email_Content_To_Outlook --> Send_Email
    Send_Email --> End_program

    Convert_To_PDF --> End_program : Error Occurred
    Attach_PDF_file_to_Outlook --> End_program : Error Occurred




