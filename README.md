### [Back to Main-Projects tab](https://github.com/B-White-M/Projects/blob/fb845cb9a92211b6517c9284dd1290c2c84d6aca/README.md)

# VBA - Excel-Macros

This repository contains two Excel VBA macro projects that automate data gathering and processing; and email communication directly within Excel. These projects are designed to streamline tasks and enhance productivity by leveraging VBA automation.

## Project 1: Data Gathering Macro
This macro allows users to efficiently gather and process information with the click of a button. It simplifies data collection workflows, enabling easy treatment and handling of large datasets.

``` VBA Code

Sub AUC_Macro()
    ' Macro to copy data from a manual entry sheet and paste it into a main worksheet

    ' Variable declarations
    Dim mews As Worksheet
    Dim last_row As Long
    Dim last_column As Long
    Dim copy_range As Range
    Dim ws As Worksheet
    
    ' Define Manual Entry worksheet
    Set mews = ThisWorkbook.Sheets("Manual_Entry")
    
    ' Find last row with data in column D (manual entry data)
    last_row = mews.Cells(mews.Rows.Count, "D").End(xlUp).Row
    
    ' Exit if no data is found
    If last_row < 6 Then
        MsgBox "Please include AUC's details to proceed", vbExclamation, "No Data"
        Exit Sub
    End If
    
    ' Define the range of data to copy based on found last row
    Set copy_range = mews.Range("D6:O" & last_row)
    
    ' Define the AUC's Main worksheet
    Set ws = ThisWorkbook.Sheets("AUC's_Main")
    
    ' Find the last available row in the AUC's Main sheet to paste the data
    last_row = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1
    
    ' Paste copied data into the main sheet
    ws.Range("A" & last_row).Resize(copy_range.Rows.Count, copy_range.Columns.Count).Value = copy_range.Value

    ' Clear the data from Manual Entry after copying
    mews.Range("E6:O" & last_row).ClearContents
    
    ' Prevent copy mode and display a success message
    Application.CutCopyMode = False
    MsgBox "AUC's details saved!", vbInformation, "Process Complete"
End Sub

```
![Intvw 5](https://github.com/user-attachments/assets/911d4a82-fad6-4158-a3b1-7ef18fd857b9)

## Project 2: Automated Email Sender Macro
This macro automates the process of sending emails based on the data entered in the Excel sheet. With a single execution, it composes and sends emails to designated recipients, reducing manual efforts and ensuring timely communication.

``` VBA Code
Sub Sent_To_DEPARTMENT()
    ' Macro to send an email with data from the "Analysis Data" worksheet

    Dim wsAnalysisData As Worksheet
    Set wsAnalysisData = Worksheets("Analysis Data")

    ' Define the range to be included in the email
    Dim emailRange As Range
    Set emailRange = wsAnalysisData.Range("A9:E31")

    ' Show the envelope for the active workbook
    ActiveWorkbook.EnvelopeVisible = True

    ' Configure and send the email
    With wsAnalysisData.MailEnvelope
        ' Set recipients
        .Item.To = wsAnalysisData.Range("C11").Value & "; " & wsAnalysisData.Range("C12").Value
        .Item.CC = "email.pay@COMPANY.com"
        .Item.BCC = "SAMPLE.EMAIL@COMPANY.com"
        
        ' Set email subject
        .Item.Subject = "RE: " & wsAnalysisData.Range("C15").Value & " " & wsAnalysisData.Range("C37").Value & " " & wsAnalysisData.Range("C20").Value & "  " & wsAnalysisData.Range("D10").Value & ": " & wsAnalysisData.Range("D11").Value
        
        ' Send the email
        .Item.Send
    End With

    ' Notify the user
    MsgBox "Email Sent"
End Sub

```

![Intvw 6](https://github.com/user-attachments/assets/cc489248-5ec6-4456-9377-e47f1372e926)

Both projects demonstrate the power of Excel VBA in automating repetitive tasks and improving productivity in everyday workflows.

### [Back to Main-Projects tab](https://github.com/B-White-M/Projects/blob/fb845cb9a92211b6517c9284dd1290c2c84d6aca/README.md)
