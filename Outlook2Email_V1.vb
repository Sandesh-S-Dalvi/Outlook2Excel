' ' Sure! I can help you with that. Here's a VBA code that you can use in Outlook to automate getting data from selected emails and storing them in an Excel workbook:

Sub ExportSelectedEmailsToExcel()
    Dim olApp As Outlook.Application
    Dim olSelection As Outlook.Selection
    Dim olItem As Outlook.MailItem
    Dim xlApp As Excel.Application
    Dim xlWorkbook As Excel.Workbook
    Dim xlWorksheet As Excel.Worksheet
    Dim rowIndex As Long
    Dim attachmentIndex As Long
    Dim attachmentRow As Long
    Dim attachmentColumn As Long
    
    ' Create a new Excel workbook
    Set xlApp = New Excel.Application
    Set xlWorkbook = xlApp.Workbooks.Add
    Set xlWorksheet = xlWorkbook.Sheets(1)
    
    ' Get the selected emails from Outlook
    Set olApp = Outlook.Application
    Set olSelection = olApp.ActiveExplorer.Selection
    
    ' Add headers to Excel columns
    xlWorksheet.Range("A1").Value = "Sender"
    xlWorksheet.Range("B1").Value = "SenderMail"
    xlWorksheet.Range("C1").Value = "MailTime"
    xlWorksheet.Range("D1").Value = "MailAttachment"
    xlWorksheet.Range("E1").Value = "MailBody"
    
    ' Loop through each selected email
    rowIndex = 2 ' Start from row 2 (below headers)
    For Each olItem In olSelection
        ' Get sender's name
        xlWorksheet.Cells(rowIndex, 1).Value = olItem.SenderName
        
        ' Get sender's email address
        xlWorksheet.Cells(rowIndex, 2).Value = olItem.SenderEmailAddress
        
        ' Get mail date and time
        xlWorksheet.Cells(rowIndex, 3).Value = olItem.ReceivedTime
        
        ' Save mail attachments
        attachmentRow = rowIndex ' Start from the current row
        attachmentColumn = 4 ' Start from column D
        attachmentIndex = 1 ' Initialize attachment index
        
        For Each attachment In olItem.Attachments
            If attachment.Type = olByValue Then ' Only save attachments that are not linked
                If attachment.FileName Like "*.xls*" Or attachment.FileName Like "*.xlsx" Or attachment.FileName Like "*.xlsb" Or attachment.FileName Like "*.msg" Then
                    ' Save attachment as embedded object in Excel
                    attachment.SaveAsFile "C:\Attachments\" & attachment.FileName ' Change the path as per your requirement
                    xlWorksheet.Cells(attachmentRow, attachmentColumn).Value = attachment.FileName
                    attachmentRow = attachmentRow + 1
                Else
                    ' Save attachment as a hyperlink in the same cell
                    xlWorksheet.Hyperlinks.Add Anchor:=xlWorksheet.Cells(attachmentRow, attachmentColumn), _
                        Address:="", _
                        SubAddress:="C:\Attachments\" & attachment.FileName, _
                        TextToDisplay:=attachment.FileName
                    attachmentRow = attachmentRow + 1
                End If
            End If
        Next attachment
        
        ' Save mail body as HTML file
        olItem.SaveAs "C:\Attachments\" & olItem.SenderName & ".html", olHTML ' Change the path as per your requirement
        xlWorksheet.Cells(rowIndex, 5).Value = olItem.SenderName & ".html"
        
        rowIndex = attachmentRow ' Update the row index for the next email
        Set olItem = Nothing ' Release memory
    Next olItem
    
    ' Save and close the Excel workbook
    xlWorkbook.SaveAs "C:\Attachments\EmailData.xlsx" ' Change the path as per your requirement
    xlWorkbook.Close
    
    ' Release memory
    Set xlWorksheet = Nothing
    Set xlWorkbook = Nothing
    Set xlApp = Nothing
    Set olSelection = Nothing


    Set olApp = Nothing
End Sub

' To use this code:

' 1. Open Outlook and press "Alt+F11" to open the VBA editor.
' 2. In the VBA editor, go to "ThisOutlookSession" under "Project - VBAProject" on the left side.
' 3. Copy and paste the above code into the module.
' 4. Modify the file paths in the code to match your desired location for saving attachments and the Excel file.
' 5. Close the VBA editor.
' 6. Select the emails you want to export in Outlook.
' 7. Go to the "Developer" tab in the Outlook ribbon (if you don't see the "Developer" tab, you may need to enable it in Outlook options).
' 8. Click on the "Macros" button in the "Code" group.
' 9. Select the "ExportSelectedEmailsToExcel" macro and click "Run".

' The code will extract the required data from the selected emails, save attachments as embedded objects or hyperlinks in the Excel file, and save the mail body as an HTML file with the sender's name. The data will be saved in rows, with the appropriate headers in the columns.

' Please note that you need to have the necessary permissions to access and save attachments in Outlook and Excel, and you may need to adjust the file paths in the code to match your specific environment.
