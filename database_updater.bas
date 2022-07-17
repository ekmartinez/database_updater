Attribute VB_Name = "DataBaseUpdater"

Sub DataBaseUpdater()
    Attribute DataBaseUpdater.VB_ProcData.VB_Invoke_Func = "U\n14"

    Dim Destiny As Workbook
    Dim sht As Worksheet
    Dim LastRow As Long
    Dim LastColumn As Long
    Dim StartCell As Range


    ScreenUpdating = True

    'Copy data from summary report
    Set sht = Worksheets("Summary")
    Set StartCell = Range("A5")


    LastRow = sht.Cells(sht.Rows.Count, StartCell.Column).End(xlUp).Row  'Find Last Row
    sht.Range("A5:K" & LastRow).Copy     'Select Range
    
    'Paste data from summary report to Database
        Set Destiny = Workbooks.Open("N:\Professional Services\Database.xlsx")
            Destiny.Sheets("Database").Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).PasteSpecial xlPasteValues
        
    CutCopyMode = False


    'Automate the calculation of total hours in database
    'LRow = Range("A" & Rows.Count).End(xlUp).Row
    'For i = 2 To LRow
    'Next i

    Unload UserForm1

End Sub


'LRow = Range("B" & Rows.Count).End(xlUp).Row







