Attribute VB_Name = "CommonFun"

Public Function IsInArray(value As Variant, arr As Variant) As Boolean
    Dim element As Variant

    For Each element In arr
        If element = value Then
            IsInArray = True
            Exit Function
        End If
    Next element

    IsInArray = False
End Function

Public Function LoadReceiptData()
    ' MsgBox "Ready",,"Kudadonbe"
    Dim receiptData As Variant
    ' Dim newReleaseNoteData(1 To 1000, 1 To 26) As String
    Dim filePath As Variant
    ' Dim templateFilePath As String
    ' Dim newReleaseNoteFilePath As String
    Dim lastRow As Long, lastColumn As Long
    Dim wb As Workbook, ws As Worksheet

    ' open file dialog to select file
    filePath = Application.GetOpenFilename(FileFilter:="Excel Files (*.xls*),*.xls*", Title:="Select Excel File")

    ' exit if no file selected
    If filePath = False Then Exit Function

        ' open workbook and worksheet
        Set wb = Workbooks.Open(filePath)
        Set ws = wb.Sheets(1)

        ' determine last row and column with data
        lastRow = ws.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
        lastColumn = ws.Cells.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column

        ' store data into array
        receiptData = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastColumn)).Value



        ' close workbook without saving
        wb.Close SaveChanges:=False

        LoadReceiptData = receiptData
End Function