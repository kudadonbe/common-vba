Attribute VB_Name = "Common"

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

Public  Sub LoadReceiptData()

    ' MsgBox "Ready",,"Kudadonbe" 
    Dim receiptData As Variant
    Dim newReleaseNoteData(1 To 1000, 1 To 26) As String
    Dim filePath As Variant
    Dim templateFilePath As String
    Dim newReleaseNoteFilePath As String
    Dim lastRow As Long, lastColumn As Long
    Dim wb As Workbook, ws As Worksheet
    
    ' open file dialog to select file
    filePath = Application.GetOpenFilename(FileFilter:="Excel Files (*.xls*),*.xls*", Title:="Select Excel File")
    
    ' exit if no file selected
    If filePath = False Then Exit Sub
    
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

    ' Filter the data based on fundCode, accountNo, and incomeCode
    Dim filteredRows As Variant ' 1D array to store the filtered rows
    Dim fundCode As Variant
    Dim accountNo As String

    Dim releaseNoteNo As String
    Dim releasedName As String
    Dim releasedDate As Date

    Dim depositerName As String
    Dim depositerDesig As String
    Dim depositerDate As String

    Dim accountantName As String
    Dim accountantDesig As String
    Dim accountantDate As String
    Dim note As String

    Dim incomeCodeformulaResult As Variant
    Dim incomeCodeformulaResultAsString As Variant
    Dim i As Long, j As Long, k As Long
    
    Dim fundCol As Integer
    Dim acctNoCol As Integer
    Dim incomeCodeCol As Integer
    Dim glCodeCol As Integer
    Dim receiptNoCol As Integer
    Dim entryTotalCol As Integer
    Dim activityCol As Integer
    Dim isCancelCol As Integer


    releaseNoteNo = Range("A" & ActiveCell.Row).Text
    fundCode = Range("B" & ActiveCell.Row).Value
    accountNo = CStr(Range("C" & ActiveCell.Row).Value)

    releasedName = Range("F" & ActiveCell.Row).Text
    releasedDate = Range("G" & ActiveCell.Row).Value

    depositerName = Range("H" & ActiveCell.Row).Text
    depositerDesig = Range("I" & ActiveCell.Row).Text
    depositedDate = Range("J" & ActiveCell.Row).Text

    accountantName = Range("K" & ActiveCell.Row).Text
    accountantDesig = Range("L" & ActiveCell.Row).Text
    accountantDate = Range("M" & ActiveCell.Row).Text
    note = Range("N" & ActiveCell.Row).Text
    ' incomeCodeformulaResult = Range("D" & ActiveCell.Row).Value
    ' incomeCodeformulaResultAsString = Replace(incomeCodeformulaResult, " ", "")
    ' incomeCode = Split(incomeCodeformulaResultAsString, ",")

    glCodeCol = 10
    fundCol = 9 'Assuming fund code is in column 9
    receiptNoCol = 5
    entryTotalCol = 21
    activityCol = 8
    isCancelCol = 23

    acctNoCol = 11 'Assuming accountNo is in column 11
    incomeCodeCol = 7 'Assuming incomeCode is in column 7   
    
    filteredRows = 0

    For i = 1 To lastRow
        ' Debug.Print receiptData(i, 10) & vbTab & receiptData(i, fundCol) & vbTab & receiptData(i, receiptNoCol) & vbTab & receiptData(i, entryTotalCol) 
        If receiptData(i, fundCol) = fundCode And receiptData(i, acctNoCol) = accountNo And receiptData(i, isCancelCol) = "No" Then
            filteredRows = filteredRows + 1
            ' Debug.Print receiptData(i, 10) & vbTab & receiptData(i, 9) & vbTab & receiptData(i, 5) & vbTab & receiptData(i, 21) 
            For j = 1 To lastColumn
                newReleaseNoteData(filteredRows, j) = receiptData(i, j)
            Next j       
        End If
    Next i

   
    templateFilePath = "C:\Users\HussainShareef\OneDrive\Documents\Custom Office Templates\Release Note.xltm"
    Set wb = Workbooks.Open(templateFilePath)
    Set ws = wb.Sheets(1)

    Range("G5").Value = releaseNoteNo
    Range("G6").Value = releasedDate

    Range("B13").Value = releasedName
    Range("E13").Value = releasedDate
    
    Range("A14").Value = note
    
    Range("D17").Value = depositerName
    Range("D18").Value = depositerDesig
    Range("D19").Value = depositerDate

    Range("G17").Value = accountantName
    Range("G18").Value = accountantDesig
    Range("G19").Value = accountantDate
    


    Dim startRow As Long
    startRow = 10 'Change to the row number where you want to start inserting
    
    Dim numRows As Long
    numRows = filteredRows ' Change to the number of rows you want to insert
    
    For i = 1 To numRows
        ws.Rows(startRow).Insert shift:=xlDown
        ws.Rows(startRow - 1).Copy Destination:=ws.Rows(startRow)
    Next i

    For i = 1 To filteredRows
        Range("A" & (i + 8)).Value = i
        Range("B" & (i + 8)).Value = newReleaseNoteData(i, glCodeCol)
        Range("C" & (i + 8)).Value = newReleaseNoteData(i, fundCol)
        Range("D" & (i + 8)).Value = newReleaseNoteData(i, receiptNoCol)
        Range("E" & (i + 8)).Value = newReleaseNoteData(i, activityCol)
        Range("G" & (i + 8)).Value = newReleaseNoteData(i, entryTotalCol)
        ' Debug.Print newReleaseNoteData(i, 10) & vbTab & newReleaseNoteData(i, 9) & vbTab & newReleaseNoteData(i, 5) & vbTab & newReleaseNoteData(i, 21) 
    Next i 
    newReleaseNoteFilePath = "S:\Co-operate Affairs\Safe\2023\safe_release\" & releaseNoteNo &".xlsx"
    wb.SaveAs newReleaseNoteFilePath
    ' wb.Close SaveChanges:=False


End Sub