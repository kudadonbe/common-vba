Public  Sub makeReleaseNote()
    Dim newReleaseNoteData(1 To 1000, 1 To 26) As String
    Dim newReleaseNoteFilePath As String
    Dim templateFilePath As String
    Dim receiptData As Variant
    receiptData = CommonFun.LoadReceiptData()
    ' Filter the data based on fundCode, accountNo, and incomeCode
    Dim filteredRows As Variant ' 1D array to store the filtered rows
    Dim fundCode As Variant
    Dim accountNo As String
    Dim payType As String

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
    Dim dateFrom As Date
    Dim dateTo As Date
    

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
    Dim dateCol As Integer
    Dim payTypeCol As Integer


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
    dateFrom = Range("O" & ActiveCell.Row).Value
    dateTo = Range("P" & ActiveCell.Row).Value
    payType = "Cash"
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
    dateCol = 4   
    payTypeCol = 25   
    
    filteredRows = 0

    For i = 1 To lastRow
        ' Debug.Print receiptData(i, 10) & vbTab & receiptData(i, fundCol) & vbTab & receiptData(i, receiptNoCol) & vbTab & receiptData(i, entryTotalCol) 
        If receiptData(i, dateCol) >= dateFrom And receiptData(i, dateCol) <= dateTo And receiptData(i, fundCol) = fundCode And receiptData(i, acctNoCol) = accountNo And receiptData(i, isCancelCol) = "No" And receiptData(i, payTypeCol) = payType Then
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
    numRows = (filteredRows - 2) ' Change to the number of rows you want to insert
    
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
