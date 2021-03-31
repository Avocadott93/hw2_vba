Attribute VB_Name = "Module1"
Sub StockCalculation1()


    Dim readName As String
    Dim nextName As String
    Dim groupNo As Long
    Dim totalSV As Double
     Dim i As Long
    Dim lastRow As Long
    Dim curSheet As Worksheet
    Dim Change As Double
    Dim percentChange As Double

 Set curSheet = ThisWorkbook.Worksheets("A")

    curSheet.Cells(1, 9).Value = "Ticker"
    curSheet.Cells(1, 10).Value = "Yearly Change"
    curSheet.Cells(1, 11).Value = "Percent Change"
    curSheet.Cells(1, 12).Value = "Total Stock Volumn"
    groupNo = 1
    totalSV = 0

lastRow = curSheet.Cells(curSheet.Rows.Count, 1).End(xlUp).Row
 
firstopen = curSheet.Cells(2, 3).Value

 For i = 2 To lastRow

        readName = curSheet.Cells(i, 1).Value
        nextName = curSheet.Cells(i + 1, 1).Value

         If nextName = readName Then
                   totalSV = totalSV + curSheet.Cells(i, 7).Value
        Else  'current group last row
            totalSV = totalSV + curSheet.Cells(i, 7).Value
            curSheet.Cells(groupNo + 1, 9).Value = readName
            curSheet.Cells(groupNo + 1, 10).Value = Change
            curSheet.Cells(groupNo + 1, 11).Value = percentChange
            curSheet.Cells(groupNo + 1, 12).Value = totalSV
            lastclose = curSheet.Cells(i, 6).Value
            Change = lastclose - firstopen
            percentChange = (Change / firstopen)


            groupNo = groupNo + 1  '!!
            totalSV = 0
            firstopen = curSheet.Cells(i + 1, 3).Value
        End If

 Next i    'increase 1 and continue the loop

    curSheet.Cells(1, 9).Interior.Color = RGB(0, 255, 0)
    
    curSheet.Cells(1, 12).Interior.Color = RGB(0, 255, 0)

    curSheet.Range("A1:C30").NumberFormat = "0.00%"

End Sub
