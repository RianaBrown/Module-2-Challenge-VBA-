Attribute VB_Name = "Module1"
Sub MethodStockData()

'Declare variables
Dim ws As Worksheet
Dim openingprice As Double
Dim closingprice As Double
Dim tickerRange As Range
Dim volRange As Range
Dim cell As Range
Dim lastRow As Long
Dim rowNum As Integer
Dim returnRange As Range
Dim highestVolume As Range
Dim tickers As String

'count tount of worksheets in the workbook
Dim wsCount As Integer
wsCount = ThisWorkbook.Worksheets.Count

'begin row counter
rowNum = 2


'add ticker,yearly change, percent change,total stock volume to columns
For i = 1 To 3

    'add the columns to the worksheet
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
Next

'------------------------------------------------------------
'loop through the sheets and add data to the columns above

For i = 1 To wsCount

    Set ws = ActiveWorkbook.Sheets(i)
    ws.Activate
    ticker = ws.Range("A2")
    lastRow = ws.Range("A2").End(xlDown).Row
    Set tickerRange = ws.Range("A2:A" & lastRow)
    Set volRange = ws.Range("G2:G" & lastRow)

     For Each cell In tickerRange
     
        If Cells(cell.Row, 1) = ticker And Cells(cell.Row, 2).Value = ws.Name & "0102" Then
        
            'add opening price to the variable
             openingprice = Cells(cell.Row, 3).Value
     
             
        ElseIf Cells(cell.Row, 1) = ticker And Cells(cell.Row, 2).Value = ws.Name & "1231" Then
        
            'add closing price to the variable
            closingprice = Cells(cell.Row, 6).Value
        
            'add the data to the columns created on the right
            Cells(rowNum, 9).Value = ticker
            Cells(rowNum, 10).Value = closingprice - openingprice
            Cells(rowNum, 11).Value = (closingprice / openingprice) - 1

            'format the percentage change column to a percentage
            Cells(rowNum, 11).NumberFormat = "0.00%"

            'format the total volume as a general number instead of scientific notation
            'https://learn.microsoft.com/en-us/office/vba/api/excel.range.numberformat
            Cells(rowNum, 12).Value = WorksheetFunction.SumIfs(volRange, tickerRange, ticker)

            Cells(rowNum, 12).NumberFormat = "General"
        
            'add 1 to the row counter
            rowNum = rowNum + 1
        
                If cell.Row = lastRow Then

                    'reset the row counter for the next sheet
                    rowNum = 2
                
                Else

                    'change the ticker to the next row's ticker
                    ticker = Cells(cell.Row + 1, 1).Value
        
                End If
        End If
      Next
Next

'-----------------------------------------------------------
'calculate the greatest % increase/decrease & greatest total volume

For i = 1 To wsCount

    Set ws = ActiveWorkbook.Sheets(i)
    ws.Activate
    lastRow = ws.Range("I2").End(xlDown).Row
    Set returnRange = ws.Range("I2:L" & lastRow)

    'sort percent change column in descending order
    returnRange.Sort key1:=Range("K2"), order1:=xlDescending

    'add column names
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"

    'add names to future calculated rows
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"

    'add sorted data from the top row to the columns for the calculated data
    Cells(2, 16).Value = Range("I2")
    Cells(2, 17).Value = Range("K2")

    'return back to the original sort by ticker
    returnRange.Sort key1:=Range("I2"), order1:=xlAscending

    'sort by greatest loss
    returnRange.Sort key1:=Range("K2"), order1:=xlAscending

    'get the greatest loss
    Cells(3, 16).Value = Range("I2")
    Cells(3, 17).Value = Range("K2")

    'return back to the original sort by ticker
    returnRange.Sort key1:=Range("I2"), order1:=xlAscending

    'get the greatest amount of volume
    returnRange.Sort key1:=Range("L2"), order1:=xlDescending

    Cells(4, 16).Value = Range("I2")
    Cells(4, 17).Value = Range("L2")

    'return back to the original sort by ticker
    returnRange.Sort key1:=Range("I2"), order1:=xlAscending
Next

'-------------------------------------------------------------------
'apply conditional formatting to the worksheets

For i = 1 To wsCount

    Set ws = ActiveWorkbook.Sheets(i)
    ws.Activate

    lastRow = ws.Range("J2").End(xlDown).Row

    Set Rng = ws.Range("J2:J" & lastRow)

        For Each cell In Rng
    
            If cell.Value >= 0 Then
        
                'apply green to positive percents
                cell.Interior.ColorIndex = 4
                
                'apply green to the percent change column too
                Cells(cell.Row, 11).Interior.ColorIndex = 4
        
            Else

                'apply red to negative percents
                cell.Interior.ColorIndex = 3
                
                'apply red to the percent change column too
                Cells(cell.Row, 11).Interior.ColorIndex = 3
        
            End If
        Next
Next

'the end of the subroutine
MsgBox ("Done")





End Sub
