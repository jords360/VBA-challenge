Attribute VB_Name = "Module1"
Sub LoopThroughStocks()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim tickerColumn As Long
    Dim yearSheet As Variant
    Dim openColumn As Long
    Dim closeColumn As Long
    Dim volumeColumn As Long
    Dim outputColumn As Long ' Column to output the data
    
    ' Specify the column indices
    tickerColumn = 1 ' Column A
    openColumn = 3 ' Column C for "open" prices
    closeColumn = 6 ' Column F for "close" prices
    volumeColumn = 7 ' Column G for "volume"
    outputColumn = 9 ' Column I for output data
    
' Array of sheet names to process
    Dim sheetNames() As Variant
    sheetNames = Array("2018", "2019", "2020")
    
    ' Loop through each sheet
    For Each yearSheet In sheetNames
        ' Set the worksheet for the current year
        Set ws = ThisWorkbook.Worksheets(yearSheet)
    
    ' Find the last used row in the ticker symbol column
    lastRow = ws.Cells(ws.Rows.Count, tickerColumn).End(xlUp).Row
    
    ' Set the column headers
    ws.Cells(1, outputColumn).Value = "Ticker"
    ws.Cells(1, outputColumn + 1).Value = "Yearly Change"
    ws.Cells(1, outputColumn + 2).Value = "Percentage Change"
    ws.Cells(1, outputColumn + 3).Value = "Total Volume"
    
    ' Initialize variables for aggregated data
    Dim ticker As String
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim volume As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    
    ' Initialize the variables with the first ticker symbol in the data
    ticker = ws.Cells(2, tickerColumn).Value
    openingPrice = ws.Cells(2, openColumn).Value
    totalVolume = 0
    
    ' Initialize output row
    Dim outputRow As Long
    outputRow = 2
    
    ' Loop through each row in the ticker symbol column
    Dim i As Long
    For i = 2 To lastRow
        ' Check if the current ticker is different from the previous ticker
        If ws.Cells(i, tickerColumn).Value <> ticker Then
            ' Output the aggregated data for the previous ticker
            ws.Cells(outputRow, outputColumn).Value = ticker
            ws.Cells(outputRow, outputColumn + 1).Value = yearlyChange
            ws.Cells(outputRow, outputColumn + 2).Value = percentChange
            ws.Cells(outputRow, outputColumn + 3).Value = totalVolume
            
            ' Apply conditional formatting to the "Yearly Change" column
            Dim yearlyChangeRange As Range
            Set yearlyChangeRange = ws.Range(ws.Cells(outputRow, outputColumn + 1), ws.Cells(outputRow, outputColumn + 1))
            
            With yearlyChangeRange
                ' Clear any existing conditional formatting
                .FormatConditions.Delete
                
                ' Define the positive change format (Green)
                Dim positiveFormat As FormatCondition
                Set positiveFormat = .FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0")
                With positiveFormat
                    .Interior.ColorIndex = 10 ' Green
                End With
                
                ' Define the negative change format (Red)
                Dim negativeFormat As FormatCondition
                Set negativeFormat = .FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="=0")
                With negativeFormat
                    .Interior.ColorIndex = 3 ' Red
                End With
            End With
            
            ' Reset the aggregated values for the new ticker
            ticker = ws.Cells(i, tickerColumn).Value
            openingPrice = ws.Cells(i, openColumn).Value
            totalVolume = 0
            
            ' Move to the next output row
            outputRow = outputRow + 1
        End If
        
        ' Accumulate the volume for the current ticker
        volume = ws.Cells(i, volumeColumn).Value
        totalVolume = totalVolume + volume
        
        ' Calculate the yearly change and percentage change
        closingPrice = ws.Cells(i, closeColumn).Value
        yearlyChange = closingPrice - openingPrice
        If openingPrice <> 0 Then
            percentChange = (yearlyChange / openingPrice) * 100
        Else
            percentChange = 0 ' Handle the case when opening price is zero
        End If
    Next i
    
    ' Output the last aggregated data for the final ticker
    ws.Cells(outputRow, outputColumn).Value = ticker
    ws.Cells(outputRow, outputColumn + 1).Value = yearlyChange
    ws.Cells(outputRow, outputColumn + 2).Value = percentChange
    ws.Cells(outputRow, outputColumn + 3).Value = totalVolume
    
    Next yearSheet
    
End Sub


