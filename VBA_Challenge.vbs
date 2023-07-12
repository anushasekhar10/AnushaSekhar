Sub Stock_Program()

'Declare Required Variables
Dim Ticker_Name As String
Dim Ticker_Table_Row As Integer
Dim LastRow As Long
Dim Ticker_Total As Long
Dim Ticker_Volume As Double
Dim closingPrice As Double
Dim openingPrice As Double
Dim YearlyChange As Double
Dim percentChange As Double
Dim summaryLastRow As Long
    
'Command to run code across all worksheets
For Each ws In Worksheets

'Below command can be used to sort data for processing, in our case the working data file is pre-sorted.
'Columns.Sort key1:=Columns("A"), Order1:=xlAscending, Key2:=Columns("B"), Order2:=xlAscending, Header:=xlYes

Dim WorksheetName As String

'Assign column headers for calculation
ws.Cells(1, 9) = "Ticker"
ws.Cells(1, 10) = "Yearly Change"
ws.Cells(1, 11) = "Percent Change"
ws.Cells(1, 12) = "Total Stock Volume"

'Find lastrow for processing
LastRow = ws.Cells(Rows.Count, 2).End(xlUp).Row
WorksheetName = ws.Name

'Initial row count as 2 to skip Header
Ticker_Table_Row = 2
  
 For i = 2 To LastRow
  
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

      Ticker_Name = ws.Cells(i, 1).Value
      ws.Range("I" & Ticker_Table_Row).Value = Ticker_Name
      
      closingPrice = ws.Cells(i, 6).Value
      
      Ticker_Volume = Ticker_Volume + Cells(i, 7).Value
      ws.Range("L" & Ticker_Table_Row).Value = Ticker_Volume

' Calculate yearly change and percent change
      YearlyChange = closingPrice - openingPrice
      If openingPrice <> 0 Then
        percentChange = YearlyChange / openingPrice
      Else
        percentChange = 0
      End If
          ws.Range("J" & Ticker_Table_Row).Value = YearlyChange
          ws.Range("K" & Ticker_Table_Row).Value = percentChange
            
' Conditional formatting for positive (green) and negative (red) changes
        If YearlyChange > 0 Then
          ws.Cells(Ticker_Table_Row, 10).Interior.Color = RGB(0, 255, 0)
        ElseIf YearlyChange < 0 Then
          ws.Cells(Ticker_Table_Row, 10).Interior.Color = RGB(255, 0, 0)
        End If
       
' Format the percent change as percentage
        ws.Cells(Ticker_Table_Row, 11).NumberFormat = "0.00%"
        

        openingPrice = ws.Cells(i + 1, 3).Value
       
' Increase counter variables by 1
          Ticker_Table_Row = Ticker_Table_Row + 1
' Reset variable for next iteration
          Ticker_Total = 0
          Ticker_Volume = 0
          
                          
    Else
      
      Ticker_Total = Ticker_Total + Cells(i, 3).Value
      Ticker_Volume = Ticker_Volume + Cells(i, 7).Value

       If openingPrice = 0 Then
          openingPrice = ws.Cells(i, 3).Value
        End If

    End If

  Next i
  
 
 'Section of code below calculate data for the summary table
 
    summaryLastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
 ' Addtional variables needed for calculations
    Dim maxPercentIncrease As Double
    Dim maxPercentDecrease As Double
    Dim maxTotalVolume As Double
    Dim maxPercentIncreaseTicker As String
    Dim maxPercentDecreaseTicker As String
    Dim maxTotalVolumeTicker As String
    Dim currentPercentChange As Double
    Dim currentTotalVolume As Double
      
     
    maxPercentIncrease = 0
    maxPercentDecrease = 0
    maxTotalVolume = 0
    
    For j = 2 To summaryLastRow
      
      currentPercentChange = ws.Cells(j, 11).Value
      currentTotalVolume = ws.Cells(j, 12).Value
      ' Check if the current percent change is greater than the previous maximum percent increase
      If currentPercentChange > maxPercentIncrease Then
        maxPercentIncrease = currentPercentChange
        maxPercentIncreaseTicker = ws.Cells(j, 9).Value
      End If
      ' Check if the current percent change is smaller than the previous maximum percent decrease
      If currentPercentChange < maxPercentDecrease Then
        maxPercentDecrease = currentPercentChange
        maxPercentDecreaseTicker = ws.Cells(j, 9).Value
      End If
      ' Check if the current total volume is greater than the previous maximum total volume
      If currentTotalVolume > maxTotalVolume Then
        maxTotalVolume = currentTotalVolume
        maxTotalVolumeTicker = ws.Cells(j, 9).Value
      End If
    Next j
      ' Assign values to the summary table
    ws.Cells(2, 16).Value = maxPercentIncreaseTicker
    ws.Cells(3, 16).Value = maxPercentDecreaseTicker
    ws.Cells(4, 16).Value = maxTotalVolumeTicker
    ws.Cells(2, 17).Value = maxPercentIncrease
    ws.Cells(3, 17).Value = maxPercentDecrease
    ws.Cells(4, 17).Value = maxTotalVolume
    ws.Cells(2, 17).NumberFormat = "0.00%"
    ws.Cells(3, 17).NumberFormat = "0.00%"
    ws.Cells(4, 17).NumberFormat = "0.00%"
    ' Format the columns in the summary table
    ws.Columns("A:Q").AutoFit
  
'Loop through the next worksheet
Next ws

End Sub

