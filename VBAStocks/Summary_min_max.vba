Attribute VB_Name = "Summary_min_max"

Sub TickerSummary()

For Each ws In Worksheets
Dim WorksheetName As String
WorksheetName = ws.Name


  ' Set headers
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
        
    ws.Columns("I:P").AutoFit
    
  ' Set an initial variable for holding the ticker symbol
  Dim ticker As String
  ticker = 1

  ' Set an initial variable for holding the total per ticker
  Dim ticker_total As Double
  ticker_total = 0
  

  ' Keep track of the location for each ticker in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
  lastrow = ws.Cells(Rows.Count, 1).End(xlUp).row
  
  LastColumn = ws.Cells(1, Columns.Count).End(xlToLeft).column
  
       
  Dim yearly_change As Double
  Dim percent_change As Double
  Dim opening_price As Double
  Dim closing_price As Double
  opening_price = ws.Cells(2, 3)
  
  
  'Min & Max section variables
  Dim max_percent As Double
  Dim min_percent As Double
  Dim max_volume As Double
  max_percent = 0
  min_percent = 0
  max_volume = 0
  
  Dim max_ticker As String
  Dim min_ticker As String
  Dim max_volume_ticker As String
  max_ticker = " "
  min_ticker = " "
  max_volume_ticker = " "
  
  
  
  
 
  ' Loop through all ticker symbols
  For I = 2 To lastrow

    ' Check if we are still within the same ticker, if it is not...
    If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then

      ' Set the ticker symbol name
      ticker = ws.Cells(I, 1).Value

      ' Set the closing price value to obtain yearly change
      closing_price = ws.Cells(I, 6)
      yearly_change = closing_price - opening_price
      
      If (yearly_change > 0) Then
        ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
            ElseIf (yearly_change <= 0) Then
                ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                End If

       ' Obtain Percent Change
      If opening_price <> 0 Then
            percent_change = (closing_price - opening_price) / opening_price * 100
                Else
                    percent_change = 0
                End If
                
      ' Add to the Ticker Volume Total
      ticker_total = ticker_total + ws.Cells(I, 7).Value
      
      ' Print the Ticker Name in the Summary Table
      ws.Range("I" & Summary_Table_Row).Value = ticker
      
      ' Print the Yearly Change in the Summary Table
      ws.Range("J" & Summary_Table_Row).Value = yearly_change
      
      ' Print the Percent Change in the Summary Table
      ws.Range("K" & Summary_Table_Row).Value = (CStr(percent_change) & "%")

      ' Print the Volume Amount to the Summary Table
      ws.Range("L" & Summary_Table_Row).Value = ticker_total

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the ticker total values
      yearly_change = 0
      closing_price = 0
      opening_price = ws.Cells(I + 1, 3).Value
      ' If the cell immediately following a row is the same ticker...
    
    
    ' Find max% and min% values
    If (percent_change > max_percent) Then
                    max_percent = percent_change
                    max_ticker = ticker
                ElseIf (percent_change < min_percent) Then
                    min_percent = percent_change
                    min_ticker = ticker
                End If
                
   ' Find max volume value
    If (ticker_total > max_volume) Then
                    max_volume = ticker_total
                    max_volume_ticker = ticker
                End If
                
    ' Reset counters
        percent_change = 0
        ticker_total = 0

      
    Else

      ' Add to the Ticker Volume Total
      ticker_total = ticker_total + ws.Cells(I, 7).Value
      

    End If
 
 
Next I

' Set headers
    ws.Cells(2, 14).Value = "Greatest % increase"
    ws.Cells(3, 14).Value = "Greatest % decrease"
    ws.Cells(4, 14).Value = "Greatest total volume"
    ws.Cells(1, 15).Value = "Ticker"
    ws.Cells(1, 16).Value = "Value"
    
    ws.Columns("N:P").AutoFit

    If Not Summary_Table_Row Then
            
                ws.Range("P2").Value = (CStr(max_percent) & "%")
                ws.Range("P3").Value = (CStr(min_percent) & "%")
                ws.Range("O2").Value = max_ticker
                ws.Range("O3").Value = min_ticker
                ws.Range("P4").Value = max_volume
                ws.Range("O4").Value = max_volume_ticker
                
            Else
                Summary_Table_Row = False
            End If


Next ws

End Sub


 


