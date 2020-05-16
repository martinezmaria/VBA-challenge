Attribute VB_Name = "Summary_loop"

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
    
    
          
    Else

      ' Add to the Ticker Volume Total
      ticker_total = ticker_total + ws.Cells(I, 7).Value
      

    End If
 
 
Next I




Next ws

End Sub


 



