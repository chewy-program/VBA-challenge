Sub tickerandvolume()

  ' Set an initial variable for holding the ticker
  Dim ticker As String
    
  ' Set an initial variable for holding the volume
  Dim volume, change, open_value, close_value, percentage_change As Double
  volume = 0
  change = 0
  open_value = 0
  close_value = 0
  percentage_change = 0
  
  Dim tickerdictionary
  Set tickerdictionary = CreateObject("Scripting.Dictionary")
   
  Dim summarydictionary
  Set summarydictionary = CreateObject("Scripting.Dictionary")
  summarydictionary.Add "volume_ticker", "None"
  summarydictionary.Add "volume_value", 0
  summarydictionary.Add "largest_change_ticker", "None"
  summarydictionary.Add "largest_change_value", 0
  summarydictionary.Add "smallest_change_ticker", "None"
  summarydictionary.Add "smallest_change_value", 0
  
  ' Keep track of the location for each ticker
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
  
  'define LastRow
  Dim LastRow As Long
    LastRow = Cells(Rows.Count, 2).End(xlUp).Row

  ' Loop through all stocks
For i = 2 To LastRow
    
    ' Set the ticker name
    ticker = Cells(i, 1).Value
      
    'find the open price
    open_price = Cells(i, 3).Value
      
    'find the close price
    close_price = Cells(i, 6).Value
    
    'create tickerdictionary to find change between first open and close
    Dim open_price_key, close_price_key As String
    open_price_key = ticker + "_first_open_price"
    close_price_key = ticker + "_last_close_price"
    
    If tickerdictionary.Exists(open_price_key) <> True Then
        tickerdictionary.Add open_price_key, open_price
        End If
    If tickerdictionary.Exists(close_price_key) <> True Then
        tickerdictionary.Add close_price_key, close_price
    Else
        tickerdictionary(close_price_key) = close_price
        End If
    ' add in the change
    change = tickerdictionary(close_price_key) - tickerdictionary(open_price_key)
    
    ' Print the change in the Summary Table
    Range("K" & Summary_Table_Row).Value = change
      
    'add in the percentage change
    percentage_change = change / tickerdictionary(open_price_key)
    
    Range("L" & Summary_Table_Row).Value = percentage_change
    
    
    ' Add to the volume
    volume = volume + Cells(i, 7).Value
      
    ' Print the ticker in the Summary Table
    Range("J" & Summary_Table_Row).Value = ticker
      
    ' Print the volume to the summary Table
    Range("M" & Summary_Table_Row).Value = volume
    
    'if statement for conditional formatting
    
    'Add in largest volume of summarytable to summarydictionary
    
    If summarydictionary("volume_value") < volume Then
        summarydictionary("volume_value") = volume
        summarydictionary("volume_ticker") = ticker
    End If
    

    
    ' Check if we are still within the same ticker value
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
      
      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the volume
      volume = 0
      
      If summarydictionary("largest_change_value") < percentage_change Then
        summarydictionary("largest_change_value") = percentage_change
        summarydictionary("largest_change_ticker") = ticker
      End If
      If summarydictionary("smallest_change_value") > percentage_change Then
        summarydictionary("smallest_change_value") = percentage_change
        summarydictionary("smallest_change_ticker") = ticker
      End If
    End If
  
Next i

    'conditional formatting
For i = 2 To Summary_Table_Row
    If Cells(i, 11).Value > 0 Then
       Cells(i, 11).Interior.ColorIndex = 4
    ElseIf Cells(i, 11).Value < 0 Then
       Cells(i, 11).Interior.ColorIndex = 3
    Else
       Cells(i, 11).Interior.ColorIndex = 0
    End If
Next i

    'insert summary table values
    Range("J1").Value = "Ticker"
    Range("K1").Value = "Yearly Change"
    Range("L1").Value = "Percentage Change"
    Range("M1").Value = "Total Volume"
    
    'insert summary dictionary values
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    
    Range("P2").Value = summarydictionary("largest_change_ticker")
    Range("Q2").Value = summarydictionary("largest_change_value")
    Range("O2").Value = "Greatest % increase"
    
    Range("P3").Value = summarydictionary("smallest_change_ticker")
    Range("Q3").Value = summarydictionary("smallest_change_value")
    Range("O3").Value = "Greatest % decrease"
    
    Range("P4").Value = summarydictionary("volume_ticker")
    Range("Q4").Value = summarydictionary("volume_value")
    Range("O4").Value = "Greatest total volume"
    
End Sub