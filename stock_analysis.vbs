Attribute VB_Name = "Module1"
Sub Stock_analysis()
    ' Set variable for holding ticker symbol
    Dim ticker As String
    
    ' Set variable for holding volume total
    Dim vol As Double
    vol = 0
    
    ' Set variable for last row
    Dim lastrow As Long
    
    ' Set Variable for open
    Dim year_open As Double
    
    ' Set Variable for close
    Dim year_close As Double
    Dim percent_change As Double
    
   ' Loop through all sheets
   For Each ws In Worksheets
   
   ' Add ticker label
   ws.Cells(1, 9) = "Ticker"
   
   ' Add total stock volume label
   ws.Cells(1, 10) = "Total Stock Volume"
   
   ' Add Yearly Change label
   ws.Cells(1, 11) = "Yearly Change"
   
   ' Addpercent change label
   ws.Cells(1, 12) = "Percent Change"
    
    ' Place each ticker symbol in ticker total table
    Dim total_table As Integer
    total_table = 2
        
    ' Find Last Row
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Loop through volume
    For i = 2 To lastrow
        
        ' Find year open
        year_open = ws.Cells(i, 3)
        
      ' Check if ticker symbol is the same if not...
      If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
      
        'Find year close
        year_close = ws.Cells(i, 6)
      
        'Set the ticker symbol
        ticker = ws.Cells(i, 1).Value
        
        ' Add to the Volume total
        vol = vol + ws.Cells(i, 7).Value
    
        ' Print the ticker in the total table
        ws.Range("I" & total_table).Value = ticker
        
        ' Print the Volume to the total table
        ws.Range("J" & total_table).Value = vol
        
        ' Print Difference in yearly open and close
        ws.Range("K" & total_table).Value = (year_close - year_open)
        
            ' If positive change make cell green or ...
                If ws.Range("K" & total_table).Value > 0 Then
                
            ' Change cell to green
            ws.Range("K" & total_table).Interior.ColorIndex = 4
            Else
            
            'Change cell to red
            ws.Range("K" & total_table).Interior.ColorIndex = 3
            End If
        
            ' Set percent change to 0 if open and close is 0
            If year_open = 0 Then
            ws.Range("L" & total_table).Value = Format(0, "Percent")
       Else
    
    
        ' Calculate percent change
          percent_change = ((year_close - year_open) / year_open)
         
        
        ' Print percent change in yearly open and close
        ws.Range("L" & total_table).Value = percent_change
        End If
        
        ' Add one to Volume in the total table
        total_table = total_table + 1
        
        ' Reset the Volume total
        vol = 0
        
    ' If the next immediate cell is the same ticker...
    Else
        
        ' Add to the Volume total
        vol = vol + ws.Cells(i, 7).Value
        
        
      End If
      
Next i

Next ws

End Sub
