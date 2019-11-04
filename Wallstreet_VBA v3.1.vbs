Sub WallStreet_VBA()

Dim ws As Worksheet
For Each ws In Sheets

    ' Set key variables
    Dim Ticker As String
    Dim Open_Price As Double
    Dim Close_Price As Double
    Dim Yearly_Change As Double
    Dim Percentage_Change As Double
    Dim Total_Volume As Double
    Dim Stock_Table_Row As Integer
       
    ' Set Stock Table Row
    Stock_Table_Row = 2
     
    ' Loop through all stocks in worksheet rows
    For i = 2 To 760192
       
    ' Pick up opening price by detecting change in Ticker with check for stocks for opening price of zero
        If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value And ws.Cells(i, 3).Value <> 0 Then
        Open_Price = ws.Cells(i, 3).Value
    
    ' If opening price is zero pick up opening price when it becomes non zero
        ElseIf ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value And ws.Cells(i, 3).Value = 0 And ws.Cells(i + 1, 3).Value <> 0 Then
        Open_Price = ws.Cells(i + 1, 3).Value
        
    ' Check if we are still within the same Ticker, if it is not then pick ticker and closing price
        ElseIf ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        Ticker = ws.Cells(i, 1).Value
        Close_Price = ws.Cells(i, 6).Value
        
    ' Calculate Total Volume for stock at change of Ticker
        Total_Volume = Total_Volume + ws.Cells(i, 7).Value
        
    ' Print Ticker, Total Volume,  Opening Price and Closing Price and to Table (Opening Price and Closing Price were not asked for but are helpful in checking for errors etc)
        ws.Range("I" & Stock_Table_Row).Value = Ticker
        ws.Range("M" & Stock_Table_Row).Value = Open_Price
        ws.Range("N" & Stock_Table_Row).Value = Close_Price
        ws.Range("L" & Stock_Table_Row).Value = Total_Volume
    
    ' Calculate and Print Yearly_Change & Percentage_Change
        Yearly_Change = (Close_Price - Open_Price)
        Percentage_Change = Yearly_Change / Open_Price
        ws.Range("J" & Stock_Table_Row).Value = Yearly_Change
        ws.Range("K" & Stock_Table_Row).Value = Percentage_Change
        
    ' Add one to Stock Table row to move to next row
        Stock_Table_Row = Stock_Table_Row + 1
    
    ' Reset the Total_Volume
        Total_Volume = 0
        
    ' If the cell immediately following a row is the same stock...
        Else
        
    ' Add to the Total_Volume and continue calculation
        Total_Volume = Total_Volume + ws.Cells(i, 7).Value
            
        End If
        
      Next i

' Conditional Formatting
    'Definining the variables:
      Dim rng As Range
      Dim condition1 As FormatCondition, condition2 As FormatCondition
      Dim lastrow As Long
    
    'Find the last non-blank cell in column J(10)
      lastrow = ws.Cells(Rows.Count, 10).End(xlUp).Row
    
     'Fixing/Setting the range on which conditional formatting is to be desired
      Set rng = ws.Range("J2", "J" & lastrow)
    
      'To delete/clear any existing conditional formatting from the range
       rng.FormatConditions.Delete
    
      'Defining and setting the criteria for each conditional format
       Set condition1 = rng.FormatConditions.Add(xlCellValue, xlGreater, "0")
       Set condition2 = rng.FormatConditions.Add(xlCellValue, xlLess, "0")
    
       'Defining and setting the format to be applied for each condition
       With condition1
        .Interior.ColorIndex = 4
       End With
    
       With condition2
         .Interior.ColorIndex = 3
          .Font.Bold = True
       End With
    
' Challenge

    ' Set challenge variables
    Dim TickerB As String
    Dim TickerC As String
    Dim TickerD As String
    Dim Greatest_Percent_Inc As Double
    Dim Greatest_Percent_Dec As Double
    Dim Greatest_Total_Vol As Double
    Dim Challenge_Table_Row As Integer
    
    ' Set initial values
    Challenge_Table_Row = 2
    TickerB = ws.Cells(2, 9).Value
    Greatest_Percent_Inc = ws.Cells(2, 11).Value
    Greatest_Percent_Dec = ws.Cells(2, 11).Value
    Greatest_Total_Vol = ws.Cells(2, 12).Value
    
    ' Loop through all stocks in worksheet rows
    For j = 2 To lastrow
    
    ' If next cell Percent Increase is greater use next cell values
        If ws.Cells(j, 11).Value > Greatest_Percent_Inc Then
        TickerB = ws.Cells(j, 9).Value
        Greatest_Percent_Inc = ws.Cells(j, 11).Value
        End If

    ' If next cell Percent Decrease is less use next cell values
        If ws.Cells(j, 11).Value < Greatest_Percent_Dec Then
        TickerC = ws.Cells(j, 9).Value
        Greatest_Percent_Dec = ws.Cells(j, 11).Value
        End If

    ' If next cell Volume is greater use next cell values
        If ws.Cells(j, 12).Value > Greatest_Total_Vol Then
        TickerD = ws.Cells(j, 9).Value
        Greatest_Total_Vol = ws.Cells(j, 12).Value
        End If

    ' Print Ticker and Greatest_Percent_Inc  to Table
        ws.Range("Q2").Value = TickerB
        ws.Range("R2").Value = Greatest_Percent_Inc
        ws.Range("Q3").Value = TickerC
        ws.Range("R3").Value = Greatest_Percent_Dec
        ws.Range("Q4").Value = TickerD
        ws.Range("R4").Value = Greatest_Total_Vol
        
    Next j

' Add Headers & Labels
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly_Change"
    ws.Range("K1").Value = "Percentage_Change"
    ws.Range("L1").Value = "Total_Volume"
    ws.Range("M1").Value = "Open_Price"
    ws.Range("N1").Value = "Close_Price"
    ws.Range("Q1").Value = "Ticker"
    ws.Range("R1").Value = "Value"
    ws.Range("P2").Value = "Greatest % Increase"
    ws.Range("P3").Value = "Greatest % Decrease"
    ws.Range("P4").Value = "Greatest Total Volume"



Next ws


        
End Sub


