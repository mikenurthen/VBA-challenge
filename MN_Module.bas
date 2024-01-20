Attribute VB_Name = "Module1"

Sub stock_data():
For Each ws In Worksheets
Dim WorksheetName As String
Dim Ticker As String
Dim Yearly_Change As Double
Dim Percent_Change As Double
Dim Total_Volume As Double
Dim i As Double
Dim j As Integer
Dim Stock_Open As Double
Dim Stock_Close As Double
Dim Start_Row As Double
Dim LastRow As Double
Dim Summary_Table_Row As Double

' Keep track of the location for each Ticker in the summary table
Summary_Table_Row = 2
Start_Row = 2
  
'Set an initial variable for holding the Total Stock Volume per Ticker
Total_Volume = 0

'Set an initial variable for holding the Yearly Change per Ticker
Yearly_Change = 0

'Set an initial variable for holding the Percent Change per Ticker
Percent_Change = 0

'Dimension Additional Variables
j = 0

'Determine the Last Row
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Add word "Ticker" to i1 Column Header
ws.Range("i1").Value = "Ticker"

'Add words "Yearly Change" to J1 Column Header
ws.Range("J1").Value = "Yearly Change"

'Add words "Percent Change" to K1 Column Header
ws.Range("K1").Value = "Percent Change"

'Add words "Percent Change" to L1 Column Header
ws.Range("L1").Value = "Total Stock Volume"

ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"


'Loop through all the ticker symbols
For i = 2 To LastRow
    'Check if we are still within the same Ticker, if it is not...
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
    'Set the Ticker name
    Ticker = ws.Cells(i, 1).Value
    
    'Add up all the stock volume
    Total_Volume = Total_Volume + ws.Cells(i, 7).Value

    'Print the Ticker name in Column "I"
    ws.Range("I" & Summary_Table_Row).Value = Ticker
    
    'Print total stock volume in Column L
    ws.Range("L" & Summary_Table_Row).Value = Total_Volume
        
    'Find the Yearly Change
    Open_Price = ws.Range("C" & Start_Row).Value
    Close_Price = ws.Range("F" & i).Value
    Yearly_Change = Close_Price - Open_Price

    'Find the Percent Change. Need to deal with a 0 in the denominator
    If Open_Price = 0 Then
        Percent_Change = 0
    Else
        Percent_Change = Yearly_Change / Open_Price
    End If
    
    'Print Values of Yearly Change and Percent Change
    ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
    ws.Range("K" & Summary_Table_Row).Value = Percent_Change
    ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
    
    
    'Conditional Formatting
    If ws.Range("J" & 2 + j).Value > 0 Then
        ws.Range("J" & 2 + j).Interior.ColorIndex = 4
    ElseIf ws.Range("J" & 2 + j).Value < 0 Then
        ws.Range("J" & 2 + j).Interior.ColorIndex = 3
    Else
        ws.Range("J" & 2 + j).Interior.ColorIndex = 0
    End If
    
    'Reset Total Stock Volume
    Total_Volume = 0
    'Increase Summary Table Row by 1 to account for next ticker placement
    Summary_Table_Row = Summary_Table_Row + 1
    'Begin at the next stock ticker
    Start_Row = i + 1
    'Reset j
    j = j + 1
      
Else
Total_Volume = Total_Volume + ws.Cells(i, 7).Value

End If
Next i

' take the max and min and place them in a separate part in the worksheet
    Range("Q2") = "%" & WorksheetFunction.Max(Range("K2:K" & LastRow)) * 100
    Range("Q3") = "%" & WorksheetFunction.Min(Range("K2:K" & LastRow)) * 100
    Range("Q4") = WorksheetFunction.Max(Range("L2:L" & LastRow))

    ' returns one less because header row not a factor
    increase_number = WorksheetFunction.Match(WorksheetFunction.Max(Range("K2:K" & LastRow)), Range("K2:K" & LastRow), 0)
    decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(Range("K2:K" & LastRow)), Range("K2:K" & LastRow), 0)
    volume_number = WorksheetFunction.Match(WorksheetFunction.Max(Range("L2:L" & LastRow)), Range("L2:L" & LastRow), 0)

    ' final ticker symbol for total, greatest % of increase and decrease, and greatest total volume
    Range("P2") = ws.Cells(increase_number + 1, 9)
    Range("P3") = ws.Cells(decrease_number + 1, 9)
    Range("P4") = ws.Cells(volume_number + 1, 9)
Next ws
End Sub

