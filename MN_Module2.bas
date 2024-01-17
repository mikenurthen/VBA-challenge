Attribute VB_Name = "Module1"
Sub stock_data()

For Each ws In Worksheets

Dim WorksheetName As String
Dim Ticker As String
Dim Yearly_Change As Double
Dim Percent_Change As Double
Dim Total_Stock_Volume As Long
Dim i As Long
Dim j As Long
Dim Total As Double
Dim change As Double



' Keep track of the location for each Ticker in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
  
'Set an initial variable for holding the Total Stock Volume per Ticker
Total_Stock_Volume = 0

'Set an initial variable for holding the Yearly Change per Ticker
Yearly_Change = 0

'Set an initial variable for holding the Percent Change per Ticker
Percent_Change = 0

'Determine the Last Row
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row


'Get the Worksheet Name
WorksheetName = ws.Name


'Add word "Ticker" to i1 Column Header
ws.Range("i1").Value = "Ticker"

'Add words "Yearly Change" to J1 Column Header
ws.Range("J1").Value = "Yearly Change"

'Add words "Percent Change" to K1 Column Header
ws.Range("K1").Value = "Percent Change"

'Add words "Percent Change" to L1 Column Header
ws.Range("L1").Value = "Total Stock Volume"

 'Loop through all the ticker symbols
 For i = 2 To LastRow

'Check if we are still within the same Ticker, if it is not...
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
    Total = Total + Cells(i, 7).Value
    
    If Total = 0 Then
    ws.Range("i" & 2 + j).Value = ws.Cells(i, 1).Value
    ws.Range("I" & 2 + j).Value = Cells(i, 1).Value
    ws.Range("J" & 2 + j).Value = 0
    ws.Range("K" & 2 + j).Value = "%" & 0
    ws.Range("L" & 2 + j).Value = 0
    
    Else
    
    
    
    'Set the Ticker name
    Ticker = ws.Cells(i, 1).Value
    
    'Print the Ticker name in Range i2
    ws.Range("I" & Summary_Table_Row).Value = Ticker
    
    ' This is the Total Calculation first guy was trying to give me "Total = Total + ws.Cells(i, 7).Value"
    
    'If ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value Then
    'Total = Total + ws.Cells(i, 7).Value
    





'Find the Yearly Change from Day Open to Day Close and Print in Range J2
ws.Cells(


'Add one to the summary table row
Summary_Table_Row = Summary_Table_Row + 1



Else
End If
Next i
Next ws

End Sub
