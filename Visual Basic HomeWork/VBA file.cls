VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Ticker()
 ' Loop through all Tickers
For Each ws In Worksheets
Dim WorksheetName As String
 

  ' Set an initial variable for holding the Ticker_Name
  Dim Ticker_Name As String

  ' Set an initial variable for holding the totalVolume per ticker
  Dim Total_Volumel As Long
  Total_Volume = 0

  ' Keep track of the location for each Ticker_Name in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
  'set headers
  ws.Cells(1, 9).Value = "Ticker"
  
   ws.Cells(1, 10).Value = "Total Stock Volume"
    ' Determine the Last Row
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
For I = 2 To LastRow
 
 ' Check if we are still within the same ticker name, if it is not...
    If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then

      ' Set the Ticker name
      Ticker_Name = ws.Cells(I, 1).Value

      ' Add to the Total Volume
      Total_Volume = Total_Volume + ws.Cells(I, 7).Value

      ' Print the Ticker_Name in the Summary Table
      ws.Range("I" & Summary_Table_Row).Value = Ticker_Name

      ' Print the Total_Volume to the Summary Table
      ws.Range("J" & Summary_Table_Row).Value = Total_Volume
      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Total_Volume
      Total_Volume = 0

    ' If the cell immediately following a row is the same Ticker...
    Else

      ' Add to the Total_Volume
      Total_Volume = Total_Volume + ws.Cells(I, 7).Value

    End If

Next I
  Next ws

End Sub



