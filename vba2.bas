Attribute VB_Name = "Module2"

Sub hw2()


'Dim ws As Worksheet
For Each ws In Worksheets

Dim Range_Percent As Range
  Dim Range_Volumen As Range
  Dim Percent_Maximo As Double
  Dim Percent_Minimo As Double
  Dim Volumen_Maximo As Double
  Dim MaxTicker As Variant
  Dim MinTicker As String
  Dim VolMaxTicker As String
lastrow = ws.Cells(Rows.Count, 10).End(xlUp).row

  'Set ranges
  Set Range_Percent = ws.Range("L2:L" & lastrow)
  Set Range_Volume = ws.Range("M2:M" & lastrow)
  
    ws.Range("o2").Value = "Greatest % Increase"
 ws.Range("o3").Value = "Greatest % Decrease"
    ws.Range("o4").Value = "Greatest Total Volume"
   ws.Range("p1").Value = "Ticker"
   ws.Range("q1").Value = "Value"

'Worksheet function MIN returns the smallest value in a range


Percent_Maximo = Application.WorksheetFunction.Max(Range_Percent)
ws.Range("q2") = Percent_Maximo
Percent_Minimo = Application.WorksheetFunction.Min(Range_Percent)
ws.Range("q3") = Percent_Minimo
Volumen_Maximo = Application.WorksheetFunction.Max(Range_Volume)
ws.Range("q4") = Volumen_Maximo

MaxTicker = Application.Index(ws.Range("J:J"), Application.Match(Percent_Maximo, ws.Range("L:L"), 0))
ws.Range("p2") = MaxTicker
MinTicker = Application.Index(ws.Range("J:J"), Application.Match(Percent_Minimo, ws.Range("L:L"), 0))
ws.Range("p3") = MinTicker
VolMaxTicker = Application.Index(ws.Range("J:J"), Application.Match(Volumen_Maximo, ws.Range("M:M"), 0))
ws.Range("p4") = VolMaxTicker


Next ws

End Sub


'
