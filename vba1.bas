Attribute VB_Name = "Module1"
Sub hw()
Dim ws As Worksheet
For Each ws In Worksheets

Dim total_vol As Double
    Dim row As Variant
    Dim yearchange As Variant

    Dim start As Variant
    Dim percentage As Variant
    Dim open_price As Variant
    Dim lastrow As Variant
    
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).row
    
      ' Searches for when the value of the next cell is different than that of the current cell
    row = 2
    start = 2
    percentage = 2
    open_price = 0
    
    ws.Range("j1").Value = "Ticker"
   ws.Range("k1").Value = "Yearly Change"
   ws.Range("l1").Value = "Percent Change"
   ws.Range("m1").Value = "Total Stock Volume"
    
    For i = 2 To lastrow
       If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
   'sum
       total_vol = total_vol + ws.Cells(i, 7)
       yearchange = ws.Cells(i, 6) - ws.Cells(start, 3)
      
    
If ws.Cells(start, 3).Value = 0 Then
percentage = 1
Else
percentage = yearchange / ws.Cells(start, 3) * 100
End If

    'location placement
    
       ws.Range("j" & row).Value = ws.Cells(i, 1)
       ws.Range("m" & row).Value = total_vol
       ws.Range("k" & row).Value = yearchange
       ws.Range("l" & row).Value = percentage
       
       total_vol = 0
       row = row + 1
       start = i + 1
       
      Else
      total_vol = total_vol + ws.Cells(i, 7)
      End If
    Next i
    
    
    Next ws
    
    
    
 End Sub
