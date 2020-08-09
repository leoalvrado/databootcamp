Attribute VB_Name = "Module3"
Sub COLOR()

Dim i As Variant
Dim lastrow As Variant
Dim ws As Worksheet
For Each ws In Worksheets

lastrow = ws.Cells(Rows.Count, 11).End(xlUp).row


    For i = 2 To lastrow

    ws.Cells(i, 11).Select
        If Selection.Value > 0 Then

        
                 Selection.Interior.COLOR = RGB(0, 255, 0)
                    Else
                     
                  Selection.Interior.COLOR = RGB(255, 0, 0)

            

        End If

    Next i
    
Next ws
    

End Sub



