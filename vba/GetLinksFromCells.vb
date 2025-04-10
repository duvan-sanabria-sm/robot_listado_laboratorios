Sub ReplaceCellWithHyperlink(sheetName As String, numColumn As String)
    Dim ws As Worksheet
    Dim cell As Range
    Dim lastRow As Long
    Set ws = ActiveSheet

    lastRow = ws.Cells(ws.Rows.Count, numColumn).End(xlUp).Row

    Dim i As Long
    For i = 2 To lastRow ' Empieza en la fila 2, asumiendo encabezado
        Set cell = ws.Range(numColumn & i)
        If cell.Hyperlinks.Count > 0 Then
            cell.Value = cell.Hyperlinks(1).Address ' Reemplaza el texto con el link real
        End If
    Next i
End Sub