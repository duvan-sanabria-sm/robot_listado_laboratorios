Sub BorrarTodoCompletoEnTodasLasHojas()
    Dim ws As Worksheet
    Dim forma As Shape

    For Each ws In ThisWorkbook.Worksheets
        ' Borrar contenido y formato de celdas
        ws.Cells.Clear

        ' Eliminar todas las formas, imágenes y gráficos
        For Each forma In ws.Shapes
            forma.Delete
        Next forma

        ' Restablecer tamaño de columnas y filas
        ws.Cells.EntireColumn.ColumnWidth = 8.43 ' Valor predeterminado de Excel
        ws.Cells.EntireRow.RowHeight = 15 ' Valor predeterminado de Excel
    Next ws
End Sub
