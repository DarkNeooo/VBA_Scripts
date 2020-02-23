Sub cellsMerger()

    Dim column As Range
    Dim cell As Range
    Dim tempRange As Range
    Dim testRange As Range
    
    Set column = Application.Selection
    
    For Each cell In column
        If cell.Address = column.Item(1).Address Then
            Set tempRange = cell
        Else
            If cell = column.Item(cell.Row - 1) Or cell.Value = "" Or cell = tempRange.Item(1) Then
                Set tempRange = Union(tempRange, cell)
                If cell.Address = column.Item(column.Count).Address Then
                    tempRange.Merge
                End If
            Else
                tempRange.Merge
                Set tempRange = cell
            End If
        End If
    Next cell
End Sub
