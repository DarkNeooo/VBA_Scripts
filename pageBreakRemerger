Sub pageBreakRemerger()
    Dim selectedRange As Range
    Dim mergedRange As Range
    Dim cell As Range
    Dim pageBreaks As HPageBreaks
    Dim pageBreakItem As HPageBreak
    Dim newMergedRange1, newMergedRange2 As Range
    
    Set selectedRange = Application.Selection

    Set pageBreaks = Worksheets(1).HPageBreaks
    
    For Each pageBreakItem In pageBreaks
        Set cell = Cells(pageBreakItem.Location.Row, 1)
        
        If cell.MergeCells Then
            Set mergedRange = cell.MergeArea
            
            If cell.Address <> mergedRange.Item(1, 1).Address Then
                mergedValue = mergedRange.Item(1, 1)
                
                Set newMergedRange1 = Range(mergedRange.Cells(1, 1), Cells(pageBreakItem.Location.Row - 1, 1))
                Set newMergedRange2 = Range(cell, mergedRange.Cells(mergedRange.Count, 1))
                
                mergedRange.UnMerge
                
                newMergedRange1.Merge
                newMergedRange1.Item(1, 1) = mergedValue
                newMergedRange2.Merge
                newMergedRange2.Item(1, 1) = mergedValue
            End If
        End If
    Next pageBreakItem
End Sub
