Attribute VB_Name = "Module1"
Public Sub AdjustDataGridColumns _
           (DG As DBGrid, _
           adoData As Data, _
           intRecord As Integer, _
           intField As Integer, _
           Optional AccForHeaders As Boolean)



    Dim row As Long, col As Long
    Dim width As Single, maxWidth As Single
    Dim saveFont As StdFont, saveScaleMode As Integer
    Dim cellText As String
    
    'If number of records = 0 then exit from the sub
    If intRecord = 0 Then Exit Sub
    'Save the form's font for DataGrid's font
    'We need this for form's TextWidth method
    Set saveFont = DG.Parent.Font
    Set DG.Parent.Font = DG.Font
    'Adjust ScaleMode to vbTwips for the form (parent).
    saveScaleMode = DG.Parent.ScaleMode
    DG.Parent.ScaleMode = vbTwips
    'Always from first record...
    adoData.Recordset.MoveFirst
    maxWidth = 0
    'We begin from the first column until the last column
    For col = 0 To intField - 1
        adoData.Recordset.MoveFirst
        'Optional param, if true, set maxWidth to
        'width of DG.Parent
        If AccForHeaders Then
            maxWidth = DG.Parent.TextWidth(DG.Columns(col).Text) + 200
        End If
        'Repeat from first record again after we have
        'finished process the last record in
        'former column...
        adoData.Recordset.MoveFirst
        For row = 0 To intRecord - 1
            'Get the text from the DataGrid's cell
            If intField = 1 Then
            Else  'If number of field more than one
                cellText = DG.Columns(col).Text
            End If
            'Fix the border...
            'Not for "multiple-line text"...
            width = DG.Parent.TextWidth(cellText) + 200
            'Update the maximum width if we found
            'the wider string...
            If width > maxWidth Then
               maxWidth = width
               DG.Columns(col).width = maxWidth
            End If
            'Process next record...
            adoData.Recordset.MoveNext
        Next row
        'Change the column width...
        DG.Columns(col).width = maxWidth 'kolom terakhir!
    Next col
    'Change the DataGrid's parent property
    Set DG.Parent.Font = saveFont
    DG.Parent.ScaleMode = saveScaleMode
    'If finished, then move pointer to first record again
    adoData.Recordset.MoveFirst
End Sub  'End of AdjustDataGridColumns

