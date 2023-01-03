# import csv files to microsoft access

Public Sub ImportCSVFile(fileName as String)

    DoCmd.TransferText acImportDelim, "Specification file", "tblName", _
            fileName, true, , 1252

End Sub


Private Sub cdSelectFile Click ()

    Dim fd As FileDialog
    
    Set fd = Application.FileDialog (msoFileDialogOpen)
    
    with fd
        .AllowMultiSelect = False
        
        .Filters.Clear
        •Filters.Add "Any file","*.*",1
        •Filters .Add "Comma separated File", "*.csv; *.txt", 2
        .FilterIndex = 2
        
        If .Show Then
        Me.txtFileName.Value = .SelectedItems.Item (1)
        End If
End With

End Sub
