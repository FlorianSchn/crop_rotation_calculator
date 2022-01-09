Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    If Target.Cells(1, 1).Address = "$B$2" Then
        Cells.Delete
        Cells.HorizontalAlignment = xlCenter
        Cells.VerticalAlignment = xlTop
        
        ' Create Reload Button
        Cells(2, 2).Value = "Aktualisieren"
        Cells(2, 2).Interior.Color = RGB(155, 194, 230)
        Cells(2, 2).Borders(xlRight).Weight = xlMedium
        Cells(2, 2).Borders(xlLeft).Weight = xlMedium
        Cells(2, 2).Borders(xlBottom).Weight = xlMedium
        Cells(2, 2).Borders(xlTop).Weight = xlMedium
        
        ' cropping
        Dim oCompostConfig As New CompostConfig
        Dim oCropsConfig As New CropsConfig
        oCompostConfig.Initialize "Kompostierung", "Betriebskonfiguration"
        oCropsConfig.Initialize "Ackerbau", "Betriebskonfiguration", oCompostConfig
        
        Dim sError As String
        sError = oCropsConfig.CheckForInvalidData
        If sError <> "" Then
            MsgBox sError
        Else
            oCropsConfig.PrintTable Cells(4, 2)
        End If
        
        ' Autofit
        For Each oRow In UsedRange.Rows
            Dim iNewlines As Integer, iTmp As Integer
            iNewlines = 0
            iTmp = 0
            For Each oCell In oRow.Columns
                iTmp = (Len(oCell.Value) - Len(Replace(oCell.Value, vbCrLf, ""))) / Len(vbCrLf)
                If iTmp > inewlinex Then
                    iNewlines = iTmp
                End If
            Next oCell
            oRow.RowHeight = 15 * (iNewlines + 1)
        Next oRow
        Cells.Columns.AutoFit
        
        Exit Sub
    End If
End Sub
