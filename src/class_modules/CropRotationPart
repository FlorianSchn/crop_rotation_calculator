Private colCropsOrFieldWork As Collection
Private dData As Scripting.Dictionary
Private dDataValues As Scripting.Dictionary
Private oCropRotation As CropRotation
Private oCropsConfig As CropsConfig

Private Sub Class_Initialize()
    Set colCropsOrFieldWork = New Collection
    Set dData = New Scripting.Dictionary
    Set dDataValues = New Scripting.Dictionary
End Sub

Public Function Initialize(arg_oCropsConfig As CropsConfig, arg_cReference As Range, arg_cCompostReference As Range, arg_oCropRotation As CropRotation)
    Set oCropRotation = arg_oCropRotation
    Set oCropsConfig = arg_oCropsConfig
    Dim i As Integer
    i = 1
    Dim oCrop As Crop, oFieldWork As FieldWork
    While arg_cReference.Offset(i, 0).Value <> ""
        If IsCrop(arg_cReference.Offset(i, 0).Value) Then
            Set oCrop = New Crop
            colCropsOrFieldWork.add oCrop
        Else
            Set oFieldWork = New FieldWork
            colCropsOrFieldWork.add oFieldWork
        End If
        i = i + 1
    Wend
    For i = 1 To colCropsOrFieldWork.Count
        Set colCropsOrFieldWork(i).NextInRot = colCropsOrFieldWork((i Mod colCropsOrFieldWork.Count) + 1)
        Set colCropsOrFieldWork(i).PrevInRot = colCropsOrFieldWork(((i + colCropsOrFieldWork.Count - 2) Mod colCropsOrFieldWork.Count) + 1)
    Next i
    
    ' initialize crops and field works
    For i = 1 To colCropsOrFieldWork.Count
        If colCropsOrFieldWork(i).IsCrop Then
            Set oCrop = colCropsOrFieldWork(i)
            oCrop.Initialize arg_oCropsConfig, arg_cReference.Offset(i, 0).Value, Me
        Else
            Set oFieldWork = colCropsOrFieldWork(i)
            oFieldWork.Initialize arg_oCropsConfig, arg_cReference.Offset(i, 0).Value, arg_cCompostReference, Me
        End If
    Next i
    For i = 1 To colCropsOrFieldWork.Count
        colCropsOrFieldWork(i).LateInitialize
    Next i
    Set oCrop = Nothing
    Set oFieldWork = Nothing
    
    ' calculate data
    dDataValues("Flächenanteil") = arg_cReference.Value
    dDataValues("Fläche") = arg_cReference.Value * CDbl(oCropRotation.DataValue("Fläche"))
    dData("Fläche") = Round(dDataValues("Fläche"), 1) & " ha"
End Function

Public Function CompostInitialize()
    For i = 1 To colCropsOrFieldWork.Count
        If Not IsCrop(colCropsOrFieldWork(i).DataValue("Frucht bzw. Feldarbeit")) Then
            colCropsOrFieldWork(i).CompostInitialize
        End If
    Next i
End Function

Public Function LateInitialize()
    dDataValues("Dauer") = 0#
    Dim i As Integer
    For i = 1 To colCropsOrFieldWork.Count
        If colCropsOrFieldWork(i).IsCrop Then
            dDataValues("Dauer") = dDataValues("Dauer") + _
                colCropsOrFieldWork(i).DataValue("Standzeit") + _
                colCropsOrFieldWork(i).DataValue("Brache danach")
        End If
    Next i
    dDataValues("Dauer") = dDataValues("Dauer") / 12
    dData("Dauer") = dDataValues("Dauer") & " Jahre"
    For i = 1 To colCropsOrFieldWork.Count
        colCropsOrFieldWork(i).LateLateInitialize
    Next i
End Function

Public Function LateLateInitialize()
    dDataValues("Deckungsbeitrag inkl. Leistungen") = 0#
    For i = 1 To colCropsOrFieldWork.Count
        dDataValues("Deckungsbeitrag inkl. Leistungen") = dDataValues("Deckungsbeitrag inkl. Leistungen") + _
            colCropsOrFieldWork(i).DataValue("Deckungsbeitrag inkl. Leistungen") * colCropsOrFieldWork(i).DataValue("Fläche")
    Next i
    dDataValues("Deckungsbeitrag inkl. Leistungen") = dDataValues("Deckungsbeitrag inkl. Leistungen") / dDataValues("Fläche")
    dData("Deckungsbeitrag inkl. Leistungen") = Round(dDataValues("Deckungsbeitrag inkl. Leistungen"), 1) & " €/ha" & vbCrLf & _
        Round(dDataValues("Deckungsbeitrag inkl. Leistungen") * dDataValues("Fläche"), 1) & " €"

    dDataValues("Arbeitszeit") = 0#
    For i = 1 To colCropsOrFieldWork.Count
        dDataValues("Arbeitszeit") = dDataValues("Arbeitszeit") + _
            colCropsOrFieldWork(i).DataValue("Arbeitszeit") * colCropsOrFieldWork(i).DataValue("Fläche")
    Next i
    dDataValues("Arbeitszeit") = dDataValues("Arbeitszeit") / dDataValues("Fläche")
    dData("Arbeitszeit") = Round(dDataValues("Arbeitszeit"), 1) & " AKh/ha" & vbCrLf & _
        Round(dDataValues("Arbeitszeit") * dDataValues("Fläche"), 1) & " AKh"

    dDataValues("Stundenlohn") = dDataValues("Deckungsbeitrag inkl. Leistungen") / dDataValues("Arbeitszeit")
    dData("Stundenlohn") = Round(dDataValues("Stundenlohn"), 1) & " €/AKh"
    
    dDataValues("Wasserbedarf") = 0#
    For i = 1 To colCropsOrFieldWork.Count
        dDataValues("Wasserbedarf") = dDataValues("Wasserbedarf") + colCropsOrFieldWork(i).DataValue("Wasserbedarf")
    Next i
    dDataValues("Wasserbedarf") = dDataValues("Wasserbedarf") / dDataValues("Dauer")
    dData("Wasserbedarf") = Round(dDataValues("Wasserbedarf"), 0) & " mm/m²"
    
    NutrientHelper "Stickstoff", "kg"
    NutrientHelper "Phosphor", "kg"
    NutrientHelper "Kalium", "kg"
    NutrientHelper "Schwefel", "kg"
    NutrientHelper "Calcium", "kg"
    NutrientHelper "Magnesium", "kg"
    NutrientHelper "Bor", "g"
    NutrientHelper "Kupfer", "g"
    NutrientHelper "Mangan", "g"
    NutrientHelper "Zink", "g"
End Function

Public Function PrintLine(arg_aRowData As Variant, arg_cReference As Range, arg_bEssential As Boolean, arg_bWantPartLine As Boolean) As Integer
    PrintLine = 0
    Dim iColumnOffset As Integer
    iColumnOffset = 0
    For Each oCropOrFieldWork In colCropsOrFieldWork
        Dim iPrintedLines As Integer
        iPrintedLines = oCropOrFieldWork.PrintLine(arg_aRowData, arg_cReference.Offset(0, iColumnOffset), (Not dData.Exists(arg_aRowData(1))) And arg_bEssential)
        PrintLine = iPrintedLines
        iColumnOffset = iColumnOffset + 1
    Next oCropOrFieldWork
    If dData.Exists(arg_aRowData(1)) And arg_bWantPartLine Then
        With arg_cReference.Offset(PrintLine, 0)
            .NumberFormat = "@"
            .WrapText = True
            .Value = dData(arg_aRowData(1))
            If InStr(.Value, vbCrLf) Then
                .Characters(InStr(.Value, vbCrLf)).Font.Color = RGB(170, 170, 170)
            End If
            Range(.Address, arg_cReference.Offset(PrintLine, NumberOfEntries - 1).Address).Merge
            .Interior.Color = RGB(226, 239, 218)
            .Borders(xlEdgeLeft).Weight = xlMedium
            .Borders(xlEdgeLeft).Color = RGB(128, 128, 128)
        End With
        PrintLine = PrintLine + 1
    End If
    arg_cReference.Borders(xlEdgeLeft).Weight = xlMedium
    arg_cReference.Borders(xlEdgeLeft).Color = RGB(128, 128, 128)
End Function

Public Function CheckForInvalidData() As String
    Dim i As Integer
    i = 1
    While i <= colCropsOrFieldWork.Count And CheckForInvalidData = ""
        CheckForInvalidData = colCropsOrFieldWork(i).CheckForInvalidData
        i = i + 1
    Wend
End Function

Public Property Get NumberOfEntries() As Integer
    NumberOfEntries = colCropsOrFieldWork.Count
End Property

Public Function IsCrop(arg_sName) As Boolean
    Dim i As Integer, bFound As Boolean
    i = 1
    IsCrop = False
    While oCropsConfig.CropReference.Offset(0, i).Value <> "" And Not bFound
        If oCropsConfig.CropReference.Offset(0, i).Value = arg_sName Then
            IsCrop = True
            bFound = True
        End If
        i = i + 1
    Wend
End Function

Public Function DataValue(arg_sDataName As String) As String
    If dDataValues.Exists(arg_sDataName) Then
        DataValue = dDataValues(arg_sDataName)
    Else
        DataValue = ""
    End If
End Function

Public Function AccDataValue(arg_sDataName As String, arg_sCropOrFieldWorkFilter As String, arg_bIsNan As Boolean) As Variant
    If Not arg_bIsNan Then
        AccDataValue = 0#
    End If
    If dDataValues.Exists(arg_sDataName) And arg_sCropOrFieldWorkFilter = "" Then
        AccDataValue = dDataValues(arg_sDataName)
    Else
        For Each oCropOrFieldWork In colCropsOrFieldWork
            If arg_sCropOrFieldWorkFilter = "" Or oCropOrFieldWork.DataValue("Frucht bzw. Feldarbeit") = arg_sCropOrFieldWorkFilter Then
                If arg_bIsNan Then
                    AccDataValue = AccDataValue + oCropOrFieldWork.DataValue(arg_sDataName)
                Else
                    AccDataValue = AccDataValue + CDbl(oCropOrFieldWork.DataValue(arg_sDataName))
                End If
            End If
        Next oCropOrFieldWork
    End If
End Function

Private Function NutrientHelper(arg_sNutrientName As String, arg_sNutrientWeightUnit As String)
    dDataValues(arg_sNutrientName) = 0#
    For i = 1 To colCropsOrFieldWork.Count
        dDataValues(arg_sNutrientName) = dDataValues(arg_sNutrientName) + _
            colCropsOrFieldWork(i).DataValue(arg_sNutrientName) * colCropsOrFieldWork(i).DataValue("Ertrag bzw. Aufwand") * colCropsOrFieldWork(i).DataValue("Fläche")
    Next i
    dDataValues(arg_sNutrientName) = dDataValues(arg_sNutrientName) / dDataValues("Fläche")
    dData(arg_sNutrientName) = Round(dDataValues(arg_sNutrientName), 1) & Chr(160) & arg_sNutrientWeightUnit & "/ha" & vbCrLf & _
        Round(dDataValues(arg_sNutrientName) * dDataValues("Fläche"), 1) & Chr(160) & arg_sNutrientWeightUnit
End Function

Public Property Get CropRotation() As CropRotation
    Set CropRotation = oCropRotation
End Property
