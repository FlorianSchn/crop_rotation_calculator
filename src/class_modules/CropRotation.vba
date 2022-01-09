Private colCropRotParts As Collection
Private dData As Scripting.Dictionary
Private dDataValues As Scripting.Dictionary
Private iCompostCount As Integer

Private Sub Class_Initialize()
    Set colCropRotParts = New Collection
    Set dData = New Scripting.Dictionary
    Set dDataValues = New Scripting.Dictionary
End Sub

Public Function Initialize(arg_oCropsConfig As CropsConfig, arg_cReference As Range, arg_cCompostReference As Range, arg_sName As String)
    ' calculate data
    dDataValues("Fläche") = arg_oCropsConfig.DataValue("Fläche [ha]")
    dData("Fruchtfolge") = arg_sName
    dData("Fläche") = dDataValues("Fläche") & " ha"
    iCompostCount = 0
    
    Dim i As Integer, j As Integer
    i = 1
    Dim colCropRotPart As CropRotationPart
    While arg_cReference.Offset(0, i).Value <> ""
        If arg_cReference.Offset(0, i).Value = arg_sName Then
            Set colCropRotPart = New CropRotationPart
            colCropRotPart.Initialize arg_oCropsConfig, arg_cReference.Offset(1, i), arg_cCompostReference, Me
            colCropRotParts.add colCropRotPart
            j = 2
            While arg_cReference.Offset(j, i).Value <> ""
                If InStr(1, arg_cReference.Offset(j, i).Value, "Kompost") = 1 Then
                    iCompostCount = iCompostCount + 1
                End If
                j = j + 1
            Wend
        End If
        i = i + 1
    Wend
    Set colCropRotPart = Nothing
    Dim colCropRotPartDuration As Collection
    Set colCropRotPartDuration = New Collection
    For i = 1 To colCropRotParts.Count
        colCropRotParts(i).LateInitialize
        colCropRotPartDuration.add colCropRotParts(i).DataValue("Dauer")
    Next i
    For i = 1 To colCropRotParts.Count
        colCropRotParts(i).CompostInitialize
    Next i
    For i = 1 To colCropRotParts.Count
        colCropRotParts(i).LateLateInitialize
    Next i
    dDataValues("Dauer") = Tools.LeastCommonMultiple(colCropRotPartDuration)
    dData("Dauer") = dDataValues("Dauer") & " Jahre"
    
    dDataValues("Deckungsbeitrag inkl. Leistungen") = 0#
    For i = 1 To colCropRotParts.Count
        dDataValues("Deckungsbeitrag inkl. Leistungen") = dDataValues("Deckungsbeitrag inkl. Leistungen") + _
            colCropRotParts(i).DataValue("Deckungsbeitrag inkl. Leistungen") * colCropRotParts(i).DataValue("Fläche")
    Next i
    dDataValues("Deckungsbeitrag inkl. Leistungen") = dDataValues("Deckungsbeitrag inkl. Leistungen") / dDataValues("Fläche")
    dData("Deckungsbeitrag inkl. Leistungen") = Round(dDataValues("Deckungsbeitrag inkl. Leistungen"), 1) & " €/ha" & vbCrLf & _
        Round(dDataValues("Deckungsbeitrag inkl. Leistungen") * dDataValues("Fläche"), 1) & " €"
    
    dDataValues("Arbeitszeit") = 0#
    For i = 1 To colCropRotParts.Count
        dDataValues("Arbeitszeit") = dDataValues("Arbeitszeit") + _
            colCropRotParts(i).DataValue("Arbeitszeit") * colCropRotParts(i).DataValue("Fläche")
    Next i
    dDataValues("Arbeitszeit") = dDataValues("Arbeitszeit") / dDataValues("Fläche")
    dData("Arbeitszeit") = Round(dDataValues("Arbeitszeit"), 1) & " AKh/ha" & vbCrLf & _
        Round(dDataValues("Arbeitszeit") * dDataValues("Fläche"), 1) & " AKh"

    dDataValues("Stundenlohn") = dDataValues("Deckungsbeitrag inkl. Leistungen") / dDataValues("Arbeitszeit")
    dData("Stundenlohn") = Round(dDataValues("Stundenlohn"), 1) & " €/AKh"
    
    dDataValues("Wasserbedarf") = 0#
    For i = 1 To colCropRotParts.Count
        dDataValues("Wasserbedarf") = dDataValues("Wasserbedarf") + _
            colCropRotParts(i).DataValue("Wasserbedarf") * colCropRotParts(i).DataValue("Flächenanteil")
    Next i
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

' return additional offset (i.e. one line printed -> return 0, two lines printed -> return 1)
Public Function PrintLine(arg_aRowData As Variant, arg_cReference As Range) As Integer
    PrintLine = 0
    Dim i As Integer, iColumnOffset As Integer
    iColumnOffset = 0
    
    arg_cReference.Value = arg_aRowData(0)
    arg_cReference.HorizontalAlignment = xlRight
    arg_cReference.Font.Italic = True
    arg_cReference.Offset(0, 1).Value = arg_aRowData(1)
    arg_cReference.Offset(0, 1).HorizontalAlignment = xlRight
    
    For Each oCropRotPart In colCropRotParts
        PrintLine = oCropRotPart.PrintLine(arg_aRowData, arg_cReference.Offset(0, 2 + iColumnOffset), Not dData.Exists(arg_aRowData(1)), (colCropRotParts.Count <> 1))
        iColumnOffset = iColumnOffset + oCropRotPart.NumberOfEntries()
    Next oCropRotPart
    If dData.Exists(arg_aRowData(1)) Then
        With arg_cReference.Offset(PrintLine, 2)
            .NumberFormat = "@"
            .WrapText = True
            .Value = dData(arg_aRowData(1))
            If InStr(.Value, vbCrLf) Then
                .Characters(InStr(.Value, vbCrLf)).Font.Color = RGB(170, 170, 170)
            End If
            Range(.Address, arg_cReference.Offset(PrintLine, 2 + NumberOfEntries - 1).Address).Merge
            .Interior.Color = RGB(213, 232, 202)
        End With
        For i = 0 To (NumberOfEntries - 1)
            arg_cReference.Offset(PrintLine, 2 + i).Borders(xlEdgeBottom).Weight = xlThin
            arg_cReference.Offset(PrintLine, 2 + i).Borders(xlEdgeBottom).Color = RGB(128, 128, 128)
        Next i
        PrintLine = PrintLine + 1
    End If
    For i = 0 To PrintLine - 1
        arg_cReference.Offset(i, 2).Borders(xlEdgeLeft).Weight = xlThick
        arg_cReference.Offset(i, 2).Borders(xlEdgeLeft).Color = RGB(128, 128, 128)
        arg_cReference.Offset(i, 2 + NumberOfEntries - 1).Borders(xlEdgeRight).Weight = xlThick
        arg_cReference.Offset(i, 2 + NumberOfEntries - 1).Borders(xlEdgeRight).Color = RGB(128, 128, 128)
    Next i
    If arg_aRowData(0) <> "" Then
        For i = 0 To (2 + NumberOfEntries - 1)
            arg_cReference.Offset(0, i).Borders(xlEdgeTop).Weight = xlMedium
            arg_cReference.Offset(0, i).Borders(xlEdgeTop).Color = RGB(128, 128, 128)
        Next i
    End If
End Function

Public Function CheckForInvalidData() As String
    Dim i As Integer
    i = 1
    While i <= colCropRotParts.Count And CheckForInvalidData = ""
        CheckForInvalidData = colCropRotParts(i).CheckForInvalidData
        i = i + 1
    Wend
    If CheckForInvalidData = "" Then
        Dim fAreaPartsSum As Double
        For i = 1 To colCropRotParts.Count
            fAreaPartsSum = fAreaPartsSum + colCropRotParts(i).DataValue("Flächenanteil")
        Next i
        If fAreaPartsSum <> 1 Then
            CheckForInvalidData = "Flächenanteile müssen in Summe 1 ergeben"
        End If
    End If
End Function

Public Property Get NumberOfEntries() As Integer
    NumberOfEntries = 0
    For i = 1 To colCropRotParts.Count
        NumberOfEntries = NumberOfEntries + colCropRotParts(i).NumberOfEntries
    Next i
End Property

Public Function Data(arg_sDataName As String) As Variant
    If dData.Exists(arg_sDataName) Then
        Data = dData(arg_sDataName)
    Else
        Data = Nothing
    End If
End Function

Public Function DataValue(arg_sDataName As String) As String
    If dDataValues.Exists(arg_sDataName) Then
        DataValue = dDataValues(arg_sDataName)
    Else
        DataValue = ""
    End If
End Function

Public Function AccDataValue(arg_sDataName As String, arg_sCropOrFieldWorkFilter As String, arg_bIsNan As Boolean) As Variant
    If dDataValues.Exists(arg_sDataName) And arg_sCropOrFieldWorkFilter = "" Then
        AccDataValue = dDataValues(arg_sDataName)
    Else
        For Each oCropRotPart In colCropRotParts
            AccDataValue = AccDataValue + oCropRotPart.AccDataValue(arg_sDataName, arg_sCropOrFieldWorkFilter, arg_bIsNan)
        Next oCropRotPart
    End If
End Function

Private Function NutrientHelper(arg_sNutrientName As String, arg_sNutrientWeightUnit As String)
    dDataValues(arg_sNutrientName) = 0#
    For i = 1 To colCropRotParts.Count
        dDataValues(arg_sNutrientName) = dDataValues(arg_sNutrientName) + _
            colCropRotParts(i).DataValue(arg_sNutrientName) * colCropRotParts(i).DataValue("Fläche")
    Next i
    dDataValues(arg_sNutrientName) = dDataValues(arg_sNutrientName) / dDataValues("Fläche")
    dData(arg_sNutrientName) = Round(dDataValues(arg_sNutrientName), 1) & Chr(160) & arg_sNutrientWeightUnit & "/ha" & vbCrLf & _
        Round(dDataValues(arg_sNutrientName) * dDataValues("Fläche"), 1) & Chr(160) & arg_sNutrientWeightUnit
End Function

Public Property Get CompostCount() As Integer
    CompostCount = iCompostCount
End Property
