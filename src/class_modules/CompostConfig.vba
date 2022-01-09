Private cCompositionReference As Range
Private cCompostReference As Range
Private colCompostParts As Collection
Private fTotalWeightCalculated As Double
Private bTotalWeightCalculated As Boolean
Private dNameReplacements As Scripting.Dictionary

Public Function Initialize(arg_sCompostConfigSheet As String, arg_sFarmConfigSheet As String)
    Set dNameReplacements = New Scripting.Dictionary
    dNameReplacements("Stickstoffbilanz") = "Stickstoff"
    dNameReplacements("Phosphorbilanz") = "Phosphor"
    dNameReplacements("Kaliumbilanz") = "Kalium"
    dNameReplacements("Schwefelbilanz") = "Schwefel"
    dNameReplacements("Calciumbilanz") = "Calcium"
    dNameReplacements("Magnesiumbilanz") = "Magnesium"
    dNameReplacements("Borbilanz") = "Bor"
    dNameReplacements("Kupferbilanz") = "Kupfer"
    dNameReplacements("Manganbilanz") = "Mangan"
    dNameReplacements("Zinkbilanz") = "Zink"
    Set colCompostParts = New Collection
    Set cCompositionReference = Sheets(arg_sCompostConfigSheet).Range("2:2").Find("Zusammensetzung")
    Set cCompostReference = Sheets(arg_sFarmConfigSheet).Range("2:2").Find("Kompostierung")
    bTotalWeightCalculated = False
    Dim i As Integer, j As Integer, oCompostPart As CompostPart
    i = 1
    j = 1
    While cCompositionReference.Offset(0, i).Value <> ""
        Set oCompostPart = New CompostPart
        colCompostParts.add oCompostPart
        Dim bFound As Boolean
        bFound = False
        j = 1
        While cCompostReference.Offset(0, j).Value <> "" And Not bFound
            If cCompostReference.Offset(0, j).Value = cCompositionReference.Offset(0, i).Value Then
                bFound = True
            Else
                j = j + 1
            End If
        Wend
        
        If bFound Then
            colCompostParts(colCompostParts.Count).Initialize Me, cCompositionReference.Offset(0, i), cCompostReference.Offset(0, j)
        Else
            colCompostParts(colCompostParts.Count).Initialize Me, cCompositionReference.Offset(0, i), Nothing
        End If
        i = i + 1
    Wend
End Function

Public Function CompostPartConfigValues(arg_cReference As Range) As Scripting.Dictionary
    Dim dCompostPartConfigValues As Scripting.Dictionary
    Set dCompostPartConfigValues = New Scripting.Dictionary
    Dim i As Integer, j As Integer, bFound As Boolean
    Dim sEntryName As String
    For Each cReference In Array(cCompositionReference, cCompostReference)
        i = 1
        j = 1
        bFound = False
        With cReference
            If arg_cReference.Worksheet.Name = .Worksheet.Name Then
                While .Offset(0, i).Value <> "" And Not bFound
                    If .Offset(0, i).Value = arg_cReference.Value Then
                        While .Offset(j, 0).Value <> ""
                            sEntryName = .Offset(j, 0).Value
                            If InStr(sEntryName, " [") > 0 Then
                                sEntryName = Left(sEntryName, InStr(sEntryName, " [") - 1)
                            End If
                            If dNameReplacements.Exists(sEntryName) Then
                                dCompostPartConfigValues(dNameReplacements(sEntryName)) = .Offset(j, i).Value
                            Else
                                dCompostPartConfigValues(sEntryName) = .Offset(j, i).Value
                            End If
                            j = j + 1
                        Wend
                        bFound = True
                    End If
                    i = i + 1
                Wend
            End If
        End With
    Next cReference
    Set CompostPartConfigValues = dCompostPartConfigValues
End Function

Public Property Get TotalWeight() As Double
    If Not bTotalWeightCalculated Then
        Dim fTotalWeight As Double
        fTotalWeight = 0#
        For Each oCompostPart In colCompostParts
            fTotalWeight = fTotalWeight + oCompostPart.DataValue("Dichte") * oCompostPart.DataValue("Volumen")
        Next oCompostPart
        fTotalWeight = fTotalWeight / 100
        fTotalWeightCalculated = fTotalWeight
        bTotalWeightCalculated = True
    End If
    TotalWeight = fTotalWeightCalculated
End Property

Public Function InterpolatedDataValue(arg_sDataName As String) As Double
    Dim fInterDataValue As Double
    fInterDataValue = 0#
    For Each oCompostPart In colCompostParts
        fInterDataValue = fInterDataValue + oCompostPart.DataValue(arg_sDataName) * oCompostPart.DataValue("Dichte") * oCompostPart.DataValue("Volumen")
    Next oCompostPart
    fInterDataValue = fInterDataValue / 100
    fInterDataValue = fInterDataValue / TotalWeight
    InterpolatedDataValue = fInterDataValue
End Function
