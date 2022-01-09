Private dData As Scripting.Dictionary
Private dDataValues As Scripting.Dictionary
Private oCropRotationPart As CropRotationPart
Private oNextInRot As Variant
Private oPrevInRot As Variant

Private Sub Class_Initialize()
    Set dData = New Scripting.Dictionary
End Sub

Public Function Initialize(arg_oCropsConfig As CropsConfig, arg_sName As String, arg_oCropRotationPart As CropRotationPart)
    Set oCropRotationPart = arg_oCropRotationPart
    
    ' get data
    Set dDataValues = arg_oCropsConfig.CropConfigValues(arg_sName)
    
    ' calculate data
    If dDataValues("Endzeit") = dDataValues("Startzeit") Then
        dDataValues("Standzeit") = 12#
    Else
        dDataValues("Standzeit") = CDbl((HalfMonthToInt(dDataValues("Endzeit")) - HalfMonthToInt(dDataValues("Startzeit")) + 24) Mod 24) / 2
    End If
    dDataValues("Deckungsbeitrag") = dDataValues("Ertrag bzw. Aufwand") * dDataValues("Verkaufs- bzw. Einkaufspreis") - dDataValues("Summe variable Kosten")
    dDataValues("Deckungsbeitrag inkl. Leistungen") = dDataValues("Deckungsbeitrag") + dDataValues("Sonstige Leistungen/Prämien")
    dDataValues("Stundenlohn") = dDataValues("Deckungsbeitrag inkl. Leistungen") / dDataValues("Arbeitszeit")
    dData("Frucht bzw. Feldarbeit") = dDataValues("Frucht bzw. Feldarbeit")
    dData("Startzeit") = dDataValues("Startzeit")
    dData("Endzeit") = dDataValues("Endzeit")
    dData("Standzeit") = dDataValues("Standzeit") & " Mon"
End Function

Public Function LateInitialize()
    Dim oCropOrFieldWork As Variant
    Set oCropOrFieldWork = oNextInRot
    While oCropOrFieldWork.IsFieldWork
        Set oCropOrFieldWork = oCropOrFieldWork.NextInRot
    Wend
    dDataValues("Brache danach") = CDbl((HalfMonthToInt(oCropOrFieldWork.DataValue("Startzeit")) - HalfMonthToInt(dDataValues("Endzeit")) + 24) Mod 24) / 2
    dData("Brache danach") = dDataValues("Brache danach") & " Mon"
End Function

Public Function LateLateInitialize()
    dDataValues("Fläche") = oCropRotationPart.DataValue("Fläche") / oCropRotationPart.DataValue("Dauer")
    dData("Fläche") = Round(dDataValues("Fläche"), 1) & " ha"
    dData("Ertrag bzw. Aufwand") = Round(dDataValues("Ertrag bzw. Aufwand"), 1) & " dt/ha" & vbCrLf & _
        Round(dDataValues("Ertrag bzw. Aufwand") * dDataValues("Fläche"), 1) & " dt"
    dData("Verkaufs- bzw. Einkaufspreis") = dDataValues("Verkaufs- bzw. Einkaufspreis") & " €/dt" & vbCrLf & _
        Round(dDataValues("Verkaufs- bzw. Einkaufspreis") * dDataValues("Ertrag bzw. Aufwand"), 1) & " €/ha" & vbCrLf & _
        Round(dDataValues("Verkaufs- bzw. Einkaufspreis") * dDataValues("Ertrag bzw. Aufwand") * dDataValues("Fläche"), 1) & " €"
    dData("Summe variable Kosten") = dDataValues("Summe variable Kosten") & " €/ha" & vbCrLf & _
        Round(dDataValues("Summe variable Kosten") * dDataValues("Fläche"), 1) & " €"
    dData("Deckungsbeitrag") = dDataValues("Deckungsbeitrag") & " €/ha" & vbCrLf & _
        Round(dDataValues("Deckungsbeitrag") * dDataValues("Fläche"), 1) & " €"
    dData("Sonstige Leistungen/Prämien") = dDataValues("Sonstige Leistungen/Prämien") & " €/ha" & vbCrLf & _
        Round(dDataValues("Sonstige Leistungen/Prämien") * dDataValues("Fläche"), 1) & " €"
    dData("Deckungsbeitrag inkl. Leistungen") = dDataValues("Deckungsbeitrag inkl. Leistungen") & " €/ha" & vbCrLf & _
        Round(dDataValues("Deckungsbeitrag inkl. Leistungen") * dDataValues("Fläche"), 1) & " €"
    dData("Arbeitszeit") = dDataValues("Arbeitszeit") & " AKh/ha" & vbCrLf & _
        Round(dDataValues("Arbeitszeit") * dDataValues("Fläche"), 1) & " AKh"
    dData("Stundenlohn") = Round(dDataValues("Stundenlohn"), 1) & " €/AKh"
    dData("Wasserbedarf") = Round(dDataValues("Wasserbedarf"), 0) & " mm/m²"
    dData("Stickstoff") = Round(dDataValues("Stickstoff"), 1) & " kg/dt" & vbCrLf & _
        Round(dDataValues("Stickstoff") * dDataValues("Ertrag bzw. Aufwand"), 1) & " kg/ha" & vbCrLf & _
        Round(dDataValues("Stickstoff") * dDataValues("Ertrag bzw. Aufwand") * dDataValues("Fläche"), 1) & " kg"
    dData("Phosphor") = Round(dDataValues("Phosphor"), 1) & " kg/dt" & vbCrLf & _
        Round(dDataValues("Phosphor") * dDataValues("Ertrag bzw. Aufwand"), 1) & " kg/ha" & vbCrLf & _
        Round(dDataValues("Phosphor") * dDataValues("Ertrag bzw. Aufwand") * dDataValues("Fläche"), 1) & " kg"
    dData("Kalium") = Round(dDataValues("Kalium"), 1) & " kg/dt" & vbCrLf & _
        Round(dDataValues("Kalium") * dDataValues("Ertrag bzw. Aufwand"), 1) & " kg/ha" & vbCrLf & _
        Round(dDataValues("Kalium") * dDataValues("Ertrag bzw. Aufwand") * dDataValues("Fläche"), 1) & " kg"
    dData("Schwefel") = Round(dDataValues("Schwefel"), 1) & " kg/dt" & vbCrLf & _
        Round(dDataValues("Schwefel") * dDataValues("Ertrag bzw. Aufwand"), 1) & " kg/ha" & vbCrLf & _
        Round(dDataValues("Schwefel") * dDataValues("Ertrag bzw. Aufwand") * dDataValues("Fläche"), 1) & " kg"
    dData("Calcium") = Round(dDataValues("Calcium"), 1) & " kg/dt" & vbCrLf & _
        Round(dDataValues("Calcium") * dDataValues("Ertrag bzw. Aufwand"), 1) & " kg/ha" & vbCrLf & _
        Round(dDataValues("Calcium") * dDataValues("Ertrag bzw. Aufwand") * dDataValues("Fläche"), 1) & " kg"
    dData("Magnesium") = Round(dDataValues("Magnesium"), 1) & " kg/dt" & vbCrLf & _
        Round(dDataValues("Magnesium") * dDataValues("Ertrag bzw. Aufwand"), 1) & " kg/ha" & vbCrLf & _
        Round(dDataValues("Magnesium") * dDataValues("Ertrag bzw. Aufwand") * dDataValues("Fläche"), 1) & " kg"
    dData("Bor") = Round(dDataValues("Bor"), 1) & " g/dt" & vbCrLf & _
        Round(dDataValues("Bor") * dDataValues("Ertrag bzw. Aufwand"), 1) & " g/ha" & vbCrLf & _
        Round(dDataValues("Bor") * dDataValues("Ertrag bzw. Aufwand") * dDataValues("Fläche"), 1) & " g"
    dData("Kupfer") = Round(dDataValues("Kupfer"), 1) & " g/dt" & vbCrLf & _
        Round(dDataValues("Kupfer") * dDataValues("Ertrag bzw. Aufwand"), 1) & " g/ha" & vbCrLf & _
        Round(dDataValues("Kupfer") * dDataValues("Ertrag bzw. Aufwand") * dDataValues("Fläche"), 1) & " g"
    dData("Mangan") = Round(dDataValues("Mangan"), 1) & " g/dt" & vbCrLf & _
        Round(dDataValues("Mangan") * dDataValues("Ertrag bzw. Aufwand"), 1) & " g/ha" & vbCrLf & _
        Round(dDataValues("Mangan") * dDataValues("Ertrag bzw. Aufwand") * dDataValues("Fläche"), 1) & " g"
    dData("Zink") = Round(dDataValues("Zink"), 1) & " g/dt" & vbCrLf & _
        Round(dDataValues("Zink") * dDataValues("Ertrag bzw. Aufwand"), 1) & " g/ha" & vbCrLf & _
        Round(dDataValues("Zink") * dDataValues("Ertrag bzw. Aufwand") * dDataValues("Fläche"), 1) & " g"
End Function

Public Function PrintLine(arg_aRowData As Variant, arg_cReference As Range, arg_bEssential As Boolean) As Integer
    PrintLine = 0
    If dData.Exists(arg_aRowData(1)) Then
        arg_cReference.NumberFormat = "@"
        arg_cReference.WrapText = True
        Dim sData As String
        sData = Chr(160) & Chr(160) & Chr(160) & dData(arg_aRowData(1)) & Chr(160) & Chr(160) & Chr(160)
        sData = Replace(sData, vbCrLf, Chr(160) & Chr(160) & Chr(160) & vbCrLf & Chr(160) & Chr(160) & Chr(160))
        arg_cReference.Value = sData
        If InStr(arg_cReference.Value, vbCrLf) Then
            arg_cReference.Characters(InStr(arg_cReference.Value, vbCrLf)).Font.Color = RGB(170, 170, 170)
        End If
        PrintLine = 1
    ElseIf arg_bEssential Then
        arg_cReference.Value = "not set"
        arg_cReference.Font.Italic = True
        arg_cReference.Font.Color = RGB(170, 170, 170)
        PrintLine = 1
    End If
    arg_cReference.Interior.Color = RGB(242, 248, 238)
    arg_cReference.Borders(xlEdgeRight).Weight = xlThin
    arg_cReference.Borders(xlEdgeRight).Color = RGB(191, 191, 191)
    arg_cReference.Borders(xlEdgeLeft).Weight = xlThin
    arg_cReference.Borders(xlEdgeLeft).Color = RGB(191, 191, 191)
End Function

Public Function CheckForInvalidData() As String
    CheckForInvalidData = ""
End Function

Public Function DataValue(arg_sDataName As String) As String
    If dDataValues.Exists(arg_sDataName) Then
        DataValue = dDataValues(arg_sDataName)
    Else
        DataValue = ""
    End If
End Function

Public Property Get IsCrop() As Boolean
    IsCrop = True
End Property

Public Property Get IsFieldWork() As Boolean
    IsFieldWork = False
End Property

Public Property Set NextInRot(arg_oNextInRot As Variant)
    Set oNextInRot = arg_oNextInRot
End Property

Public Property Set PrevInRot(arg_oPrevInRot As Variant)
    Set oPrevInRot = arg_oPrevInRot
End Property

Public Property Get NextInRot() As Variant
    Set NextInRot = oNextInRot
End Property

Public Property Get PrevInRot() As Variant
    Set PrevInRot = oPrevInRot
End Property
