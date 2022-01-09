Private oCropRotation As CropRotation
Private aTableEntries As Variant
Private dDataValues As Scripting.Dictionary
Private cGeneralReference As Range
Private cCropReference As Range
Private cFieldWorkReference As Range
Private dNameReplacements As Scripting.Dictionary
Private oCompostConfig As CompostConfig

Private Sub Class_Initialize()
    aTableEntries = Array( _
        Array("Fruchtfolgedaten", "Fruchtfolge"), _
        Array("", "Frucht bzw. Feldarbeit"), _
        Array("", "Dauer"), _
        Array("", "Fläche"), _
        Array("Fruchtfolge Zeiten", "Startzeit"), _
        Array("", "Endzeit"), _
        Array("", "Standzeit"), _
        Array("", "Brache danach"), _
        Array("Erträge und Leistungen", "Ertrag bzw. Aufwand"), _
        Array("", "Verkaufs- bzw. Einkaufspreis"), _
        Array("Deckungsbeitrag", "Summe variable Kosten"), _
        Array("", "Deckungsbeitrag"), _
        Array("", "Sonstige Leistungen/Prämien"), _
        Array("", "Deckungsbeitrag inkl. Leistungen"), _
        Array("Arbeit", "Arbeitszeit"), _
        Array("", "Stundenlohn"), _
        Array("Bilanz und Bedarf", "Wasserbedarf"), _
        Array("", "Stickstoff"), Array("", "Phosphor"), Array("", "Kalium"), _
        Array("", "Schwefel"), Array("", "Calcium"), Array("", "Magnesium"), _
        Array("", "Bor"), Array("", "Kupfer"), Array("", "Mangan"), Array("", "Zink"))
    Set dDataValues = New Scripting.Dictionary
    Set dNameReplacements = New Scripting.Dictionary
    dNameReplacements("Anwendungszeit") = "Startzeit"
    dNameReplacements("Ertrag") = "Ertrag bzw. Aufwand"
    dNameReplacements("Aufwand") = "Ertrag bzw. Aufwand"
    dNameReplacements("Verkaufspreis") = "Verkaufs- bzw. Einkaufspreis"
    dNameReplacements("Einkaufspreis") = "Verkaufs- bzw. Einkaufspreis"
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
End Sub

Public Function Initialize(arg_sCropConfigSheet As String, arg_sFarmConfigSheet As String, arg_oCompostConfig As CompostConfig)
    Set cGeneralReference = Sheets(arg_sFarmConfigSheet).Range("2:2").Find("Allgemeines")
    Set cCropReference = Sheets(arg_sCropConfigSheet).Range("2:2").Find("Früchte")
    Set cFieldWorkReference = Sheets(arg_sCropConfigSheet).Range("2:2").Find("Feldarbeit")
    Set oCompostConfig = arg_oCompostConfig
    Dim i As Integer
    i = 1
    While cGeneralReference.Offset(i, 0).Value <> ""
        dDataValues(cGeneralReference.Offset(i, 0).Value) = cGeneralReference.Offset(i, 1).Value
        i = i + 1
    Wend

    Set oCropRotation = New CropRotation
    oCropRotation.Initialize Me, Sheets(arg_sFarmConfigSheet).Range("2:2").Find("FF Möglichkeiten"), _
        Sheets(arg_sFarmConfigSheet).Range("2:2").Find("Kompostierung"), dDataValues("Fruchtfolge")
End Function

Public Function PrintTable(arg_cReference As Range)
    Dim i As Integer
    Dim iRowOffset As Integer
    iRowOffset = 0
    For i = 0 To Tools.ArrayLength(aTableEntries) - 1
        Dim iPrintedRows As Integer
        iPrintedRows = oCropRotation.PrintLine(aTableEntries(i), arg_cReference.Offset(iRowOffset, 0))
        iRowOffset = iRowOffset + iPrintedRows
    Next i
End Function

Public Function CheckForInvalidData() As String
    CheckForInvalidData = oCropRotation.CheckForInvalidData
End Function

Public Function DataValue(arg_sDataName As String) As Variant
    If dDataValues.Exists(arg_sDataName) Then
        DataValue = dDataValues(arg_sDataName)
    Else
        DataValue = ""
    End If
End Function

Public Property Get CropReference() As Range
    Set CropReference = cCropReference
End Property

Public Function CropConfigValues(arg_sCropName As String) As Scripting.Dictionary
    Dim dCropConfigValues As Scripting.Dictionary
    Set dCropConfigValues = New Scripting.Dictionary
    dCropConfigValues("Frucht bzw. Feldarbeit") = arg_sCropName
    Dim i As Integer, j As Integer, bFound As Boolean
    i = 1
    j = 1
    bFound = False
    Dim sEntryName As String
    With cCropReference
        While .Offset(0, i).Value <> "" And Not bFound
            If .Offset(0, i).Value = arg_sCropName Then
                While .Offset(j, 0).Value <> ""
                    sEntryName = .Offset(j, 0).Value
                    If InStr(sEntryName, " [") > 0 Then
                        sEntryName = Left(sEntryName, InStr(sEntryName, " [") - 1)
                    End If
                    If dNameReplacements.Exists(sEntryName) Then
                        dCropConfigValues(dNameReplacements(sEntryName)) = .Offset(j, i).Value
                    Else
                        dCropConfigValues(sEntryName) = .Offset(j, i).Value
                    End If
                    j = j + 1
                Wend
                bFound = True
            End If
            i = i + 1
        Wend
    End With
    Set CropConfigValues = dCropConfigValues
End Function

Public Function FieldWorkConfigValues(arg_sFieldWorkName As String) As Scripting.Dictionary
    Dim dFieldWorkConfigValues As Scripting.Dictionary
    Set dFieldWorkConfigValues = New Scripting.Dictionary
    If InStr(1, arg_sFieldWorkName, "Kompost") = 1 Then
        dFieldWorkConfigValues("Frucht bzw. Feldarbeit") = "Kompost"
    Else
        dFieldWorkConfigValues("Frucht bzw. Feldarbeit") = arg_sFieldWorkName
    End If
    Dim i As Integer, j As Integer, bFound As Boolean
    i = 1
    j = 1
    bFound = False
    Dim sEntryName As String
    With cFieldWorkReference
        While .Offset(0, i).Value <> "" And Not bFound
            If .Offset(0, i).Value = arg_sFieldWorkName Then
                While .Offset(j, 0).Value <> ""
                    sEntryName = .Offset(j, 0).Value
                    If InStr(sEntryName, " [") > 0 Then
                        sEntryName = Left(sEntryName, InStr(sEntryName, " [") - 1)
                    End If
                    If dNameReplacements.Exists(sEntryName) Then
                        dFieldWorkConfigValues(dNameReplacements(sEntryName)) = .Offset(j, i).Value
                    Else
                        dFieldWorkConfigValues(sEntryName) = .Offset(j, i).Value
                    End If
                    j = j + 1
                Wend
                bFound = True
            End If
            i = i + 1
        Wend
    End With
    Set FieldWorkConfigValues = dFieldWorkConfigValues
End Function

Public Property Get CompostConfig() As CompostConfig
    Set CompostConfig = oCompostConfig
End Property
