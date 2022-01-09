Dim dDataValues As Scripting.Dictionary
Dim sName As String

Public Function Initialize(arg_oCompostConfig As CompostConfig, arg_cReference1 As Range, arg_cReference2 As Range)
    sName = arg_cReference1.Value
    Dim dDataValuesTmp As Scripting.Dictionary
    Set dDataValues = arg_oCompostConfig.CompostPartConfigValues(arg_cReference1)
    If Not arg_cReference2 Is Nothing Then
        Set dDataValuesTmp = arg_oCompostConfig.CompostPartConfigValues(arg_cReference2)
        For Each sKey In dDataValuesTmp.Keys()
            dDataValues(sKey) = dDataValuesTmp(sKey)
        Next sKey
    End If
End Function

Public Function DataValue(arg_sDataName As String) As Variant
    If dDataValues.Exists(arg_sDataName) Then
        DataValue = dDataValues(arg_sDataName)
    Else
        DataValue = ""
    End If
End Function

Public Property Get Name() As String
    Name = sName
End Property
