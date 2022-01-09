Public Function LeastCommonMultiple2(a As Integer, b As Integer) As Integer
    Dim lcm_a As Integer, lcm_b As Integer
    lcm_a = a
    lcm_b = b
    While lcm_a <> lcm_b
        If lcm_a < lcm_b Then
            lcm_a = lcm_a + a
        Else
            lcm_b = lcm_b + b
        End If
    Wend
    LeastCommonMultiple2 = lcm_a
End Function

Public Function LeastCommonMultiple(col As Collection) As Integer
    Dim lcm As Integer
    lcm = col(1)
    For i = 2 To col.Count
        lcm = LeastCommonMultiple2(lcm, col(i))
    Next i
    LeastCommonMultiple = lcm
End Function

Public Function ArrayLength(a As Variant) As Integer
   If IsEmpty(a) Then
      ArrayLength = 0
   Else
      ArrayLength = UBound(a) - LBound(a) + 1
   End If
End Function

Public Function HalfMonthToInt(str As String) As Integer
    Dim splits As Variant
    splits = Split(str, " ")
    Select Case splits(0)
        Case "Jan"
            HalfMonthToInt = 0
        Case "Feb"
            HalfMonthToInt = 2
        Case "Mar"
            HalfMonthToInt = 4
        Case "Apr"
            HalfMonthToInt = 6
        Case "Mai"
            HalfMonthToInt = 8
        Case "Jun"
            HalfMonthToInt = 10
        Case "Jul"
            HalfMonthToInt = 12
        Case "Aug"
            HalfMonthToInt = 14
        Case "Sep"
            HalfMonthToInt = 16
        Case "Okt"
            HalfMonthToInt = 18
        Case "Nov"
            HalfMonthToInt = 20
        Case "Dez"
            HalfMonthToInt = 22
    End Select
    If splits(1) = "2" Then
        HalfMonthToInt = HalfMonthToInt + 1
    End If
End Function
