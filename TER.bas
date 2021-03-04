Attribute VB_Name = "Módulo4"
Function TER(ByVal string1 As String, ByVal string2 As String) As Long

Dim i As Long, j As Long, string1_length As Long, string2_length As Long
Dim distance(0 To 60, 0 To 50) As Long, smStr1(1 To 60) As Long, smStr2(1 To 50) As Long
Dim min1 As Long, min2 As Long, min3 As Long, minmin As Long, MaxL As Long

string1_length = Len(string1): string2_length = Len(string2)

distance(0, 0) = 0
For i = 1 To string1_length: distance(i, 0) = i: smStr1(i) = Asc(LCase(Mid$(string1, i, 1))): Next
For j = 1 To string2_length: distance(0, j) = j: smStr2(j) = Asc(LCase(Mid$(string2, j, 1))): Next
For i = 1 To string1_length
    For j = 1 To string2_length
        If smStr1(i) = smStr2(j) Then
            distance(i, j) = distance(i - 1, j - 1)
        Else
            min1 = distance(i - 1, j) + 1
            min2 = distance(i, j - 1) + 1
            min3 = distance(i - 1, j - 1) + 1
            If min2 < min1 Then
                If min2 < min3 Then minmin = min2 Else minmin = min3
            Else
                If min1 < min3 Then minmin = min1 Else minmin = min3
            End If
            distance(i, j) = minmin
        End If
    Next
Next

MaxL = string1_length: If string2_length > MaxL Then MaxL = string2_length
TER = CLng((distance(string1_length, string2_length) * 100) / MaxL)

End Function

