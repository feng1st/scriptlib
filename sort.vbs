Option Explicit

Function SortStrObjDict(strObjDict, isAsc)
    Dim keys(), values()
    Dim key, value
    Dim i, j

    Dim count : count = strObjDict.Count
    If count > 1 Then
        ReDim keys(count)
        ReDim values(count)

        i = 0
        For Each key In strObjDict
            keys(i) = key
            Set values(i) = strObjDict(key)
            i = i + 1
        Next

        For i = 0 To (count - 2)
            For j = i To (count - 1)
                Dim comp : comp = StrComp(keys(i), keys(j), vbTextCompare)
                If (isAsc And comp > 0) Or (Not isAsc And comp < 0) Then
                    key = keys(i) : keys(i) = keys(j) : keys(j) = key
                    Set value = values(i) : Set values(i) = values(j) : Set values(j) = value
                End If
            Next
        Next

        strObjDict.RemoveAll

        For i = 0 To (count - 1)
            strObjDict.Add keys(i), values(i)
        Next
    End If
End Function
