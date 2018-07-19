Option Explicit

Class ExcelBook
    Private obj

    Public Sub Attach(book)
        Set obj = book
    End Sub

    Public Sub Save()
        obj.Save
    End Sub

    Public Sub SaveAs(filename)
        Dim displayAlerts : displayAlerts = obj.Application.DisplayAlerts
        obj.Application.DisplayAlerts = False
        obj.SaveAs filename
        obj.Application.DisplayAlerts = displayAlerts
    End Sub

    Public Sub Close()
        obj.Close
    End Sub

    Public Function GetLastRow(sheetNameOrIndex)
        Dim usedRange : Set usedRange = obj.Worksheets(sheetNameOrIndex).UsedRange
        GetLastRow = usedRange.Row + usedRange.Rows.Count - 1
    End Function

    Public Function GetLastColumn(sheetNameOrIndex)
        Dim usedRange : Set usedRange = obj.Worksheets(sheetNameOrIndex).UsedRange
        GetLastColumn = usedRange.Column + usedRange.Columns.Count - 1
    End Function

    Public Function Read(sheetNameOrIndex, row, col)
        Read = obj.Worksheets(sheetNameOrIndex).Cells(row, col).Value
    End Function

    Public Sub Write(sheetNameOrIndex, row, col, value)
        obj.Worksheets(sheetNameOrIndex).Cells(row, col).Value = value
    End Sub

    Public Function FindRow(sheetNameOrIndex, col, rowFrom, rowTo, text)
        text = Trim(CStr(text))
        Dim sheet : Set sheet = obj.Worksheets(sheetNameOrIndex)
        If rowTo <= 0 Then
            rowTo = GetLastRow(sheetNameOrIndex)
        End If
        FindRow = 0
        Dim row
        For row = rowFrom To rowTo
            If Trim(CStr(sheet.Cells(row, col).Value)) = text Then
                FindRow = row
                Exit For
            End If
        Next
    End Function

    Public Function FindColumn(sheetNameOrIndex, row, colFrom, colTo, text)
        text = Trim(CStr(text))
        Dim sheet : Set sheet = obj.Worksheets(sheetNameOrIndex)
        If colTo <= 0 Then
            colTo = GetLastColumn(sheetNameOrIndex)
        End If
        FindColumn = 0
        Dim col
        For col = colFrom To colTo
            If Trim(CStr(sheet.Cells(row, col).Value)) = text Then
                FindColumn = col
                Exit For
            End If
        Next
    End Function
End Class

Class ExcelApp
    Private obj

    Private Sub Class_Initialize()
        Set obj = CreateObject("Excel.Application")
    End Sub

    Private Sub Class_Terminate()
        obj.Workbooks.Close
        obj.Quit
    End Sub

    Public Function Create()
        Set Create = New ExcelBook
        Create.Attach obj.Workbooks.Add()
    End Function

    Public Function Open(filename)
        Set Open = New ExcelBook
        Open.Attach obj.Workbooks.Open(filename)
    End Function
End Class
