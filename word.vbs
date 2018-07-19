Option Explicit

Const wdFindStop = 0
Const wdFindContinue = 1
Const wdFindAsk = 2

Const wdReplaceNone = 0
Const wdReplaceOne = 1
Const wdReplaceAll = 2

Class WordDoc
    Private obj

    Public Sub Attach(doc)
        Set obj = doc
    End Sub

    Public Sub Save()
        obj.Save
    End Sub

    Public Sub SaveAs(filename)
        obj.SaveAs filename
    End Sub

    Public Sub Close()
        obj.Close
    End Sub

    Sub TypeText(text)
        obj.Activate
        obj.Application.Selection.TypeText text
    End Sub

    Function Find(findText, matchCase, matchWholeWord, matchWildcards, _
                  forward, wrap, replaceWith, replace)
        Const matchSoundsLike = False
        Const matchAllWordForms = False
        Const format = False
        obj.Activate
        With obj.Application.Selection.Find
            .ClearAllFuzzyOptions
            .ClearFormatting
            .Replacement.ClearFormatting
            Find = .Execute(findText, _
                    matchCase, matchWholeWord, matchWildcards, _
                    matchSoundsLike, matchAllWordForms, _
                    forward, wrap, format, replaceWith, replace)
        End With
    End Function

    Sub ReplaceAll(findText, replaceWith, _
                   matchCase, matchWholeWord, matchWildcards)
        Const forward = True
        Find findText, matchCase, matchWholeWord, matchWildcards, _
                forward, wdFindContinue, replaceWith, wdReplaceAll
    End Sub
End Class

Class WordApp
    Private obj

    Private Sub Class_Initialize()
        Set obj = CreateObject("Word.Application")
    End Sub

    Private Sub Class_Terminate()
        obj.Quit
    End Sub

    Public Function Create()
        Set Create = New WordDoc
        Create.Attach obj.Documents.Add()
    End Function

    Public Function Open(filename)
        Set Open = New WordDoc
        Open.Attach obj.Documents.Open(filename)
    End Function
End Class
