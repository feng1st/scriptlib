Option Explicit

Const READYSTATE_UNINITIALIZED = 0
Const READYSTATE_LOADING = 1
Const READYSTATE_LOADED = 2
Const READYSTATE_INTERACTIVE = 3
Const READYSTATE_COMPLETE = 4

Class IEApp
    Private obj

    Public Function Find(regex)
        Dim re, shell, windows, i, window
        Set re = New RegExp
        re.Pattern = regex
        re.IgnoreCase = True
        Set shell = CreateObject("Shell.Application")
        Set windows = shell.Windows
        Find = False
        If Not windows Is Nothing Then
            For i = windows.Count - 1 To 0 Step -1 ' reverse order
                Set window = windows(i)
                If Not window Is Nothing Then
                    If InStr(1, window.FullName, "iexplore.exe", vbTextCompare) > 0 Then
                        If re.Test(window.LocationURL) Or _
                                re.Test(window.LocationName) Then
                            Set obj = window
                            Find = True
                            Exit For
                        End If
                    End If
                End If
            Next
        End If
    End Function

    Public Function WaitPage(regex, timeout)
        Dim start
        start = Timer()
        While Timer() - start < timeout
            If Find(regex) Then
                WaitPage = True
                Exit Function
            End If
        Wend
        WaitPage = False
    End Function

    Public Sub OpenNew(url)
        Set obj = CreateObject("InternetExplorer.Application")
        Open url
    End Sub

    Public Sub Open(url)
        On Error Resume Next
        obj.Visible = True
        obj.Navigate url
        While obj.Busy Or obj.ReadyState <> READYSTATE_COMPLETE
            WScript.Sleep 100
        Wend
        On Error GoTo 0
    End Sub

    Public Sub Wait()
        On Error Resume Next
        While obj.Busy Or obj.ReadyState <> READYSTATE_COMPLETE
            WScript.Sleep 100
        Wend
        On Error GoTo 0
    End Sub

    Public Sub Quit()
        On Error Resume Next
        obj.Stop
        obj.Quit
        On Error GoTo 0
    End Sub

    Public Function Exist(path)
        Exist = False
        On Error Resume Next
        Exist = Not GetElement(path) Is Nothing
        On Error GoTo 0
    End Function

    Public Function WaitElement(path, timeout)
        Dim start
        start = Timer()
        While Timer() - start < timeout
            If Exist(path) Then
                WaitElement = True
                Exit Function
            End If
        Wend
        WaitElement = False
    End Function

    Public Function GetOuterHTML(path)
        On Error Resume Next
        GetOuterHTML = GetElement(path).outerHTML
        On Error GoTo 0
    End Function

    Public Function GetInnerHTML(path)
        On Error Resume Next
        GetInnerHTML = GetElement(path).innerHTML
        On Error GoTo 0
    End Function

    Public Function GetValue(path)
        On Error Resume Next
        GetValue = GetElement(path).value
        On Error GoTo 0
    End Function

    Public Sub SetValue(path, value)
        On Error Resume Next
        GetElement(path).value = value
        On Error GoTo 0
    End Sub

    Public Sub Click(path)
        On Error Resume Next
        GetElement(path).click
        On Error GoTo 0
    End Sub

    Public Sub Check(path, checked)
        On Error Resume Next
        GetElement(path).checked = checked
        On Error GoTo 0
    End Sub

    Public Sub SelectValue(path, value)
        Dim parent, child, i
        On Error Resume Next
        Set parent = GetElement(path)
        For i = 0 To parent.options.length - 1
            Set child = parent.options(i)
            If child.value = value Then
                parent.selectedIndex = i
                Exit For
            End If
        Next
        On Error GoTo 0
    End Sub

    Public Sub SelectText(path, text)
        Dim parent, child, i
        On Error Resume Next
        Set parent = GetElement(path)
        For i = 0 To parent.options.length - 1
            Set child = parent.options(i)
            If child.text = text Then
                parent.selectedIndex = i
                Exit For
            End If
        Next
        On Error GoTo 0
    End Sub

    Public Function GetOptionValue(path)
        Dim element
        GetOptionValue = ""
        On Error Resume Next
        Set element = GetElement(path)
        GetOptionValue = element.options(element.selectedIndex).value
        On Error GoTo 0
    End Function

    Public Sub SetOptionValue(path, value)
        Dim element
        On Error Resume Next
        Set element = GetElement(path)
        element.options(element.selectedIndex).value = value
        On Error GoTo 0
    End Sub

    Public Function GetOptionText(path)
        Dim element
        GetOptionText = ""
        On Error Resume Next
        Set element = GetElement(path)
        GetOptionText = element.options(element.selectedIndex).text
        On Error GoTo 0
    End Function

    Public Sub SetOptionText(path, text)
        Dim element
        On Error Resume Next
        Set element = GetElement(path)
        element.options(element.selectedIndex).text = text
        On Error GoTo 0
    End Sub

    Public Function RunScript(frameNames, script, returnType)
        Dim document, headElement, scriptElement
        On Error Resume Next
        Set document = GetFrameDocument(frameNames)
        Set headElement = document.getElementsByTagName("HEAD")(0)
        Set scriptElement = document.createElement("SCRIPT")
        scriptElement.type = "Text/JavaScript"
        scriptElement.text = "function wsfTmpFunc() { " & script & " }"
        headElement.appendChild scriptElement
        If returnType = 2 Then
            Set RunScript = document.parentWindow.wsfTmpFunc()
        ElseIf returnType = 1 Then
            RunScript = document.parentWindow.wsfTmpFunc()
        Else
            document.parentWindow.wsfTmpFunc()
        End If
        headElement.removeChild scriptElement
    End Function

    Public Function ListIEs()
        Dim shell, windows, i, window
        Set shell = CreateObject("Shell.Application")
        Set windows = shell.Windows
        ListIEs = ""
        If Not windows Is Nothing Then
            For i = windows.Count - 1 To 0 Step -1 ' reverse order
                Set window = windows(i)
                If Not window Is Nothing Then
                    If InStr(1, window.FullName, "iexplore.exe", vbTextCompare) > 0 Then
                        ListIEs = ListIEs + window.LocationURL + vbCrLf + window.LocationName + vbCrLf + vbCrLf
                    End If
                End If
            Next
        End If
    End Function

    Public Function FindText(text)
        FindText = FindTextIn(obj.Document, text, "")
    End Function

    Private Function GetElement(path)
        Dim attrs, keyValue, key, value
        Dim frameNames, id, name, tagName, index
        Dim document, elements, element, i, match
        Set GetElement = Nothing
        On Error Resume Next
        Set attrs = CreateObject("Scripting.Dictionary")
        For Each keyValue In Split(path, ",")
            key = Left(keyValue, InStr(keyValue, ":") - 1)
            value = Mid(keyValue, InStr(keyValue, ":") + 1)
            If key = "frame" Then
                frameNames = value
            ElseIf key = "id" Then
                id = value
            ElseIf key = "name" Then
                name = value
            ElseIf key = "tagName" Then
                tagName = value
            ElseIf key = "index" Then
                index = CInt(value)
            Else
                attrs.Add key, value
            End If
        Next

        If IsEmpty(frameNames) Then
            Set document = obj.Document
        Else
            Set document = GetFrameDocument(frameNames)
        End If
        If Not document Is Nothing Then
            If Not IsEmpty(id) Then
                Set GetElement = document.getElementById(id)
            ElseIf IsEmpty(name) And IsEmpty(tagName) Then
                Set GetElement = document.body
            Else
                If Not IsEmpty(name) Then
                    Set elements = document.getElementsByName(name)
                Else
                    Set elements = document.getElementsByTagName(tagName)
                End If
                If Not elements Is Nothing Then
                    If IsEmpty(index) Then
                        index = 0
                    End If
                    If attrs.Count = 0 Then
                        Set GetElement = elements(index)
                    Else
                        i = 0
                        For Each element In elements
                            match = True
                            For Each key In attrs
                                If GetProperty(element, key, "") <> attrs(key) Then
                                    match = False
                                    Exit For
                                End If
                            Next
                            If match Then
                                If i = index Then
                                    Set GetElement = element
                                    Exit For
                                End If
                                i = i + 1
                            End If
                        Next
                    End If
                End If
            End If
        End If
        On Error GoTo 0
    End Function

    Private Function GetFrameDocument(names)
        Dim frame, lastFrame, name
        Set GetFrameDocument = Nothing
        On Error Resume Next
        Set frame = Nothing
        Set frame = obj.Document.parentWindow
        If Not frame Is Nothing Then
            For Each name In Split(names, "|")
                If IsNumeric(name) Then
                    name = CInt(name)
                End If
                Set lastFrame = frame
                Set frame = Nothing
                Set frame = lastFrame.frames(name)
                Set frame = lastFrame.window.frames(name)
                If frame Is Nothing Then
                    Exit For
                End If
            Next
        End If
        If Not frame Is Nothing Then
            Set GetFrameDocument = frame.document
            Set GetFrameDocument = frame.frameElement.contentDocument
        End If
        On Error GoTo 0
    End Function

    Private Function FindTextIn(document, text, frameNames)
        Dim paths, element, match, path
        Dim frames, i, frame, frameName
        On Error Resume Next
        FindTextIn = ""
        paths = ""
        For Each element In document.all
            match = False
            match = (element.innerText = text)
            If Not match Then
                match = (GetProperty(element, "value", "") = text)
            End If
            If match Then
                If element.tagName = "OPTION" Then
                    Set element = element.parentElement
                End If
                path = MakePair("frame", frameNames)
                path = Join(path, MakePair("id", element.id), ",")
                If GetProperty(element, "name", "") <> "" Then
                    path = Join(path, MakePair("name", element.name), ",")
                End If
                path = Join(path, MakePair("tagName", element.tagName), ",")
                path = Join(path, MakePair("innerText", element.innerText), ",")
                If GetProperty(element, "value", "") <> "" Then
                    path = Join(path, MakePair("value", element.value), ",")
                End If
                If GetProperty(element, "text", "") <> "" Then
                    path = Join(path, MakePair("text", element.text), ",")
                End If
                paths = paths & path & vbCrLf
            End If
        Next
        Set frames = document.parentWindow.frames
        Set frames = document.parentWindow.window.frames
        For i = 0 To frames.length - 1
            Set frame = frames(i)
            frameName = Empty
            frameName = frame.name
            If IsEmpty(frameName) Or frameName = "" Then
                frameName = CStr(i)
            End If
            paths = paths & FindTextIn(frame.document, text, Join(frameNames, frameName, "|"))
        Next
        FindTextIn = paths
    End Function

    Private Function GetProperty(object, key, defaultValue)
        On Error Resume Next
        If IsObject(defaultValue) Then
            Set GetProperty = defaultValue
            Set GetProperty = Eval("object." & key)
        Else
            GetProperty = defaultValue
            GetProperty = Eval("object." & key)
        End If
    End Function

    Private Function MakePair(key, value)
        If Not IsEmpty(value) And value <> "" Then
            MakePair = key & ":" & value
        Else
            MakePair = ""
        End If
    End Function

    Private Function Join(str1, str2, separator)
        If Not IsEmpty(str1) And str1 <> "" And _
                Not IsEmpty(str2) And str2 <> "" Then
            Join = str1 & separator & str2
        Else
            Join = str1 & str2
        End If
    End Function
End Class
