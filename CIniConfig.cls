Option Explicit

Private mMap As Object ' Scripting.Dictionary

Private Sub Class_Initialize()
    Set mMap = CreateObject("Scripting.Dictionary")
    mMap.CompareMode = 1 ' vbTextCompare
End Sub

Public Sub LoadFromFile(ByVal iniPath As String)
    Dim f As Integer
    Dim lineText As String
    Dim sectionName As String
    Dim pos As Long
    Dim key As String, value As String
    Dim fullKey As String

    If Len(Dir$(iniPath)) = 0 Then
        Err.Raise vbObjectError + 1000, "CIniConfig.LoadFromFile", _
                  "INIファイルが見つかりません: " & iniPath
    End If

    f = FreeFile
    Open iniPath For Input As #f

    sectionName = ""

    Do While Not EOF(f)
        Line Input #f, lineText
        lineText = Trim$(lineText)

        If Len(lineText) = 0 Then GoTo ContinueLoop
        If Left$(lineText, 1) = ";" Or Left$(lineText, 1) = "#" Then GoTo ContinueLoop

        If Left$(lineText, 1) = "[" And Right$(lineText, 1) = "]" Then
            sectionName = Trim$(Mid$(lineText, 2, Len(lineText) - 2))
            GoTo ContinueLoop
        End If

        pos = InStr(1, lineText, "=", 1)
        If pos > 0 Then
            key = Trim$(Left$(lineText, pos - 1))
            value = Trim$(Mid$(lineText, pos + 1))
            fullKey = BuildKey(sectionName, key)
            mMap(fullKey) = value
        End If

ContinueLoop:
    Loop

    Close #f
End Sub

Public Function GetString(ByVal sectionName As String, ByVal key As String, _
                          Optional ByVal defaultValue As String = "") As String
    Dim fullKey As String
    fullKey = BuildKey(sectionName, key)

    If mMap.Exists(fullKey) Then
        GetString = CStr(mMap(fullKey))
    Else
        GetString = defaultValue
    End If
End Function

Private Function BuildKey(ByVal sectionName As String, ByVal key As String) As String
    BuildKey = Trim$(sectionName) & "." & Trim$(key)
End Function
