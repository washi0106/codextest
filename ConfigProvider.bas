Option Explicit

Private gConfig As CIniConfig
Private gLoaded As Boolean

Public Sub InitializeConfig(ByVal iniPath As String)
    Set gConfig = New CIniConfig
    gConfig.LoadFromFile iniPath
    gLoaded = True
End Sub

Public Function ConfigString(ByVal sectionName As String, ByVal key As String, _
                             Optional ByVal defaultValue As String = "") As String
    EnsureLoaded
    ConfigString = gConfig.GetString(sectionName, key, defaultValue)
End Function

Private Sub EnsureLoaded()
    If Not gLoaded Or gConfig Is Nothing Then
        Err.Raise vbObjectError + 1001, "ConfigProvider", _
                  "設定が初期化されていません。InitializeConfig を先に呼んでください。"
    End If
End Sub
