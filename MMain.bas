Option Explicit

' ===== 設定 =====
' CATIA VBA から「VBAプロジェクトファイル自身の配置フォルダ」を安定して取得する
' 標準APIは実行形態依存で不安定なため、INIフォルダを明示指定する方式にする。
Private Const INI_FOLDER_PATH As String = "C:\\CATIA\\MacroConfig"
Private Const INI_FILE_NAME As String = "project.ini"

Public Sub CATMain()
    Dim iniPath As String

    iniPath = GetIniPath(INI_FILE_NAME)
    InitializeConfig iniPath

    Dim baseDir As String
    baseDir = ConfigString("Path", "BaseDir", "C:\\Temp")

    MsgBox "INI: " & iniPath & vbCrLf & "BaseDir: " & baseDir
End Sub

Public Function GetIniPath(ByVal iniFileName As String) As String
    Dim basePath As String

    basePath = Trim$(INI_FOLDER_PATH)
    If Len(basePath) = 0 Then
        Err.Raise vbObjectError + 1100, "MMain.GetIniPath", _
                  "INI_FOLDER_PATH が未設定です。INI配置フォルダを設定してください。"
    End If

    If Right$(basePath, 1) <> "\\" Then
        basePath = basePath & "\\"
    End If

    GetIniPath = basePath & iniFileName
End Function
