Option Explicit

' ===== 設定 =====
' CATIA VBA から「VBAプロジェクトファイル自身の配置フォルダ」を安定して取得する
' 標準APIは実行形態依存で不安定なため、INIフォルダを明示指定する方式にする。
Private Const BASE_DIR As String = "C:\workspace\TB\MNLRec"
Private Const INI_FILE_NAME As String = "project.ini"

Public Sub CATMain()
    
    ' iniファイル読み込み
    Dim iniPath As String: iniPath = BASE_DIR & "\" & INI_FILE_NAME
    MConfigProvider.InitializeConfig iniPath
    
    ' メイン画面表示
    UFMain.Show
    
End Sub
