Option Explicit

' 作業開始画面の初期化
Private Sub UserForm_Initialize()

    ' 作業区分コンボボックスに値を設定
    With CBWorkClassification
        .AddItem "新規登録"
        .ListIndex = 0 ' デフォルトで最初の項目を選択
    End With

End Sub

' 「登録」ボタン押下時の処理
Private Sub CmdRegister()


End Sub
