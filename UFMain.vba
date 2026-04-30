Option Explicit

' 作業開始ボタン押下時
Private Sub CmdStartWork_Click()
    ' 自身を閉じる
    Unload Me
    ' 作業開始画面を表示
    UFStartWork.Show
End Sub

’ 履歴確認ボタン押下時
Private Sub CmdCheckHistory_Click()
    ' 自身を閉じる
    Unload Me
    ' 履歴確認画面を表示
    UFCheckHistory.Show
End Sub

