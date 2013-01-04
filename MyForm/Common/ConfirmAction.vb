Namespace MVVMCalc.Common
    ''' <summary>
    ''' 確認ダイアログを表示するアクション
    ''' </summary>
    Public Class ConfirmAction
        'Inherits TriggerAction<DependencyObject>

        Protected Sub Invoke(ByVal parameter As Object)
            Dim args As MessageEventArgs

            ' MessageEventArgs以外の場合は何もしない
            Try
                args = DirectCast(parameter, MessageEventArgs)
            Catch ex As InvalidCastException
                Return
            End Try

            ' メッセージボックスを表示して
            Dim result As DialogResult = MessageBox.Show( _
                args.Message.Body.ToString(), _
                "確認", _
                MessageBoxButtons.OKCancel)

            ' ボタンの押された結果をResponseに格納して
            args.Message.Response = (result = DialogResult.OK)
            ' コールバックを呼ぶ (一度変数に格納しないと呼び出せない)
            Dim callback As Action_Message = args.Callback
            callback(args.Message)
        End Sub

#Region "イベントハンドラ"

        Public Shared Function CreateEventHandler() As EventHandler
            Dim action As ConfirmAction = New ConfirmAction
            Return AddressOf action.ExecuteEventHandler
        End Function

        Private Sub ExecuteEventHandler(ByVal sender As System.Object, ByVal e As EventArgs)
            Invoke(e)
        End Sub

#End Region

    End Class
End Namespace
