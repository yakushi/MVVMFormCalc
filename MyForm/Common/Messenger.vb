Namespace MVVMCalc.Common

    ''' <summary>
    ''' Messageを送信するクラス
    ''' </summary>
    Public Class Messenger
        ''' <summary>
        ''' メッセージが送信されたことを通知するイベント
        ''' </summary>
        Public Event Raised As EventHandler

        ''' <summary>
        ''' 指定したメッセージとコールバックでメッセージを送信する
        ''' </summary>
        ''' <param name="message">メッセージ</param>
        ''' <param name="callback">コールバック</param>
        Public Sub Raise(ByVal message As Message, ByVal callback As Action_Message)
            RaiseEvent Raised(Me, New MessageEventArgs(message, callback))
        End Sub
    End Class
End Namespace
