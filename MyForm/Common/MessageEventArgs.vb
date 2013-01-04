Namespace MVVMCalc.Common
    ''' <summary>
    ''' Messengerの通知イベント用のイベント引数
    ''' </summary>
    Public Class MessageEventArgs
        Inherits EventArgs

        ''' <summary>
        ''' 送信するメッセージ
        ''' </summary>
        Private _message As Message
        Public ReadOnly Property Message() As Message
            Get
                Return _message
            End Get
        End Property

        ''' <summary>
        ''' ViewModelのコールバック
        ''' </summary>
        Private _callback As Action_Message
        Public ReadOnly Property Callback() As Action_Message
            Get
                Return _callback
            End Get
        End Property

        ''' <summary>
        ''' メッセージとコールバックを指定してイベント引数を作成する
        ''' </summary>
        ''' <param name="message">メッセージ</param>
        ''' <param name="callback">コールバック</param>
        Public Sub New(ByVal message As Message, ByVal callback As Action_Message)
            Me._message = message
            Me._callback = callback
        End Sub
    End Class
End Namespace
