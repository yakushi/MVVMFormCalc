Namespace MVVMCalc.Common

    ''' <summary>
    ''' ViewModelとViewの間での情報をやり取りを行うメッセージ
    ''' </summary>
    Public Class Message
        ''' <summary>
        ''' メッセージの本体
        ''' </summary>
        Private _body As Object
        Public ReadOnly Property Body() As Object
            Get
                Return _body
            End Get
        End Property

        ''' <summary>
        ''' ViewからViewModeへのメッセージのレスポンス
        ''' </summary>
        Private _response As Object
        Public Property Response() As Object
            Get
                Return _response
            End Get
            Set(ByVal Value As Object)
                _response = Value
            End Set
        End Property

        ''' <summary>
        ''' Bodyを指定してMessageを作成する
        ''' </summary>
        ''' <param name="body"></param>
        Public Sub New(ByVal body As Object)
            Me._body = body
        End Sub

    End Class
End Namespace
