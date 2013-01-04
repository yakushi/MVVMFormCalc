Namespace MVVMCalc.Common

    ''' <summary>
    ''' MessengerのRaisedイベントを受信するトリガー
    ''' </summary>
    Public Class MessageTrigger
        'Inherits EventTrigger

        Protected Function GetEventName() As String
            Return "Raised"
        End Function
    End Class
End Namespace
