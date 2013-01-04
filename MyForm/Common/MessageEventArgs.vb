Namespace MVVMCalc.Common
    ''' <summary>
    ''' Messenger�̒ʒm�C�x���g�p�̃C�x���g����
    ''' </summary>
    Public Class MessageEventArgs
        Inherits EventArgs

        ''' <summary>
        ''' ���M���郁�b�Z�[�W
        ''' </summary>
        Private _message As Message
        Public ReadOnly Property Message() As Message
            Get
                Return _message
            End Get
        End Property

        ''' <summary>
        ''' ViewModel�̃R�[���o�b�N
        ''' </summary>
        Private _callback As Action_Message
        Public ReadOnly Property Callback() As Action_Message
            Get
                Return _callback
            End Get
        End Property

        ''' <summary>
        ''' ���b�Z�[�W�ƃR�[���o�b�N���w�肵�ăC�x���g�������쐬����
        ''' </summary>
        ''' <param name="message">���b�Z�[�W</param>
        ''' <param name="callback">�R�[���o�b�N</param>
        Public Sub New(ByVal message As Message, ByVal callback As Action_Message)
            Me._message = message
            Me._callback = callback
        End Sub
    End Class
End Namespace
