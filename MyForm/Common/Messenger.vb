Namespace MVVMCalc.Common

    ''' <summary>
    ''' Message�𑗐M����N���X
    ''' </summary>
    Public Class Messenger
        ''' <summary>
        ''' ���b�Z�[�W�����M���ꂽ���Ƃ�ʒm����C�x���g
        ''' </summary>
        Public Event Raised As EventHandler

        ''' <summary>
        ''' �w�肵�����b�Z�[�W�ƃR�[���o�b�N�Ń��b�Z�[�W�𑗐M����
        ''' </summary>
        ''' <param name="message">���b�Z�[�W</param>
        ''' <param name="callback">�R�[���o�b�N</param>
        Public Sub Raise(ByVal message As Message, ByVal callback As Action_Message)
            RaiseEvent Raised(Me, New MessageEventArgs(message, callback))
        End Sub
    End Class
End Namespace
