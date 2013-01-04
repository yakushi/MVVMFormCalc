Namespace MVVMCalc.Common
    ''' <summary>
    ''' �m�F�_�C�A���O��\������A�N�V����
    ''' </summary>
    Public Class ConfirmAction
        'Inherits TriggerAction<DependencyObject>

        Protected Sub Invoke(ByVal parameter As Object)
            Dim args As MessageEventArgs

            ' MessageEventArgs�ȊO�̏ꍇ�͉������Ȃ�
            Try
                args = DirectCast(parameter, MessageEventArgs)
            Catch ex As InvalidCastException
                Return
            End Try

            ' ���b�Z�[�W�{�b�N�X��\������
            Dim result As DialogResult = MessageBox.Show( _
                args.Message.Body.ToString(), _
                "�m�F", _
                MessageBoxButtons.OKCancel)

            ' �{�^���̉����ꂽ���ʂ�Response�Ɋi�[����
            args.Message.Response = (result = DialogResult.OK)
            ' �R�[���o�b�N���Ă� (��x�ϐ��Ɋi�[���Ȃ��ƌĂяo���Ȃ�)
            Dim callback As Action_Message = args.Callback
            callback(args.Message)
        End Sub

#Region "�C�x���g�n���h��"

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
