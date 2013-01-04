Namespace MVVMCalc.Common

    ''' <summary>
    ''' �f���Q�[�g���󂯎��ICommand�̎���
    ''' </summary>
    Public Class DelegateCommand
        'Implements ICommand

        Private _execute As Action

        Private _canExecute As Predicate

        ''' <summary>
        ''' �R�}���h��Execute���\�b�h�Ŏ��s���鏈�����w�肵��DelegateCommand�̃C���X�^���X��
        ''' �쐬���܂��B
        ''' </summary>
        ''' <param name="execute">Execute���\�b�h�Ŏ��s���鏈��</param>
        Public Sub New(ByVal execute As Action)
            Me.New(execute, AddressOf Func_True)
        End Sub

        ''' <summary>
        ''' �R�}���h��Execute���\�b�h�Ŏ��s���鏈����CanExecute���\�b�h�Ŏ��s���鏈�����w�肵��
        ''' DelegateCommand�̃C���X�^���X���쐬���܂��B
        ''' </summary>
        ''' <param name="execute">Execute���\�b�h�Ŏ��s���鏈��</param>
        ''' <param name="canExecute">CanExecute���\�b�h�Ŏ��s���鏈��</param>
        Public Sub New(ByVal execute As Action, ByVal canExecute As Predicate)
            If execute Is Nothing Then
                Throw New ArgumentNullException("execute")
            End If
            If canExecute Is Nothing Then
                Throw New ArgumentNullException("canExecute")
            End If

            Me._execute = execute
            Me._canExecute = canExecute
        End Sub

        ''' <summary>
        ''' �R�}���h�����s���܂��B
        ''' </summary>
        Public Sub Execute()
            Me._execute()
        End Sub

        ''' <summary>
        ''' �R�}���h�����s�\�ȏ�Ԃ��ǂ����₢���킹�܂��B
        ''' </summary>
        ''' <remarks>
        ''' �v���p�e�B XX �ɕύX���������ꍇ�̃C�x���g�� XXChanged �ł��邱�Ƃ�
        ''' �t�H�[���ɂ����� Binding �ł͊��҂���邽�߁A������ CanExecute ��
        ''' �v���p�e�B�Ŏ��������B
        ''' </remarks>
        Public ReadOnly Property CanExecute() As Boolean
            Get
                Return Me._canExecute()
            End Get
        End Property

        ''' <summary>
        ''' CanExecute�̌��ʂɕύX�����������Ƃ�ʒm����C�x���g�B
        ''' </summary>
        ''' <remarks>
        ''' �t�H�[���ŏ���� Binding ���Ă���邱�Ƃ����҂��ĉ������Ă��܂���B
        ''' </remarks>
        Public Event CanExecuteChanged As EventHandler

#Region "�C�x���g�n���h��"

        ''' <summary>
        ''' DelegateCommand ���ĂԂ��߂� EventHandler ��Ԃ��B
        ''' </summary>
        ''' <returns>EventHandler �̃C���X�^���X</returns>
        Public Function CreateEventHandler() As EventHandler
            Return AddressOf ExecuteEventHandler
        End Function

        ''' <summary>
        ''' �t�H�[����̃R���g���[������C�x���g�o�R�� DelegateCommand ���ĂԂ��߂̃O���\
        ''' </summary>
        ''' <param name="sender">�C�x���g�̃\�[�X</param>
        ''' <param name="e">�C�x���g �f�[�^���i�[���Ă��� EventArgs</param>
        Private Sub ExecuteEventHandler(ByVal sender As System.Object, ByVal e As System.EventArgs)
            Me.Execute()
        End Sub

#End Region

        ''' <summary>
        ''' CanExecuteChanged�C�x���g���O������N������C���^�t�F�[�X
        ''' </summary>
        Public Overridable Sub RaiseCanExecuteChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
            RaiseEvent CanExecuteChanged(sender, e)
        End Sub

        ''' <summary>
        ''' () => True �̎���
        ''' </summary>
        Private Shared Function Func_True() As Boolean
            Return True
        End Function

    End Class
End Namespace
