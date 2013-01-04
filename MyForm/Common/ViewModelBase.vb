Imports System.ComponentModel
Imports System.Reflection

Namespace MVVMCalc.Common

    ''' <summary>
    ''' ViewModel�̊�{�N���X�BINotifyPropertyChanged�̎�����񋟂��܂��B
    ''' </summary>
    Public Class ViewModelBase
        Implements IDataErrorInfo
        'Implements INotifyPropertyChanged, IDataErrorInfo

        ''' <summary>
        ''' �v���p�e�B�ɕR�Â����G���[���b�Z�[�W���i�[���܂��B
        ''' </summary>
        Private _Errors As Hashtable = New Hashtable

        ''' <summary>
        ''' �v���p�e�B�̕ύX�����������ɔ��s����܂��B
        ''' </summary>
        Public Event PropertyChanged As PropertyChangedEventHandler

        ''' <summary>
        ''' PropertyChanged�C�x���g�𔭍s���܂��B
        ''' </summary>
        ''' <param name="propertyName">�v���p�e�B��</param>
        Protected Overridable Sub RaisePropertyChanged(ByVal propertyName As String)
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(propertyName))

            ' XxxxChanged �`���̃C�x���g���Ăяo��
            RaiseChanged(propertyName)
        End Sub

        ''' <summary>
        ''' ���g�p
        ''' </summary>
        ReadOnly Property [Error]() As String Implements IDataErrorInfo.Error
            Get
                Throw New NotImplementedException
            End Get
        End Property

        ''' <summary>
        ''' columnName�Ŏw�肵���v���p�e�B�̃G���[��Ԃ��܂��B
        ''' </summary>
        ''' <param name="columnName">�v���p�e�B��</param>
        ''' <returns>�G���[���b�Z�[�W</returns>
        Default Public ReadOnly Property Errors(ByVal columnName As String) As String Implements IDataErrorInfo.Item
            Get
                If Me._Errors.ContainsKey(columnName) Then
                    Return Me._Errors(columnName)
                End If
                Return Nothing
            End Get
        End Property

        ''' <summary>
        ''' �v���p�e�B�ɃG���[���b�Z�[�W��ݒ肵�܂��B
        ''' </summary>
        ''' <param name="propertyName">�v���p�e�B��</param>
        ''' <param name="errorMessage">�G���[���b�Z�[�W</param>
        Protected Sub SetError(ByVal propertyName As String, ByVal errorMessage As String)
            Me._Errors(propertyName) = errorMessage
            Me.RaisePropertyChanged("HasError")
        End Sub

        ''' <summary>
        ''' �v���p�e�B�̃G���[���N���A���܂��B
        ''' </summary>
        ''' <param name="propertyName">�v���p�e�B��</param>
        Protected Sub ClearError(ByVal propertyName As String)
            If Me._Errors.ContainsKey(propertyName) Then
                Me._Errors.Remove(propertyName)
                Me.RaisePropertyChanged("HasError")
            End If
        End Sub

        ''' <summary>
        ''' �S�ẴG���[���N���A���܂��B
        ''' </summary>
        Protected Sub ClearErrors()
            Me._Errors.Clear()
            Me.RaisePropertyChanged("HasError")
        End Sub

        ''' <summary>
        ''' �G���[�̗L�����擾���܂��B
        ''' </summary>
        Public ReadOnly Property HasError() As Boolean
            Get
                Return Me._Errors.Count <> 0
            End Get
        End Property

#Region "WinForm�p"

        ''' <summary>
        ''' �G���[�����������ɔ��s����܂��B
        ''' </summary>
        Public Event HasErrorChanged As EventHandler

        ''' <summary>
        ''' XxxxChanged �`���̃C�x���g�𔭍s���܂��B
        ''' </summary>
        Private Sub RaiseChanged(ByVal propertyName As String)
            Dim this As Type = Me.GetType()
            Dim changedEvent As FieldInfo = Nothing

            While (Not this Is Nothing) And (changedEvent Is Nothing)
                changedEvent = this.GetField(propertyName & "ChangedEvent", _
                    BindingFlags.Instance Or BindingFlags.NonPublic Or BindingFlags.IgnoreCase _
                Or BindingFlags.Public Or BindingFlags.Static)
                this = this.BaseType
            End While
            If changedEvent Is Nothing Then Return

            Dim changedDelegate As MulticastDelegate = changedEvent.GetValue(Me)
            If changedDelegate Is Nothing Then Return
            changedDelegate.GetType().GetMethod("Invoke").Invoke(changedDelegate, New Object() {Me, EventArgs.Empty})
        End Sub

#End Region

    End Class

End Namespace