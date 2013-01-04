Namespace MVVMCalc.ViewModel

    ''' <summary>
    ''' MainView��ViewModel
    ''' </summary>
    Public Class MainViewModel
        Inherits Common.ViewModelBase

        Private _Lhs As String

        Private _Rhs As String

        Private _Answer As Double

        Private _CalculateTypes As ArrayList

        Private _SelectedCalculateType As CalculateTypeViewModel

        Private _CalculateCommand As Common.DelegateCommand

        Private _ErrorMessenger As Common.Messenger = New Common.Messenger

        Public Sub New()
            Me.CalculateTypes = ViewModel.CalculateTypeViewModel.Create
            Me._CalculateCommand = New Common.DelegateCommand(AddressOf CalculateExecute, AddressOf CanCalculateExecute)

            ' �v���p�e�B�����������đÓ������؂��s��
            Me.InitializeProperties()
        End Sub

        ''' <summary>
        ''' �v�Z����
        ''' </summary>
        Public Property CalculateTypes() As IEnumerable
            Get
                Return _CalculateTypes
            End Get
            Set(ByVal Value As IEnumerable)
                _CalculateTypes = Value
            End Set
        End Property

        ''' <summary>
        ''' ���ݑI������Ă���v�Z����
        ''' </summary>
        Public Property SelectedCalculateType() As CalculateTypeViewModel
            Get
                Return Me._SelectedCalculateType
            End Get
            Set(ByVal Value As CalculateTypeViewModel)
                Me.SelectedCalculateTypeValue = Value.CalculateType
                Me.RaisePropertyChanged("SelectedCalculateType")
            End Set
        End Property

        ''' <summary>
        ''' ���ݑI������Ă���v�Z����(�l)
        ''' </summary>
        ''' <remarks>
        ''' �l�ɂ����o�C���h�ł��Ȃ����ߒl�^�̃v���p�e�B��ǉ�
        ''' </remarks>
        Public Property SelectedCalculateTypeValue() As Model.CalclateType
            Get
                Return Me._SelectedCalculateType.CalculateType
            End Get
            Set(ByVal Value As Model.CalclateType)
                Me._SelectedCalculateType = Me._CalculateTypes(Value)
                Me.RaisePropertyChanged("SelectedCalculateTypeValue")
                Me._CalculateCommand.RaiseCanExecuteChanged(Me, EventArgs.Empty)
            End Set
        End Property

        ''' <summary>
        ''' �v�Z�̍��Ӓl
        ''' </summary>
        Public Property Lhs() As String
            Get
                Return Me._Lhs
            End Get
            Set(ByVal Value As String)
                Me._Lhs = Value
                If Not Me.IsDouble(Value) Then
                    Me.SetError("Lhs", "��������͂��Ă�������")
                Else
                    Me.ClearError("Lhs")
                End If
                Me.RaisePropertyChanged("Lhs")
                Me._CalculateCommand.RaiseCanExecuteChanged(Me, EventArgs.Empty)
            End Set
        End Property

        ''' <summary>
        ''' �v�Z�̉E�Ӓl
        ''' </summary>
        Public Property Rhs() As String
            Get
                Return Me._Rhs
            End Get
            Set(ByVal Value As String)
                Me._Rhs = Value
                If Not Me.IsDouble(Value) Then
                    Me.SetError("Rhs", "��������͂��Ă�������")
                Else
                    Me.ClearError("Rhs")
                End If
                Me.RaisePropertyChanged("Rhs")
                Me._CalculateCommand.RaiseCanExecuteChanged(Me, EventArgs.Empty)
            End Set
        End Property

        ''' <summary>
        ''' �v�Z����
        ''' </summary>
        Public Property Answer() As Double
            Get
                Return Me._Answer
            End Get
            Set(ByVal Value As Double)
                Me._Answer = Value
                Me.RaisePropertyChanged("Answer")
            End Set
        End Property

        ''' <summary>
        ''' �v�Z�����̃R�}���h
        ''' </summary>
        Public ReadOnly Property CalculateCommand() As Common.DelegateCommand
            Get
                ' �R���X�g���N�^�ŏ��������Ă���
                Return Me._CalculateCommand
            End Get
        End Property

        ''' <summary>
        ''' �v�Z���ʂɃG���[�����������Ƃ�ʒm���郁�b�Z�[�W�𑗐M���郁�b�Z���W���[���擾����B
        ''' </summary>
        Public ReadOnly Property ErrorMessenger() As Common.Messenger
            Get
                Return Me._ErrorMessenger
            End Get
        End Property

        ''' <summary>
        ''' �v�Z�����̃R�}���h�̎��s���s���܂��B
        ''' </summary>
        Private Sub CalculateExecute()
            ' ���݂̓��͒l�����ƂɌv�Z���s��
            Dim calc As Model.Calculator = New Model.Calculator
            Me.Answer = calc.Execute( _
                Double.Parse(Me.Lhs), _
                Double.Parse(Me.Rhs), _
                Me.SelectedCalculateType.CalculateType)

            If IsInvalidAnswer() Then
                ' �v�Z���ʂ������͈̔͂���O��Ă���ꍇ��View�ɒʒm����
                Me.ErrorMessenger.Raise( _
                    New Common.Message("�v�Z���ʂ������͈̔͂𒴂��܂����B���͒l�����������܂���?"), _
                    AddressOf CalculateExecuteConfirmActionCallback)
            End If
        End Sub

        ''' <summary>
        ''' �v�Z�������ʂ��͈͊O�������ꍇ��View����̃R�[���o�b�N
        ''' </summary>
        Private Sub CalculateExecuteConfirmActionCallback(ByVal m As Common.Message)
            ' View������͂�����������Ǝw�肳�ꂽ�ꍇ�̓v���p�e�B�̏��������s��
            If Not CType(m.Response, Boolean) Then
                Return
            End If

            InitializeProperties()
        End Sub

        ''' <summary>
        ''' �v�Z���������s�\���ǂ����̔�����s���܂��B
        ''' </summary>
        Private Function CanCalculateExecute() As Boolean
            ' ���ݑI������Ă���v�Z���@��None�ȊO�����͂ɃG���[��������΃R�}���h�̎��s���\
            Return (Me.SelectedCalculateType.CalculateType <> Model.CalclateType.None) And (Not Me.HasError)
        End Function

        ''' Answer�������Ȓl���m�F����B
        Private Function IsInvalidAnswer() As Boolean
            Return Double.IsInfinity(Me.Answer) Or Double.IsNaN(Me.Answer)
        End Function

        ''' �v���p�e�B�̏��������s��
        Private Sub InitializeProperties()
            Me.Lhs = String.Empty
            Me.Rhs = String.Empty
            Me.Answer = 0.0
            Me.SelectedCalculateType = CType(Me.CalculateTypes, ArrayList).Item(0)
        End Sub

        ''' value��Double�^�ɕϊ��ł��邩���؂��܂��B
        Private Function IsDouble(ByVal value As String) As Boolean
            Dim temp As Double
            Return Double.TryParse(value, Globalization.NumberStyles.Float, Globalization.NumberFormatInfo.CurrentInfo, temp)
        End Function

#Region "WinForm�p"

        ' �t�H�[�����玩���o�C���h����ύX�ʒm�C�x���g
        ' �v���p�e�B��: Xxxx  ��  �C�x���g��: XxxxChanged
        Public Event LhsChanged As EventHandler
        Public Event RhsChanged As EventHandler
        Public Event AnswerChanged As EventHandler
        Public Event SelectedCalculateTypeValueChanged As EventHandler

#End Region

    End Class
End Namespace
