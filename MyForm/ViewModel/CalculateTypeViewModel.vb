Namespace MVVMCalc.ViewModel

    Public Class CalculateTypeViewModel

        ''' <summary>
        ''' �v�Z���@
        ''' </summary>
        Dim _caclulateType As Model.CalclateType
        Public Property CalculateType() As Model.CalclateType
            Get
                Return _caclulateType
            End Get
            Set(ByVal Value As Model.CalclateType)
                _caclulateType = Value
            End Set
        End Property

        ''' <summary>
        ''' �\����
        ''' </summary>
        Dim _label As String
        Public Property Label() As String
            Get
                Return _label
            End Get
            Set(ByVal Value As String)
                _label = Value
            End Set
        End Property

        ''' <summary>
        ''' �v�Z���@�ƕ\������ݒ肵�ăC���X�^���X�𐶐����܂��B
        ''' </summary>
        ''' <param name="_calculateType">�v�Z���@</param>
        ''' <param name="_label">�\����</param>
        Public Sub New(ByVal calculateType As Model.CalclateType, ByVal label As String)
            Me.CalculateType = calculateType
            Me.Label = label
        End Sub

        ''' <summary>
        ''' CalculateType�̒l�ƕ\���p���x���̃}�b�v
        ''' </summary>
        Private Class TypeLabelDictionary
            Inherits System.Collections.DictionaryBase
            Default Public Property Item(ByVal key As Model.CalclateType) As String
                Get
                    Return CStr(Dictionary(key))
                End Get
                Set(ByVal Value As String)
                    Dictionary(key) = Value
                End Set
            End Property
            Public Sub Add(ByVal key As Model.CalclateType, ByVal value As String)
                Dictionary.Add(key, value)
            End Sub
        End Class
        Private Shared ReadOnly _typeLabelMap As TypeLabelDictionary = New TypeLabelDictionary
        Private Shared Property TypeLabelMap(ByVal key As Model.CalclateType)
            Get
                If _typeLabelMap.Count = 0 Then
                    _typeLabelMap.Add(Model.CalclateType.None, "���I��")
                    _typeLabelMap.Add(Model.CalclateType.Add, "�����Z")
                    _typeLabelMap.Add(Model.CalclateType.Subs, "�����Z")
                    _typeLabelMap.Add(Model.CalclateType.Mul, "�|���Z")
                    _typeLabelMap.Add(Model.CalclateType.Div, "����Z")
                End If
                Return _typeLabelMap(key)
            End Get
            Set(ByVal Value)
                _typeLabelMap.Item(key) = Value
            End Set
        End Property

        ''' <summary>
        ''' �v�Z���@���烉�x�����������I�Ɉ������Ă�CalculateTypeViewModel�̃C���X�^���X�𐶐�����B
        ''' </summary>
        ''' <param name="type">�v�Z���@</param>
        ''' <returns>�������ꂽ�C���X�^���X</returns>
        Public Shared Function Create(ByVal type As Model.CalclateType) As CalculateTypeViewModel
            Return New CalculateTypeViewModel(type, TypeLabelMap(type))
        End Function

        ''' <summary>
        ''' CalculateType�̑S�l�ɑΉ�����CalculateTypeViewModel���쐬����B
        ''' </summary>
        ''' <returns>�������ꂽ�C���X�^���X</returns>
        ''' <remarks>
        ''' yield ���g���Ȃ����߁A�R���N�V�����S�̂�Ԃ��B
        ''' </remarks>
        Public Shared Function Create() As IEnumerable
            Dim viewModelList As ArrayList = New ArrayList
            For Each e As Model.CalclateType In System.Enum.GetValues(GetType(Model.CalclateType))
                viewModelList.Add(Create(e))
            Next
            Return viewModelList
        End Function

    End Class
End Namespace
