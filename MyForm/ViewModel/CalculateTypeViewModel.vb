Namespace MVVMCalc.ViewModel

    Public Class CalculateTypeViewModel

        ''' <summary>
        ''' 計算方法
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
        ''' 表示名
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
        ''' 計算方法と表示名を設定してインスタンスを生成します。
        ''' </summary>
        ''' <param name="_calculateType">計算方法</param>
        ''' <param name="_label">表示名</param>
        Public Sub New(ByVal calculateType As Model.CalclateType, ByVal label As String)
            Me.CalculateType = calculateType
            Me.Label = label
        End Sub

        ''' <summary>
        ''' CalculateTypeの値と表示用ラベルのマップ
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
                    _typeLabelMap.Add(Model.CalclateType.None, "未選択")
                    _typeLabelMap.Add(Model.CalclateType.Add, "足し算")
                    _typeLabelMap.Add(Model.CalclateType.Subs, "引き算")
                    _typeLabelMap.Add(Model.CalclateType.Mul, "掛け算")
                    _typeLabelMap.Add(Model.CalclateType.Div, "割り算")
                End If
                Return _typeLabelMap(key)
            End Get
            Set(ByVal Value)
                _typeLabelMap.Item(key) = Value
            End Set
        End Property

        ''' <summary>
        ''' 計算方法からラベル名を自動的に引き当ててCalculateTypeViewModelのインスタンスを生成する。
        ''' </summary>
        ''' <param name="type">計算方法</param>
        ''' <returns>生成されたインスタンス</returns>
        Public Shared Function Create(ByVal type As Model.CalclateType) As CalculateTypeViewModel
            Return New CalculateTypeViewModel(type, TypeLabelMap(type))
        End Function

        ''' <summary>
        ''' CalculateTypeの全値に対応するCalculateTypeViewModelを作成する。
        ''' </summary>
        ''' <returns>生成されたインスタンス</returns>
        ''' <remarks>
        ''' yield が使えないため、コレクション全体を返す。
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
