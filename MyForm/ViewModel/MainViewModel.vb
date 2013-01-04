Namespace MVVMCalc.ViewModel

    ''' <summary>
    ''' MainViewのViewModel
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

            ' プロパティを初期化して妥当性検証を行う
            Me.InitializeProperties()
        End Sub

        ''' <summary>
        ''' 計算方式
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
        ''' 現在選択されている計算方式
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
        ''' 現在選択されている計算方式(値)
        ''' </summary>
        ''' <remarks>
        ''' 値にしかバインドできないため値型のプロパティを追加
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
        ''' 計算の左辺値
        ''' </summary>
        Public Property Lhs() As String
            Get
                Return Me._Lhs
            End Get
            Set(ByVal Value As String)
                Me._Lhs = Value
                If Not Me.IsDouble(Value) Then
                    Me.SetError("Lhs", "数字を入力してください")
                Else
                    Me.ClearError("Lhs")
                End If
                Me.RaisePropertyChanged("Lhs")
                Me._CalculateCommand.RaiseCanExecuteChanged(Me, EventArgs.Empty)
            End Set
        End Property

        ''' <summary>
        ''' 計算の右辺値
        ''' </summary>
        Public Property Rhs() As String
            Get
                Return Me._Rhs
            End Get
            Set(ByVal Value As String)
                Me._Rhs = Value
                If Not Me.IsDouble(Value) Then
                    Me.SetError("Rhs", "数字を入力してください")
                Else
                    Me.ClearError("Rhs")
                End If
                Me.RaisePropertyChanged("Rhs")
                Me._CalculateCommand.RaiseCanExecuteChanged(Me, EventArgs.Empty)
            End Set
        End Property

        ''' <summary>
        ''' 計算結果
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
        ''' 計算処理のコマンド
        ''' </summary>
        Public ReadOnly Property CalculateCommand() As Common.DelegateCommand
            Get
                ' コンストラクタで初期化しておく
                Return Me._CalculateCommand
            End Get
        End Property

        ''' <summary>
        ''' 計算結果にエラーがあったことを通知するメッセージを送信するメッセンジャーを取得する。
        ''' </summary>
        Public ReadOnly Property ErrorMessenger() As Common.Messenger
            Get
                Return Me._ErrorMessenger
            End Get
        End Property

        ''' <summary>
        ''' 計算処理のコマンドの実行を行います。
        ''' </summary>
        Private Sub CalculateExecute()
            ' 現在の入力値をもとに計算を行う
            Dim calc As Model.Calculator = New Model.Calculator
            Me.Answer = calc.Execute( _
                Double.Parse(Me.Lhs), _
                Double.Parse(Me.Rhs), _
                Me.SelectedCalculateType.CalculateType)

            If IsInvalidAnswer() Then
                ' 計算結果が実数の範囲から外れている場合はViewに通知する
                Me.ErrorMessenger.Raise( _
                    New Common.Message("計算結果が実数の範囲を超えました。入力値を初期化しますか?"), _
                    AddressOf CalculateExecuteConfirmActionCallback)
            End If
        End Sub

        ''' <summary>
        ''' 計算処理結果が範囲外だった場合のViewからのコールバック
        ''' </summary>
        Private Sub CalculateExecuteConfirmActionCallback(ByVal m As Common.Message)
            ' Viewから入力を初期化すると指定された場合はプロパティの初期化を行う
            If Not CType(m.Response, Boolean) Then
                Return
            End If

            InitializeProperties()
        End Sub

        ''' <summary>
        ''' 計算処理が実行可能かどうかの判定を行います。
        ''' </summary>
        Private Function CanCalculateExecute() As Boolean
            ' 現在選択されている計算方法がNone以外かつ入力にエラーが無ければコマンドの実行が可能
            Return (Me.SelectedCalculateType.CalculateType <> Model.CalclateType.None) And (Not Me.HasError)
        End Function

        ''' Answerが無効な値か確認する。
        Private Function IsInvalidAnswer() As Boolean
            Return Double.IsInfinity(Me.Answer) Or Double.IsNaN(Me.Answer)
        End Function

        ''' プロパティの初期化を行う
        Private Sub InitializeProperties()
            Me.Lhs = String.Empty
            Me.Rhs = String.Empty
            Me.Answer = 0.0
            Me.SelectedCalculateType = CType(Me.CalculateTypes, ArrayList).Item(0)
        End Sub

        ''' valueがDouble型に変換できるか検証します。
        Private Function IsDouble(ByVal value As String) As Boolean
            Dim temp As Double
            Return Double.TryParse(value, Globalization.NumberStyles.Float, Globalization.NumberFormatInfo.CurrentInfo, temp)
        End Function

#Region "WinForm用"

        ' フォームから自動バインドする変更通知イベント
        ' プロパティ名: Xxxx  →  イベント名: XxxxChanged
        Public Event LhsChanged As EventHandler
        Public Event RhsChanged As EventHandler
        Public Event AnswerChanged As EventHandler
        Public Event SelectedCalculateTypeValueChanged As EventHandler

#End Region

    End Class
End Namespace
