Namespace MVVMCalc.Common

    ''' <summary>
    ''' デリゲートを受け取るICommandの実装
    ''' </summary>
    Public Class DelegateCommand
        'Implements ICommand

        Private _execute As Action

        Private _canExecute As Predicate

        ''' <summary>
        ''' コマンドのExecuteメソッドで実行する処理を指定してDelegateCommandのインスタンスを
        ''' 作成します。
        ''' </summary>
        ''' <param name="execute">Executeメソッドで実行する処理</param>
        Public Sub New(ByVal execute As Action)
            Me.New(execute, AddressOf Func_True)
        End Sub

        ''' <summary>
        ''' コマンドのExecuteメソッドで実行する処理とCanExecuteメソッドで実行する処理を指定して
        ''' DelegateCommandのインスタンスを作成します。
        ''' </summary>
        ''' <param name="execute">Executeメソッドで実行する処理</param>
        ''' <param name="canExecute">CanExecuteメソッドで実行する処理</param>
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
        ''' コマンドを実行します。
        ''' </summary>
        Public Sub Execute()
            Me._execute()
        End Sub

        ''' <summary>
        ''' コマンドが実行可能な状態かどうか問い合わせます。
        ''' </summary>
        ''' <remarks>
        ''' プロパティ XX に変更があった場合のイベントは XXChanged であることが
        ''' フォームにおける Binding では期待されるため、あえて CanExecute は
        ''' プロパティで実装した。
        ''' </remarks>
        Public ReadOnly Property CanExecute() As Boolean
            Get
                Return Me._canExecute()
            End Get
        End Property

        ''' <summary>
        ''' CanExecuteの結果に変更があったことを通知するイベント。
        ''' </summary>
        ''' <remarks>
        ''' フォームで勝手に Binding してくれることを期待して何もしていません。
        ''' </remarks>
        Public Event CanExecuteChanged As EventHandler

#Region "イベントハンドラ"

        ''' <summary>
        ''' DelegateCommand を呼ぶための EventHandler を返す。
        ''' </summary>
        ''' <returns>EventHandler のインスタンス</returns>
        Public Function CreateEventHandler() As EventHandler
            Return AddressOf ExecuteEventHandler
        End Function

        ''' <summary>
        ''' フォーム上のコントロールからイベント経由で DelegateCommand を呼ぶためのグル―
        ''' </summary>
        ''' <param name="sender">イベントのソース</param>
        ''' <param name="e">イベント データを格納している EventArgs</param>
        Private Sub ExecuteEventHandler(ByVal sender As System.Object, ByVal e As System.EventArgs)
            Me.Execute()
        End Sub

#End Region

        ''' <summary>
        ''' CanExecuteChangedイベントを外部から起動するインタフェース
        ''' </summary>
        Public Overridable Sub RaiseCanExecuteChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
            RaiseEvent CanExecuteChanged(sender, e)
        End Sub

        ''' <summary>
        ''' () => True の実装
        ''' </summary>
        Private Shared Function Func_True() As Boolean
            Return True
        End Function

    End Class
End Namespace
