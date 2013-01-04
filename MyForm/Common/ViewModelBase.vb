Imports System.ComponentModel
Imports System.Reflection

Namespace MVVMCalc.Common

    ''' <summary>
    ''' ViewModelの基本クラス。INotifyPropertyChangedの実装を提供します。
    ''' </summary>
    Public Class ViewModelBase
        Implements IDataErrorInfo
        'Implements INotifyPropertyChanged, IDataErrorInfo

        ''' <summary>
        ''' プロパティに紐づいたエラーメッセージを格納します。
        ''' </summary>
        Private _Errors As Hashtable = New Hashtable

        ''' <summary>
        ''' プロパティの変更があった時に発行されます。
        ''' </summary>
        Public Event PropertyChanged As PropertyChangedEventHandler

        ''' <summary>
        ''' PropertyChangedイベントを発行します。
        ''' </summary>
        ''' <param name="propertyName">プロパティ名</param>
        Protected Overridable Sub RaisePropertyChanged(ByVal propertyName As String)
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(propertyName))

            ' XxxxChanged 形式のイベントも呼び出す
            RaiseChanged(propertyName)
        End Sub

        ''' <summary>
        ''' 未使用
        ''' </summary>
        ReadOnly Property [Error]() As String Implements IDataErrorInfo.Error
            Get
                Throw New NotImplementedException
            End Get
        End Property

        ''' <summary>
        ''' columnNameで指定したプロパティのエラーを返します。
        ''' </summary>
        ''' <param name="columnName">プロパティ名</param>
        ''' <returns>エラーメッセージ</returns>
        Default Public ReadOnly Property Errors(ByVal columnName As String) As String Implements IDataErrorInfo.Item
            Get
                If Me._Errors.ContainsKey(columnName) Then
                    Return Me._Errors(columnName)
                End If
                Return Nothing
            End Get
        End Property

        ''' <summary>
        ''' プロパティにエラーメッセージを設定します。
        ''' </summary>
        ''' <param name="propertyName">プロパティ名</param>
        ''' <param name="errorMessage">エラーメッセージ</param>
        Protected Sub SetError(ByVal propertyName As String, ByVal errorMessage As String)
            Me._Errors(propertyName) = errorMessage
            Me.RaisePropertyChanged("HasError")
        End Sub

        ''' <summary>
        ''' プロパティのエラーをクリアします。
        ''' </summary>
        ''' <param name="propertyName">プロパティ名</param>
        Protected Sub ClearError(ByVal propertyName As String)
            If Me._Errors.ContainsKey(propertyName) Then
                Me._Errors.Remove(propertyName)
                Me.RaisePropertyChanged("HasError")
            End If
        End Sub

        ''' <summary>
        ''' 全てのエラーをクリアします。
        ''' </summary>
        Protected Sub ClearErrors()
            Me._Errors.Clear()
            Me.RaisePropertyChanged("HasError")
        End Sub

        ''' <summary>
        ''' エラーの有無を取得します。
        ''' </summary>
        Public ReadOnly Property HasError() As Boolean
            Get
                Return Me._Errors.Count <> 0
            End Get
        End Property

#Region "WinForm用"

        ''' <summary>
        ''' エラーがあった時に発行されます。
        ''' </summary>
        Public Event HasErrorChanged As EventHandler

        ''' <summary>
        ''' XxxxChanged 形式のイベントを発行します。
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