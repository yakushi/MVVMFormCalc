Namespace MVVMCalc
    Public Class Form1
        Inherits System.Windows.Forms.Form

#Region " Windows フォーム デザイナで生成されたコード "

        Dim vm As ViewModel.MainViewModel

        Public Sub New()
            MyBase.New()

            ' この呼び出しは Windows フォーム デザイナで必要です。
            InitializeComponent()

            ' InitializeComponent() 呼び出しの後に初期化を追加します。
            Me.vm = New ViewModel.MainViewModel
            Me.TextBox1.DataBindings.Add("Text", Me.vm, "Lhs")
            AddHandler TextBox1.Validated, AddressOf TextBox_Validated
            Me.TextBox2.DataBindings.Add("Text", Me.vm, "Rhs")
            AddHandler TextBox2.Validated, AddressOf TextBox_Validated
            Me.Label4.DataBindings.Add("Text", Me.vm, "Answer")
            Me.ComboBox1.DataSource = Me.vm.CalculateTypes
            Me.ComboBox1.DisplayMember = "Label"
            Me.ComboBox1.ValueMember = "CalculateType"
            Me.ComboBox1.DataBindings.Add("SelectedValue", Me.vm, "SelectedCalculateTypeValue")
            Me.Button1.DataBindings.Add("Enabled", Me.vm.CalculateCommand, "CanExecute")
            AddHandler Button1.Click, Me.vm.CalculateCommand.CreateEventHandler
            AddHandler vm.ErrorMessenger.Raised, Common.ConfirmAction.CreateEventHandler
        End Sub

        ' Form は、コンポーネント一覧に後処理を実行するために dispose をオーバーライドします。
        Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
            If disposing Then
                If Not (components Is Nothing) Then
                    components.Dispose()
                End If
            End If
            MyBase.Dispose(disposing)
        End Sub

        ' Windows フォーム デザイナで必要です。
        Private components As System.ComponentModel.IContainer

        ' メモ : 以下のプロシージャは、Windows フォーム デザイナで必要です。
        'Windows フォーム デザイナを使って変更してください。  
        ' コード エディタを使って変更しないでください。
        Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
        Friend WithEvents Button1 As System.Windows.Forms.Button
        Friend WithEvents Label1 As System.Windows.Forms.Label
        Friend WithEvents Label2 As System.Windows.Forms.Label
        Friend WithEvents Label3 As System.Windows.Forms.Label
        Friend WithEvents ComboBox1 As System.Windows.Forms.ComboBox
        Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
        Friend WithEvents Label4 As System.Windows.Forms.Label
        Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
        <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
            Me.components = New System.ComponentModel.Container
            Me.TextBox1 = New System.Windows.Forms.TextBox
            Me.Button1 = New System.Windows.Forms.Button
            Me.Label1 = New System.Windows.Forms.Label
            Me.Label2 = New System.Windows.Forms.Label
            Me.Label3 = New System.Windows.Forms.Label
            Me.ComboBox1 = New System.Windows.Forms.ComboBox
            Me.TextBox2 = New System.Windows.Forms.TextBox
            Me.Label4 = New System.Windows.Forms.Label
            Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
            Me.SuspendLayout()
            '
            'TextBox1
            '
            Me.TextBox1.Location = New System.Drawing.Point(104, 0)
            Me.TextBox1.Name = "TextBox1"
            Me.TextBox1.Size = New System.Drawing.Size(176, 19)
            Me.TextBox1.TabIndex = 0
            Me.TextBox1.Text = ""
            '
            'Button1
            '
            Me.Button1.Location = New System.Drawing.Point(0, 72)
            Me.Button1.Name = "Button1"
            Me.Button1.Size = New System.Drawing.Size(280, 23)
            Me.Button1.TabIndex = 3
            Me.Button1.Text = "計算実行"
            '
            'Label1
            '
            Me.Label1.Location = New System.Drawing.Point(0, 0)
            Me.Label1.Name = "Label1"
            Me.Label1.TabIndex = 2
            Me.Label1.Text = "左辺値:"
            '
            'Label2
            '
            Me.Label2.Location = New System.Drawing.Point(0, 24)
            Me.Label2.Name = "Label2"
            Me.Label2.TabIndex = 3
            Me.Label2.Text = "計算方法:"
            '
            'Label3
            '
            Me.Label3.Location = New System.Drawing.Point(0, 48)
            Me.Label3.Name = "Label3"
            Me.Label3.TabIndex = 4
            Me.Label3.Text = "右辺値:"
            '
            'ComboBox1
            '
            Me.ComboBox1.DisplayMember = "Label"
            Me.ComboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
            Me.ComboBox1.Location = New System.Drawing.Point(104, 24)
            Me.ComboBox1.Name = "ComboBox1"
            Me.ComboBox1.Size = New System.Drawing.Size(176, 20)
            Me.ComboBox1.TabIndex = 1
            Me.ComboBox1.ValueMember = "CalculateType"
            '
            'TextBox2
            '
            Me.TextBox2.Location = New System.Drawing.Point(104, 48)
            Me.TextBox2.Name = "TextBox2"
            Me.TextBox2.Size = New System.Drawing.Size(176, 19)
            Me.TextBox2.TabIndex = 2
            Me.TextBox2.Text = ""
            '
            'Label4
            '
            Me.Label4.Location = New System.Drawing.Point(0, 96)
            Me.Label4.Name = "Label4"
            Me.Label4.Size = New System.Drawing.Size(280, 23)
            Me.Label4.TabIndex = 7
            Me.Label4.TextAlign = System.Drawing.ContentAlignment.TopCenter
            '
            'Form1
            '
            Me.AutoScaleBaseSize = New System.Drawing.Size(5, 12)
            Me.ClientSize = New System.Drawing.Size(284, 123)
            Me.Controls.Add(Me.Label4)
            Me.Controls.Add(Me.TextBox2)
            Me.Controls.Add(Me.TextBox1)
            Me.Controls.Add(Me.ComboBox1)
            Me.Controls.Add(Me.Label3)
            Me.Controls.Add(Me.Label2)
            Me.Controls.Add(Me.Label1)
            Me.Controls.Add(Me.Button1)
            Me.Name = "Form1"
            Me.Text = "Form1"
            Me.ResumeLayout(False)

        End Sub

#End Region

        Private Sub TextBox_Validated(ByVal sender As Object, ByVal e As System.EventArgs)
            Dim target As TextBox = DirectCast(sender, TextBox)
            Dim propertyName As String = target.DataBindings("Text").BindingMemberInfo.BindingField

            ' バリデーションエラーがある場合に
            If Not vm(propertyName) Is Nothing Then
                ' ツールチップにエラーメッセージと
                Me.ToolTip1.SetToolTip(target, vm(propertyName))
                ' 背景色を設定する
                target.BackColor = Color.Pink
            Else
                ' ツールチップと背景色を戻す
                Me.ToolTip1.SetToolTip(target, "")
                target.BackColor = System.Drawing.SystemColors.Window
            End If
        End Sub
    End Class
End Namespace
