Namespace MVVMCalc.Common

    ''' <summary>
    ''' ViewModel��View�̊Ԃł̏����������s�����b�Z�[�W
    ''' </summary>
    Public Class Message
        ''' <summary>
        ''' ���b�Z�[�W�̖{��
        ''' </summary>
        Private _body As Object
        Public ReadOnly Property Body() As Object
            Get
                Return _body
            End Get
        End Property

        ''' <summary>
        ''' View����ViewMode�ւ̃��b�Z�[�W�̃��X�|���X
        ''' </summary>
        Private _response As Object
        Public Property Response() As Object
            Get
                Return _response
            End Get
            Set(ByVal Value As Object)
                _response = Value
            End Set
        End Property

        ''' <summary>
        ''' Body���w�肵��Message���쐬����
        ''' </summary>
        ''' <param name="body"></param>
        Public Sub New(ByVal body As Object)
            Me._body = body
        End Sub

    End Class
End Namespace
