Namespace MVVMCalc.Model

    ''' <summary>
    ''' �v�Z���s���N���X
    ''' </summary>
    Public Class Calculator

        ''' <summary>
        ''' �v�Z���@��\��CalculateType�Ǝ��ۂ̏����̃}�b�v
        ''' </summary>
        Private Shared ReadOnly calcMap As CalculateDictionary = New CalculateDictionary
        Public Sub New()
            ''' CalculateDictionary �̏�����
            If calcMap.Count = 0 Then
                ' ���w��
                calcMap.Add(CalclateType.None, New DelegateCalculator(AddressOf CalculatorFunction.None))
                ' �����Z
                calcMap.Add(CalclateType.Add, New DelegateCalculator(AddressOf CalculatorFunction.Add))
                ' �����Z
                calcMap.Add(CalclateType.Subs, New DelegateCalculator(AddressOf CalculatorFunction.Subs))
                ' �|���Z
                calcMap.Add(CalclateType.Mul, New DelegateCalculator(AddressOf CalculatorFunction.Mul))
                ' ����Z
                calcMap.Add(CalclateType.Div, New DelegateCalculator(AddressOf CalculatorFunction.Div))
            End If
        End Sub
        ''' <summary>
        ''' �v�Z������\�� DelegateCalculator
        ''' </summary>
        Private Delegate Function DelegateCalculator(ByVal lhs As Double, ByVal rhs As Double) As Double
        ''' <summary>
        ''' �v�Z���@��\�� CalculateType �Ǝ��ۂ̏������i�[����J�X�^�� Dictionary
        ''' </summary>
        Private Class CalculateDictionary
            Inherits System.Collections.DictionaryBase

            Default Public Property Item(ByVal key As CalclateType) As DelegateCalculator
                Get
                    Return DirectCast(Dictionary(key), DelegateCalculator)
                End Get
                Set(ByVal Value As DelegateCalculator)
                    Dictionary(key) = Value
                End Set
            End Property
            Public Sub Add(ByVal key As CalclateType, ByVal value As DelegateCalculator)
                Dictionary.Add(key, value)
            End Sub
        End Class
        ''' <summary>
        ''' �v�Z�̏���
        ''' static �ɂ��Ȃ����߁Aprivate class �Ŏ����B
        ''' </summary>
        Private Class CalculatorFunction
            ' ���w��
            Public Shared Function None(ByVal x As Double, ByVal y As Double) As Double
                Throw New InvalidOperationException
            End Function
            ' �����Z
            Public Shared Function Add(ByVal x As Double, ByVal y As Double) As Double
                Return x + y
            End Function
            ' �����Z
            Public Shared Function Subs(ByVal x As Double, ByVal y As Double) As Double
                Return x - y
            End Function
            ' �|���Z
            Public Shared Function Mul(ByVal x As Double, ByVal y As Double) As Double
                Return x * y
            End Function
            ' ����Z
            Public Shared Function Div(ByVal x As Double, ByVal y As Double) As Double
                Return x / y
            End Function
        End Class

        ''' <summary>
        ''' �n���ꂽ�l�̎w�肳�ꂽ�v�Z���ʂ�Ԃ��B
        ''' </summary>
        ''' <param name="x">���Ӓl</param>
        ''' <param name="y">�E�Ӓl</param>
        ''' <param name="op">�v�Z���@</param>
        ''' <returns>�v�Z����</returns>
        Public Function Execute(ByVal x As Double, ByVal y As Double, ByVal op As CalclateType) As Double
            Return calcMap(op)(x, y)
        End Function

    End Class
End Namespace
