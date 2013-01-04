Namespace MVVMCalc.Model

    ''' <summary>
    ''' 計算を行うクラス
    ''' </summary>
    Public Class Calculator

        ''' <summary>
        ''' 計算方法を表すCalculateTypeと実際の処理のマップ
        ''' </summary>
        Private Shared ReadOnly calcMap As CalculateDictionary = New CalculateDictionary
        Public Sub New()
            ''' CalculateDictionary の初期化
            If calcMap.Count = 0 Then
                ' 未指定
                calcMap.Add(CalclateType.None, New DelegateCalculator(AddressOf CalculatorFunction.None))
                ' 足し算
                calcMap.Add(CalclateType.Add, New DelegateCalculator(AddressOf CalculatorFunction.Add))
                ' 引き算
                calcMap.Add(CalclateType.Subs, New DelegateCalculator(AddressOf CalculatorFunction.Subs))
                ' 掛け算
                calcMap.Add(CalclateType.Mul, New DelegateCalculator(AddressOf CalculatorFunction.Mul))
                ' 割り算
                calcMap.Add(CalclateType.Div, New DelegateCalculator(AddressOf CalculatorFunction.Div))
            End If
        End Sub
        ''' <summary>
        ''' 計算処理を表す DelegateCalculator
        ''' </summary>
        Private Delegate Function DelegateCalculator(ByVal lhs As Double, ByVal rhs As Double) As Double
        ''' <summary>
        ''' 計算方法を表す CalculateType と実際の処理を格納するカスタム Dictionary
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
        ''' 計算の処理
        ''' static にしないため、private class で実装。
        ''' </summary>
        Private Class CalculatorFunction
            ' 未指定
            Public Shared Function None(ByVal x As Double, ByVal y As Double) As Double
                Throw New InvalidOperationException
            End Function
            ' 足し算
            Public Shared Function Add(ByVal x As Double, ByVal y As Double) As Double
                Return x + y
            End Function
            ' 引き算
            Public Shared Function Subs(ByVal x As Double, ByVal y As Double) As Double
                Return x - y
            End Function
            ' 掛け算
            Public Shared Function Mul(ByVal x As Double, ByVal y As Double) As Double
                Return x * y
            End Function
            ' 割り算
            Public Shared Function Div(ByVal x As Double, ByVal y As Double) As Double
                Return x / y
            End Function
        End Class

        ''' <summary>
        ''' 渡された値の指定された計算結果を返す。
        ''' </summary>
        ''' <param name="x">左辺値</param>
        ''' <param name="y">右辺値</param>
        ''' <param name="op">計算方法</param>
        ''' <returns>計算結果</returns>
        Public Function Execute(ByVal x As Double, ByVal y As Double, ByVal op As CalclateType) As Double
            Return calcMap(op)(x, y)
        End Function

    End Class
End Namespace
