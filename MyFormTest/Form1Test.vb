Imports NUnit.Framework

<TestFixture()> Public Class Form1Test
    <Test()> Public Sub Add()
        Assert.AreEqual(6, 2 + 4)
    End Sub

End Class
