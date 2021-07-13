Option Explicit On
Option Strict On

Imports NUnit.Framework

Namespace CompuMaster.Test.Data

    <TestFixture(Category:="TextTables")> Public Class TextRowTest

        <Test> Public Sub SpaceCharHebrew()
            Assert.AreEqual(" "c, "השר הנגבי".Chars(3))
            Assert.AreEqual(32, AscW("השר הנגבי".Chars(3)))
        End Sub

        <Test> Public Sub EvenVsOdd()
            Assert.AreEqual(True, (8 And 1) = 0)
            Assert.AreEqual(False, (7 And 1) = 0)
            Assert.AreEqual(4, 7 \ 2 + 1)
            Assert.AreEqual(3, 7 \ 2)
        End Sub

    End Class

End Namespace