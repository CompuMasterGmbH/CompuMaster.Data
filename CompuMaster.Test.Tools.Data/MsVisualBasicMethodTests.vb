Option Strict On
Option Explicit On

Imports NUnit.Framework
Imports NUnit.Framework.Legacy
'Imports Microsoft.VisualBasic
Imports CompuMaster.Data

Namespace CompuMaster.Test.Data

    <TestFixture(Category:="Common Utils")> Public Class MsVisualBasicMethodTests

        <Test> Public Sub IsDbNullTest()
            ClassicAssert.IsTrue(Information.IsDBNull(DBNull.Value))
            ClassicAssert.IsFalse(Information.IsDBNull(Nothing))
            ClassicAssert.IsFalse(Information.IsDBNull(String.Empty))
            ClassicAssert.IsFalse(Information.IsDBNull(2.0))
            ClassicAssert.IsFalse(Information.IsDBNull(New Object))
        End Sub

        <Test> Public Sub IsNothingTest()
            ClassicAssert.IsFalse(Information.IsNothing(DBNull.Value))
            ClassicAssert.IsTrue(Information.IsNothing(Nothing))
            ClassicAssert.IsTrue(Information.IsNothing(CType(Nothing, String)))
            ClassicAssert.IsFalse(Information.IsNothing(String.Empty))
            ClassicAssert.IsFalse(Information.IsNothing(0))
            ClassicAssert.IsFalse(Information.IsNothing(New Object))
            ClassicAssert.IsFalse(Information.IsNothing(New Object()))
            ClassicAssert.IsTrue(Information.IsNothing(CType(Nothing, Object())))
        End Sub

        <Test> Public Sub IsNumericTest()
            ClassicAssert.IsFalse(Information.IsNumeric(DBNull.Value))
            ClassicAssert.IsFalse(Information.IsNumeric(Nothing))
            ClassicAssert.IsFalse(Information.IsNumeric(CType(Nothing, String)))
            ClassicAssert.IsFalse(Information.IsNumeric(String.Empty))
            ClassicAssert.IsTrue(Information.IsNumeric(True))
            ClassicAssert.IsTrue(Information.IsNumeric(0))
            ClassicAssert.IsTrue(Information.IsNumeric(0S))
            ClassicAssert.IsTrue(Information.IsNumeric(Byte.MaxValue))
            ClassicAssert.IsTrue(Information.IsNumeric(2.0!))
            ClassicAssert.IsTrue(Information.IsNumeric(2.0))
            ClassicAssert.IsTrue(Information.IsNumeric(20L))
            ClassicAssert.IsTrue(Information.IsNumeric(-200D))
            ClassicAssert.IsFalse(Information.IsNumeric(New Object))
            ClassicAssert.IsFalse(Information.IsNumeric(New Byte() {200}))
        End Sub

        <Test> Public Sub IsDateTest()
            ClassicAssert.IsFalse(Information.IsDate(DBNull.Value))
            ClassicAssert.IsFalse(Information.IsDate(Nothing))
            ClassicAssert.IsFalse(Information.IsDate(CType(Nothing, String)))
            ClassicAssert.IsFalse(Information.IsDate(String.Empty))
            ClassicAssert.IsFalse(Information.IsDate(0))
            ClassicAssert.IsFalse(Information.IsDate(New Object))
            ClassicAssert.IsTrue(Information.IsDate(New DateTime))
            ClassicAssert.IsFalse(Information.IsDate(New TimeSpan))
        End Sub

        <Test> Public Sub ControlCharsTest()
            ClassicAssert.AreEqual(ChrW(13) & ChrW(10), ControlChars.CrLf)
            ClassicAssert.AreEqual(ChrW(13), ControlChars.Cr)
            ClassicAssert.AreEqual(ChrW(10), ControlChars.Lf)
            ClassicAssert.AreEqual(ChrW(9), ControlChars.Tab)
        End Sub

        <Test> Public Sub TriStateTest()
            ClassicAssert.AreEqual(-2, CType(TriState.UseDefault, Integer))
            ClassicAssert.AreEqual(-1, CType(TriState.True, Integer))
            ClassicAssert.AreEqual(0, CType(TriState.False, Integer))
        End Sub

        <Test> Public Sub MidTest()
            ClassicAssert.AreEqual(Nothing, Strings.Mid(Nothing, 2))
            ClassicAssert.AreEqual("", Strings.Mid(Nothing, 2, 2))
            ClassicAssert.AreEqual("bcdef", Strings.Mid("abcdef", 2))
            ClassicAssert.AreEqual("bc", Strings.Mid("abcdef", 2, 2))
            ClassicAssert.AreEqual("", Strings.Mid("", 2))
            ClassicAssert.AreEqual("", Strings.Mid("", 2, 2))
        End Sub

        <Test> Public Sub ReplaceTest()
            ClassicAssert.AreEqual(Nothing, Strings.Replace(Nothing, "kj", "DD"))
            ClassicAssert.AreEqual("abcDEfDEf", Strings.Replace("abcdefdef", "de", "DE"))
            ClassicAssert.AreEqual("abcdefdef", Strings.Replace("abcdefdef", "kj", "DD"))
            'Assert.AreEqual(Nothing, Strings.Replace(String.Empty, "kj", "DD")) 'MS VisualBasic behaviour
            ClassicAssert.AreEqual("", Strings.Replace(String.Empty, "kj", "DD")) 'CM.Data behaviour
        End Sub

        <Test> Public Sub SpaceTest()
            ClassicAssert.AreEqual("", Strings.Space(0))
            ClassicAssert.AreEqual("    ", Strings.Space(4))
        End Sub

        <Test> Public Sub TrimTest()
            ClassicAssert.AreEqual("", Strings.Trim(Nothing))
            ClassicAssert.AreEqual("", Strings.Trim(String.Empty))
            ClassicAssert.AreEqual("", Strings.Trim("    "))
            ClassicAssert.AreEqual("ÖLKJ", Strings.Trim("  ÖLKJ  "))
        End Sub

        <Test> Public Sub LenTest()
            ClassicAssert.AreEqual(0, Strings.Len(CType(Nothing, String)))
            ClassicAssert.AreEqual(0, Strings.Len(String.Empty))
            ClassicAssert.AreEqual(4, Strings.Len("    "))
            ClassicAssert.AreEqual(8, Strings.Len("  ÖLKJ  "))
            ClassicAssert.AreEqual(0, Strings.Len(CType(Nothing, Object)))
            ClassicAssert.AreEqual(4, Strings.Len(0))
            ClassicAssert.AreEqual(1, Strings.Len(Byte.MaxValue))
            ClassicAssert.AreEqual(4, Strings.Len(0!))
            ClassicAssert.AreEqual(8, Strings.Len(0D))
            ClassicAssert.AreEqual(8, Strings.Len(0L))
            ClassicAssert.AreEqual(8, Strings.Len(New DateTime))
            ClassicAssert.AreEqual(8, Strings.Len(New DateTime(2020, 1, 1)))
            ClassicAssert.AreEqual(0, Strings.Len(New TimeSpan))
            ClassicAssert.AreEqual(0, Strings.Len(New TimeSpan(20, 10, 15)))
            'Assert.AreEqual(0, Strings.Len(New Byte() {}))
            'Assert.AreEqual(2, Strings.Len(New Integer() {1, 2}))
        End Sub

        <Test> Public Sub StrDupTest()
            ClassicAssert.AreEqual("", Strings.StrDup(0, "K"c))
            ClassicAssert.AreEqual("    ", Strings.StrDup(4, " "c))
            ClassicAssert.AreEqual("KKKK", Strings.StrDup(4, "K"c))
            ClassicAssert.AreEqual("", Strings.StrDup(0, "K"c))
            'Assert.AreEqual("    ", Strings.StrDup(4, " "))
            'Assert.AreEqual("KKKK", Strings.StrDup(4, "K"))
            'Assert.AreEqual("KKKK", Strings.StrDup(4, "KK"))
        End Sub

        <Test> Public Sub InStrTest()
            ClassicAssert.AreEqual(0, Strings.InStr(Nothing, Nothing))
            ClassicAssert.AreEqual(0, Strings.InStr(Nothing, String.Empty))
            ClassicAssert.AreEqual(0, Strings.InStr(String.Empty, String.Empty))
            ClassicAssert.AreEqual(1, Strings.InStr("abcdefdef", String.Empty))
            ClassicAssert.AreEqual(0, Strings.InStr("abcdefdef", "DE"))
            ClassicAssert.AreEqual(4, Strings.InStr("abcdefdef", "de"))
        End Sub

        <Test> Public Sub LSetTest()
            ClassicAssert.AreEqual("", Strings.LSet("text", 0))
            ClassicAssert.AreEqual("tex", Strings.LSet("text", 3))
            ClassicAssert.AreEqual("text", Strings.LSet("text", 4))
            ClassicAssert.AreEqual("text    ", Strings.LSet("text", 8))
            ClassicAssert.AreEqual("", Strings.LSet("", 0))
            ClassicAssert.AreEqual("   ", Strings.LSet("", 3))
            ClassicAssert.AreEqual("    ", Strings.LSet("", 4))
            ClassicAssert.AreEqual("        ", Strings.LSet("", 8))
            ClassicAssert.AreEqual("", Strings.LSet(Nothing, 0))
            ClassicAssert.AreEqual("   ", Strings.LSet(Nothing, 3))
            ClassicAssert.AreEqual("    ", Strings.LSet(Nothing, 4))
            ClassicAssert.AreEqual("        ", Strings.LSet(Nothing, 8))
        End Sub

        <Test> Public Sub RSetTest()
            ClassicAssert.AreEqual("", Strings.RSet("text", 0))
            ClassicAssert.AreEqual("tex", Strings.RSet("text", 3))
            ClassicAssert.AreEqual("text", Strings.RSet("text", 4))
            ClassicAssert.AreEqual("    text", Strings.RSet("text", 8))
            ClassicAssert.AreEqual("", Strings.RSet("", 0))
            ClassicAssert.AreEqual("   ", Strings.RSet("", 3))
            ClassicAssert.AreEqual("    ", Strings.RSet("", 4))
            ClassicAssert.AreEqual("        ", Strings.RSet("", 8))
            ClassicAssert.AreEqual("", Strings.RSet(Nothing, 0))
            ClassicAssert.AreEqual("   ", Strings.RSet(Nothing, 3))
            ClassicAssert.AreEqual("    ", Strings.RSet(Nothing, 4))
            ClassicAssert.AreEqual("        ", Strings.RSet(Nothing, 8))
        End Sub

        <Test> Public Sub LCaseTest()
            'Assert.AreEqual(ChrW(0), Strings.LCase(Nothing)) 'MS.VisualBasic behaviour
            ClassicAssert.AreEqual("", Strings.LCase(Nothing)) 'CM.Data behaviour
            ClassicAssert.AreEqual("", Strings.LCase(""))
            ClassicAssert.AreEqual("abcdefdef", Strings.LCase("abcDEfdef"))
            ClassicAssert.AreEqual("abcdefdefäöüß", Strings.LCase("abcDEfdefäÖüß"))
        End Sub

        <Test> Public Sub UCaseTest()
            'Assert.AreEqual(ChrW(0), Strings.UCase(Nothing)) 'MS.VisualBasic behaviour
            ClassicAssert.AreEqual("", Strings.UCase(Nothing)) 'CM.Data behaviour
            ClassicAssert.AreEqual("", Strings.UCase(""))
            ClassicAssert.AreEqual("ABCDEFDEF", Strings.UCase("abcDEfdef"))
            ClassicAssert.AreEqual("ABCDEFDEFÄÖÜß", Strings.UCase("abcDEfdefäÖüß"))
        End Sub

    End Class

End Namespace