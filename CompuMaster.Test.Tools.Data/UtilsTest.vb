Imports NUnit.Framework
Imports NUnit.Framework.Legacy

Namespace CompuMaster.Test.Data

    <TestFixture(Category:="Common Utils")> Public Class UtilsTest

        <Test> Public Sub TripleState()
            ClassicAssert.AreEqual(CompuMaster.Data.Utils.TripleState.Undefined, CType(Nothing, CompuMaster.Data.Utils.TripleState))
        End Sub

        <Test> Public Sub NoDbNull()
            'Object type
            ClassicAssert.AreEqual(Me, CompuMaster.Data.Utils.NoDBNull(Me))
            ClassicAssert.AreEqual(Nothing, CompuMaster.Data.Utils.NoDBNull(DBNull.Value))
            ClassicAssert.AreNotEqual(Nothing, CompuMaster.Data.Utils.NoDBNull(New Object, CType(Nothing, Object)))
            ClassicAssert.AreNotEqual(Me, CompuMaster.Data.Utils.NoDBNull(New Object, CType(Me, Object)))
            ClassicAssert.AreEqual(Me, CompuMaster.Data.Utils.NoDBNull(Me, CType(Nothing, Object)))
            ClassicAssert.AreEqual(Nothing, CompuMaster.Data.Utils.NoDBNull(DBNull.Value, CType(Nothing, Object)))
            ClassicAssert.AreEqual(Me, CompuMaster.Data.Utils.NoDBNull(DBNull.Value, CType(Me, Object)))

            'Value types incl. strings
            ClassicAssert.AreEqual(1, CompuMaster.Data.Utils.NoDBNull(1, -1))
            ClassicAssert.AreEqual(-1, CompuMaster.Data.Utils.NoDBNull(DBNull.Value, -1))
            ClassicAssert.AreEqual(1L, CompuMaster.Data.Utils.NoDBNull(1L, -1L))
            ClassicAssert.AreEqual(-1L, CompuMaster.Data.Utils.NoDBNull(DBNull.Value, -1L))
            ClassicAssert.AreEqual(1S, CompuMaster.Data.Utils.NoDBNull(1S, -1S))
            ClassicAssert.AreEqual(-1S, CompuMaster.Data.Utils.NoDBNull(DBNull.Value, -1S))
            ClassicAssert.AreEqual(1D, CompuMaster.Data.Utils.NoDBNull(1D, -1D))
            ClassicAssert.AreEqual(-1D, CompuMaster.Data.Utils.NoDBNull(DBNull.Value, -1D))
            ClassicAssert.AreEqual(1.0!, CompuMaster.Data.Utils.NoDBNull(1.0!, -1.0!))
            ClassicAssert.AreEqual(-1.0!, CompuMaster.Data.Utils.NoDBNull(DBNull.Value, -1.0!))
            ClassicAssert.AreEqual(1.0R, CompuMaster.Data.Utils.NoDBNull(1.0R, -1.0R))
            ClassicAssert.AreEqual(-1.0R, CompuMaster.Data.Utils.NoDBNull(DBNull.Value, -1.0R))
            ClassicAssert.AreEqual(True, CompuMaster.Data.Utils.NoDBNull(True, False))
            ClassicAssert.AreEqual(False, CompuMaster.Data.Utils.NoDBNull(DBNull.Value, False))
            ClassicAssert.AreEqual(CType(1, Byte), CompuMaster.Data.Utils.NoDBNull(CType(1, Byte), CType(2, Byte)))
            ClassicAssert.AreEqual(CType(2, Byte), CompuMaster.Data.Utils.NoDBNull(DBNull.Value, CType(2, Byte)))
            ClassicAssert.AreEqual(CType(1, UInt16), CompuMaster.Data.Utils.NoDBNull(CType(1, UInt16), CType(2, UInt16)))
            ClassicAssert.AreEqual(CType(2, UInt16), CompuMaster.Data.Utils.NoDBNull(DBNull.Value, CType(2, UInt16)))
            ClassicAssert.AreEqual(CType(1, UInt32), CompuMaster.Data.Utils.NoDBNull(CType(1, UInt32), CType(2, UInt32)))
            ClassicAssert.AreEqual(CType(2, UInt32), CompuMaster.Data.Utils.NoDBNull(DBNull.Value, CType(2, UInt32)))
            ClassicAssert.AreEqual(CType(1, UInt64), CompuMaster.Data.Utils.NoDBNull(CType(1, UInt64), CType(2, UInt64)))
            ClassicAssert.AreEqual(CType(2, UInt64), CompuMaster.Data.Utils.NoDBNull(DBNull.Value, CType(2, UInt64)))
            ClassicAssert.AreEqual(New DateTime(1), CompuMaster.Data.Utils.NoDBNull(New DateTime(1), New DateTime(2)))
            ClassicAssert.AreEqual(New DateTime(2), CompuMaster.Data.Utils.NoDBNull(DBNull.Value, New DateTime(2)))
            ClassicAssert.AreEqual("1", CompuMaster.Data.Utils.NoDBNull("1", "-1"))
            ClassicAssert.AreEqual("-1", CompuMaster.Data.Utils.NoDBNull(DBNull.Value, "-1"))

            'Generics
            ClassicAssert.AreEqual(1, CompuMaster.Data.Utils.NoDBNull(Of Long?)(New Long?(1)))
            If Type.GetType("Mono.Runtime") Is Nothing Then
                'Running in Mono .NET Framework which is not fully implemented/bug-free at this point
                ClassicAssert.AreEqual(1, CompuMaster.Data.Utils.NoDBNull(Of Long)(1))
            End If
            ClassicAssert.AreEqual(1, CompuMaster.Data.Utils.NoDBNull(Of Long?)(1L))
            ClassicAssert.AreEqual(1, CompuMaster.Data.Utils.NoDBNull(Of Long?)(1))
            'Assert.Catch(Of InvalidCastException)(Sub() CompuMaster.Data.Utils.NoDBNull(Of Long?)(1))
            ClassicAssert.AreEqual("Text", CompuMaster.Data.Utils.NoDBNull(Of String)("Text"))
            ClassicAssert.AreEqual(True, CompuMaster.Data.Utils.NoDBNull(Of Boolean?)(True))
            ClassicAssert.AreEqual(New Byte() {1}, CompuMaster.Data.Utils.NoDBNull(Of Byte())(New Byte() {1}))
            ClassicAssert.AreEqual(1D, CompuMaster.Data.Utils.NoDBNull(Of Decimal?)(1D))
            ClassicAssert.AreEqual(New DateTime(1, 1, 1), CompuMaster.Data.Utils.NoDBNull(Of DateTime?)(New DateTime(1, 1, 1)))
            ClassicAssert.AreEqual(Nothing, CompuMaster.Data.Utils.NoDBNull(Of Long?)(DBNull.Value))
            ClassicAssert.AreEqual(Nothing, CompuMaster.Data.Utils.NoDBNull(Of String)(DBNull.Value))
            ClassicAssert.AreEqual(Nothing, CompuMaster.Data.Utils.NoDBNull(Of Boolean?)(DBNull.Value))
            ClassicAssert.AreEqual(Nothing, CompuMaster.Data.Utils.NoDBNull(Of Byte())(DBNull.Value))
            ClassicAssert.AreEqual(Nothing, CompuMaster.Data.Utils.NoDBNull(Of Decimal?)(DBNull.Value))
            ClassicAssert.AreEqual(Nothing, CompuMaster.Data.Utils.NoDBNull(Of DateTime?)(DBNull.Value))
            ClassicAssert.AreEqual(False, CompuMaster.Data.Utils.NoDBNull(Of Long?)(DBNull.Value).HasValue)
            ClassicAssert.AreEqual(False, CompuMaster.Data.Utils.NoDBNull(Of Boolean?)(DBNull.Value).HasValue)
            ClassicAssert.AreEqual(False, CompuMaster.Data.Utils.NoDBNull(Of Decimal?)(DBNull.Value).HasValue)
            ClassicAssert.AreEqual(False, CompuMaster.Data.Utils.NoDBNull(Of DateTime?)(DBNull.Value).HasValue)
            ClassicAssert.AreEqual(CType(1, UInt16), CompuMaster.Data.Utils.NoDBNull(Of UInt16)(CType(1, UInt16)))
            ClassicAssert.AreEqual(CType(Nothing, UInt16), CompuMaster.Data.Utils.NoDBNull(Of UInt16)(DBNull.Value))
            ClassicAssert.AreEqual(CType(1, UInt16), CompuMaster.Data.Utils.NoDBNull(Of UInt16)(CType(1, UInt16), CType(2, UInt16)))
            ClassicAssert.AreEqual(CType(2, UInt16), CompuMaster.Data.Utils.NoDBNull(Of UInt16)(DBNull.Value, CType(2, UInt16)))
            'Nullable Generics
            ClassicAssert.AreEqual(CType(1, UInt16?), CompuMaster.Data.Utils.NoDBNull(Of UInt16?)(CType(1, UInt16)))
            ClassicAssert.AreEqual(CType(Nothing, UInt16?), CompuMaster.Data.Utils.NoDBNull(Of UInt16?)(DBNull.Value))
            ClassicAssert.AreEqual(New UInt16?, CompuMaster.Data.Utils.NoDBNull(Of UInt16?)(DBNull.Value))
            ClassicAssert.AreEqual(CType(1, UInt16?), CompuMaster.Data.Utils.NoDBNull(Of UInt16?)(CType(1, UInt16), CType(2, UInt16)))
            ClassicAssert.AreEqual(CType(2, UInt16?), CompuMaster.Data.Utils.NoDBNull(Of UInt16?)(DBNull.Value, CType(2, UInt16)))

        End Sub

        <Test> <CodeAnalysis.SuppressMessage("Performance", "CA1825:Avoid zero-length array allocations.", Justification:="<Ausstehend>")>
        Public Sub NoDBNullArrayOrListFromString_Generics()
            'DBNull.Value
            ClassicAssert.AreEqual(Nothing, CompuMaster.Data.Utils.NoDBNullArrayFromString(Of String)(DBNull.Value, ","c))
            ClassicAssert.AreEqual(Nothing, CompuMaster.Data.Utils.NoDBNullArrayFromString(Of Object)(DBNull.Value, ","c))
            ClassicAssert.AreEqual(Nothing, CompuMaster.Data.Utils.NoDBNullArrayFromString(Of Integer)(DBNull.Value, ","c))
            ClassicAssert.AreEqual(Nothing, CompuMaster.Data.Utils.NoDBNullArrayFromString(Of Integer?)(DBNull.Value, ","c))
            'Empty String
            ClassicAssert.AreEqual(New String() {}, CompuMaster.Data.Utils.NoDBNullArrayFromString(Of String)("", ","c))
            ClassicAssert.AreEqual(New Object() {}, CompuMaster.Data.Utils.NoDBNullArrayFromString(Of Object)("", ","c))
            ClassicAssert.AreEqual(New Integer() {}, CompuMaster.Data.Utils.NoDBNullArrayFromString(Of Integer)("", ","c))
            ClassicAssert.AreEqual(New Integer?() {}, CompuMaster.Data.Utils.NoDBNullArrayFromString(Of Integer?)("", ","c))
            ClassicAssert.AreEqual(New String() {}, CompuMaster.Data.Utils.NoDBNullArrayFromString(Of String)("", New Char() {","c, ";"c}))
            ClassicAssert.AreEqual(New Object() {}, CompuMaster.Data.Utils.NoDBNullArrayFromString(Of Object)("", New Char() {","c, ";"c}))
            ClassicAssert.AreEqual(New Integer() {}, CompuMaster.Data.Utils.NoDBNullArrayFromString(Of Integer)("", New Char() {","c, ";"c}))
            ClassicAssert.AreEqual(New Integer?() {}, CompuMaster.Data.Utils.NoDBNullArrayFromString(Of Integer?)("", New Char() {","c, ";"c}))
            'Single String
            ClassicAssert.AreEqual(New String() {"Test"}, CompuMaster.Data.Utils.NoDBNullArrayFromString(Of String)("Test", ","c))
            'Assert.AreEqual(New Object() {New Object}, CompuMaster.Data.Utils.NoDBNullArrayFromString(Of Object)(New Object.ToString, ","))
            If Type.GetType("Mono.Runtime") Is Nothing Then
                'Running in Mono .NET Framework which is not fully implemented/bug-free at this point
                ClassicAssert.AreEqual(New Integer() {1}, CompuMaster.Data.Utils.NoDBNullArrayFromString(Of Integer)("1", ","c))
            End If
            ClassicAssert.AreEqual(New Integer?() {1}, CompuMaster.Data.Utils.NoDBNullArrayFromString(Of Integer?)("1", ","c))
            'Separatable String
            ClassicAssert.AreEqual(New String() {"Test1", "Test2", "Test3"}, CompuMaster.Data.Utils.NoDBNullArrayFromString(Of String)("Test1,Test2,Test3", ","c))
            'Assert.AreEqual(New Object() {}, CompuMaster.Data.Utils.NoDBNullArrayFromString(Of Object)("", ","))
            If Type.GetType("Mono.Runtime") Is Nothing Then
                'Running in Mono .NET Framework which is not fully implemented/bug-free at this point
                ClassicAssert.AreEqual(New Integer() {1, 2, 3}, CompuMaster.Data.Utils.NoDBNullArrayFromString(Of Integer)("1,2,3", ","c))
            End If
            ClassicAssert.AreEqual(New Integer?() {1, 2, 3}, CompuMaster.Data.Utils.NoDBNullArrayFromString(Of Integer?)("1,2,3", ","c))
        End Sub

        <Test> <CodeAnalysis.SuppressMessage("Performance", "CA1825:Avoid zero-length array allocations.", Justification:="<Ausstehend>")>
        Public Sub NoDBNullArrayOrListFromString_NoGenerics()
            'DBNull.Value
            ClassicAssert.AreEqual(Nothing, CompuMaster.Data.Utils.NoDBNullArrayFromString(DBNull.Value, ","c, CType(Nothing, String())))
            ClassicAssert.AreEqual(Nothing, CompuMaster.Data.Utils.NoDBNullArrayFromString(DBNull.Value, ","c, CType(Nothing, Object())))
            ClassicAssert.AreEqual(Nothing, CompuMaster.Data.Utils.NoDBNullArrayFromString(DBNull.Value, ","c, CType(Nothing, Integer())))
            'Assert.AreEqual(Nothing, CompuMaster.Data.Utils.NoDBNullArrayFromString(DBNull.Value, ","c, CType(Nothing, Integer?())))
            'Empty String
            ClassicAssert.AreEqual(New String() {}, CompuMaster.Data.Utils.NoDBNullArrayFromString("", ","c, CType(Nothing, String())))
            ClassicAssert.AreEqual(New Object() {}, CompuMaster.Data.Utils.NoDBNullArrayFromString("", ","c, CType(Nothing, Object())))
            ClassicAssert.AreEqual(New Integer() {}, CompuMaster.Data.Utils.NoDBNullArrayFromString("", ","c, CType(Nothing, Integer())))
            'Assert.AreEqual(New Integer?() {}, CompuMaster.Data.Utils.NoDBNullArrayFromString("", ","c, CType(Nothing, Integer?())))
            'Single String
            ClassicAssert.AreEqual(New String() {"Test"}, CompuMaster.Data.Utils.NoDBNullArrayFromString("Test", ","c, CType(Nothing, String())))
            'Assert.AreEqual(New Object() {New Object}, CompuMaster.Data.Utils.NoDBNullArrayFromString(New Object.ToString, ","c, CType(Nothing, Object())))
            ClassicAssert.AreEqual(New Integer() {1}, CompuMaster.Data.Utils.NoDBNullArrayFromString("1", ","c, CType(Nothing, Integer())))
            'Assert.AreEqual(New Integer?() {1}, CompuMaster.Data.Utils.NoDBNullArrayFromString("1", ","c, CType(Nothing, Integer?())))
            'Separatable String
            ClassicAssert.AreEqual(New String() {"Test1", "Test2", "Test3"}, CompuMaster.Data.Utils.NoDBNullArrayFromString("Test1,Test2,Test3", ","c, CType(Nothing, String())))
            'Assert.AreEqual(New Object() {}, CompuMaster.Data.Utils.NoDBNullArrayFromString("", ","c, CType(Nothing, Object())))
            ClassicAssert.AreEqual(New Integer() {1, 2, 3}, CompuMaster.Data.Utils.NoDBNullArrayFromString("1,2,3", ","c, CType(Nothing, Integer())))
            'Assert.AreEqual(New Integer?() {1, 2, 3}, CompuMaster.Data.Utils.NoDBNullArrayFromString("1,2,3", ","c, CType(Nothing, Integer?())))
        End Sub

        <Test> Public Sub NullableTypeWithItsValueOrDBNull()
            ClassicAssert.AreNotEqual(GetType(Integer?), CompuMaster.Data.Utils.NullableTypeWithItsValueOrDBNull(Of Integer)(New Integer?(1)).GetType)
            ClassicAssert.AreEqual(GetType(Integer), CompuMaster.Data.Utils.NullableTypeWithItsValueOrDBNull(Of Integer)(New Integer?(1)).GetType)
            ClassicAssert.AreEqual(1, CompuMaster.Data.Utils.NullableTypeWithItsValueOrDBNull(Of Integer)(New Integer?(1)))
            ClassicAssert.AreEqual(1, CompuMaster.Data.Utils.NullableTypeWithItsValueOrDBNull(Of Integer)(1))
            ClassicAssert.AreEqual(DBNull.Value, CompuMaster.Data.Utils.NullableTypeWithItsValueOrDBNull(Of Integer)(New Integer?))
            ClassicAssert.AreEqual(DBNull.Value, CompuMaster.Data.Utils.NullableTypeWithItsValueOrDBNull(Of Integer)(Nothing))
        End Sub

        <Test> <CodeAnalysis.SuppressMessage("Performance", "CA1825:Avoid zero-length array allocations.", Justification:="<Ausstehend>")>
        Sub ArrayNotNothingOrDBNull()
            ClassicAssert.AreEqual(DBNull.Value, CompuMaster.Data.Utils.ArrayNotNothingOrDBNull(Nothing))
            Dim arr As Byte() = Nothing
            ClassicAssert.AreEqual(DBNull.Value, CompuMaster.Data.Utils.ArrayNotNothingOrDBNull(arr))
            ClassicAssert.AreEqual(New Byte() {}, CompuMaster.Data.Utils.ArrayNotNothingOrDBNull(New Byte() {}))
            ClassicAssert.AreEqual(New Byte() {1}, CompuMaster.Data.Utils.ArrayNotNothingOrDBNull(New Byte() {1}))
        End Sub

        <Test> <CodeAnalysis.SuppressMessage("Performance", "CA1825:Avoid zero-length array allocations.", Justification:="<Ausstehend>")>
        Sub ArrayNotEmptyOrDBNull()
            ClassicAssert.AreEqual(DBNull.Value, CompuMaster.Data.Utils.ArrayNotEmptyOrDBNull(Nothing))
            Dim arr As Byte() = Nothing
            ClassicAssert.AreEqual(DBNull.Value, CompuMaster.Data.Utils.ArrayNotEmptyOrDBNull(arr))
            ClassicAssert.AreEqual(DBNull.Value, CompuMaster.Data.Utils.ArrayNotEmptyOrDBNull(New Byte() {}))
            ClassicAssert.AreEqual(New Byte() {1}, CompuMaster.Data.Utils.ArrayNotEmptyOrDBNull(New Byte() {1}))
        End Sub

        <Test> Public Sub ReadStringFromUri()
            Dim Url As String = CsvTest.CSV_ONLINE_TEST_RESOURCE_EN_US_URL
            ClassicAssert.NotZero(CompuMaster.Data.Utils.ReadByteDataFromUri(Url).Length)
            ClassicAssert.NotZero(CompuMaster.Data.Utils.ReadStringDataFromUri(Url, System.Text.Encoding.ASCII.WebName).Length)
            ClassicAssert.NotZero(CompuMaster.Data.Utils.ReadStringDataFromUri(Url, System.Text.Encoding.UTF8.WebName).Length)
            ClassicAssert.NotZero(CompuMaster.Data.Utils.ReadStringDataFromUri(Url, Nothing).Length)
        End Sub
    End Class

End Namespace
