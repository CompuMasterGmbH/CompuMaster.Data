﻿Imports NUnit.Framework

Namespace CompuMaster.Test.Data

    <TestFixture(Category:="Common Utils")> Public Class Utils

        <Test> Public Sub NoDbNull()
            'Object type
            Assert.AreEqual(Me, CompuMaster.Data.Utils.NoDBNull(Me))
            Assert.AreEqual(Nothing, CompuMaster.Data.Utils.NoDBNull(DBNull.Value))
            Assert.AreNotEqual(Nothing, CompuMaster.Data.Utils.NoDBNull(New Object, CType(Nothing, Object)))
            Assert.AreNotEqual(Me, CompuMaster.Data.Utils.NoDBNull(New Object, CType(Me, Object)))
            Assert.AreEqual(Me, CompuMaster.Data.Utils.NoDBNull(Me, CType(Nothing, Object)))
            Assert.AreEqual(Nothing, CompuMaster.Data.Utils.NoDBNull(DBNull.Value, CType(Nothing, Object)))
            Assert.AreEqual(Me, CompuMaster.Data.Utils.NoDBNull(DBNull.Value, CType(Me, Object)))

            'Value types incl. strings
            Assert.AreEqual(1, CompuMaster.Data.Utils.NoDBNull(1, -1))
            Assert.AreEqual(-1, CompuMaster.Data.Utils.NoDBNull(DBNull.Value, -1))
            Assert.AreEqual(1L, CompuMaster.Data.Utils.NoDBNull(1L, -1L))
            Assert.AreEqual(-1L, CompuMaster.Data.Utils.NoDBNull(DBNull.Value, -1L))
            Assert.AreEqual(1S, CompuMaster.Data.Utils.NoDBNull(1S, -1S))
            Assert.AreEqual(-1S, CompuMaster.Data.Utils.NoDBNull(DBNull.Value, -1S))
            Assert.AreEqual(1D, CompuMaster.Data.Utils.NoDBNull(1D, -1D))
            Assert.AreEqual(-1D, CompuMaster.Data.Utils.NoDBNull(DBNull.Value, -1D))
            Assert.AreEqual(1.0!, CompuMaster.Data.Utils.NoDBNull(1.0!, -1.0!))
            Assert.AreEqual(-1.0!, CompuMaster.Data.Utils.NoDBNull(DBNull.Value, -1.0!))
            Assert.AreEqual(1.0R, CompuMaster.Data.Utils.NoDBNull(1.0R, -1.0R))
            Assert.AreEqual(-1.0R, CompuMaster.Data.Utils.NoDBNull(DBNull.Value, -1.0R))
            Assert.AreEqual(True, CompuMaster.Data.Utils.NoDBNull(True, False))
            Assert.AreEqual(False, CompuMaster.Data.Utils.NoDBNull(DBNull.Value, False))
            Assert.AreEqual(CType(1, Byte), CompuMaster.Data.Utils.NoDBNull(CType(1, Byte), CType(2, Byte)))
            Assert.AreEqual(CType(2, Byte), CompuMaster.Data.Utils.NoDBNull(DBNull.Value, CType(2, Byte)))
            Assert.AreEqual(CType(1, UInt16), CompuMaster.Data.Utils.NoDBNull(CType(1, UInt16), CType(2, UInt16)))
            Assert.AreEqual(CType(2, UInt16), CompuMaster.Data.Utils.NoDBNull(DBNull.Value, CType(2, UInt16)))
            Assert.AreEqual(CType(1, UInt32), CompuMaster.Data.Utils.NoDBNull(CType(1, UInt32), CType(2, UInt32)))
            Assert.AreEqual(CType(2, UInt32), CompuMaster.Data.Utils.NoDBNull(DBNull.Value, CType(2, UInt32)))
            Assert.AreEqual(CType(1, UInt64), CompuMaster.Data.Utils.NoDBNull(CType(1, UInt64), CType(2, UInt64)))
            Assert.AreEqual(CType(2, UInt64), CompuMaster.Data.Utils.NoDBNull(DBNull.Value, CType(2, UInt64)))
            Assert.AreEqual(New DateTime(1), CompuMaster.Data.Utils.NoDBNull(New DateTime(1), New DateTime(2)))
            Assert.AreEqual(New DateTime(2), CompuMaster.Data.Utils.NoDBNull(DBNull.Value, New DateTime(2)))
            Assert.AreEqual("1", CompuMaster.Data.Utils.NoDBNull("1", "-1"))
            Assert.AreEqual("-1", CompuMaster.Data.Utils.NoDBNull(DBNull.Value, "-1"))

            'Generics
            Assert.AreEqual(1, CompuMaster.Data.Utils.NoDBNull(Of Long?)(New Long?(1)))
            Assert.AreEqual(1, CompuMaster.Data.Utils.NoDBNull(Of Long)(1))
            Assert.AreEqual(1, CompuMaster.Data.Utils.NoDBNull(Of Long?)(1L))
            Assert.AreEqual(1, CompuMaster.Data.Utils.NoDBNull(Of Long?)(1))
            'Assert.Catch(Of InvalidCastException)(Sub() CompuMaster.Data.Utils.NoDBNull(Of Long?)(1))
            Assert.AreEqual("Text", CompuMaster.Data.Utils.NoDBNull(Of String)("Text"))
            Assert.AreEqual(True, CompuMaster.Data.Utils.NoDBNull(Of Boolean?)(True))
            Assert.AreEqual(New Byte() {1}, CompuMaster.Data.Utils.NoDBNull(Of Byte())(New Byte() {1}))
            Assert.AreEqual(1D, CompuMaster.Data.Utils.NoDBNull(Of Decimal?)(1D))
            Assert.AreEqual(New DateTime(1, 1, 1), CompuMaster.Data.Utils.NoDBNull(Of DateTime?)(New DateTime(1, 1, 1)))
            Assert.AreEqual(Nothing, CompuMaster.Data.Utils.NoDBNull(Of Long?)(DBNull.Value))
            Assert.AreEqual(Nothing, CompuMaster.Data.Utils.NoDBNull(Of String)(DBNull.Value))
            Assert.AreEqual(Nothing, CompuMaster.Data.Utils.NoDBNull(Of Boolean?)(DBNull.Value))
            Assert.AreEqual(Nothing, CompuMaster.Data.Utils.NoDBNull(Of Byte())(DBNull.Value))
            Assert.AreEqual(Nothing, CompuMaster.Data.Utils.NoDBNull(Of Decimal?)(DBNull.Value))
            Assert.AreEqual(Nothing, CompuMaster.Data.Utils.NoDBNull(Of DateTime?)(DBNull.Value))
            Assert.AreEqual(False, CompuMaster.Data.Utils.NoDBNull(Of Long?)(DBNull.Value).HasValue)
            Assert.AreEqual(False, CompuMaster.Data.Utils.NoDBNull(Of Boolean?)(DBNull.Value).HasValue)
            Assert.AreEqual(False, CompuMaster.Data.Utils.NoDBNull(Of Decimal?)(DBNull.Value).HasValue)
            Assert.AreEqual(False, CompuMaster.Data.Utils.NoDBNull(Of DateTime?)(DBNull.Value).HasValue)
            Assert.AreEqual(CType(1, UInt16), CompuMaster.Data.Utils.NoDBNull(Of UInt16)(CType(1, UInt16)))
            Assert.AreEqual(CType(Nothing, UInt16), CompuMaster.Data.Utils.NoDBNull(Of UInt16)(DBNull.Value))
            Assert.AreEqual(CType(1, UInt16), CompuMaster.Data.Utils.NoDBNull(Of UInt16)(CType(1, UInt16), CType(2, UInt16)))
            Assert.AreEqual(CType(2, UInt16), CompuMaster.Data.Utils.NoDBNull(Of UInt16)(DBNull.Value, CType(2, UInt16)))
            'Nullable Generics
            Assert.AreEqual(CType(1, UInt16?), CompuMaster.Data.Utils.NoDBNull(Of UInt16?)(CType(1, UInt16)))
            Assert.AreEqual(CType(Nothing, UInt16?), CompuMaster.Data.Utils.NoDBNull(Of UInt16?)(DBNull.Value))
            Assert.AreEqual(New UInt16?, CompuMaster.Data.Utils.NoDBNull(Of UInt16?)(DBNull.Value))
            Assert.AreEqual(CType(1, UInt16?), CompuMaster.Data.Utils.NoDBNull(Of UInt16?)(CType(1, UInt16), CType(2, UInt16)))
            Assert.AreEqual(CType(2, UInt16?), CompuMaster.Data.Utils.NoDBNull(Of UInt16?)(DBNull.Value, CType(2, UInt16)))

        End Sub

        <Test> <CodeAnalysis.SuppressMessage("Performance", "CA1825:Avoid zero-length array allocations.", Justification:="<Ausstehend>")>
        Public Sub NoDBNullArrayOrListFromString_Generics()
            'DBNull.Value
            Assert.AreEqual(Nothing, CompuMaster.Data.Utils.NoDBNullArrayFromString(Of String)(DBNull.Value, ","))
            Assert.AreEqual(Nothing, CompuMaster.Data.Utils.NoDBNullArrayFromString(Of Object)(DBNull.Value, ","))
            Assert.AreEqual(Nothing, CompuMaster.Data.Utils.NoDBNullArrayFromString(Of Integer)(DBNull.Value, ","))
            Assert.AreEqual(Nothing, CompuMaster.Data.Utils.NoDBNullArrayFromString(Of Integer?)(DBNull.Value, ","))
            'Empty String
            Assert.AreEqual(New String() {}, CompuMaster.Data.Utils.NoDBNullArrayFromString(Of String)("", ","))
            Assert.AreEqual(New Object() {}, CompuMaster.Data.Utils.NoDBNullArrayFromString(Of Object)("", ","))
            Assert.AreEqual(New Integer() {}, CompuMaster.Data.Utils.NoDBNullArrayFromString(Of Integer)("", ","))
            Assert.AreEqual(New Integer?() {}, CompuMaster.Data.Utils.NoDBNullArrayFromString(Of Integer?)("", ","))
            'Single String
            Assert.AreEqual(New String() {"Test"}, CompuMaster.Data.Utils.NoDBNullArrayFromString(Of String)("Test", ","))
            'Assert.AreEqual(New Object() {New Object}, CompuMaster.Data.Utils.NoDBNullArrayFromString(Of Object)(New Object.ToString, ","))
            Assert.AreEqual(New Integer() {1}, CompuMaster.Data.Utils.NoDBNullArrayFromString(Of Integer)("1", ","))
            Assert.AreEqual(New Integer?() {1}, CompuMaster.Data.Utils.NoDBNullArrayFromString(Of Integer?)("1", ","))
            'Separatable String
            Assert.AreEqual(New String() {"Test1", "Test2", "Test3"}, CompuMaster.Data.Utils.NoDBNullArrayFromString(Of String)("Test1,Test2,Test3", ","))
            'Assert.AreEqual(New Object() {}, CompuMaster.Data.Utils.NoDBNullArrayFromString(Of Object)("", ","))
            Assert.AreEqual(New Integer() {1, 2, 3}, CompuMaster.Data.Utils.NoDBNullArrayFromString(Of Integer)("1,2,3", ","))
            Assert.AreEqual(New Integer?() {1, 2, 3}, CompuMaster.Data.Utils.NoDBNullArrayFromString(Of Integer?)("1,2,3", ","))
        End Sub

        <Test> <CodeAnalysis.SuppressMessage("Performance", "CA1825:Avoid zero-length array allocations.", Justification:="<Ausstehend>")>
        Public Sub NoDBNullArrayOrListFromString_NoGenerics()
            'DBNull.Value
            Assert.AreEqual(Nothing, CompuMaster.Data.Utils.NoDBNullArrayFromString(DBNull.Value, ",", CType(Nothing, String())))
            Assert.AreEqual(Nothing, CompuMaster.Data.Utils.NoDBNullArrayFromString(DBNull.Value, ",", CType(Nothing, Object())))
            Assert.AreEqual(Nothing, CompuMaster.Data.Utils.NoDBNullArrayFromString(DBNull.Value, ",", CType(Nothing, Integer())))
            'Assert.AreEqual(Nothing, CompuMaster.Data.Utils.NoDBNullArrayFromString(DBNull.Value, ",", CType(Nothing, Integer?())))
            'Empty String
            Assert.AreEqual(New String() {}, CompuMaster.Data.Utils.NoDBNullArrayFromString("", ",", CType(Nothing, String())))
            Assert.AreEqual(New Object() {}, CompuMaster.Data.Utils.NoDBNullArrayFromString("", ",", CType(Nothing, Object())))
            Assert.AreEqual(New Integer() {}, CompuMaster.Data.Utils.NoDBNullArrayFromString("", ",", CType(Nothing, Integer())))
            'Assert.AreEqual(New Integer?() {}, CompuMaster.Data.Utils.NoDBNullArrayFromString("", ",", CType(Nothing, Integer?())))
            'Single String
            Assert.AreEqual(New String() {"Test"}, CompuMaster.Data.Utils.NoDBNullArrayFromString("Test", ",", CType(Nothing, String())))
            'Assert.AreEqual(New Object() {New Object}, CompuMaster.Data.Utils.NoDBNullArrayFromString(New Object.ToString, ",", CType(Nothing, Object())))
            Assert.AreEqual(New Integer() {1}, CompuMaster.Data.Utils.NoDBNullArrayFromString("1", ",", CType(Nothing, Integer())))
            'Assert.AreEqual(New Integer?() {1}, CompuMaster.Data.Utils.NoDBNullArrayFromString("1", ",", CType(Nothing, Integer?())))
            'Separatable String
            Assert.AreEqual(New String() {"Test1", "Test2", "Test3"}, CompuMaster.Data.Utils.NoDBNullArrayFromString("Test1,Test2,Test3", ",", CType(Nothing, String())))
            'Assert.AreEqual(New Object() {}, CompuMaster.Data.Utils.NoDBNullArrayFromString("", ",", CType(Nothing, Object())))
            Assert.AreEqual(New Integer() {1, 2, 3}, CompuMaster.Data.Utils.NoDBNullArrayFromString("1,2,3", ",", CType(Nothing, Integer())))
            'Assert.AreEqual(New Integer?() {1, 2, 3}, CompuMaster.Data.Utils.NoDBNullArrayFromString("1,2,3", ",", CType(Nothing, Integer?())))
        End Sub

        <Test> Public Sub NullableTypeWithItsValueOrDBNull()
            Assert.AreNotEqual(GetType(Integer?), CompuMaster.Data.Utils.NullableTypeWithItsValueOrDBNull(Of Integer)(New Integer?(1)).GetType)
            Assert.AreEqual(GetType(Integer), CompuMaster.Data.Utils.NullableTypeWithItsValueOrDBNull(Of Integer)(New Integer?(1)).GetType)
            Assert.AreEqual(1, CompuMaster.Data.Utils.NullableTypeWithItsValueOrDBNull(Of Integer)(New Integer?(1)))
            Assert.AreEqual(1, CompuMaster.Data.Utils.NullableTypeWithItsValueOrDBNull(Of Integer)(1))
            Assert.AreEqual(DBNull.Value, CompuMaster.Data.Utils.NullableTypeWithItsValueOrDBNull(Of Integer)(New Integer?))
            Assert.AreEqual(DBNull.Value, CompuMaster.Data.Utils.NullableTypeWithItsValueOrDBNull(Of Integer)(Nothing))
        End Sub

        <Test> <CodeAnalysis.SuppressMessage("Performance", "CA1825:Avoid zero-length array allocations.", Justification:="<Ausstehend>")>
        Sub ArrayNotNothingOrDBNull()
            Assert.AreEqual(DBNull.Value, CompuMaster.Data.Utils.ArrayNotNothingOrDBNull(Nothing))
            Dim arr As Byte() = Nothing
            Assert.AreEqual(DBNull.Value, CompuMaster.Data.Utils.ArrayNotNothingOrDBNull(arr))
            Assert.AreEqual(New Byte() {}, CompuMaster.Data.Utils.ArrayNotNothingOrDBNull(New Byte() {}))
            Assert.AreEqual(New Byte() {1}, CompuMaster.Data.Utils.ArrayNotNothingOrDBNull(New Byte() {1}))
        End Sub

        <Test> <CodeAnalysis.SuppressMessage("Performance", "CA1825:Avoid zero-length array allocations.", Justification:="<Ausstehend>")>
        Sub ArrayNotEmptyOrDBNull()
            Assert.AreEqual(DBNull.Value, CompuMaster.Data.Utils.ArrayNotEmptyOrDBNull(Nothing))
            Dim arr As Byte() = Nothing
            Assert.AreEqual(DBNull.Value, CompuMaster.Data.Utils.ArrayNotEmptyOrDBNull(arr))
            Assert.AreEqual(DBNull.Value, CompuMaster.Data.Utils.ArrayNotEmptyOrDBNull(New Byte() {}))
            Assert.AreEqual(New Byte() {1}, CompuMaster.Data.Utils.ArrayNotEmptyOrDBNull(New Byte() {1}))
        End Sub

    End Class

End Namespace
