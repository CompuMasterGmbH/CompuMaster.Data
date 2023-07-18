Option Explicit On 
Option Strict On

Imports CompuMaster.Data.Information
Imports CompuMaster.Data.Strings

Namespace CompuMaster.Data

    ''' <summary>
    ''' Utils for converting and handling database data
    ''' </summary>
    Public NotInheritable Class Utils

        ''' <summary>
        ''' A triple state defaulting to Undefined
        ''' </summary>
        Friend Enum TripleState As Byte
            Undefined = 0
            [True] = 1
            [False] = 2
        End Enum

#Region "NoDBNull"
        ''' <summary>
        '''     Checks for DBNull and returns null (Nothing in VisualBasic) in that case
        ''' </summary>
        ''' <param name="checkValueIfDBNull">The value to be checked</param>
        ''' <returns>A value which is not DBNull; a DBNull as input will return null (Nothing in VisualBasic)</returns>
        <DebuggerHidden()> Public Shared Function NoDBNull(ByVal checkValueIfDBNull As Object) As Object
            If IsDBNull(checkValueIfDBNull) Then
                Return Nothing
            Else
                Return (checkValueIfDBNull)
            End If
        End Function

        ''' <summary>
        '''     Checks for DBNull and returns null (Nothing in VisualBasic) in that case
        ''' </summary>
        ''' <param name="replaceWithThis">The value to be checked</param>
        ''' <returns>A value which is not DBNull; a DBNull as input will return null (Nothing in VisualBasic)</returns>
        ''' <remarks>
        ''' </remarks>
        <DebuggerHidden()> Public Shared Function NoDBNull(ByVal value As Object, ByVal replaceWithThis As Char) As Char
            If IsDBNull(value) Then
                Return replaceWithThis
            Else
                Return CType(value, Char)
            End If
        End Function

        ''' <summary>
        '''     Checks for DBNull and returns the second value alternatively
        ''' </summary>
        ''' <param name="checkValueIfDBNull">The value to be checked</param>
        ''' <param name="replaceWithThis">The alternative value, null (Nothing in VisualBasic) if not defined</param>
        ''' <returns>A value which is not DBNull</returns>
        <DebuggerHidden()> Public Shared Function NoDBNull(ByVal checkValueIfDBNull As Object, ByVal replaceWithThis As Object) As Object
            If IsDBNull(checkValueIfDBNull) Then
                Return (replaceWithThis)
            Else
                Return (checkValueIfDBNull)
            End If
        End Function

        ''' <summary>
        '''     Checks for DBNull and returns the second value alternatively
        ''' </summary>
        ''' <param name="checkValueIfDBNull">The value to be checked</param>
        ''' <param name="replaceWithThis">The alternative value, null (Nothing in VisualBasic) if not defined</param>
        ''' <returns>A value which is not DBNull</returns>
        <DebuggerHidden()> Public Shared Function NoDBNull(ByVal checkValueIfDBNull As Object, ByVal replaceWithThis As Integer) As Integer
            If IsDBNull(checkValueIfDBNull) Then
                Return (replaceWithThis)
            Else
                Return CType(checkValueIfDBNull, Integer)
            End If
        End Function

        ''' <summary>
        '''     Checks for DBNull and returns the second value alternatively
        ''' </summary>
        ''' <param name="checkValueIfDBNull">The value to be checked</param>
        ''' <param name="replaceWithThis">The alternative value, null (Nothing in VisualBasic) if not defined</param>
        ''' <returns>A value which is not DBNull</returns>
        <DebuggerHidden()> Public Shared Function NoDBNull(ByVal checkValueIfDBNull As Object, ByVal replaceWithThis As Long) As Long
            If IsDBNull(checkValueIfDBNull) Then
                Return (replaceWithThis)
            Else
                Return CType(checkValueIfDBNull, Long)
            End If
        End Function

        ''' <summary>
        '''     Checks for DBNull and returns the second value alternatively
        ''' </summary>
        ''' <param name="checkValueIfDBNull">The value to be checked</param>
        ''' <param name="replaceWithThis">The alternative value, null (Nothing in VisualBasic) if not defined</param>
        ''' <returns>A value which is not DBNull</returns>
        <DebuggerHidden()> Public Shared Function NoDBNull(ByVal checkValueIfDBNull As Object, ByVal replaceWithThis As Decimal) As Decimal
            If IsDBNull(checkValueIfDBNull) Then
                Return (replaceWithThis)
            Else
                Return CType(checkValueIfDBNull, Decimal)
            End If
        End Function

        ''' <summary>
        '''     Checks for DBNull and returns the second value alternatively
        ''' </summary>
        ''' <param name="checkValueIfDBNull">The value to be checked</param>
        ''' <param name="replaceWithThis">The alternative value, null (Nothing in VisualBasic) if not defined</param>
        ''' <returns>A value which is not DBNull</returns>
        <DebuggerHidden()> Public Shared Function NoDBNull(ByVal checkValueIfDBNull As Object, ByVal replaceWithThis As Short) As Short
            If IsDBNull(checkValueIfDBNull) Then
                Return (replaceWithThis)
            Else
                Return CType(checkValueIfDBNull, Short)
            End If
        End Function

        ''' <summary>
        '''     Checks for DBNull and returns the second value alternatively
        ''' </summary>
        ''' <param name="checkValueIfDBNull">The value to be checked</param>
        ''' <param name="replaceWithThis">The alternative value, null (Nothing in VisualBasic) if not defined</param>
        ''' <returns>A value which is not DBNull</returns>
        <DebuggerHidden()> Public Shared Function NoDBNull(ByVal checkValueIfDBNull As Object, ByVal replaceWithThis As Single) As Single
            If IsDBNull(checkValueIfDBNull) Then
                Return (replaceWithThis)
            Else
                Return CType(checkValueIfDBNull, Single)
            End If
        End Function

        ''' <summary>
        '''     Checks for DBNull and returns the second value alternatively
        ''' </summary>
        ''' <param name="checkValueIfDBNull">The value to be checked</param>
        ''' <param name="replaceWithThis">The alternative value, null (Nothing in VisualBasic) if not defined</param>
        ''' <returns>A value which is not DBNull</returns>
        <DebuggerHidden()> Public Shared Function NoDBNull(ByVal checkValueIfDBNull As Object, ByVal replaceWithThis As Boolean) As Boolean
            If IsDBNull(checkValueIfDBNull) Then
                Return (replaceWithThis)
            Else
                Return CType(checkValueIfDBNull, Boolean)
            End If
        End Function

        ''' <summary>
        '''     Checks for DBNull and returns the second value alternatively
        ''' </summary>
        ''' <param name="checkValueIfDBNull">The value to be checked</param>
        ''' <param name="replaceWithThis">The alternative value, null (Nothing in VisualBasic) if not defined</param>
        ''' <returns>A value which is not DBNull</returns>
        <DebuggerHidden()> Public Shared Function NoDBNull(ByVal checkValueIfDBNull As Object, ByVal replaceWithThis As DateTime) As DateTime
            If IsDBNull(checkValueIfDBNull) Then
                Return (replaceWithThis)
            Else
                Return CType(checkValueIfDBNull, DateTime)
            End If
        End Function

        ''' <summary>
        '''     Checks for DBNull and returns the second value alternatively
        ''' </summary>
        ''' <param name="checkValueIfDBNull">The value to be checked</param>
        ''' <param name="replaceWithThis">The alternative value, null (Nothing in VisualBasic) if not defined</param>
        ''' <returns>A value which is not DBNull</returns>
        <DebuggerHidden()> Public Shared Function NoDBNull(ByVal checkValueIfDBNull As Object, ByVal replaceWithThis As Double) As Double
            If IsDBNull(checkValueIfDBNull) Then
                Return (replaceWithThis)
            Else
                Return CType(checkValueIfDBNull, Double)
            End If
        End Function
        ''' <summary>
        '''     Checks for DBNull and returns the second value alternatively
        ''' </summary>
        ''' <param name="checkValueIfDBNull">The value to be checked</param>
        ''' <param name="replaceWithThis">The alternative value, null (Nothing in VisualBasic) if not defined</param>
        ''' <returns>A value which is not DBNull</returns>
        <DebuggerHidden()> Public Shared Function NoDBNull(ByVal checkValueIfDBNull As Object, ByVal replaceWithThis As Byte) As Byte
            If IsDBNull(checkValueIfDBNull) Then
                Return (replaceWithThis)
            Else
                Return CType(checkValueIfDBNull, Byte)
            End If
        End Function

        ''' <summary>
        '''     Checks for DBNull and returns null (Nothing in VisualBasic) in that case
        ''' </summary>
        ''' <param name="replaceWithThis">The value to be checked</param>
        ''' <returns>A value which is not DBNull; a DBNull as input will return null (Nothing in VisualBasic)</returns>
        ''' <remarks>
        ''' </remarks>
        <DebuggerHidden()> Public Shared Function NoDBNull(ByVal value As Object, ByVal replaceWithThis As Char?) As Char?
            If IsDBNull(value) Then
                Return replaceWithThis
            Else
                Return New Char?(CType(value, Char))
            End If
        End Function

        ''' <summary>
        '''     Checks for DBNull and returns the second value alternatively
        ''' </summary>
        ''' <param name="checkValueIfDBNull">The value to be checked</param>
        ''' <param name="replaceWithThis">The alternative value, null (Nothing in VisualBasic) if not defined</param>
        ''' <returns>A value which is not DBNull</returns>
        <DebuggerHidden()> Public Shared Function NoDBNull(ByVal checkValueIfDBNull As Object, ByVal replaceWithThis As Integer?) As Integer?
            If IsDBNull(checkValueIfDBNull) Then
                Return (replaceWithThis)
            Else
                Return New Integer?(CType(checkValueIfDBNull, Integer))
            End If
        End Function

        ''' <summary>
        '''     Checks for DBNull and returns the second value alternatively
        ''' </summary>
        ''' <param name="checkValueIfDBNull">The value to be checked</param>
        ''' <param name="replaceWithThis">The alternative value, null (Nothing in VisualBasic) if not defined</param>
        ''' <returns>A value which is not DBNull</returns>
        <DebuggerHidden()> Public Shared Function NoDBNull(ByVal checkValueIfDBNull As Object, ByVal replaceWithThis As Long?) As Long?
            If IsDBNull(checkValueIfDBNull) Then
                Return (replaceWithThis)
            Else
                Return New Long?(CType(checkValueIfDBNull, Long))
            End If
        End Function

        ''' <summary>
        '''     Checks for DBNull and returns the second value alternatively
        ''' </summary>
        ''' <param name="checkValueIfDBNull">The value to be checked</param>
        ''' <param name="replaceWithThis">The alternative value, null (Nothing in VisualBasic) if not defined</param>
        ''' <returns>A value which is not DBNull</returns>
        <DebuggerHidden()> Public Shared Function NoDBNull(ByVal checkValueIfDBNull As Object, ByVal replaceWithThis As Decimal?) As Decimal?
            If IsDBNull(checkValueIfDBNull) Then
                Return (replaceWithThis)
            Else
                Return New Decimal?(CType(checkValueIfDBNull, Decimal))
            End If
        End Function

        ''' <summary>
        '''     Checks for DBNull and returns the second value alternatively
        ''' </summary>
        ''' <param name="checkValueIfDBNull">The value to be checked</param>
        ''' <param name="replaceWithThis">The alternative value, null (Nothing in VisualBasic) if not defined</param>
        ''' <returns>A value which is not DBNull</returns>
        <DebuggerHidden()> Public Shared Function NoDBNull(ByVal checkValueIfDBNull As Object, ByVal replaceWithThis As Short?) As Short?
            If IsDBNull(checkValueIfDBNull) Then
                Return (replaceWithThis)
            Else
                Return New Short?(CType(checkValueIfDBNull, Short))
            End If
        End Function

        ''' <summary>
        '''     Checks for DBNull and returns the second value alternatively
        ''' </summary>
        ''' <param name="checkValueIfDBNull">The value to be checked</param>
        ''' <param name="replaceWithThis">The alternative value, null (Nothing in VisualBasic) if not defined</param>
        ''' <returns>A value which is not DBNull</returns>
        <DebuggerHidden()> Public Shared Function NoDBNull(ByVal checkValueIfDBNull As Object, ByVal replaceWithThis As Single?) As Single?
            If IsDBNull(checkValueIfDBNull) Then
                Return (replaceWithThis)
            Else
                Return New Single?(CType(checkValueIfDBNull, Single))
            End If
        End Function

        ''' <summary>
        '''     Checks for DBNull and returns the second value alternatively
        ''' </summary>
        ''' <param name="checkValueIfDBNull">The value to be checked</param>
        ''' <param name="replaceWithThis">The alternative value, null (Nothing in VisualBasic) if not defined</param>
        ''' <returns>A value which is not DBNull</returns>
        <DebuggerHidden()> Public Shared Function NoDBNull(ByVal checkValueIfDBNull As Object, ByVal replaceWithThis As Boolean?) As Boolean?
            If IsDBNull(checkValueIfDBNull) Then
                Return (replaceWithThis)
            Else
                Return New Boolean?(CType(checkValueIfDBNull, Boolean))
            End If
        End Function

        ''' <summary>
        '''     Checks for DBNull and returns the second value alternatively
        ''' </summary>
        ''' <param name="checkValueIfDBNull">The value to be checked</param>
        ''' <param name="replaceWithThis">The alternative value, null (Nothing in VisualBasic) if not defined</param>
        ''' <returns>A value which is not DBNull</returns>
        <DebuggerHidden()> Public Shared Function NoDBNull(ByVal checkValueIfDBNull As Object, ByVal replaceWithThis As DateTime?) As DateTime?
            If IsDBNull(checkValueIfDBNull) Then
                Return (replaceWithThis)
            Else
                Return New DateTime?(CType(checkValueIfDBNull, DateTime))
            End If
        End Function

        ''' <summary>
        '''     Checks for DBNull and returns the second value alternatively
        ''' </summary>
        ''' <param name="checkValueIfDBNull">The value to be checked</param>
        ''' <param name="replaceWithThis">The alternative value, null (Nothing in VisualBasic) if not defined</param>
        ''' <returns>A value which is not DBNull</returns>
        <DebuggerHidden()> Public Shared Function NoDBNull(ByVal checkValueIfDBNull As Object, ByVal replaceWithThis As Double?) As Double?
            If IsDBNull(checkValueIfDBNull) Then
                Return (replaceWithThis)
            Else
                Return New Double?(CType(checkValueIfDBNull, Double))
            End If
        End Function
        ''' <summary>
        '''     Checks for DBNull and returns the second value alternatively
        ''' </summary>
        ''' <param name="checkValueIfDBNull">The value to be checked</param>
        ''' <param name="replaceWithThis">The alternative value, null (Nothing in VisualBasic) if not defined</param>
        ''' <returns>A value which is not DBNull</returns>
        <DebuggerHidden()> Public Shared Function NoDBNull(ByVal checkValueIfDBNull As Object, ByVal replaceWithThis As Byte?) As Byte?
            If IsDBNull(checkValueIfDBNull) Then
                Return (replaceWithThis)
            Else
                Return New Byte?(CType(checkValueIfDBNull, Byte))
            End If
        End Function

        ''' <summary>
        '''     Checks for DBNull and returns the second value alternatively
        ''' </summary>
        ''' <param name="checkValueIfDBNull">The value to be checked</param>
        ''' <param name="replaceWithThis">The alternative value, null (Nothing in VisualBasic) if not defined</param>
        ''' <returns>A value which is not DBNull</returns>
        <DebuggerHidden()> Public Shared Function NoDBNull(ByVal checkValueIfDBNull As Object, ByVal replaceWithThis As Byte()) As Byte()
            If IsDBNull(checkValueIfDBNull) Then
                Return (replaceWithThis)
            Else
                Return CType(checkValueIfDBNull, Byte())
            End If
        End Function

        ''' <summary>
        '''     Checks for DBNull and returns the second value alternatively
        ''' </summary>
        ''' <param name="checkValueIfDBNull">The value to be checked</param>
        ''' <param name="replaceWithThis">The alternative value, null (Nothing in VisualBasic) if not defined</param>
        ''' <returns>A value which is not DBNull</returns>
        <DebuggerHidden()> Public Shared Function NoDBNull(ByVal checkValueIfDBNull As Object, ByVal replaceWithThis As String) As String
            If IsDBNull(checkValueIfDBNull) Then
                Return (replaceWithThis)
            Else
                Return CType(checkValueIfDBNull, String)
            End If
        End Function

        ''' <summary>
        '''     Checks for DBNull and returns null (Nothing in VisualBasic) alternatively
        ''' </summary>
        ''' <param name="checkValueIfDBNull">The value to be checked</param>
        ''' <returns>A value which is not DBNull, otherwise null (Nothing in VisualBasic)</returns>
        <DebuggerHidden()> Public Shared Function NoDBNull(Of T)(ByVal checkValueIfDBNull As Object) As T
            If IsDBNull(checkValueIfDBNull) Then
                Return Nothing
            ElseIf checkValueIfDBNull Is Nothing Then
                Return CType(Nothing, T)
            ElseIf Nullable.GetUnderlyingType(GetType(T)) IsNot Nothing AndAlso checkValueIfDBNull.GetType.IsValueType AndAlso Nullable.GetUnderlyingType(GetType(T)) IsNot checkValueIfDBNull.GetType Then
                'Dim UnderlyingType As Type = Nullable.GetUnderlyingType(GetType(T))
                Dim Result As T = CType(Activator.CreateInstance(GetType(T), checkValueIfDBNull), T)
                Return Result
            Else
                Return CType(checkValueIfDBNull, T)
            End If
        End Function

        ''' <summary>
        '''     Checks for DBNull and returns the second value alternatively
        ''' </summary>
        ''' <param name="checkValueIfDBNull">The value to be checked</param>
        ''' <param name="replaceWithThis">The alternative value, null (Nothing in VisualBasic) if not defined</param>
        ''' <returns>A value which is not DBNull</returns>
        <DebuggerHidden()> Public Shared Function NoDBNull(Of T)(ByVal checkValueIfDBNull As Object, ByVal replaceWithThis As T) As T
            If IsDBNull(checkValueIfDBNull) Then
                Return (replaceWithThis)
            Else
                Return CType(checkValueIfDBNull, T)
            End If
        End Function

        ''' <summary>
        ''' Check for DBNull and return null (Nothing in VisualBasic) alternatively, second check for empty string and return empty array, third split the string and fill the array with all elements
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="arrayData"></param>
        ''' <param name="splitChar"></param>
        ''' <returns></returns>
        Public Shared Function NoDBNullArrayFromString(Of T)(ByVal arrayData As Object, splitChar As Char) As T()
            Dim ListResult As Generic.List(Of T) = NoDBNullListFromString(Of T)(arrayData, splitChar)
            If ListResult Is Nothing Then
                Return Nothing
            Else
                Return ListResult.ToArray
            End If
        End Function

        ''' <summary>
        ''' Check for DBNull and return null (Nothing in VisualBasic) alternatively, second check for empty string and return empty array, third split the string and fill the array with all elements
        ''' </summary>
        ''' <param name="arrayData"></param>
        ''' <param name="splitChar"></param>
        ''' <param name="alternativeValue"></param>
        ''' <returns></returns>
        Public Shared Function NoDBNullArrayFromString(ByVal arrayData As Object, splitChar As Char, alternativeValue As String()) As String()
            Dim ListResult As Generic.List(Of String)
            If alternativeValue Is Nothing Then
                ListResult = NoDBNullListFromString(arrayData, splitChar, CType(Nothing, Generic.List(Of String)))
            Else
                ListResult = NoDBNullListFromString(arrayData, splitChar, New Generic.List(Of String)(alternativeValue))
            End If
            If ListResult Is Nothing Then
                Return alternativeValue
            Else
                Return ListResult.ToArray
            End If
        End Function

        ''' <summary>
        ''' Check for DBNull and return null (Nothing in VisualBasic) alternatively, second check for empty string and return empty array, third split the string and fill the array with all elements
        ''' </summary>
        ''' <param name="arrayData"></param>
        ''' <param name="splitChar"></param>
        ''' <param name="alternativeValue"></param>
        ''' <returns></returns>
        Public Shared Function NoDBNullArrayFromString(ByVal arrayData As Object, splitChar As Char, alternativeValue As Boolean()) As Boolean()
            Dim ListResult As Generic.List(Of Boolean)
            If alternativeValue Is Nothing Then
                ListResult = NoDBNullListFromString(arrayData, splitChar, CType(Nothing, Generic.List(Of Boolean)))
            Else
                ListResult = NoDBNullListFromString(arrayData, splitChar, New Generic.List(Of Boolean)(alternativeValue))
            End If
            If ListResult Is Nothing Then
                Return alternativeValue
            Else
                Return ListResult.ToArray
            End If
        End Function

        ''' <summary>
        ''' Check for DBNull and return null (Nothing in VisualBasic) alternatively, second check for empty string and return empty array, third split the string and fill the array with all elements
        ''' </summary>
        ''' <param name="arrayData"></param>
        ''' <param name="splitChar"></param>
        ''' <param name="alternativeValue"></param>
        ''' <returns></returns>
        Public Shared Function NoDBNullArrayFromString(ByVal arrayData As Object, splitChar As Char, alternativeValue As Byte()) As Byte()
            Dim ListResult As Generic.List(Of Byte)
            If alternativeValue Is Nothing Then
                ListResult = NoDBNullListFromString(arrayData, splitChar, CType(Nothing, Generic.List(Of Byte)))
            Else
                ListResult = NoDBNullListFromString(arrayData, splitChar, New Generic.List(Of Byte)(alternativeValue))
            End If
            If ListResult Is Nothing Then
                Return alternativeValue
            Else
                Return ListResult.ToArray
            End If
        End Function

        ''' <summary>
        ''' Check for DBNull and return null (Nothing in VisualBasic) alternatively, second check for empty string and return empty array, third split the string and fill the array with all elements
        ''' </summary>
        ''' <param name="arrayData"></param>
        ''' <param name="splitChar"></param>
        ''' <param name="alternativeValue"></param>
        ''' <returns></returns>
        Public Shared Function NoDBNullArrayFromString(ByVal arrayData As Object, splitChar As Char, alternativeValue As DateTime()) As DateTime()
            Dim ListResult As Generic.List(Of DateTime)
            If alternativeValue Is Nothing Then
                ListResult = NoDBNullListFromString(arrayData, splitChar, CType(Nothing, Generic.List(Of DateTime)))
            Else
                ListResult = NoDBNullListFromString(arrayData, splitChar, New Generic.List(Of DateTime)(alternativeValue))
            End If
            If ListResult Is Nothing Then
                Return alternativeValue
            Else
                Return ListResult.ToArray
            End If
        End Function

        ''' <summary>
        ''' Check for DBNull and return null (Nothing in VisualBasic) alternatively, second check for empty string and return empty array, third split the string and fill the array with all elements
        ''' </summary>
        ''' <param name="arrayData"></param>
        ''' <param name="splitChar"></param>
        ''' <param name="alternativeValue"></param>
        ''' <returns></returns>
        Public Shared Function NoDBNullArrayFromString(ByVal arrayData As Object, splitChar As Char, alternativeValue As Double()) As Double()
            Dim ListResult As Generic.List(Of Double)
            If alternativeValue Is Nothing Then
                ListResult = NoDBNullListFromString(arrayData, splitChar, CType(Nothing, Generic.List(Of Double)))
            Else
                ListResult = NoDBNullListFromString(arrayData, splitChar, New Generic.List(Of Double)(alternativeValue))
            End If
            If ListResult Is Nothing Then
                Return alternativeValue
            Else
                Return ListResult.ToArray
            End If
        End Function

        ''' <summary>
        ''' Check for DBNull and return null (Nothing in VisualBasic) alternatively, second check for empty string and return empty array, third split the string and fill the array with all elements
        ''' </summary>
        ''' <param name="arrayData"></param>
        ''' <param name="splitChar"></param>
        ''' <param name="alternativeValue"></param>
        ''' <returns></returns>
        Public Shared Function NoDBNullArrayFromString(ByVal arrayData As Object, splitChar As Char, alternativeValue As Integer()) As Integer()
            Dim ListResult As Generic.List(Of Integer)
            If alternativeValue Is Nothing Then
                ListResult = NoDBNullListFromString(arrayData, splitChar, CType(Nothing, Generic.List(Of Integer)))
            Else
                ListResult = NoDBNullListFromString(arrayData, splitChar, New Generic.List(Of Integer)(alternativeValue))
            End If
            If ListResult Is Nothing Then
                Return alternativeValue
            Else
                Return ListResult.ToArray
            End If
        End Function

        ''' <summary>
        ''' Check for DBNull and return null (Nothing in VisualBasic) alternatively, second check for empty string and return empty list, third split the string and fill the list with all elements
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="arrayData"></param>
        ''' <param name="splitChar"></param>
        ''' <returns></returns>
        Public Shared Function NoDBNullListFromString(Of T)(ByVal arrayData As Object, splitChar As Char) As Generic.List(Of T)
            If IsDBNull(arrayData) OrElse arrayData Is Nothing Then
                Return Nothing
            ElseIf CType(arrayData, String) = "" Then
                Return New Generic.List(Of T)
            Else
                Dim Result As New Generic.List(Of T)
                For Each SplitValue As String In CType(arrayData, String).Split(splitChar)
                    If Nullable.GetUnderlyingType(GetType(T)) IsNot Nothing AndAlso GetType(T).IsValueType AndAlso Nullable.GetUnderlyingType(GetType(T)) IsNot GetType(T) Then
                        Dim UnderlyingType As Type = Nullable.GetUnderlyingType(GetType(T)) 'e.g.: T = Nullable(Of Integer) -> UnderlyingType = Integer
                        Dim SplitValueOrNothing As String = StringNotEmptyOrNothing(SplitValue)
                        Dim SplitValueAsT As T = CType(Convert.ChangeType(SplitValueOrNothing, UnderlyingType, Threading.Thread.CurrentThread.CurrentCulture), T)
                        Result.Add(CType(Activator.CreateInstance(GetType(T), SplitValueAsT), T))
                    Else
                        Result.Add(CType(CType(SplitValue, Object), T))
                    End If
                Next
                Return Result
            End If
        End Function

        ''' <summary>
        ''' Check for DBNull and return null (Nothing in VisualBasic) alternatively, second check for empty string and return empty list, third split the string and fill the list with all elements
        ''' </summary>
        ''' <param name="arrayData"></param>
        ''' <param name="splitChar"></param>
        ''' <param name="alternativeValue"></param>
        ''' <returns></returns>
        Public Shared Function NoDBNullListFromString(ByVal arrayData As Object, splitChar As Char, alternativeValue As Generic.List(Of DateTime)) As Generic.List(Of DateTime)
            If IsDBNull(arrayData) OrElse arrayData Is Nothing Then
                Return alternativeValue
            ElseIf CType(arrayData, String) = "" Then
                Return New Generic.List(Of DateTime)
            Else
                Dim Result As New Generic.List(Of DateTime)
                For Each SplitValue As String In CType(arrayData, String).Split(splitChar)
                    If Nullable.GetUnderlyingType(GetType(DateTime)) IsNot Nothing AndAlso GetType(DateTime).IsValueType AndAlso Nullable.GetUnderlyingType(GetType(DateTime)) IsNot GetType(DateTime) Then
                        Dim UnderlyingType As Type = Nullable.GetUnderlyingType(GetType(DateTime)) 'e.g.: T = Nullable(Of Integer) -> UnderlyingType = Integer
                        Dim SplitValueOrNothing As String = StringNotEmptyOrNothing(SplitValue)
                        Dim SplitValueAsT As DateTime = CType(Convert.ChangeType(SplitValueOrNothing, UnderlyingType, Threading.Thread.CurrentThread.CurrentCulture), DateTime)
                        Result.Add(CType(Activator.CreateInstance(GetType(DateTime), SplitValueAsT), DateTime))
                    Else
                        Result.Add(CType(CType(SplitValue, Object), DateTime))
                    End If
                Next
                Return Result
            End If
        End Function

        ''' <summary>
        ''' Check for DBNull and return null (Nothing in VisualBasic) alternatively, second check for empty string and return empty list, third split the string and fill the list with all elements
        ''' </summary>
        ''' <param name="arrayData"></param>
        ''' <param name="splitChar"></param>
        ''' <param name="alternativeValue"></param>
        ''' <returns></returns>
        Public Shared Function NoDBNullListFromString(ByVal arrayData As Object, splitChar As Char, alternativeValue As Generic.List(Of Boolean)) As Generic.List(Of Boolean)
            If IsDBNull(arrayData) OrElse arrayData Is Nothing Then
                Return alternativeValue
            ElseIf CType(arrayData, String) = "" Then
                Return New Generic.List(Of Boolean)
            Else
                Dim Result As New Generic.List(Of Boolean)
                For Each SplitValue As String In CType(arrayData, String).Split(splitChar)
                    If Nullable.GetUnderlyingType(GetType(Boolean)) IsNot Nothing AndAlso GetType(Boolean).IsValueType AndAlso Nullable.GetUnderlyingType(GetType(Boolean)) IsNot GetType(Boolean) Then
                        Dim UnderlyingType As Type = Nullable.GetUnderlyingType(GetType(Boolean)) 'e.g.: T = Nullable(Of Integer) -> UnderlyingType = Integer
                        Dim SplitValueOrNothing As String = StringNotEmptyOrNothing(SplitValue)
                        Dim SplitValueAsT As Boolean = CType(Convert.ChangeType(SplitValueOrNothing, UnderlyingType, Threading.Thread.CurrentThread.CurrentCulture), Boolean)
                        Result.Add(CType(Activator.CreateInstance(GetType(Boolean), SplitValueAsT), Boolean))
                    Else
                        Result.Add(CType(CType(SplitValue, Object), Boolean))
                    End If
                Next
                Return Result
            End If
        End Function

        ''' <summary>
        ''' Check for DBNull and return null (Nothing in VisualBasic) alternatively, second check for empty string and return empty list, third split the string and fill the list with all elements
        ''' </summary>
        ''' <param name="arrayData"></param>
        ''' <param name="splitChar"></param>
        ''' <param name="alternativeValue"></param>
        ''' <returns></returns>
        Public Shared Function NoDBNullListFromString(ByVal arrayData As Object, splitChar As Char, alternativeValue As Generic.List(Of Byte)) As Generic.List(Of Byte)
            If IsDBNull(arrayData) OrElse arrayData Is Nothing Then
                Return alternativeValue
            ElseIf CType(arrayData, String) = "" Then
                Return New Generic.List(Of Byte)
            Else
                Dim Result As New Generic.List(Of Byte)
                For Each SplitValue As String In CType(arrayData, String).Split(splitChar)
                    If Nullable.GetUnderlyingType(GetType(Byte)) IsNot Nothing AndAlso GetType(Byte).IsValueType AndAlso Nullable.GetUnderlyingType(GetType(Byte)) IsNot GetType(Byte) Then
                        Dim UnderlyingType As Type = Nullable.GetUnderlyingType(GetType(Byte)) 'e.g.: T = Nullable(Of Integer) -> UnderlyingType = Integer
                        Dim SplitValueOrNothing As String = StringNotEmptyOrNothing(SplitValue)
                        Dim SplitValueAsT As Byte = CType(Convert.ChangeType(SplitValueOrNothing, UnderlyingType, Threading.Thread.CurrentThread.CurrentCulture), Byte)
                        Result.Add(CType(Activator.CreateInstance(GetType(Byte), SplitValueAsT), Byte))
                    Else
                        Result.Add(CType(CType(SplitValue, Object), Byte))
                    End If
                Next
                Return Result
            End If
        End Function

        ''' <summary>
        ''' Check for DBNull and return null (Nothing in VisualBasic) alternatively, second check for empty string and return empty list, third split the string and fill the list with all elements
        ''' </summary>
        ''' <param name="arrayData"></param>
        ''' <param name="splitChar"></param>
        ''' <param name="alternativeValue"></param>
        ''' <returns></returns>
        Public Shared Function NoDBNullListFromString(ByVal arrayData As Object, splitChar As Char, alternativeValue As Generic.List(Of Double)) As Generic.List(Of Double)
            If IsDBNull(arrayData) OrElse arrayData Is Nothing Then
                Return alternativeValue
            ElseIf CType(arrayData, String) = "" Then
                Return New Generic.List(Of Double)
            Else
                Dim Result As New Generic.List(Of Double)
                For Each SplitValue As String In CType(arrayData, String).Split(splitChar)
                    If Nullable.GetUnderlyingType(GetType(Double)) IsNot Nothing AndAlso GetType(Double).IsValueType AndAlso Nullable.GetUnderlyingType(GetType(Double)) IsNot GetType(Double) Then
                        Dim UnderlyingType As Type = Nullable.GetUnderlyingType(GetType(Double)) 'e.g.: T = Nullable(Of Integer) -> UnderlyingType = Integer
                        Dim SplitValueOrNothing As String = StringNotEmptyOrNothing(SplitValue)
                        Dim SplitValueAsT As Double = CType(Convert.ChangeType(SplitValueOrNothing, UnderlyingType, Threading.Thread.CurrentThread.CurrentCulture), Double)
                        Result.Add(CType(Activator.CreateInstance(GetType(Double), SplitValueAsT), Double))
                    Else
                        Result.Add(CType(CType(SplitValue, Object), Double))
                    End If
                Next
                Return Result
            End If
        End Function

        ''' <summary>
        ''' Check for DBNull and return null (Nothing in VisualBasic) alternatively, second check for empty string and return empty list, third split the string and fill the list with all elements
        ''' </summary>
        ''' <param name="arrayData"></param>
        ''' <param name="splitChar"></param>
        ''' <param name="alternativeValue"></param>
        ''' <returns></returns>
        Public Shared Function NoDBNullListFromString(ByVal arrayData As Object, splitChar As Char, alternativeValue As Generic.List(Of Integer)) As Generic.List(Of Integer)
            If IsDBNull(arrayData) OrElse arrayData Is Nothing Then
                Return alternativeValue
            ElseIf CType(arrayData, String) = "" Then
                Return New Generic.List(Of Integer)
            Else
                Dim Result As New Generic.List(Of Integer)
                For Each SplitValue As String In CType(arrayData, String).Split(splitChar)
                    If Nullable.GetUnderlyingType(GetType(Integer)) IsNot Nothing AndAlso GetType(Integer).IsValueType AndAlso Nullable.GetUnderlyingType(GetType(Integer)) IsNot GetType(Integer) Then
                        Dim UnderlyingType As Type = Nullable.GetUnderlyingType(GetType(Integer)) 'e.g.: T = Nullable(Of Integer) -> UnderlyingType = Integer
                        Dim SplitValueOrNothing As String = StringNotEmptyOrNothing(SplitValue)
                        Dim SplitValueAsT As Integer = CType(Convert.ChangeType(SplitValueOrNothing, UnderlyingType, Threading.Thread.CurrentThread.CurrentCulture), Integer)
                        Result.Add(CType(Activator.CreateInstance(GetType(Integer), SplitValueAsT), Integer))
                    Else
                        Result.Add(CType(CType(SplitValue, Object), Integer))
                    End If
                Next
                Return Result
            End If
        End Function

        ''' <summary>
        ''' Check for DBNull and return null (Nothing in VisualBasic) alternatively, second check for empty string and return empty list, third split the string and fill the list with all elements
        ''' </summary>
        ''' <param name="arrayData"></param>
        ''' <param name="splitChar"></param>
        ''' <param name="alternativeValue"></param>
        ''' <returns></returns>
        Public Shared Function NoDBNullListFromString(ByVal arrayData As Object, splitChar As Char, alternativeValue As Generic.List(Of String)) As Generic.List(Of String)
            If IsDBNull(arrayData) OrElse arrayData Is Nothing Then
                Return alternativeValue
            ElseIf CType(arrayData, String) = "" Then
                Return New Generic.List(Of String)
            Else
                Dim Result As New Generic.List(Of String)
                For Each SplitValue As String In CType(arrayData, String).Split(splitChar)
                    If Nullable.GetUnderlyingType(GetType(String)) IsNot Nothing AndAlso GetType(String).IsValueType AndAlso Nullable.GetUnderlyingType(GetType(String)) IsNot GetType(String) Then
                        Dim UnderlyingType As Type = Nullable.GetUnderlyingType(GetType(String)) 'e.g.: T = Nullable(Of Integer) -> UnderlyingType = Integer
                        Dim SplitValueOrNothing As String = StringNotEmptyOrNothing(SplitValue)
                        Dim SplitValueAsT As String = CType(Convert.ChangeType(SplitValueOrNothing, UnderlyingType, Threading.Thread.CurrentThread.CurrentCulture), String)
                        Result.Add(CType(Activator.CreateInstance(GetType(String), SplitValueAsT), String))
                    Else
                        Result.Add(CType(CType(SplitValue, Object), String))
                    End If
                Next
                Return Result
            End If
        End Function

#End Region

#Region "Type conversions"

        ''' <summary>
        ''' Return a double which is not NaN (double's special constant &quot;not a number&quot;)
        ''' </summary>
        ''' <param name="value"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <DebuggerHidden()> Public Shared Function DoubleNotNaNOrNothing(ByVal value As Double) As Double
            If Double.IsNaN(value) Then
                Return Nothing
            Else
                Return value
            End If
        End Function
        ''' <summary>
        ''' Return a double which is not NaN (double's special constant &quot;not a number&quot;)
        ''' </summary>
        ''' <param name="value"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <DebuggerHidden()> Public Shared Function DoubleNotNaNOrDBNull(ByVal value As Double) As Object
            If Double.IsNaN(value) Then
                Return DBNull.Value
            Else
                Return value
            End If
        End Function
        ''' <summary>
        ''' Return a double which is not NaN (double's special constant &quot;not a number&quot;)
        ''' </summary>
        ''' <param name="value"></param>
        ''' <param name="alternativeValue"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <DebuggerHidden()> Public Shared Function DoubleNotNaNOrAlternativeValue(ByVal value As Double, ByVal alternativeValue As Double) As Double
            If Double.IsNaN(value) Then
                Return alternativeValue
            Else
                Return value
            End If
        End Function

        ''' <summary>
        '''     Return the string which is not nothing or else String.Empty
        ''' </summary>
        ''' <param name="value">The string to be validated</param>
        ''' <returns></returns>
        <DebuggerHidden()> Public Shared Function StringNotEmptyOrNothing(ByVal value As String) As String
            If value = Nothing Then
                Return Nothing
            Else
                Return value
            End If
        End Function

        ''' <summary>
        '''     Return the string which is not nothing or else String.Empty
        ''' </summary>
        ''' <param name="value">The string to be validated</param>
        ''' <returns></returns>
        <DebuggerHidden()> Public Shared Function StringNotNothingOrEmpty(ByVal value As String) As String
            If value Is Nothing Then
                Return String.Empty
            Else
                Return value
            End If
        End Function

        ''' <summary>
        '''     Return the string which is not nothing or else the alternative value
        ''' </summary>
        ''' <param name="value">The string to be validated</param>
        ''' <param name="alternativeValue">An alternative value if the first value is nothing</param>
        ''' <returns></returns>
        <DebuggerHidden()> Public Shared Function StringNotNothingOrAlternativeValue(ByVal value As String, ByVal alternativeValue As String) As String
            If value Is Nothing Then
                Return alternativeValue
            Else
                Return value
            End If
        End Function

        ''' <summary>
        '''     Return the string which is not empty or else the alternative value
        ''' </summary>
        ''' <param name="value">The string to be validated</param>
        ''' <param name="alternativeValue">An alternative value if the first value is empty</param>
        ''' <returns></returns>
        <DebuggerHidden()> Public Shared Function StringNotEmptyOrAlternativeValue(ByVal value As String, ByVal alternativeValue As String) As String
            If value = Nothing Then
                Return alternativeValue
            Else
                Return value
            End If
        End Function

        ''' <summary>
        '''     Return the string which is not empty or otherwise return DBNull.Value 
        ''' </summary>
        ''' <param name="value">The string to be validated</param>
        ''' <returns></returns>
        <DebuggerHidden()> Public Shared Function StringNotEmptyOrDBNull(ByVal value As String) As Object
            If value = Nothing Then
                Return DBNull.Value
            Else
                Return value
            End If
        End Function

        ''' <summary>
        '''     Return the value which is not nothing/null/zero or otherwise return DBNull.Value 
        ''' </summary>
        ''' <param name="value">The value to be validated</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <DebuggerHidden()> Public Shared Function ValueNotNothingOrDBNull(ByVal value As Double) As Object
            If value = Nothing Then
                Return DBNull.Value
            Else
                Return value
            End If
        End Function

        ''' <summary>
        '''     Return the value which is not nothing/null/zero or otherwise return DBNull.Value 
        ''' </summary>
        ''' <param name="value">The value to be validated</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <DebuggerHidden()> Public Shared Function ValueNotNothingOrDBNull(ByVal value As Integer) As Object
            If value = Nothing Then
                Return DBNull.Value
            Else
                Return value
            End If
        End Function

        ''' <summary>
        '''     Return the value which is not nothing/null/zero or otherwise return DBNull.Value 
        ''' </summary>
        ''' <param name="value">The value to be validated</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <DebuggerHidden()> Public Shared Function ValueNotNothingOrDBNull(ByVal value As Long) As Object
            If value = Nothing Then
                Return DBNull.Value
            Else
                Return value
            End If
        End Function

        ''' <summary>
        '''     Return the value which is not nothing/null/zero or otherwise return DBNull.Value 
        ''' </summary>
        ''' <param name="value">The value to be validated</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <DebuggerHidden()> Public Shared Function ValueNotNothingOrDBNull(ByVal value As Decimal) As Object
            If value = Nothing Then
                Return DBNull.Value
            Else
                Return value
            End If
        End Function

        ''' <summary>
        '''     Return the value which is not nothing/null/zero or otherwise return DBNull.Value 
        ''' </summary>
        ''' <param name="value">The value to be validated</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <DebuggerHidden()> Public Shared Function ValueNotNothingOrDBNull(ByVal value As DateTime) As Object
            If value = Nothing Then
                Return DBNull.Value
            Else
                Return value
            End If
        End Function

        ''' <summary>
        '''     Return the value which is not nothing/null/zero or otherwise return DBNull.Value 
        ''' </summary>
        ''' <param name="value">The value to be validated</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <DebuggerHidden()> Public Shared Function ValueNotNothingOrDBNull(ByVal value As Single) As Object
            If value = Nothing Then
                Return DBNull.Value
            Else
                Return value
            End If
        End Function

        ''' <summary>
        '''     Return the value which is not nothing/null/zero or otherwise return DBNull.Value 
        ''' </summary>
        ''' <param name="value">The value to be validated</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <DebuggerHidden()> Public Shared Function ValueNotNothingOrDBNull(ByVal value As Byte) As Object
            If value = Nothing Then
                Return DBNull.Value
            Else
                Return value
            End If
        End Function

        ''' <summary>
        '''     Return the value which is not nothing/null/zero or otherwise return DBNull.Value 
        ''' </summary>
        ''' <param name="value">The value to be validated</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <DebuggerHidden(), CLSCompliant(False)> Public Shared Function ValueNotNothingOrDBNull(ByVal value As UInt16) As Object
            If value = Nothing Then
                Return DBNull.Value
            Else
                Return value
            End If
        End Function
        ''' <summary>
        '''     Return the value which is not nothing/null/zero or otherwise return DBNull.Value 
        ''' </summary>
        ''' <param name="value">The value to be validated</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <DebuggerHidden(), CLSCompliant(False)> Public Shared Function ValueNotNothingOrDBNull(ByVal value As UInt32) As Object
            If value = Nothing Then
                Return DBNull.Value
            Else
                Return value
            End If
        End Function
        ''' <summary>
        '''     Return the value which is not nothing/null/zero or otherwise return DBNull.Value 
        ''' </summary>
        ''' <param name="value">The value to be validated</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <DebuggerHidden(), CLSCompliant(False)> Public Shared Function ValueNotNothingOrDBNull(ByVal value As UInt64) As Object
            If value = Nothing Then
                Return DBNull.Value
            Else
                Return value
            End If
        End Function

        ''' <summary>
        '''     Return the value which is not nothing/null/zero or otherwise return DBNull.Value 
        ''' </summary>
        ''' <param name="value">The value to be validated</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <DebuggerHidden()> Public Shared Function ValueNotNothingOrDBNull(ByVal value As Short) As Object
            If value = Nothing Then
                Return DBNull.Value
            Else
                Return value
            End If
        End Function

        ''' <summary>
        '''     Return the value which is not nothing/null/zero or otherwise return DBNull.Value 
        ''' </summary>
        ''' <param name="value">The value to be validated</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <DebuggerHidden()> Public Shared Function ValueNotNothingOrAlternativeValue(ByVal value As Double, ByVal alternativeValue As Double) As Double
            If value = Nothing Then
                Return alternativeValue
            Else
                Return value
            End If
        End Function

        ''' <summary>
        '''     Return the value which is not nothing/null/zero or otherwise return DBNull.Value 
        ''' </summary>
        ''' <param name="value">The value to be validated</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <DebuggerHidden()> Public Shared Function ValueNotNothingOrAlternativeValue(ByVal value As Integer, ByVal alternativeValue As Integer) As Integer
            If value = Nothing Then
                Return alternativeValue
            Else
                Return value
            End If
        End Function

        ''' <summary>
        '''     Return the value which is not nothing/null/zero or otherwise return DBNull.Value 
        ''' </summary>
        ''' <param name="value">The value to be validated</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <DebuggerHidden()> Public Shared Function ValueNotNothingOrAlternativeValue(ByVal value As Long, ByVal alternativeValue As Long) As Long
            If value = Nothing Then
                Return alternativeValue
            Else
                Return value
            End If
        End Function

        ''' <summary>
        '''     Return the value which is not nothing/null/zero or otherwise return DBNull.Value 
        ''' </summary>
        ''' <param name="value">The value to be validated</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <DebuggerHidden()> Public Shared Function ValueNotNothingOrAlternativeValue(ByVal value As Decimal, ByVal alternativeValue As Decimal) As Decimal
            If value = Nothing Then
                Return alternativeValue
            Else
                Return value
            End If
        End Function

        ''' <summary>
        '''     Return the value which is not nothing/null/zero or otherwise return DBNull.Value 
        ''' </summary>
        ''' <param name="value">The value to be validated</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <DebuggerHidden()> Public Shared Function ValueNotNothingOrAlternativeValue(ByVal value As DateTime, ByVal alternativeValue As DateTime) As DateTime
            If value = Nothing Then
                Return alternativeValue
            Else
                Return value
            End If
        End Function

        ''' <summary>
        '''     Return the value which is not nothing/null/zero or otherwise return DBNull.Value 
        ''' </summary>
        ''' <param name="value">The value to be validated</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <DebuggerHidden()> Public Shared Function ValueNotNothingOrAlternativeValue(ByVal value As TimeSpan, ByVal alternativeValue As TimeSpan) As TimeSpan
            If value = Nothing Then
                Return alternativeValue
            Else
                Return value
            End If
        End Function

        ''' <summary>
        '''     Return the value which is not nothing/null/zero or otherwise return DBNull.Value 
        ''' </summary>
        ''' <param name="value">The value to be validated</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <DebuggerHidden()> Public Shared Function ValueNotNothingOrAlternativeValue(ByVal value As Single, ByVal alternativeValue As Single) As Single
            If value = Nothing Then
                Return alternativeValue
            Else
                Return value
            End If
        End Function

        ''' <summary>
        '''     Return the value which is not nothing/null/zero or otherwise return DBNull.Value 
        ''' </summary>
        ''' <param name="value">The value to be validated</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <DebuggerHidden()> Public Shared Function ValueNotNothingOrAlternativeValue(ByVal value As Byte, ByVal alternativeValue As Byte) As Byte
            If value = Nothing Then
                Return alternativeValue
            Else
                Return value
            End If
        End Function

        ''' <summary>
        '''     Return the value which is not nothing/null/zero or otherwise return DBNull.Value 
        ''' </summary>
        ''' <param name="value">The value to be validated</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <DebuggerHidden(), CLSCompliant(False)> Public Shared Function ValueNotNothingOrAlternativeValue(ByVal value As Byte, ByVal alternativeValue As UInt16) As UInt16
            If value = Nothing Then
                Return alternativeValue
            Else
                Return value
            End If
        End Function
        ''' <summary>
        '''     Return the value which is not nothing/null/zero or otherwise return DBNull.Value 
        ''' </summary>
        ''' <param name="value">The value to be validated</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <DebuggerHidden(), CLSCompliant(False)> Public Shared Function ValueNotNothingOrAlternativeValue(ByVal value As Byte, ByVal alternativeValue As UInt32) As UInt32
            If value = Nothing Then
                Return alternativeValue
            Else
                Return value
            End If
        End Function
        ''' <summary>
        '''     Return the value which is not nothing/null/zero or otherwise return DBNull.Value 
        ''' </summary>
        ''' <param name="value">The value to be validated</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <DebuggerHidden(), CLSCompliant(False)> Public Shared Function ValueNotNothingOrAlternativeValue(ByVal value As Byte, ByVal alternativeValue As UInt64) As UInt64
            If value = Nothing Then
                Return alternativeValue
            Else
                Return value
            End If
        End Function
        ''' <summary>
        '''     Return the value which is not nothing/null/zero or otherwise return DBNull.Value 
        ''' </summary>
        ''' <param name="value">The value to be validated</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <DebuggerHidden()> Public Shared Function ValueNotNothingOrAlternativeValue(ByVal value As Short, ByVal alternativeValue As Short) As Short
            If value = Nothing Then
                Return alternativeValue
            Else
                Return value
            End If
        End Function

        ''' <summary>
        '''     Return the object which is not nothing or otherwise return DBNull.Value 
        ''' </summary>
        ''' <param name="value">The string to be validated</param>
        ''' <returns></returns>
        <DebuggerHidden()> Public Shared Function ObjectNotNothingOrEmptyString(ByVal value As Object) As Object
            If value Is Nothing Then
                Return String.Empty
            Else
                Return value
            End If
        End Function

        ''' <summary>
        '''     Return the object which is not nothing or otherwise return DBNull.Value 
        ''' </summary>
        ''' <param name="value">The string to be validated</param>
        ''' <returns></returns>
        <DebuggerHidden()> Public Shared Function ObjectNotNothingOrDBNull(ByVal value As Object) As Object
            If value Is Nothing Then
                Return DBNull.Value
            Else
                Return value
            End If
        End Function

        ''' <summary>
        '''     Return the object which is not an empty string or otherwise return Nothing
        ''' </summary>
        ''' <param name="value">The object to be validated</param>
        ''' <returns>A string with length > 0 (the value) or nothing</returns>
        <DebuggerHidden()> Public Shared Function ObjectNotEmptyStringOrNothing(ByVal value As Object) As Object
            If value Is Nothing Then
                Return Nothing
            ElseIf value.GetType Is GetType(String) AndAlso CType(value, String) = "" Then
                Return Nothing
            Else
                Return value
            End If
        End Function

        ''' <summary>
        '''     Return the value if there is a value or otherwise return DBNull.Value 
        ''' </summary>
        ''' <param name="value">The nullable type value to be validated</param>
        Public Shared Function NullableTypeWithItsValueOrDBNull(Of T As Structure)(ByVal value As Nullable(Of T)) As Object
            If value.HasValue = False Then
                Return DBNull.Value
            Else
                Return value.Value
            End If
        End Function

        ''' <summary>
        '''     Return the array which is not nothing or otherwise return DBNull.Value 
        ''' </summary>
        ''' <param name="values">The array to be validated</param>
        Public Shared Function ArrayNotNothingOrDBNull(ByVal values As Array) As Object
            If values Is Nothing Then
                Return DBNull.Value
            Else
                Return values
            End If
        End Function

        ''' <summary>
        '''     Return the array with at least 1 element or otherwise return DBNull.Value 
        ''' </summary>
        ''' <param name="values">The array to be validated</param>
        Public Shared Function ArrayNotEmptyOrDBNull(ByVal values As Array) As Object
            If values Is Nothing OrElse values.Length = 0 Then
                Return DBNull.Value
            Else
                Return values
            End If
        End Function
        ''' <summary>
        '''     Return the array with at least 1 element or otherwise return Nothing
        ''' </summary>
        ''' <param name="values">The array to be validated</param>
        Public Shared Function ArrayNotEmptyOrNothing(Of T)(ByVal values As T()) As T()
            If values Is Nothing OrElse values.Length = 0 Then
                Return Nothing
            Else
                Return values
            End If
        End Function
        ''' <summary>
        '''     Return the array with at least 0 elements in case it's Nothing
        ''' </summary>
        ''' <param name="values">The array to be validated</param>
        Public Shared Function ArrayNotNothingOrEmpty(Of T)(ByVal values As T()) As T()
            If values Is Nothing Then
                Return Array.Empty(Of T)()
            Else
                Return values
            End If
        End Function

        ''' <summary>
        '''     Return the string which is not nothing or otherwise return DBNull.Value 
        ''' </summary>
        ''' <param name="value">The string to be validated</param>
        ''' <returns></returns>
        <DebuggerHidden()> Public Shared Function StringNotNothingOrDBNull(ByVal value As String) As Object
            If value Is Nothing Then
                Return DBNull.Value
            Else
                Return value
            End If
        End Function

#End Region

#Region "ConnectionString without sensitive data"
        ''' <summary>
        ''' Prepare a connection string for transmission to users without sensitive password information
        ''' </summary>
        ''' <param name="fullConnectionString">The regular ConnectionString</param>
        ''' <returns>The first part of the ConnectionString till the password position</returns>
        ''' <remarks>
        ''' All information after the password position will be removed, too. So, you can hide the user name by positioning it after the password (UID=user;PWD=xxxx vs. PWD=xxxx;UID=user).
        ''' </remarks>
        Public Shared Function ConnectionStringWithoutPasswords(ByVal fullConnectionString As String) As String
            Dim PWDPos As Integer
            PWDPos = InStr(UCase(fullConnectionString), "PWD=")
            If PWDPos > 0 Then
                fullConnectionString = Mid(fullConnectionString, 1, PWDPos + 3) & "..."
            End If
            PWDPos = InStr(UCase(fullConnectionString), "PASSWORD=")
            If PWDPos > 0 Then
                fullConnectionString = Mid(fullConnectionString, 1, PWDPos + 8) & "..."
            End If
            Return fullConnectionString
        End Function
#End Region

#Region "ReadString/ByteDataFromUri"

        Public Shared Function ReadByteDataFromUri(ByVal uri As String) As Byte()
            Dim client As New System.Net.WebClient
            Return client.DownloadData(uri)
        End Function

        Public Shared Function ReadStringDataFromUri(ByVal uri As String, ByVal encodingName As String) As String
            Return ReadStringDataFromUri(CType(Nothing, System.Net.WebClient), uri, encodingName)
        End Function

        Public Shared Function ReadStringDataFromUri(ByVal uri As String, ByVal encodingName As String, ByVal ignoreSslValidationExceptions As Boolean) As String
            Return ReadStringDataFromUri(CType(Nothing, System.Net.WebClient), uri, encodingName, ignoreSslValidationExceptions)
        End Function

        Public Shared Function ReadStringDataFromUri(ByVal client As System.Net.WebClient, ByVal uri As String, ByVal encodingName As String) As String
            Return ReadStringDataFromUri(client, uri, encodingName, False)
        End Function

        Public Shared Function ReadStringDataFromUri(ByVal client As System.Net.WebClient, ByVal uri As String, ByVal encodingName As String, ByVal ignoreSslValidationExceptions As Boolean) As String
            Return ReadStringDataFromUri(client, uri, encodingName, ignoreSslValidationExceptions, CType(Nothing, String))
        End Function

        Public Shared Function ReadStringDataFromUri(ByVal client As System.Net.WebClient, ByVal uri As String, ByVal encodingName As String, ByVal ignoreSslValidationExceptions As Boolean, ByVal postData As String) As String
            If client Is Nothing Then client = New System.Net.WebClient
            'https://compumaster.dyndns.biz/.....asmx without trusted certificate
            Dim CurrentValidationCallback As System.Net.Security.RemoteCertificateValidationCallback = System.Net.ServicePointManager.ServerCertificateValidationCallback
            Try
#Disable Warning CA5359 ' Deaktivieren Sie die Zertifikatüberprüfung nicht
                If ignoreSslValidationExceptions Then System.Net.ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf OnValidationCallback)
#Enable Warning CA5359 ' Deaktivieren Sie die Zertifikatüberprüfung nicht
                If encodingName <> Nothing Then
                    Dim bytes As Byte()
                    If postData Is Nothing Then
                        bytes = client.DownloadData(uri)
                    Else
                        bytes = client.UploadData(uri, System.Text.Encoding.GetEncoding(encodingName).GetBytes(postData))
                    End If
                    Return System.Text.Encoding.GetEncoding(encodingName).GetString(bytes)
                Else
                    Dim Result As String
                    If postData Is Nothing Then
                        Result = client.DownloadString(uri)
                    Else
                        Result = client.UploadString(uri, postData)
                    End If
                    If client.ResponseHeaders("Content-Type") IsNot Nothing Then
                        'HACK: download twice, but now with 1st response's charset encoding information
                        Dim ResultCharsetEncodingName As String = New System.Net.Mime.ContentType(client.ResponseHeaders("Content-Type")).CharSet
                        If ResultCharsetEncodingName = Nothing Then ResultCharsetEncodingName = "utf-8"
                        Dim bytes As Byte()
                        If postData Is Nothing Then
                            bytes = client.DownloadData(uri)
                        Else
                            bytes = client.UploadData(uri, System.Text.Encoding.GetEncoding(ResultCharsetEncodingName).GetBytes(postData))
                        End If
                        Return System.Text.Encoding.GetEncoding(ResultCharsetEncodingName).GetString(bytes)
                    Else
                        Return Result 'no content encoding information available, return downloaded string as is
                    End If
                End If
            Finally
                System.Net.ServicePointManager.ServerCertificateValidationCallback = CurrentValidationCallback
            End Try
        End Function

        ''' <summary>
        ''' Suppress all SSL certification requirements - just use the webservice SSL URL
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="cert"></param>
        ''' <param name="chain"></param>
        ''' <param name="errors"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function OnValidationCallback(ByVal sender As Object, ByVal cert As System.Security.Cryptography.X509Certificates.X509Certificate, ByVal chain As System.Security.Cryptography.X509Certificates.X509Chain, ByVal errors As System.Net.Security.SslPolicyErrors) As Boolean
            Return True
        End Function

#End Region

        Public Shared Function ArrangeTableBlocksBesides(leftTableOutput As String, rightTableOutput As String) As String
            Return ArrangeTableBlocksBesides(leftTableOutput, rightTableOutput, "     ")
        End Function

        Public Shared Function ArrangeTableBlocksBesides(leftTableOutput As String, rightTableOutput As String, blockSeparator As String) As String
            Dim MaxWidthLeftTable As TextBlockLineData = EvaluateTextBlockLineData(leftTableOutput)
            Dim MaxWidthRightTable As TextBlockLineData = EvaluateTextBlockLineData(rightTableOutput)
            Dim Result As New System.Text.StringBuilder
            Dim MaxLines As Integer = System.Math.Max(MaxWidthLeftTable.Lines.Length, MaxWidthRightTable.Lines.Length)
            For MyCounter As Integer = 0 To MaxLines
                'Left table block
                If MyCounter < MaxWidthLeftTable.Lines.Length Then
                    Result.Append(MaxWidthLeftTable.Lines(MyCounter))
                    Result.Append(Space(MaxWidthLeftTable.MaxWidth - MaxWidthLeftTable.Lines(MyCounter).Length))
                Else
                    Result.Append(Space(MaxWidthLeftTable.MaxWidth))
                End If
                'Separator
                Result.Append(blockSeparator)
                'Right table block
                If MyCounter < MaxWidthRightTable.Lines.Length Then
                    Result.Append(MaxWidthRightTable.Lines(MyCounter))
                    Result.Append(Space(MaxWidthRightTable.MaxWidth - MaxWidthRightTable.Lines(MyCounter).Length))
                Else
                    Result.Append(Space(MaxWidthRightTable.MaxWidth))
                End If
                Result.AppendLine()
            Next
            Return Result.ToString
        End Function

        Private Class TextBlockLineData
            Public Sub New(maxWidth As Integer, lines As String())
                Me.MaxWidth = maxWidth
                Me.Lines = lines
            End Sub
            Public MaxWidth As Integer
            Public Lines As String()
        End Class

        ''' <summary>
        ''' Evaluate the maximum width of a text block
        ''' </summary>
        ''' <param name="text"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Shared Function EvaluateTextBlockLineData(text As String) As TextBlockLineData
            If text = Nothing Then Return Nothing
            Dim Result As Integer = 0
            Dim TextWithoutLastLineBreak As String = text.Replace(ControlChars.CrLf, ControlChars.Cr).Replace(ControlChars.Lf, ControlChars.Cr)
            If TextWithoutLastLineBreak.EndsWith(ControlChars.Cr, StringComparison.Ordinal) Then
                'Remove last CR because it doesn't count here (will usually be added automatically again in further steps)
                TextWithoutLastLineBreak = TextWithoutLastLineBreak.Substring(0, TextWithoutLastLineBreak.Length - 1)
            End If
            Dim Lines As String() = TextWithoutLastLineBreak.Split(ControlChars.Cr)
            For MyCounter As Integer = 0 To Lines.Length - 1
                Result = System.Math.Max(Result, Lines(MyCounter).Length)
            Next
            Return New TextBlockLineData(Result, Lines)
        End Function
    End Class

End Namespace