Option Explicit On 
Option Strict On

Namespace CompuMaster.Data

    ''' -----------------------------------------------------------------------------
    ''' Project	 : CompuMaster.Data
    ''' Class	 : camm.WebManager.Utils
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Utils for converting and handling database data
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[wezel]	21.11.2008	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Class Utils

#Region "NoDBNull"
        ''' -----------------------------------------------------------------------------
        ''' <summary>
        '''     Checks for DBNull and returns null (Nothing in VisualBasic) in that case
        ''' </summary>
        ''' <param name="checkValueIfDBNull">The value to be checked</param>
        ''' <returns>A value which is not DBNull; a DBNull as input will return null (Nothing in VisualBasic)</returns>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[adminwezel]	06.07.2004	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
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

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        '''     Checks for DBNull and returns the second value alternatively
        ''' </summary>
        ''' <param name="checkValueIfDBNull">The value to be checked</param>
        ''' <param name="replaceWithThis">The alternative value, null (Nothing in VisualBasic) if not defined</param>
        ''' <returns>A value which is not DBNull</returns>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[adminwezel]	06.07.2004	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        <DebuggerHidden()> Public Shared Function NoDBNull(ByVal checkValueIfDBNull As Object, ByVal replaceWithThis As Object) As Object
            If IsDBNull(checkValueIfDBNull) Then
                Return (replaceWithThis)
            Else
                Return (checkValueIfDBNull)
            End If
        End Function
        ''' -----------------------------------------------------------------------------
        ''' <summary>
        '''     Checks for DBNull and returns the second value alternatively
        ''' </summary>
        ''' <param name="checkValueIfDBNull">The value to be checked</param>
        ''' <param name="replaceWithThis">The alternative value, null (Nothing in VisualBasic) if not defined</param>
        ''' <returns>A value which is not DBNull</returns>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[adminwezel]	06.07.2004	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        <DebuggerHidden()> Public Shared Function NoDBNull(ByVal checkValueIfDBNull As Object, ByVal replaceWithThis As Integer) As Integer
            If IsDBNull(checkValueIfDBNull) Then
                Return (replaceWithThis)
            Else
                Return CType(checkValueIfDBNull, Integer)
            End If
        End Function
        ''' -----------------------------------------------------------------------------
        ''' <summary>
        '''     Checks for DBNull and returns the second value alternatively
        ''' </summary>
        ''' <param name="checkValueIfDBNull">The value to be checked</param>
        ''' <param name="replaceWithThis">The alternative value, null (Nothing in VisualBasic) if not defined</param>
        ''' <returns>A value which is not DBNull</returns>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[adminwezel]	06.07.2004	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        <DebuggerHidden()> Public Shared Function NoDBNull(ByVal checkValueIfDBNull As Object, ByVal replaceWithThis As Long) As Long
            If IsDBNull(checkValueIfDBNull) Then
                Return (replaceWithThis)
            Else
                Return CType(checkValueIfDBNull, Long)
            End If
        End Function
        ''' -----------------------------------------------------------------------------
        ''' <summary>
        '''     Checks for DBNull and returns the second value alternatively
        ''' </summary>
        ''' <param name="checkValueIfDBNull">The value to be checked</param>
        ''' <param name="replaceWithThis">The alternative value, null (Nothing in VisualBasic) if not defined</param>
        ''' <returns>A value which is not DBNull</returns>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[adminwezel]	06.07.2004	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        <DebuggerHidden()> Public Shared Function NoDBNull(ByVal checkValueIfDBNull As Object, ByVal replaceWithThis As Decimal) As Decimal
            If IsDBNull(checkValueIfDBNull) Then
                Return (replaceWithThis)
            Else
                Return CType(checkValueIfDBNull, Decimal)
            End If
        End Function
        ''' -----------------------------------------------------------------------------
        ''' <summary>
        '''     Checks for DBNull and returns the second value alternatively
        ''' </summary>
        ''' <param name="checkValueIfDBNull">The value to be checked</param>
        ''' <param name="replaceWithThis">The alternative value, null (Nothing in VisualBasic) if not defined</param>
        ''' <returns>A value which is not DBNull</returns>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[adminwezel]	06.07.2004	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        <DebuggerHidden()> Public Shared Function NoDBNull(ByVal checkValueIfDBNull As Object, ByVal replaceWithThis As Short) As Short
            If IsDBNull(checkValueIfDBNull) Then
                Return (replaceWithThis)
            Else
                Return CType(checkValueIfDBNull, Short)
            End If
        End Function
        ''' -----------------------------------------------------------------------------
        ''' <summary>
        '''     Checks for DBNull and returns the second value alternatively
        ''' </summary>
        ''' <param name="checkValueIfDBNull">The value to be checked</param>
        ''' <param name="replaceWithThis">The alternative value, null (Nothing in VisualBasic) if not defined</param>
        ''' <returns>A value which is not DBNull</returns>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[adminwezel]	06.07.2004	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        <DebuggerHidden()> Public Shared Function NoDBNull(ByVal checkValueIfDBNull As Object, ByVal replaceWithThis As Single) As Single
            If IsDBNull(checkValueIfDBNull) Then
                Return (replaceWithThis)
            Else
                Return CType(checkValueIfDBNull, Single)
            End If
        End Function
        ''' -----------------------------------------------------------------------------
        ''' <summary>
        '''     Checks for DBNull and returns the second value alternatively
        ''' </summary>
        ''' <param name="checkValueIfDBNull">The value to be checked</param>
        ''' <param name="replaceWithThis">The alternative value, null (Nothing in VisualBasic) if not defined</param>
        ''' <returns>A value which is not DBNull</returns>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[adminwezel]	06.07.2004	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        <DebuggerHidden()> Public Shared Function NoDBNull(ByVal checkValueIfDBNull As Object, ByVal replaceWithThis As Boolean) As Boolean
            If IsDBNull(checkValueIfDBNull) Then
                Return (replaceWithThis)
            Else
                Return CType(checkValueIfDBNull, Boolean)
            End If
        End Function
        ''' -----------------------------------------------------------------------------
        ''' <summary>
        '''     Checks for DBNull and returns the second value alternatively
        ''' </summary>
        ''' <param name="checkValueIfDBNull">The value to be checked</param>
        ''' <param name="replaceWithThis">The alternative value, null (Nothing in VisualBasic) if not defined</param>
        ''' <returns>A value which is not DBNull</returns>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[adminwezel]	06.07.2004	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        <DebuggerHidden()> Public Shared Function NoDBNull(ByVal checkValueIfDBNull As Object, ByVal replaceWithThis As DateTime) As DateTime
            If IsDBNull(checkValueIfDBNull) Then
                Return (replaceWithThis)
            Else
                Return CType(checkValueIfDBNull, DateTime)
            End If
        End Function
        ''' -----------------------------------------------------------------------------
        ''' <summary>
        '''     Checks for DBNull and returns the second value alternatively
        ''' </summary>
        ''' <param name="checkValueIfDBNull">The value to be checked</param>
        ''' <param name="replaceWithThis">The alternative value, null (Nothing in VisualBasic) if not defined</param>
        ''' <returns>A value which is not DBNull</returns>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[adminwezel]	06.07.2004	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        <DebuggerHidden()> Public Shared Function NoDBNull(ByVal checkValueIfDBNull As Object, ByVal replaceWithThis As Double) As Double
            If IsDBNull(checkValueIfDBNull) Then
                Return (replaceWithThis)
            Else
                Return CType(checkValueIfDBNull, Double)
            End If
        End Function
        ''' -----------------------------------------------------------------------------
        ''' <summary>
        '''     Checks for DBNull and returns the second value alternatively
        ''' </summary>
        ''' <param name="checkValueIfDBNull">The value to be checked</param>
        ''' <param name="replaceWithThis">The alternative value, null (Nothing in VisualBasic) if not defined</param>
        ''' <returns>A value which is not DBNull</returns>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[adminwezel]	06.07.2004	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        <DebuggerHidden()> Public Shared Function NoDBNull(ByVal checkValueIfDBNull As Object, ByVal replaceWithThis As Byte()) As Byte()
            If IsDBNull(checkValueIfDBNull) Then
                Return (replaceWithThis)
            Else
                Return CType(checkValueIfDBNull, Byte())
            End If
        End Function
        ''' -----------------------------------------------------------------------------
        ''' <summary>
        '''     Checks for DBNull and returns the second value alternatively
        ''' </summary>
        ''' <param name="checkValueIfDBNull">The value to be checked</param>
        ''' <param name="replaceWithThis">The alternative value, null (Nothing in VisualBasic) if not defined</param>
        ''' <returns>A value which is not DBNull</returns>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[adminwezel]	06.07.2004	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        <DebuggerHidden()> Public Shared Function NoDBNull(ByVal checkValueIfDBNull As Object, ByVal replaceWithThis As String) As String
            If IsDBNull(checkValueIfDBNull) Then
                Return (replaceWithThis)
            Else
                Return CType(checkValueIfDBNull, String)
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

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        '''     Return the string which is not nothing or else String.Empty
        ''' </summary>
        ''' <param name="value">The string to be validated</param>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[adminsupport]	09.11.2004	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        <DebuggerHidden()> Public Shared Function StringNotEmptyOrNothing(ByVal value As String) As String
            If value = Nothing Then
                Return Nothing
            Else
                Return value
            End If
        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        '''     Return the string which is not nothing or else String.Empty
        ''' </summary>
        ''' <param name="value">The string to be validated</param>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[adminsupport]	09.11.2004	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        <DebuggerHidden()> Public Shared Function StringNotNothingOrEmpty(ByVal value As String) As String
            If value Is Nothing Then
                Return String.Empty
            Else
                Return value
            End If
        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        '''     Return the string which is not nothing or else the alternative value
        ''' </summary>
        ''' <param name="value">The string to be validated</param>
        ''' <param name="alternativeValue">An alternative value if the first value is nothing</param>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[adminsupport]	09.11.2004	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        <DebuggerHidden()> Public Shared Function StringNotNothingOrAlternativeValue(ByVal value As String, ByVal alternativeValue As String) As String
            If value Is Nothing Then
                Return alternativeValue
            Else
                Return value
            End If
        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        '''     Return the string which is not empty or else the alternative value
        ''' </summary>
        ''' <param name="value">The string to be validated</param>
        ''' <param name="alternativeValue">An alternative value if the first value is empty</param>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[adminsupport]	09.11.2004	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        <DebuggerHidden()> Public Shared Function StringNotEmptyOrAlternativeValue(ByVal value As String, ByVal alternativeValue As String) As String
            If value = Nothing Then
                Return alternativeValue
            Else
                Return value
            End If
        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        '''     Return the string which is not empty or otherwise return DBNull.Value 
        ''' </summary>
        ''' <param name="value">The string to be validated</param>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[adminsupport]	09.11.2004	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
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
        <DebuggerHidden()> Public Shared Function ValueNotNothingOrAlternativeValue(ByVal value As Short, ByVal alternativeValue As Short) As Short
            If value = Nothing Then
                Return alternativeValue
            Else
                Return value
            End If
        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        '''     Return the object which is not nothing or otherwise return DBNull.Value 
        ''' </summary>
        ''' <param name="value">The string to be validated</param>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[adminsupport]	09.11.2004	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        <DebuggerHidden()> Public Shared Function ObjectNotNothingOrEmptyString(ByVal value As Object) As Object
            If value Is Nothing Then
                Return String.Empty
            Else
                Return value
            End If
        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        '''     Return the object which is not nothing or otherwise return DBNull.Value 
        ''' </summary>
        ''' <param name="value">The string to be validated</param>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[adminsupport]	09.11.2004	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        <DebuggerHidden()> Public Shared Function ObjectNotNothingOrDBNull(ByVal value As Object) As Object
            If value Is Nothing Then
                Return DBNull.Value
            Else
                Return value
            End If
        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        '''     Return the object which is not an empty string or otherwise return Nothing
        ''' </summary>
        ''' <param name="value">The object to be validated</param>
        ''' <returns>A string with length > 0 (the value) or nothing</returns>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[adminsupport]	09.11.2004	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        <DebuggerHidden()> Public Shared Function ObjectNotEmptyStringOrNothing(ByVal value As Object) As Object
            If value Is Nothing Then
                Return Nothing
            ElseIf value.GetType Is GetType(String) AndAlso CType(value, String) = "" Then
                Return Nothing
            Else
                Return value
            End If
        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        '''     Return the string which is not nothing or otherwise return DBNull.Value 
        ''' </summary>
        ''' <param name="value">The string to be validated</param>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[adminsupport]	09.11.2004	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        <DebuggerHidden()> Public Shared Function StringNotNothingOrDBNull(ByVal value As String) As Object
            If value Is Nothing Then
                Return DBNull.Value
            Else
                Return value
            End If
        End Function

#End Region

#Region "ConnectionString without sensitive data"
        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Prepare a connection string for transmission to users without sensitive password information
        ''' </summary>
        ''' <param name="fullConnectionString">The regular ConnectionString</param>
        ''' <returns>The first part of the ConnectionString till the password position</returns>
        ''' <remarks>
        ''' All information after the password position will be removed, too. So, you can hide the user name by positioning it after the password (UID=user;PWD=xxxx vs. PWD=xxxx;UID=user).
        ''' </remarks>
        ''' <history>
        ''' 	[wezel]	25.06.2009	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Shared Function ConnectionStringWithoutPasswords(ByVal fullConnectionString As String) As String
            Dim PWDPos As Integer
            PWDPos = InStr(UCase(fullConnectionString), "PWD=")
            If PWDPos > 0 Then
                fullConnectionString = Mid(fullConnectionString, 1, PWDPos + 3) & "..."
            Else
                fullConnectionString = fullConnectionString
            End If
            PWDPos = InStr(UCase(fullConnectionString), "PASSWORD=")
            If PWDPos > 0 Then
                fullConnectionString = Mid(fullConnectionString, 1, PWDPos + 8) & "..."
            Else
                fullConnectionString = fullConnectionString
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
            Return ReadStringDataFromUri(CType(Nothing, System.Net.WebClient), uri, encodingName, False)
        End Function

        Public Shared Function ReadStringDataFromUri(ByVal client As System.Net.WebClient, ByVal uri As String, ByVal encodingName As String) As String
            Return ReadStringDataFromUri(client, uri, encodingName, False)
        End Function

        Public Shared Function ReadStringDataFromUri(ByVal client As System.Net.WebClient, ByVal uri As String, ByVal encodingName As String, ByVal ignoreSslValidationExceptions As Boolean) As String
            Return ReadStringDataFromUri(client, uri, encodingName, False, CType(Nothing, String))
        End Function

        Public Shared Function ReadStringDataFromUri(ByVal client As System.Net.WebClient, ByVal uri As String, ByVal encodingName As String, ByVal ignoreSslValidationExceptions As Boolean, ByVal postData As String) As String
            If client Is Nothing Then client = New System.Net.WebClient
            'https://compumaster.dyndns.biz/.....asmx without trusted certificate
#If Not NET_1_1 Then
            Dim CurrentValidationCallback As System.Net.Security.RemoteCertificateValidationCallback = System.Net.ServicePointManager.ServerCertificateValidationCallback
            Try
            If ignoreSslValidationExceptions Then System.Net.ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf OnValidationCallback)
#End If
            If encodingName <> Nothing Then
                Dim bytes As Byte()
                If postData Is Nothing Then
                    bytes = client.DownloadData(uri)
                Else
                    bytes = client.UploadData(uri, System.Text.Encoding.GetEncoding(encodingName).GetBytes(postData))
                End If
                Return System.Text.Encoding.GetEncoding(encodingName).GetString(bytes)
            Else
#If NET_1_1 Then
                Dim encoding As System.Text.Encoding
                Try
                    Dim encName As String = client.ResponseHeaders("Content-Type")
                    If encName <> "" And encName.IndexOf("charset=") > -1 Then
                        encName = encName.Substring(encName.IndexOf("charset=") + "charset=".Length)
                        encoding = System.Text.Encoding.GetEncoding(encName)
                    Else
                        encoding = System.Text.Encoding.Default
                    End If
                Catch
                    encoding = System.Text.Encoding.Default
                End Try
                Dim bytes As Byte()
                If postData Is Nothing Then
                    bytes = client.DownloadData(uri)
                Else
                    bytes = client.UploadData(uri, encoding.GetBytes(postData))
                End If
                Return encoding.GetString(bytes)
#Else
                If postData Is Nothing Then
                    Return client.DownloadString(uri)
                Else
                    Return client.UploadString(uri, postData)
                End If
#End If
            End If
#If Not NET_1_1 Then
            Finally
                System.Net.ServicePointManager.ServerCertificateValidationCallback = CurrentValidationCallback
            End Try
#End If
        End Function

#If Not NET_1_1 Then
        ''' <summary>
        ''' Suppress all SSL certification requirements - just use the webservice SSL URL
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="cert"></param>
        ''' <param name="chain"></param>
        ''' <param name="errors"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public shared Function OnValidationCallback(ByVal sender As Object, ByVal cert As System.Security.Cryptography.X509Certificates.X509Certificate, ByVal chain As System.Security.Cryptography.X509Certificates.X509Chain, ByVal errors As System.Net.Security.SslPolicyErrors) As Boolean
            Return True
        End Function
#End If

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
            If TextWithoutLastLineBreak.EndsWith(ControlChars.Cr) Then
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