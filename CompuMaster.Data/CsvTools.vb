Option Explicit On 
Option Strict On

Imports System.IO
Imports System.Data
Imports CompuMaster.Data.Strings

Namespace CompuMaster.Data

    ''' <summary>
    '''     Provides simplified access to CSV files
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    <CodeAnalysis.SuppressMessage("Major Code Smell", "S3385:""Exit"" statements should not be used", Justification:="<Ausstehend>")>
    Friend Class CsvTools

#Region "Read data"

#Region "Fixed columns"
        ''' <summary>
        '''     Read from a CSV table
        ''' </summary>
        ''' <param name="reader">A stream reader targetting CSV data</param>
        ''' <param name="includesColumnHeaders">Indicates wether column headers are present</param>
        ''' <param name="columnWidths">An array of column widths in their order</param>
        ''' <param name="convertEmptyStringsToDBNull">Convert values with empty strings automatically to DbNull</param>
        ''' <param name="cultureFormatProvider">A culture for all conversions</param>
        ''' <param name="lineEncodings">Encoding style for linebreaks</param>
        ''' <param name="lineEncodingAutoConversions">Change linebreak encodings on reading</param>
        ''' <returns></returns>
        ''' <remarks>
        ''' In case of duplicate column names, all additional occurances of the same column name will be modified to use a unique column name
        ''' </remarks>
        Private Shared Function ReadDataTableFromCsvReader(ByVal reader As StreamReader, ByVal includesColumnHeaders As Boolean, startAtLineIndex As Integer, ByVal cultureFormatProvider As System.Globalization.CultureInfo, ByVal columnWidths As Integer(), ByVal convertEmptyStringsToDBNull As Boolean, lineEncodings As CompuMaster.Data.Csv.ReadLineEncodings, lineEncodingAutoConversions As CompuMaster.Data.Csv.ReadLineEncodingAutoConversion) As DataTable
            'Throw New NotSupportedException
            If cultureFormatProvider Is Nothing Then
#Disable Warning IDE0059 ' Unnötige Zuweisung eines Werts.
                cultureFormatProvider = System.Globalization.CultureInfo.InvariantCulture
#Enable Warning IDE0059 ' Unnötige Zuweisung eines Werts.
            End If
            If columnWidths Is Nothing Then
                columnWidths = New Integer() {Integer.MaxValue}
            End If

            Dim Result As New DataTable
            Dim rdStr As String
            Dim RowCounter As Integer

            'Read file content
            rdStr = reader.ReadToEnd
            If rdStr = Nothing Then
                'simply return the empty table when there is no input data
                Return Result
            End If

            'Read the file char by char and add row by row
            Dim CharPosition As Integer = 0
            For MyCounter As Integer = 0 To startAtLineIndex - 1
                CharPosition = rdStr.IndexOfAny(New Char() {ControlChars.Cr, ControlChars.Lf}, CharPosition) + 1
                If CharPosition > rdStr.Length Then
                    'simply return the empty table when there is no input data
                    Return Result
                End If
                If rdStr.Chars(CharPosition - 1) = ControlChars.Cr AndAlso rdStr.Chars(CharPosition) = ControlChars.Lf Then
                    CharPosition += 1
                End If
            Next

            While CharPosition < rdStr.Length

                'Read the next csv row
                Dim ColValues As New System.Collections.Generic.List(Of String)
                SplitFixedCsvLineIntoCellValues(rdStr, ColValues, CharPosition, columnWidths, lineEncodings, lineEncodingAutoConversions)

                'Add it as a new data row (respectively add the columns definition)
                RowCounter += 1
                If RowCounter = 1 AndAlso includesColumnHeaders Then
                    'Read first line as column names
                    For ColCounter As Integer = 0 To ColValues.Count - 1
                        Dim colName As String = Trim(CType(ColValues(ColCounter), String))
                        If Result.Columns.Contains(colName) Then
                            colName = String.Empty
                        End If
                        If Result.Columns.Contains(colName) = False Then
                            Result.Columns.Add(New DataColumn(colName, GetType(String)))
                        Else
                            Result.Columns.Add(New DataColumn(DataTables.LookupUniqueColumnName(Result, colName), GetType(String)))
                        End If
                    Next
                Else
                    'Read line as data and automatically add required additional columns on the fly
                    Dim MyRow As DataRow = Result.NewRow
                    For ColCounter As Integer = 0 To ColValues.Count - 1
                        Dim colValue As String = Trim(CType(ColValues(ColCounter), String))
                        If Result.Columns.Count <= ColCounter Then
                            Result.Columns.Add(New DataColumn(Nothing, GetType(String)))
                        End If
                        MyRow(ColCounter) = colValue
                    Next
                    Result.Rows.Add(MyRow)
                End If

            End While

            If convertEmptyStringsToDBNull Then
                ConvertEmptyStringsToDBNullValue(Result)
            Else
                ConvertDBNullValuesToEmptyStrings(Result)
            End If

            Return Result

        End Function

        ''' <summary>
        '''     Read from a CSV file
        ''' </summary>
        ''' <param name="path">The path of the file</param>
        ''' <param name="includesColumnHeaders">Indicates wether column headers are present</param>
        ''' <param name="columnWidths">An array of column widths in their order</param>
        ''' <param name="encoding">The text encoding of the file</param>
        ''' <param name="convertEmptyStringsToDBNull">Convert values with empty strings automatically to DbNull</param>
        ''' <param name="lineEncodings">Encoding style for linebreaks</param>
        ''' <param name="lineEncodingAutoConversions">Change linebreak encodings on reading</param>
        ''' <returns></returns>
        ''' <remarks>
        ''' In case of duplicate column names, all additional occurances of the same column name will be modified to use a unique column name
        ''' </remarks>
        Friend Shared Function ReadDataTableFromCsvFile(ByVal path As String, ByVal includesColumnHeaders As Boolean, startAtLineIndex As Integer, ByVal columnWidths As Integer(), ByVal encoding As String, ByVal convertEmptyStringsToDBNull As Boolean, lineEncodings As CompuMaster.Data.Csv.ReadLineEncodings, lineEncodingAutoConversions As CompuMaster.Data.Csv.ReadLineEncodingAutoConversion) As DataTable

            Dim Result As New DataTable

            If File.Exists(path) Then
                'do nothing for now
            ElseIf path.ToLowerInvariant.StartsWith("http://", StringComparison.Ordinal) OrElse path.ToLowerInvariant.StartsWith("https://", StringComparison.Ordinal) Then
                Dim LocalCopyOfFileContentFromRemoteUri As String = Utils.ReadStringDataFromUri(path, encoding)
                Result = ReadDataTableFromCsvString(LocalCopyOfFileContentFromRemoteUri, includesColumnHeaders, startAtLineIndex, columnWidths, convertEmptyStringsToDBNull, lineEncodings, lineEncodingAutoConversions)
                Result.TableName = System.IO.Path.GetFileNameWithoutExtension(path)
                Return Result
            Else
                Throw New System.IO.FileNotFoundException("File not found", path)
            End If

            Dim reader As StreamReader = Nothing
            Try
                If encoding = "" Then
                    reader = New StreamReader(path, System.Text.Encoding.Default)
                Else
                    reader = New StreamReader(path, System.Text.Encoding.GetEncoding(encoding))
                End If
                Result = ReadDataTableFromCsvReader(reader, includesColumnHeaders, startAtLineIndex, System.Globalization.CultureInfo.CurrentCulture, columnWidths, convertEmptyStringsToDBNull, lineEncodings, lineEncodingAutoConversions)
                Result.TableName = System.IO.Path.GetFileNameWithoutExtension(path)
            Finally
                reader?.Close()
            End Try

            Return Result

        End Function

        ''' <summary>
        '''     Read from a CSV file
        ''' </summary>
        ''' <param name="path">The path of the file</param>
        ''' <param name="includesColumnHeaders">Indicates wether column headers are present</param>
        ''' <param name="columnWidths">An array of column widths in their order</param>
        ''' <param name="encoding">The text encoding of the file</param>
        ''' <param name="cultureFormatProvider">A culture for all conversions</param>
        ''' <param name="convertEmptyStringsToDBNull">Convert values with empty strings automatically to DbNull</param>
        ''' <param name="lineEncodings">Encoding style for linebreaks</param>
        ''' <param name="lineEncodingAutoConversions">Change linebreak encodings on reading</param>
        ''' <returns></returns>
        ''' <remarks>
        ''' In case of duplicate column names, all additional occurances of the same column name will be modified to use a unique column name
        ''' </remarks>
        Friend Shared Function ReadDataTableFromCsvFile(ByVal path As String, ByVal includesColumnHeaders As Boolean, startAtLineIndex As Integer, ByVal columnWidths As Integer(), ByVal encoding As System.Text.Encoding, ByVal cultureFormatProvider As System.Globalization.CultureInfo, ByVal convertEmptyStringsToDBNull As Boolean, lineEncodings As CompuMaster.Data.Csv.ReadLineEncodings, lineEncodingAutoConversions As CompuMaster.Data.Csv.ReadLineEncodingAutoConversion) As DataTable

            Dim Result As New DataTable

            If File.Exists(path) Then
                'do nothing for now
            ElseIf path.ToLowerInvariant.StartsWith("http://", StringComparison.Ordinal) OrElse path.ToLowerInvariant.StartsWith("https://", StringComparison.Ordinal) Then
                Dim EncodingWebName As String
                If encoding Is Nothing Then
                    EncodingWebName = Nothing
                Else
                    EncodingWebName = encoding.WebName
                End If
                Dim LocalCopyOfFileContentFromRemoteUri As String = Utils.ReadStringDataFromUri(path, EncodingWebName)
                Result = ReadDataTableFromCsvString(LocalCopyOfFileContentFromRemoteUri, includesColumnHeaders, startAtLineIndex, columnWidths, convertEmptyStringsToDBNull, lineEncodings, lineEncodingAutoConversions)
                Result.TableName = System.IO.Path.GetFileNameWithoutExtension(path)
                Return Result
            Else
                Throw New System.IO.FileNotFoundException("File not found", path)
            End If

            Dim reader As StreamReader = Nothing
            Try
                If encoding Is Nothing Then
                    reader = New StreamReader(path, System.Text.Encoding.Default)
                Else
                    reader = New StreamReader(path, (encoding))
                End If
                Result = ReadDataTableFromCsvReader(reader, includesColumnHeaders, startAtLineIndex, cultureFormatProvider, columnWidths, convertEmptyStringsToDBNull, lineEncodings, lineEncodingAutoConversions)
                Result.TableName = System.IO.Path.GetFileNameWithoutExtension(path)
            Finally
                reader?.Close()
            End Try

            Return Result

        End Function

        ''' <summary>
        '''     Read from a CSV table
        ''' </summary>
        ''' <param name="data">The content of a CSV file</param>
        ''' <param name="includesColumnHeaders">Indicates wether column headers are present</param>
        ''' <param name="columnWidths">An array of column widths in their order</param>
        ''' <param name="convertEmptyStringsToDBNull">Convert values with empty strings automatically to DbNull</param>
        ''' <param name="lineEncodings">Encoding style for linebreaks</param>
        ''' <param name="lineEncodingAutoConversions">Change linebreak encodings on reading</param>
        ''' <returns></returns>
        ''' <remarks>
        ''' In case of duplicate column names, all additional occurances of the same column name will be modified to use a unique column name
        ''' </remarks>
        Friend Shared Function ReadDataTableFromCsvString(ByVal data As String, ByVal includesColumnHeaders As Boolean, startAtLineIndex As Integer, ByVal columnWidths As Integer(), ByVal convertEmptyStringsToDBNull As Boolean, lineEncodings As CompuMaster.Data.Csv.ReadLineEncodings, lineEncodingAutoConversions As CompuMaster.Data.Csv.ReadLineEncodingAutoConversion) As DataTable

            Dim Result As New DataTable
            Dim reader As StreamReader = Nothing
            Try
                reader = New StreamReader(New MemoryStream(System.Text.Encoding.Unicode.GetBytes(data)), System.Text.Encoding.Unicode, False)
                Result = ReadDataTableFromCsvReader(reader, includesColumnHeaders, startAtLineIndex, System.Globalization.CultureInfo.CurrentCulture, columnWidths, convertEmptyStringsToDBNull, lineEncodings, lineEncodingAutoConversions)
            Finally
                reader?.Close()
            End Try

            Return Result

        End Function

        ''' <summary>
        '''     Read from a CSV table
        ''' </summary>
        ''' <param name="data">The content of a CSV file</param>
        ''' <param name="includesColumnHeaders">Indicates wether column headers are present</param>
        ''' <param name="columnWidths">An array of column widths in their order</param>
        ''' <param name="convertEmptyStringsToDBNull">Convert values with empty strings automatically to DbNull</param>
        ''' <param name="cultureFormatProvider">A culture for all conversions</param>
        ''' <param name="lineEncodings">Encoding style for linebreaks</param>
        ''' <param name="lineEncodingAutoConversions">Change linebreak encodings on reading</param>
        ''' <returns></returns>
        ''' <remarks>
        ''' In case of duplicate column names, all additional occurances of the same column name will be modified to use a unique column name
        ''' </remarks>
        Friend Shared Function ReadDataTableFromCsvString(ByVal data As String, ByVal includesColumnHeaders As Boolean, startAtLineIndex As Integer, ByVal cultureFormatProvider As System.Globalization.CultureInfo, ByVal columnWidths As Integer(), ByVal convertEmptyStringsToDBNull As Boolean, lineEncodings As CompuMaster.Data.Csv.ReadLineEncodings, lineEncodingAutoConversions As CompuMaster.Data.Csv.ReadLineEncodingAutoConversion) As DataTable

            Dim Result As New DataTable
            Dim reader As StreamReader = Nothing
            Try
                reader = New StreamReader(New MemoryStream(System.Text.Encoding.Unicode.GetBytes(data)), System.Text.Encoding.Unicode, False)
                Result = ReadDataTableFromCsvReader(reader, includesColumnHeaders, startAtLineIndex, cultureFormatProvider, columnWidths, convertEmptyStringsToDBNull, lineEncodings, lineEncodingAutoConversions)
            Finally
                reader?.Close()
            End Try

            Return Result

        End Function

        ''' <summary>
        '''     Split a line content into separate column values and add them to the output list
        ''' </summary>
        ''' <param name="lineContent">The line content as it has been read from the CSV file</param>
        ''' <param name="outputList">An array list which shall hold the separated column values</param>
        ''' <param name="startPosition">The start position to which the columnWidhts are related to</param>
        ''' <param name="columnWidths">An array of column widths in their order</param>
        ''' <param name="lineEncodings">Encoding style for linebreaks</param>
        ''' <param name="lineEncodingAutoConversions">Change linebreak encodings on reading</param>
        ''' <remarks>
        ''' </remarks>
        Private Shared Sub SplitFixedCsvLineIntoCellValues(ByRef lineContent As String, ByVal outputList As System.Collections.Generic.List(Of String), ByRef startposition As Integer, ByVal columnWidths As Integer(), lineEncodings As CompuMaster.Data.Csv.ReadLineEncodings, lineEncodingAutoConversions As CompuMaster.Data.Csv.ReadLineEncodingAutoConversion)

            Dim CurrentColumnValue As System.Text.StringBuilder = Nothing
            Dim CharPositionCounter As Integer

            For CharPositionCounter = startposition To lineContent.Length - 1
                If CharPositionCounter = startposition Then
                    'Prepare the new value for the first column
                    CurrentColumnValue = New System.Text.StringBuilder
                ElseIf SplitFixedCsvLineIntoCellValuesIsNewColumnPosition(CharPositionCounter, startposition, columnWidths) Then
                    'A new column has been found
                    'Save the previous column value 
                    outputList.Add(CurrentColumnValue.ToString)
                    'Prepare the new value for  the next column
                    CurrentColumnValue = New System.Text.StringBuilder
                End If
                'TODO: consider line encoding arguments and support cell line breaks by considering the cell line break encoding
                Select Case lineContent.Chars(CharPositionCounter)
                    Case ControlChars.Lf
                        'now it's a line separator
                        Exit For
                    Case ControlChars.Cr
                        'now it's a line separator
                        If CharPositionCounter + 1 < lineContent.Length AndAlso lineContent.Chars(CharPositionCounter + 1) = ControlChars.Lf Then
                            'Found a CrLf occurance; handle it as one line break!
                            CharPositionCounter += 1
                        End If
                        Exit For
                    Case Else
                        'just add the character as it is because it's inside of a cell text
                        CurrentColumnValue.Append(lineContent.Chars(CharPositionCounter))
                End Select
            Next

            'Add the last column value to the collection
            If CurrentColumnValue IsNot Nothing AndAlso CurrentColumnValue.Length <> 0 Then
                outputList.Add(CurrentColumnValue.ToString)
            End If

            'Next start position is the next char after the last read one
            startposition = CharPositionCounter + 1

        End Sub

        ''' <summary>
        '''     Calculate if the current position is the first position of a new column
        ''' </summary>
        ''' <param name="currentPosition">The current position in the whole document</param>
        ''' <param name="startPosition">The start position to which the columnWidhts are related to</param>
        ''' <param name="columnWidths">An array containing the definitions of the column widths</param>
        ''' <returns>True if the current position identifies a new column value, otherwise False</returns>
        Private Shared Function SplitFixedCsvLineIntoCellValuesIsNewColumnPosition(ByVal currentPosition As Integer, ByVal startPosition As Integer, ByVal columnWidths As Integer()) As Boolean
            Dim positionDifference As Integer = currentPosition - startPosition
            For MyCounter As Integer = 0 To columnWidths.Length - 1
                Dim ColumnStartPosition As Integer
                ColumnStartPosition += columnWidths(MyCounter)
                If positionDifference = ColumnStartPosition Then
                    Return True
                End If
            Next
            Return False
        End Function

        'Private Shared Function SumOfIntegerValues(ByVal array As Integer(), ByVal sumUpToElementIndex As Integer) As Integer
        '    Dim Result As Integer
        '    For MyCounter As Integer = 0 To sumUpToElementIndex
        '        Result += array(MyCounter)
        '    Next
        '    Return Result
        'End Function
#End Region

#Region "Separator separation"

        ''' <summary>
        '''     Read from a CSV file
        ''' </summary>
        ''' <param name="path">The path of the file</param>
        ''' <param name="includesColumnHeaders">Indicates wether column headers are present</param>
        ''' <param name="encoding">The text encoding of the file</param>
        ''' <param name="columnSeparator">Choose the required character for splitting the columns. Set to null (Nothing in VisualBasic) to enable fixed column widths mode</param>
        ''' <param name="recognizeTextBy">A character indicating the start and end of text strings</param>
        ''' <param name="recognizeMultipleColumnSeparatorCharsAsOne">Specifies whether we should treat multiple column seperators as one</param>
        ''' <param name="convertEmptyStringsToDBNull">Convert values with empty strings automatically to DbNull</param>
        ''' <param name="lineEncodings">Encoding style for linebreaks</param>
        ''' <param name="lineEncodingAutoConversions">Change linebreak encodings on reading</param>
        ''' <returns></returns>
        ''' <remarks>
        ''' In case of duplicate column names, all additional occurances of the same column name will be modified to use a unique column name
        ''' </remarks>
        Friend Shared Function ReadDataTableFromCsvFile(ByVal path As String, ByVal includesColumnHeaders As Boolean, startAtLineIndex As Integer, ByVal encoding As String, ByVal columnSeparator As Char, ByVal recognizeTextBy As Char, ByVal recognizeMultipleColumnSeparatorCharsAsOne As Boolean, ByVal convertEmptyStringsToDBNull As Boolean, lineEncodings As CompuMaster.Data.Csv.ReadLineEncodings, lineEncodingAutoConversions As CompuMaster.Data.Csv.ReadLineEncodingAutoConversion) As DataTable

            Dim Result As New DataTable

            If File.Exists(path) Then
                'do nothing for now
            ElseIf path.ToLowerInvariant.StartsWith("http://", StringComparison.Ordinal) OrElse path.ToLowerInvariant.StartsWith("https://", StringComparison.Ordinal) Then
                Dim LocalCopyOfFileContentFromRemoteUri As String = Utils.ReadStringDataFromUri(path, encoding)
                Result = ReadDataTableFromCsvString(LocalCopyOfFileContentFromRemoteUri, includesColumnHeaders, startAtLineIndex, columnSeparator, recognizeTextBy, recognizeMultipleColumnSeparatorCharsAsOne, convertEmptyStringsToDBNull, lineEncodings, lineEncodingAutoConversions)
                Result.TableName = System.IO.Path.GetFileNameWithoutExtension(path)
                Return Result
            Else
                Throw New System.IO.FileNotFoundException("File not found", path)
            End If

            Dim fs As FileStream = Nothing
            Dim reader As StreamReader = Nothing
            Try
                fs = New FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
                If encoding = "" Then
                    reader = New StreamReader(fs, System.Text.Encoding.Default)
                Else
                    reader = New StreamReader(fs, System.Text.Encoding.GetEncoding(encoding))
                End If
                Result = ReadDataTableFromCsvReader(reader, includesColumnHeaders, startAtLineIndex, System.Globalization.CultureInfo.CurrentCulture, columnSeparator, recognizeTextBy, recognizeMultipleColumnSeparatorCharsAsOne, convertEmptyStringsToDBNull, lineEncodings, lineEncodingAutoConversions)
                Result.TableName = System.IO.Path.GetFileNameWithoutExtension(path)
            Finally
                If reader IsNot Nothing Then
                    reader.Close()
                    reader.Dispose()
                End If
                If fs IsNot Nothing Then
                    fs.Close()
                    fs.Dispose()
                End If
            End Try

            Return Result

        End Function

        ''' <summary>
        '''     Read from a CSV file
        ''' </summary>
        ''' <param name="Path">The path of the file</param>
        ''' <param name="IncludesColumnHeaders">Indicates wether column headers are present</param>
        ''' <param name="Encoding">The text encoding of the file</param>
        ''' <param name="RecognizeTextBy">A character indicating the start and end of text strings</param>
        ''' <param name="recognizeMultipleColumnSeparatorCharsAsOne">Specifies whether we should treat multiple column seperators as one</param>
        ''' <param name="convertEmptyStringsToDBNull">Convert values with empty strings automatically to DbNull</param>
        ''' <param name="cultureFormatProvider">A culture for all conversions</param>
        ''' <param name="lineEncodings">Encoding style for linebreaks</param>
        ''' <param name="lineEncodingAutoConversions">Change linebreak encodings on reading</param>
        ''' <returns></returns>
        ''' <remarks>
        ''' In case of duplicate column names, all additional occurances of the same column name will be modified to use a unique column name
        ''' </remarks>
        Friend Shared Function ReadDataTableFromCsvFile(ByVal path As String, ByVal includesColumnHeaders As Boolean, startAtLineIndex As Integer, ByVal encoding As System.Text.Encoding, ByVal cultureFormatProvider As System.Globalization.CultureInfo, ByVal recognizeTextBy As Char, ByVal recognizeMultipleColumnSeparatorCharsAsOne As Boolean, ByVal convertEmptyStringsToDBNull As Boolean, lineEncodings As CompuMaster.Data.Csv.ReadLineEncodings, lineEncodingAutoConversions As CompuMaster.Data.Csv.ReadLineEncodingAutoConversion) As DataTable

            Dim Result As New DataTable

            If File.Exists(path) Then
                'do nothing for now
            ElseIf path.ToLowerInvariant.StartsWith("http://", StringComparison.Ordinal) OrElse path.ToLowerInvariant.StartsWith("https://", StringComparison.Ordinal) Then
                Dim EncodingWebName As String
                If encoding Is Nothing Then
                    EncodingWebName = Nothing
                Else
                    EncodingWebName = encoding.WebName
                End If
                Dim LocalCopyOfFileContentFromRemoteUri As String = Utils.ReadStringDataFromUri(path, EncodingWebName)
                Result = ReadDataTableFromCsvString(LocalCopyOfFileContentFromRemoteUri, includesColumnHeaders, startAtLineIndex, cultureFormatProvider, recognizeTextBy, recognizeMultipleColumnSeparatorCharsAsOne, convertEmptyStringsToDBNull, lineEncodings, lineEncodingAutoConversions)
                Result.TableName = System.IO.Path.GetFileNameWithoutExtension(path)
                Return Result
            Else
                Throw New System.IO.FileNotFoundException("File not found", path)
            End If

            Dim reader As StreamReader = Nothing
            Try
                If encoding Is Nothing Then
                    reader = New StreamReader(path, System.Text.Encoding.Default)
                Else
                    reader = New StreamReader(path, encoding)
                End If
                Result = ReadDataTableFromCsvReader(reader, includesColumnHeaders, startAtLineIndex, cultureFormatProvider, Nothing, recognizeTextBy, recognizeMultipleColumnSeparatorCharsAsOne, convertEmptyStringsToDBNull, lineEncodings, lineEncodingAutoConversions)
                Result.TableName = System.IO.Path.GetFileNameWithoutExtension(path)
            Finally
                reader?.Close()
            End Try

            Return Result

        End Function

        ''' <summary>
        '''     Read from a CSV table
        ''' </summary>
        ''' <param name="reader">A stream reader targetting CSV data</param>
        ''' <param name="includesColumnHeaders">Indicates wether column headers are present</param>
        ''' <param name="columnSeparator">Choose the required character for splitting the columns. Set to null (Nothing in VisualBasic) to enable fixed column widths mode</param>
        ''' <param name="recognizeTextBy">A character indicating the start and end of text strings</param>
        ''' <param name="recognizeMultipleColumnSeparatorCharsAsOne">Specifies whether we should treat multiple column seperators as one</param>
        ''' <param name="convertEmptyStringsToDBNull">Convert values with empty strings automatically to DbNull</param>
        ''' <param name="cultureFormatProvider">A culture for all conversions</param>
        ''' <param name="lineEncodings">Encoding style for linebreaks</param>
        ''' <param name="lineEncodingAutoConversions">Change linebreak encodings on reading</param>
        ''' <param name="startAtLineIndex">Start reading of table data at specified line index (most often 0 for very first line)</param>
        ''' <returns></returns>
        ''' <remarks>
        ''' In case of duplicate column names, all additional occurances of the same column name will be modified to use a unique column name
        ''' </remarks>
        Private Shared Function ReadDataTableFromCsvReader(ByVal reader As StreamReader, ByVal includesColumnHeaders As Boolean, startAtLineIndex As Integer, ByVal cultureFormatProvider As System.Globalization.CultureInfo, ByVal columnSeparator As Char, ByVal recognizeTextBy As Char, ByVal recognizeMultipleColumnSeparatorCharsAsOne As Boolean, ByVal convertEmptyStringsToDBNull As Boolean, lineEncodings As CompuMaster.Data.Csv.ReadLineEncodings, lineEncodingAutoConversions As CompuMaster.Data.Csv.ReadLineEncodingAutoConversion) As DataTable

            If cultureFormatProvider Is Nothing Then
                cultureFormatProvider = System.Globalization.CultureInfo.InvariantCulture
            End If

            If columnSeparator = Nothing OrElse columnSeparator = vbNullChar Then
                'Attention: list separator is a string, but columnSeparator is implemented as char! Might be a bug in some specal cultures
                If cultureFormatProvider.TextInfo.ListSeparator.Length > 1 Then
                    Throw New NotSupportedException("No column separator has been defined and the current culture declares a list separator with more than 1 character. Column separators with more than 1 characters are currenlty not supported.")
                End If
                columnSeparator = cultureFormatProvider.TextInfo.ListSeparator.Chars(0)
            End If

            Dim Result As New DataTable
            Dim rdStr As String
            Dim RowCounter As Integer
            Dim detectCompletedRowLineBasedOnRequiredColumnCount As Integer = 0
            If lineEncodings = Csv.ReadLineEncodings.Auto Then
                Select Case System.Environment.NewLine
                    Case ControlChars.CrLf
                        'Windows platforms
                        lineEncodings = Csv.ReadLineEncodings.RowBreakCrLf_CellLineBreakLf
                    Case ControlChars.Cr
                        'Mac platforms
                        lineEncodings = Csv.ReadLineEncodings.RowBreakCr_CellLineBreakLf
                    Case ControlChars.Lf
                        'Linux platforms
                        lineEncodings = Csv.ReadLineEncodings.RowBreakLf_CellLineBreakCr
                    Case Else
                        Throw New NotImplementedException
                End Select
            End If
            If lineEncodings = Csv.ReadLineEncodings.RowBreakCrLfOrCrOrLf_CellLineBreakCrLfOrCrOrLf AndAlso includesColumnHeaders = False Then
                Throw New ArgumentException("Line endings setting RowBreakCrLfOrCrOrLf_CellLineBreakCrLfOrCrOrLf requires the CSV data to provide column headers")
            End If

            'Read file content
            rdStr = reader.ReadToEnd 'WARNING: might cause System.OutOfMemoryException on too large files
            If rdStr = Nothing Then
                'simply return the empty table when there is no input data
                Return Result
            End If

            'Read the file char by char and add row by row
            Dim CharPosition As Integer = 0
            While CharPosition < rdStr.Length

#Disable Warning S1066 ' Collapsible "if" statements should be merged
                If lineEncodings = Csv.ReadLineEncodings.RowBreakCrLfOrCrOrLf_CellLineBreakCrLfOrCrOrLf AndAlso detectCompletedRowLineBasedOnRequiredColumnCount = 0 Then
                    If RowCounter <> 0 Then 'already includesColumnHeaders required since lineEncodings check on method head
                        Throw New ArgumentNullException("detectCompletedRowLineBasedOnRequiredColumnCount", "Argument detectCompletedRowLineBasedOnRequiredColumnCount required for reading with line endings RowBreakCrLfOrCrOrLf_CellLineBreakCrLfOrCrOrLf")
                    End If
                End If
#Enable Warning S1066 ' Collapsible "if" statements should be merged

                'Read the next csv row
                Dim ColValues As New List(Of String)
                Dim CurrentRowStartPosition As Integer = CharPosition
                SplitCsvLineIntoCellValues(rdStr, ColValues, If(CharPosition = 0, startAtLineIndex, 0), CharPosition, columnSeparator, recognizeTextBy, recognizeMultipleColumnSeparatorCharsAsOne, lineEncodings, lineEncodingAutoConversions, detectCompletedRowLineBasedOnRequiredColumnCount)

                'Add it as a new data row (respectively add the columns definition)
                RowCounter += 1
                If RowCounter = 1 AndAlso includesColumnHeaders Then
                    'Read first line as column names
                    For ColCounter As Integer = 0 To ColValues.Count - 1
                        Dim colName As String = Trim(ColValues(ColCounter))
                        If Result.Columns.Contains(colName) = False Then
                            Result.Columns.Add(New DataColumn(colName, GetType(String)))
                        Else
                            Result.Columns.Add(New DataColumn(DataTables.LookupUniqueColumnName(Result, colName), GetType(String)))
                        End If
                    Next
                    'Save current column count
                    detectCompletedRowLineBasedOnRequiredColumnCount = Result.Columns.Count
                Else
                    'Read line as data and automatically add required additional columns on the fly
                    Dim MyRow As DataRow = Result.NewRow
                    For ColCounter As Integer = 0 To ColValues.Count - 1
                        Dim colValue As String = Trim(ColValues(ColCounter))
                        If Result.Columns.Count <= ColCounter Then
                            If lineEncodings = Csv.ReadLineEncodings.RowBreakCrLfOrCrOrLf_CellLineBreakCrLfOrCrOrLf Then
                                Throw New InvalidOperationException("Line endings setting RowBreakCrLfOrCrOrLf_CellLineBreakCrLfOrCrOrLf requires the CSV data to provide the same column count in each row: error reading record row " & Result.Rows.Count + 1 & " and cell """ & colValue & """ - full raw row data:" & System.Environment.NewLine & rdStr.Substring(CurrentRowStartPosition, CharPosition - CurrentRowStartPosition))
                            Else
                                Result.Columns.Add(New DataColumn(Nothing, GetType(String)))
                            End If
                        End If
                        MyRow(ColCounter) = colValue
                    Next
                    Result.Rows.Add(MyRow)
                End If

            End While

            If convertEmptyStringsToDBNull Then
                ConvertEmptyStringsToDBNullValue(Result)
            Else
                ConvertDBNullValuesToEmptyStrings(Result)
            End If

            Return Result

        End Function

        ''' <summary>
        '''     Split a line content into separate column values and add them to the output list
        ''' </summary>
        ''' <param name="lineContent">The line content as it has been read from the CSV file</param>
        ''' <param name="outputList">An array list which shall hold the separated column values</param>
        ''' <param name="startAtLineIndex">Start reading of table data at specified line index (most often 0 for very first line)</param>
        ''' <param name="startposition">An index for the start position</param>
        ''' <param name="columnSeparator">Choose the required character for splitting the columns. Set to null (Nothing in VisualBasic) to enable fixed column widths mode</param>
        ''' <param name="recognizeTextBy">A character indicating the start and end of text string</param>
        ''' <param name="recognizeMultipleColumnSeparatorCharsAsOne">Specifies whether we should treat multiple column seperators as one</param>
        ''' <param name="detectCompletedRowLineBasedOnRequiredColumnCount">When reading CSV files with equal line break and cell break encoding, detect full row lines by column count</param>
        ''' <param name="lineEncodings">Encoding style for linebreaks</param>
        ''' <param name="lineEncodingAutoConversions">Change linebreak encodings on reading</param>
        Private Shared Sub SplitCsvLineIntoCellValues(ByRef lineContent As String, ByVal outputList As List(Of String), startAtLineIndex As Integer, ByRef startposition As Integer, ByVal columnSeparator As Char, ByVal recognizeTextBy As Char, ByVal recognizeMultipleColumnSeparatorCharsAsOne As Boolean, lineEncodings As CompuMaster.Data.Csv.ReadLineEncodings, lineEncodingAutoConversions As CompuMaster.Data.Csv.ReadLineEncodingAutoConversion, detectCompletedRowLineBasedOnRequiredColumnCount As Integer)
            Dim CurrentColumnValue As New System.Text.StringBuilder
            Dim InQuotationMarks As Boolean
            Dim CharPositionCounter As Integer
            If lineEncodings = Csv.ReadLineEncodings.Auto Then
                Select Case System.Environment.NewLine
                    Case ControlChars.CrLf
                        'Windows platforms
                        lineEncodings = Csv.ReadLineEncodings.RowBreakCrLf_CellLineBreakLf
                    Case ControlChars.Cr
                        'Mac platforms
                        lineEncodings = Csv.ReadLineEncodings.RowBreakCr_CellLineBreakLf
                    Case ControlChars.Lf
                        'Linux platforms
                        lineEncodings = Csv.ReadLineEncodings.RowBreakLf_CellLineBreakCr
                    Case Else
                        Throw New NotImplementedException
                End Select
            End If

            Dim CurrentRowIndex As Integer = 0
            For CharPositionCounter = startposition To lineContent.Length - 1
                Dim IsIgnoreLine As Boolean = (CurrentRowIndex < startAtLineIndex)
                Select Case lineContent.Chars(CharPositionCounter)
                    Case columnSeparator
                        If IsIgnoreLine Then
                            'do nothing
                        ElseIf InQuotationMarks Then
                            'just add the character as it is because it's inside of a cell text
                            CurrentColumnValue.Append(lineContent.Chars(CharPositionCounter))
                        Else
                            'now it's a column separator
                            'implementation follows to the handling of recognizeMultipleColumnSeparatorCharsAsOne as Excel does
                            If Not (recognizeMultipleColumnSeparatorCharsAsOne = True AndAlso lineContent.Chars(CharPositionCounter - 1) = columnSeparator) Then
                                outputList.Add(CurrentColumnValue.ToString)
                                CurrentColumnValue = New System.Text.StringBuilder
                            End If
                        End If
                    Case recognizeTextBy
                        If InQuotationMarks = False Then
                            InQuotationMarks = Not InQuotationMarks
                        Else
                            'Switch between state of in- our out-of quotation marks
                            If CharPositionCounter + 1 < lineContent.Length AndAlso lineContent.Chars(CharPositionCounter + 1) = recognizeTextBy Then
                                'doubled quotation marks lead to one single quotation mark
                                If IsIgnoreLine Then
                                    'do nothing
                                Else
                                    CurrentColumnValue.Append(recognizeTextBy)
                                End If
                                'fix the position to be now after the second quotation marks
                                CharPositionCounter += 1
                            Else
                                InQuotationMarks = Not InQuotationMarks
                            End If
                        End If
                    Case ControlChars.Lf
                        If InQuotationMarks OrElse (lineEncodings = Csv.ReadLineEncodings.RowBreakCrLfOrCrOrLf_CellLineBreakCrLfOrCrOrLf AndAlso outputList.Count < detectCompletedRowLineBasedOnRequiredColumnCount - 1) Then
                            'it's a line separator within logical cell
                            If IsIgnoreLine Then
                                'do nothing
                            Else
                                'TODO: read cell line breaks correctly
                                'Select Case lineEncodings
                                '    Case Csv.ReadLineEncodings.None
                                '    Case Csv.ReadLineEncodings.RowBreakCrLfOrCr_CellLineBreakLf
                                '    Case Csv.ReadLineEncodings.RowBreakCrLfOrLf_CellLineBreakCr
                                '    Case Csv.ReadLineEncodings.RowBreakCrLf_CellLineBreakCr
                                '    Case Csv.ReadLineEncodings.RowBreakCrLf_CellLineBreakLf
                                '    Case Csv.ReadLineEncodings.RowBreakCr_CellLineBreakLf
                                '    Case Csv.ReadLineEncodings.RowBreakLf_CellLineBreakCr
                                '    Case Else
                                '        Throw New NotImplementedException("Invalid lineEncoding")
                                'End Select

                                'just add the line-break because it's inside of a cell text
                                Select Case lineEncodingAutoConversions
                                    Case Csv.ReadLineEncodingAutoConversion.NoAutoConversion
                                        CurrentColumnValue.Append(lineContent.Chars(CharPositionCounter))
                                    Case Csv.ReadLineEncodingAutoConversion.AutoConvertLineBreakToCrLf
                                        CurrentColumnValue.Append(ControlChars.CrLf)
                                    Case Csv.ReadLineEncodingAutoConversion.AutoConvertLineBreakToCr
                                        CurrentColumnValue.Append(ControlChars.Cr)
                                    Case Csv.ReadLineEncodingAutoConversion.AutoConvertLineBreakToLf
                                        CurrentColumnValue.Append(ControlChars.Lf)
                                    Case Csv.ReadLineEncodingAutoConversion.AutoConvertLineBreakToSystemEnvironmentNewLine
                                        CurrentColumnValue.Append(System.Environment.NewLine)
                                    Case Else
                                        Throw New NotImplementedException("Invalid lineEncoding")
                                End Select
                            End If
                        Else
                            'now it's a row line separator

                            'TODO: read cell line breaks correctly
                            'Select Case lineEncodings
                            '    Case Csv.ReadLineEncodings.None
                            '    Case Csv.ReadLineEncodings.RowBreakCrLfOrCr_CellLineBreakLf
                            '    Case Csv.ReadLineEncodings.RowBreakCrLfOrLf_CellLineBreakCr
                            '    Case Csv.ReadLineEncodings.RowBreakCrLf_CellLineBreakCr
                            '    Case Csv.ReadLineEncodings.RowBreakCrLf_CellLineBreakLf
                            '    Case Csv.ReadLineEncodings.RowBreakCr_CellLineBreakLf
                            '    Case Csv.ReadLineEncodings.RowBreakLf_CellLineBreakCr
                            '    Case Else
                            '        Throw New NotImplementedException("Invalid lineEncoding")
                            'End Select

                            If IsIgnoreLine Then
                                'effectively, do nothing - except for counting up already-ignored lines
                                CurrentRowIndex += 1
                            Else
                                'Add previously collected data as column value
                                outputList.Add(CurrentColumnValue.ToString)
                                CurrentColumnValue = New System.Text.StringBuilder
                                'Leave this method because the reading of one csv row has been completed
                                Exit For
                            End If
                        End If
                    Case ControlChars.Cr
                        If InQuotationMarks OrElse (lineEncodings = Csv.ReadLineEncodings.RowBreakCrLfOrCrOrLf_CellLineBreakCrLfOrCrOrLf AndAlso outputList.Count < detectCompletedRowLineBasedOnRequiredColumnCount - 1) Then
                            'it's a line separator within logical cell
                            If IsIgnoreLine Then
                                'do nothing
                            Else
                                'TODO: read cell line breaks correctly
                                'Select Case lineEncodings
                                '    Case Csv.ReadLineEncodings.None
                                '    Case Csv.ReadLineEncodings.RowBreakCrLfOrCr_CellLineBreakLf
                                '    Case Csv.ReadLineEncodings.RowBreakCrLfOrLf_CellLineBreakCr
                                '    Case Csv.ReadLineEncodings.RowBreakCrLf_CellLineBreakCr
                                '    Case Csv.ReadLineEncodings.RowBreakCrLf_CellLineBreakLf
                                '    Case Csv.ReadLineEncodings.RowBreakCr_CellLineBreakLf
                                '    Case Csv.ReadLineEncodings.RowBreakLf_CellLineBreakCr
                                '    Case Else
                                '        Throw New NotImplementedException("Invalid lineEncoding")
                                'End Select

                                'just add the character as it is because it's inside of a cell text
                                Select Case lineEncodingAutoConversions
                                    Case Csv.ReadLineEncodingAutoConversion.NoAutoConversion
                                        CurrentColumnValue.Append(lineContent.Chars(CharPositionCounter))
                                    Case Csv.ReadLineEncodingAutoConversion.AutoConvertLineBreakToCrLf
                                        CurrentColumnValue.Append(ControlChars.CrLf)
                                    Case Csv.ReadLineEncodingAutoConversion.AutoConvertLineBreakToCr
                                        CurrentColumnValue.Append(ControlChars.Cr)
                                    Case Csv.ReadLineEncodingAutoConversion.AutoConvertLineBreakToLf
                                        CurrentColumnValue.Append(ControlChars.Lf)
                                    Case Csv.ReadLineEncodingAutoConversion.AutoConvertLineBreakToSystemEnvironmentNewLine
                                        CurrentColumnValue.Append(System.Environment.NewLine)
                                    Case Else
                                        Throw New NotImplementedException("Invalid lineEncoding")
                                End Select
                            End If
                        Else
                            'now it's a row line separator

                            'TODO: read cell line breaks correctly
                            'Select Case lineEncodings
                            '    Case Csv.ReadLineEncodings.None
                            '    Case Csv.ReadLineEncodings.RowBreakCrLfOrCr_CellLineBreakLf
                            '    Case Csv.ReadLineEncodings.RowBreakCrLfOrLf_CellLineBreakCr
                            '    Case Csv.ReadLineEncodings.RowBreakCrLf_CellLineBreakCr
                            '    Case Csv.ReadLineEncodings.RowBreakCrLf_CellLineBreakLf
                            '    Case Csv.ReadLineEncodings.RowBreakCr_CellLineBreakLf
                            '    Case Csv.ReadLineEncodings.RowBreakLf_CellLineBreakCr
                            '    Case Else
                            '        Throw New NotImplementedException("Invalid lineEncoding")
                            'End Select

                            If CharPositionCounter + 1 < lineContent.Length AndAlso lineContent.Chars(CharPositionCounter + 1) = ControlChars.Lf Then
                                'Found a CrLf occurance; handle it as one line break!
                                CharPositionCounter += 1
                            End If
                            If IsIgnoreLine Then
                                'effectively, do nothing - except for counting up already-ignored lines
                                CurrentRowIndex += 1
                            Else
                                'Add previously collected data as column value
                                outputList.Add(CurrentColumnValue.ToString)
                                CurrentColumnValue = New System.Text.StringBuilder
                                'Leave this method because the reading of one csv row has been completed
                                Exit For
                            End If
                        End If
                    Case Else
                        If IsIgnoreLine Then
                            'do nothing
                        Else
                            'just add the character as it is because it's inside of a cell text
                            CurrentColumnValue.Append(lineContent.Chars(CharPositionCounter))
                        End If
                End Select
            Next

            'Add the last column value to the collection
            If CurrentColumnValue.Length <> 0 Then
                outputList.Add(CurrentColumnValue.ToString)
            End If

            'Next start position is the next char after the last read one
            startposition = CharPositionCounter + 1

        End Sub

        ''' <summary>
        '''     Read from a CSV table
        ''' </summary>
        ''' <param name="data">The content of a CSV file</param>
        ''' <param name="includesColumnHeaders">Indicates wether column headers are present</param>
        ''' <param name="columnSeparator">Choose the required character for splitting the columns. Set to null (Nothing in VisualBasic) to enable fixed column widths mode</param>
        ''' <param name="recognizeTextBy">A character indicating the start and end of text strings</param>
        ''' <param name="recognizeMultipleColumnSeparatorCharsAsOne">Currently without purpose</param>
        ''' <param name="convertEmptyStringsToDBNull">Convert values with empty strings automatically to DbNull</param>
        ''' <param name="lineEncodings">Encoding style for linebreaks</param>
        ''' <param name="lineEncodingAutoConversions">Change linebreak encodings on reading</param>
        ''' <returns></returns>
        ''' <remarks>
        ''' In case of duplicate column names, all additional occurances of the same column name will be modified to use a unique column name
        ''' </remarks>
        Friend Shared Function ReadDataTableFromCsvString(ByVal data As String, ByVal includesColumnHeaders As Boolean, startAtLineIndex As Integer, ByVal columnSeparator As Char, ByVal recognizeTextBy As Char, ByVal recognizeMultipleColumnSeparatorCharsAsOne As Boolean, ByVal convertEmptyStringsToDBNull As Boolean, lineEncodings As CompuMaster.Data.Csv.ReadLineEncodings, lineEncodingAutoConversions As CompuMaster.Data.Csv.ReadLineEncodingAutoConversion) As DataTable

            Dim Result As New DataTable
            Dim reader As StreamReader = Nothing
            Try
                reader = New StreamReader(New MemoryStream(System.Text.Encoding.Unicode.GetBytes(data)), System.Text.Encoding.Unicode, False)
                Result = ReadDataTableFromCsvReader(reader, includesColumnHeaders, startAtLineIndex, System.Globalization.CultureInfo.CurrentCulture, columnSeparator, recognizeTextBy, recognizeMultipleColumnSeparatorCharsAsOne, convertEmptyStringsToDBNull, lineEncodings, lineEncodingAutoConversions)
            Finally
                reader?.Close()
            End Try

            Return Result

        End Function

        ''' <summary>
        '''     Read from a CSV table
        ''' </summary>
        ''' <param name="data">The content of a CSV file</param>
        ''' <param name="IncludesColumnHeaders">Indicates wether column headers are present</param>
        ''' <param name="cultureFormatProvider">A culture for all conversions</param>
        ''' <param name="RecognizeTextBy">A character indicating the start and end of text strings</param>
        ''' <param name="recognizeMultipleColumnSeparatorCharsAsOne">Currently without purpose</param>
        ''' <param name="convertEmptyStringsToDBNull">Convert values with empty strings automatically to DbNull</param>
        ''' <param name="lineEncodings">Encoding style for linebreaks</param>
        ''' <param name="lineEncodingAutoConversions">Change linebreak encodings on reading</param>
        ''' <returns></returns>
        ''' <remarks>
        ''' In case of duplicate column names, all additional occurances of the same column name will be modified to use a unique column name
        ''' </remarks>
        Friend Shared Function ReadDataTableFromCsvString(ByVal data As String, ByVal includesColumnHeaders As Boolean, startAtLineIndex As Integer, ByVal cultureFormatProvider As System.Globalization.CultureInfo, ByVal recognizeTextBy As Char, ByVal recognizeMultipleColumnSeparatorCharsAsOne As Boolean, ByVal convertEmptyStringsToDBNull As Boolean, lineEncodings As CompuMaster.Data.Csv.ReadLineEncodings, lineEncodingAutoConversions As CompuMaster.Data.Csv.ReadLineEncodingAutoConversion) As DataTable

            Dim Result As New DataTable
            Dim reader As StreamReader = Nothing
            Try
                reader = New StreamReader(New MemoryStream(System.Text.Encoding.Unicode.GetBytes(data)), System.Text.Encoding.Unicode, False)
                Result = ReadDataTableFromCsvReader(reader, includesColumnHeaders, startAtLineIndex, cultureFormatProvider, Nothing, recognizeTextBy, recognizeMultipleColumnSeparatorCharsAsOne, convertEmptyStringsToDBNull, lineEncodings, lineEncodingAutoConversions)
            Finally
                reader?.Close()
            End Try

            Return Result

        End Function

        ''' <summary>
        '''     Convert DBNull values to empty strings
        ''' </summary>
        ''' <param name="data">The data which might contain DBNull values</param>
        Private Shared Sub ConvertDBNullValuesToEmptyStrings(ByVal data As DataTable)

            'Parameter validation
            If data Is Nothing Then
                Throw New ArgumentNullException(NameOf(data))
            End If

            'Ensure that only string columns are here
            For ColCounter As Integer = 0 To data.Columns.Count - 1
                If data.Columns(ColCounter).DataType IsNot GetType(String) Then
                    Throw New InvalidOperationException("All columns must be of data type System.String")
                End If
            Next

            'Update content
            For RowCounter As Integer = 0 To data.Rows.Count - 1
                Dim MyRow As DataRow = data.Rows(RowCounter)
                For ColCounter As Integer = 0 To data.Columns.Count - 1
                    If MyRow(ColCounter).GetType Is GetType(DBNull) Then
                        MyRow(ColCounter) = ""
                    End If
                Next
            Next

        End Sub

        ''' <summary>
        '''     Convert empty string values to DBNull
        ''' </summary>
        ''' <param name="data">The data which might contain empty strings</param>
        Private Shared Sub ConvertEmptyStringsToDBNullValue(ByVal data As DataTable)

            'Parameter validation
            If data Is Nothing Then
                Throw New ArgumentNullException(NameOf(data))
            End If

            'Ensure that only string columns are here
            For ColCounter As Integer = 0 To data.Columns.Count - 1
                If data.Columns(ColCounter).DataType IsNot GetType(String) Then
                    Throw New InvalidOperationException("All columns must be of data type System.String")
                End If
            Next

            'Update content
            For RowCounter As Integer = 0 To data.Rows.Count - 1
                Dim MyRow As DataRow = data.Rows(RowCounter)
                For ColCounter As Integer = 0 To data.Columns.Count - 1
                    Try
                        If MyRow(ColCounter).GetType Is GetType(String) AndAlso CType(MyRow(ColCounter), String) = "" Then
                            MyRow(ColCounter) = DBNull.Value
                        End If
                    Catch
                        'Ignore any conversion errors since we only want to change string columns
                    End Try
                Next
            Next

        End Sub
#End Region

#End Region

#Region "Write data"
        Private Const WriteStandardBlockSizeInChars As Integer = CInt(2 ^ 23) 'Either the remaining string or the next 2^23 chars = 8 M chars = 16 MB in RAM (1 unicode-char = 2 bytes in RAM)

        Private Shared Sub WriteTextStringBuilderToStreamWriter(writer As StreamWriter, textStringBuilder As System.Text.StringBuilder)
            Dim byteIndexWritten As Integer = 0
            Do
                Dim bytesToWrite As Integer = textStringBuilder.Length - byteIndexWritten
                bytesToWrite = System.Math.Min(bytesToWrite, WriteStandardBlockSizeInChars) 'Either the remaining string or the next full block size
                writer.Write(textStringBuilder.ToString(byteIndexWritten, bytesToWrite))
                byteIndexWritten += bytesToWrite
            Loop While byteIndexWritten < textStringBuilder.Length - 1
        End Sub

        Friend Shared Sub WriteDataTableToCsvFile(ByVal path As String, ByVal dataTable As System.Data.DataTable, lineEncodings As CompuMaster.Data.Csv.WriteLineEncodings)
            WriteDataTableToCsvFile(path, dataTable, True, System.Globalization.CultureInfo.InvariantCulture, "UTF-8", vbNullChar, """"c, lineEncodings)
        End Sub

        Friend Shared Sub WriteDataTableToCsvFile(ByVal path As String, ByVal dataTable As System.Data.DataTable, ByVal includesColumnHeaders As Boolean, ByVal columnWidths As Integer(), ByVal cultureFormatProvider As System.Globalization.CultureInfo, ByVal encoding As String, lineEncodings As CompuMaster.Data.Csv.WriteLineEncodings)

            'Create stream writer
            Dim writer As StreamWriter = Nothing
            Try
                writer = New StreamWriter(path, False, System.Text.Encoding.GetEncoding(encoding))
                Dim textStringBuilder As System.Text.StringBuilder = ConvertDataTableToCsv(dataTable, includesColumnHeaders, cultureFormatProvider, columnWidths, lineEncodings)
                WriteTextStringBuilderToStreamWriter(writer, textStringBuilder)
            Finally
                writer?.Close()
            End Try

        End Sub

        Friend Shared Sub WriteDataTableToCsvFile(ByVal path As String, ByVal dataTable As System.Data.DataTable, ByVal includesColumnHeaders As Boolean, ByVal cultureFormatProvider As System.Globalization.CultureInfo, ByVal encoding As String, ByVal columnSeparator As String, ByVal recognizeTextBy As Char, lineEncodings As CompuMaster.Data.Csv.WriteLineEncodings)

            'Create stream writer
            Dim writer As StreamWriter = Nothing
            Try
                writer = New StreamWriter(path, False, System.Text.Encoding.GetEncoding(encoding))
                Dim textStringBuilder As System.Text.StringBuilder = ConvertDataTableToCsv(dataTable, includesColumnHeaders, cultureFormatProvider, columnSeparator, recognizeTextBy, lineEncodings)
                WriteTextStringBuilderToStreamWriter(writer, textStringBuilder)
            Finally
                writer?.Close()
            End Try

        End Sub

        ''' <summary>
        '''     Trims a string to exactly the required fix size
        ''' </summary>
        ''' <param name="text"></param>
        ''' <param name="fixedLengthSize"></param>
        ''' <param name="alignedRight">Add additionally required spaces on the left (True) or on the right (False)</param>
        ''' <returns></returns>
        Private Shared Function FixedLengthText(ByVal text As String, ByVal fixedLengthSize As Integer, ByVal alignedRight As Boolean) As String
            Dim Result As String = Mid(text, 1, fixedLengthSize)
            If Result.Length < fixedLengthSize Then
                'Add some spaces to the string
                If alignedRight = False Then
                    Result &= Strings.Space(fixedLengthSize - Result.Length)
                Else
                    Result = Strings.Space(fixedLengthSize - Result.Length) & Result
                End If
            End If
            Return Result
        End Function

        ''' <summary>
        '''     Convert the datatable to a string based, fixed-column format
        ''' </summary>
        ''' <param name="dataTable">A datatable which shall be exported</param>
        ''' <param name="includesColumnHeaders">Indicates wether column headers are present</param>
        ''' <param name="cultureFormatProvider">A culture for all conversions</param>
        ''' <param name="columnWidths">An array of columns widths in chars</param>
        ''' <param name="lineEncodings">Encoding style for linebreaks</param>
        Friend Shared Function ConvertDataTableToCsv(ByVal dataTable As System.Data.DataTable, ByVal includesColumnHeaders As Boolean, ByVal cultureFormatProvider As System.Globalization.CultureInfo, ByVal columnWidths As Integer(), lineEncodings As CompuMaster.Data.Csv.WriteLineEncodings) As System.Text.StringBuilder

            If cultureFormatProvider Is Nothing Then
                cultureFormatProvider = System.Globalization.CultureInfo.InvariantCulture
            End If

            If lineEncodings = Csv.WriteLineEncodings.Auto Then
                Select Case System.Environment.NewLine
                    Case ControlChars.CrLf
                        'Windows platforms
                        lineEncodings = Csv.WriteLineEncodings.RowBreakCrLf_CellLineBreakLf
                    Case ControlChars.Cr
                        'Mac platforms
                        lineEncodings = Csv.WriteLineEncodings.RowBreakCr_CellLineBreakLf
                    Case ControlChars.Lf
                        'Linux platforms
                        lineEncodings = Csv.WriteLineEncodings.RowBreakLf_CellLineBreakCr
                    Case Else
                        Throw New NotImplementedException
                End Select
            End If
            Dim RequestedRowLineBreak As String = RowLineBreak(lineEncodings)

            Dim writer As New System.Text.StringBuilder

            'Column headers
            If includesColumnHeaders Then
                For ColCounter As Integer = 0 To System.Math.Min(columnWidths.Length, dataTable.Columns.Count) - 1
                    writer.Append(FixedLengthText(dataTable.Columns(ColCounter).ColumnName, columnWidths(ColCounter), False))
                Next
                writer.Append(RequestedRowLineBreak)
            End If

            'Data values
            For RowCounter As Integer = 0 To dataTable.Rows.Count - 1
                For ColCounter As Integer = 0 To System.Math.Min(columnWidths.Length, dataTable.Columns.Count) - 1
                    If dataTable.Rows(RowCounter)(ColCounter) Is DBNull.Value Then
                        writer.Append(FixedLengthText(String.Empty, columnWidths(ColCounter), False))
                    ElseIf dataTable.Columns(ColCounter).DataType Is GetType(String) Then
                        'Strings
                        If dataTable.Rows(RowCounter)(ColCounter) IsNot DBNull.Value Then
                            writer.Append(FixedLengthText(CsvEncode(CType(dataTable.Rows(RowCounter)(ColCounter), String), Nothing, lineEncodings), columnWidths(ColCounter), False))
                        End If
                    ElseIf dataTable.Columns(ColCounter).DataType Is GetType(System.Double) Then
                        'Doubles
                        If dataTable.Rows(RowCounter)(ColCounter) IsNot DBNull.Value Then
                            'Other data types which do not require textual handling
                            writer.Append(FixedLengthText(CType(dataTable.Rows(RowCounter)(ColCounter), Double).ToString(cultureFormatProvider), columnWidths(ColCounter), True))
                        End If
                    ElseIf dataTable.Columns(ColCounter).DataType Is GetType(System.Decimal) Then
                        'Decimals
                        If dataTable.Rows(RowCounter)(ColCounter) IsNot DBNull.Value Then
                            'Other data types which do not require textual handling
                            writer.Append(FixedLengthText(CType(dataTable.Rows(RowCounter)(ColCounter), Decimal).ToString(cultureFormatProvider), columnWidths(ColCounter), True))
                        End If
                    ElseIf dataTable.Columns(ColCounter).DataType Is GetType(System.DateTime) Then
                        'Datetime
                        If dataTable.Rows(RowCounter)(ColCounter) IsNot DBNull.Value Then
                            'Other data types which do not require textual handling
                            writer.Append(FixedLengthText(CType(dataTable.Rows(RowCounter)(ColCounter), DateTime).ToString(cultureFormatProvider), columnWidths(ColCounter), False))
                        End If
                    ElseIf dataTable.Columns(ColCounter).DataType Is GetType(System.Int16) OrElse dataTable.Columns(ColCounter).DataType Is GetType(System.Int32) OrElse dataTable.Columns(ColCounter).DataType Is GetType(System.Int64) Then
                        'Intxx
                        If dataTable.Rows(RowCounter)(ColCounter) IsNot DBNull.Value Then
                            'Other data types which do not require textual handling
                            writer.Append(FixedLengthText(CType(dataTable.Rows(RowCounter)(ColCounter), System.Int64).ToString(cultureFormatProvider), columnWidths(ColCounter), True))
                        End If
                    ElseIf dataTable.Columns(ColCounter).DataType Is GetType(System.UInt16) OrElse dataTable.Columns(ColCounter).DataType Is GetType(System.UInt32) OrElse dataTable.Columns(ColCounter).DataType Is GetType(System.UInt64) Then
                        'UIntxx
                        If dataTable.Rows(RowCounter)(ColCounter) IsNot DBNull.Value Then
                            'Other data types which do not require textual handling
                            writer.Append(FixedLengthText(CType(dataTable.Rows(RowCounter)(ColCounter), System.UInt64).ToString(cultureFormatProvider), columnWidths(ColCounter), True))
                        End If
                    Else
                        'Other data types
                        If dataTable.Rows(RowCounter)(ColCounter) IsNot DBNull.Value Then
                            'Other data types which do not require textual handling
                            writer.Append(FixedLengthText(CType(dataTable.Rows(RowCounter)(ColCounter), String), columnWidths(ColCounter), False))
                        End If
                    End If
                Next
                writer.Append(RequestedRowLineBreak)
            Next
            Return writer

        End Function

        ''' <summary>
        '''     Convert the datatable to a string based, comma-separated format
        ''' </summary>
        ''' <param name="dataTable">A datatable which shall be exported</param>
        ''' <param name="includesColumnHeaders">Indicates wether column headers are present</param>
        ''' <param name="cultureFormatProvider">A culture for all conversions</param>
        ''' <param name="columnSeparator">Choose the required character for splitting the columns. Set to null (Nothing in VisualBasic) to enable fixed column widths mode</param>
        ''' <param name="recognizeTextBy">A character indicating the start and end of text strings</param>
        ''' <param name="lineEncodings">Encoding style for linebreaks</param>
        ''' <returns></returns>
        Friend Shared Function ConvertDataTableToCsv(ByVal dataTable As System.Data.DataTable, ByVal includesColumnHeaders As Boolean, ByVal cultureFormatProvider As System.Globalization.CultureInfo, ByVal columnSeparator As String, ByVal recognizeTextBy As Char, lineEncodings As CompuMaster.Data.Csv.WriteLineEncodings) As System.Text.StringBuilder

            If cultureFormatProvider Is Nothing Then
                cultureFormatProvider = System.Globalization.CultureInfo.InvariantCulture
            End If

            If columnSeparator = Nothing OrElse columnSeparator = vbNullChar Then
                columnSeparator = cultureFormatProvider.TextInfo.ListSeparator
            End If

            If lineEncodings = Csv.WriteLineEncodings.Auto Then
                Select Case System.Environment.NewLine
                    Case ControlChars.CrLf
                        'Windows platforms
                        lineEncodings = Csv.WriteLineEncodings.RowBreakCrLf_CellLineBreakLf
                    Case ControlChars.Cr
                        'Mac platforms
                        lineEncodings = Csv.WriteLineEncodings.RowBreakCr_CellLineBreakLf
                    Case ControlChars.Lf
                        'Linux platforms
                        lineEncodings = Csv.WriteLineEncodings.RowBreakLf_CellLineBreakCr
                    Case Else
                        Throw New NotImplementedException
                End Select
            End If
            Dim RequestedRowLineBreak As String = RowLineBreak(lineEncodings)

            Dim writer As New System.Text.StringBuilder

            'Column headers
            If includesColumnHeaders Then
                For ColCounter As Integer = 0 To dataTable.Columns.Count - 1
                    If ColCounter <> 0 Then
                        writer.Append(columnSeparator)
                    End If
                    If recognizeTextBy <> Nothing Then writer.Append(recognizeTextBy)
                    writer.Append(CsvEncode(dataTable.Columns(ColCounter).ColumnName, recognizeTextBy, lineEncodings))
                    If recognizeTextBy <> Nothing Then writer.Append(recognizeTextBy)
                Next
                writer.Append(RequestedRowLineBreak)
            End If

            'Data values
            For RowCounter As Integer = 0 To dataTable.Rows.Count - 1
                For ColCounter As Integer = 0 To dataTable.Columns.Count - 1
                    If ColCounter <> 0 Then
                        writer.Append(columnSeparator)
                    End If
                    WriteCellValue(dataTable.Columns(ColCounter).DataType, dataTable.Rows(RowCounter)(ColCounter), recognizeTextBy, columnSeparator, cultureFormatProvider, lineEncodings, Nothing, writer)
                Next
                writer.Append(RequestedRowLineBreak)
            Next
            Return writer

        End Function

        ''' <summary>
        '''     Write to a CSV file
        ''' </summary>
        ''' <param name="path">The path of the file</param>
        ''' <param name="dataTable">A datatable which shall be exported</param>
        ''' <param name="includesColumnHeaders">Indicates wether column headers are present</param>
        ''' <param name="encoding">The text encoding of the file</param>
        ''' <param name="columnSeparator">Choose the required character for splitting the columns. Set to null (Nothing in VisualBasic) to enable fixed column widths mode</param>
        ''' <param name="recognizeTextBy">A character indicating the start and end of text strings</param>
        ''' <param name="decimalSeparator">A character indicating the decimal separator in the text string</param>
        ''' <param name="lineEncodings">Encoding style for linebreaks</param>
        Friend Shared Sub WriteDataTableToCsvFile(ByVal path As String, ByVal dataTable As System.Data.DataTable, ByVal includesColumnHeaders As Boolean, ByVal encoding As String, ByVal columnSeparator As String, ByVal recognizeTextBy As Char, ByVal decimalSeparator As Char, lineEncodings As CompuMaster.Data.Csv.WriteLineEncodings)

            Dim RequestedRowLineBreak As String = RowLineBreak(lineEncodings)

            Dim cultureFormatProvider As New System.Globalization.CultureInfo("")
            cultureFormatProvider.NumberFormat.CurrencyDecimalSeparator = decimalSeparator
            cultureFormatProvider.NumberFormat.NumberDecimalSeparator = decimalSeparator
            cultureFormatProvider.NumberFormat.PercentDecimalSeparator = decimalSeparator

            'Create stream writer
            Dim writer As StreamWriter = Nothing
            Try
                writer = New StreamWriter(path, False, System.Text.Encoding.GetEncoding(encoding))

                'Column headers
                If includesColumnHeaders Then
                    For ColCounter As Integer = 0 To dataTable.Columns.Count - 1
                        If ColCounter <> 0 Then
                            writer.Write(columnSeparator)
                        End If
                        If recognizeTextBy <> Nothing Then writer.Write(recognizeTextBy)
                        writer.Write(CsvEncode(dataTable.Columns(ColCounter).ColumnName, recognizeTextBy, lineEncodings))
                        If recognizeTextBy <> Nothing Then writer.Write(recognizeTextBy)
                    Next
                    writer.Write(RequestedRowLineBreak)
                End If

                'Data values
                For RowCounter As Integer = 0 To dataTable.Rows.Count - 1
                    For ColCounter As Integer = 0 To dataTable.Columns.Count - 1
                        If ColCounter <> 0 Then
                            writer.Write(columnSeparator)
                        End If
                        WriteCellValue(dataTable.Columns(ColCounter).DataType, dataTable.Rows(RowCounter)(ColCounter), recognizeTextBy, columnSeparator, cultureFormatProvider, lineEncodings, writer, Nothing)
                    Next
                    writer.Write(RequestedRowLineBreak)
                Next

            Finally
                writer?.Close()
            End Try

        End Sub

        ''' <summary>
        '''     Create a CSV table
        ''' </summary>
        ''' <param name="dataTable">A datatable which shall be exported</param>
        ''' <param name="includesColumnHeaders">Indicates wether column headers are present</param>
        ''' <param name="encoding">The text encoding of the file</param>
        ''' <param name="columnSeparator">Choose the required character for splitting the columns. Set to null (Nothing in VisualBasic) to enable fixed column widths mode</param>
        ''' <param name="recognizeTextBy">A character indicating the start and end of text strings</param>
        ''' <param name="decimalSeparator">A character indicating the decimal separator in the text string</param>
        ''' <param name="lineEncodings">Encoding style for linebreaks</param>
        ''' <returns>A string containing the CSV table</returns>
        Friend Shared Function WriteDataTableToCsvBytes(ByVal dataTable As System.Data.DataTable, ByVal includesColumnHeaders As Boolean, ByVal encoding As String, ByVal columnSeparator As Char, ByVal recognizeTextBy As Char, ByVal decimalSeparator As Char, lineEncodings As CompuMaster.Data.Csv.WriteLineEncodings) As Byte()
            Dim MyStream As MemoryStream = WriteDataTableToCsvMemoryStream(dataTable, includesColumnHeaders, encoding, columnSeparator, recognizeTextBy, decimalSeparator, lineEncodings)
            Return MyStream.ToArray
        End Function

        ''' <summary>
        '''     Create a CSV table
        ''' </summary>
        ''' <param name="dataTable">A datatable which shall be exported</param>
        ''' <param name="includesColumnHeaders">Indicates wether column headers are present</param>
        ''' <param name="encoding">The text encoding of the file</param>
        ''' <param name="cultureFormatProvider">A globalization information object for the conversion of all data to strings</param>
        ''' <param name="columnSeparator">Choose the required character for splitting the columns. Set to null (Nothing in VisualBasic) to enable fixed column widths mode</param>
        ''' <param name="recognizeTextBy">A character indicating the start and end of text strings</param>
        ''' <param name="lineEncodings">Encoding style for linebreaks</param>
        ''' <returns>A string containing the CSV table</returns>
        Friend Shared Function WriteDataTableToCsvBytes(ByVal dataTable As System.Data.DataTable, ByVal includesColumnHeaders As Boolean, ByVal encoding As System.Text.Encoding, ByVal cultureFormatProvider As System.Globalization.CultureInfo, ByVal columnSeparator As Char, ByVal recognizeTextBy As Char, lineEncodings As CompuMaster.Data.Csv.WriteLineEncodings) As Byte()
            Dim MyStream As MemoryStream = WriteDataTableToCsvMemoryStream(dataTable, includesColumnHeaders, encoding, cultureFormatProvider, columnSeparator, recognizeTextBy, lineEncodings)
            Return MyStream.ToArray
        End Function

        ''' <summary>
        '''     Create a CSV table
        ''' </summary>
        ''' <param name="dataTable">A datatable which shall be exported</param>
        ''' <param name="includesColumnHeaders">Indicates wether column headers are present</param>
        ''' <param name="encoding">The text encoding of the file</param>
        ''' <param name="columnSeparator">Choose the required character for splitting the columns. Set to null (Nothing in VisualBasic) to enable fixed column widths mode</param>
        ''' <param name="recognizeTextBy">A character indicating the start and end of text strings</param>
        ''' <param name="decimalSeparator">A character indicating the decimal separator in the text string</param>
        ''' <param name="lineEncodings">Encoding style for linebreaks</param>
        ''' <returns>A memory stream containing all texts as bytes in Unicode format</returns>
        Friend Shared Function WriteDataTableToCsvMemoryStream(ByVal dataTable As System.Data.DataTable, ByVal includesColumnHeaders As Boolean, ByVal encoding As String, ByVal columnSeparator As String, ByVal recognizeTextBy As Char, ByVal decimalSeparator As Char, lineEncodings As CompuMaster.Data.Csv.WriteLineEncodings) As System.IO.MemoryStream
            Dim cultureFormatProvider As System.Globalization.CultureInfo = CType(System.Globalization.CultureInfo.InvariantCulture.Clone, System.Globalization.CultureInfo)
            cultureFormatProvider.NumberFormat.CurrencyDecimalSeparator = decimalSeparator
            cultureFormatProvider.NumberFormat.NumberDecimalSeparator = decimalSeparator
            cultureFormatProvider.NumberFormat.PercentDecimalSeparator = decimalSeparator
            Return WriteDataTableToCsvMemoryStream(dataTable, includesColumnHeaders, System.Text.Encoding.GetEncoding(encoding), cultureFormatProvider, columnSeparator, recognizeTextBy, lineEncodings)
        End Function

        ''' <summary>
        '''     Create a CSV table
        ''' </summary>
        ''' <param name="dataTable">A datatable which shall be exported</param>
        ''' <param name="includesColumnHeaders">Indicates wether column headers are present</param>
        ''' <param name="encoding">The text encoding of the file</param>
        ''' <param name="cultureFormatProvider">A globalization information object for the conversion of all data to strings</param>
        ''' <param name="columnSeparator">Choose the required character for splitting the columns. Set to null (Nothing in VisualBasic) to enable fixed column widths mode</param>
        ''' <param name="recognizeTextBy">A character indicating the start and end of text strings</param>
        ''' <returns>A memory stream containing all texts as bytes in Unicode format</returns>
        ''' <remarks>
        ''' </remarks>
        <Obsolete("Better use overload with parameter lineEncoding")> Friend Shared Function WriteDataTableToCsvMemoryStream(ByVal dataTable As System.Data.DataTable, ByVal includesColumnHeaders As Boolean, ByVal encoding As System.Text.Encoding, ByVal cultureFormatProvider As System.Globalization.CultureInfo, ByVal columnSeparator As String, ByVal recognizeTextBy As Char) As System.IO.MemoryStream
            Return WriteDataTableToCsvMemoryStream(dataTable, includesColumnHeaders, encoding, cultureFormatProvider, columnSeparator, recognizeTextBy, Csv.WriteLineEncodings.Default)
        End Function

        ''' <summary>
        '''     Create a CSV table
        ''' </summary>
        ''' <param name="dataTable">A datatable which shall be exported</param>
        ''' <param name="includesColumnHeaders">Indicates wether column headers are present</param>
        ''' <param name="encoding">The text encoding of the file</param>
        ''' <param name="cultureFormatProvider">A globalization information object for the conversion of all data to strings</param>
        ''' <param name="columnSeparator">Choose the required character for splitting the columns. Set to null (Nothing in VisualBasic) to enable fixed column widths mode</param>
        ''' <param name="recognizeTextBy">A character indicating the start and end of text strings</param>
        ''' <param name="lineEncodings">Encoding style for linebreaks</param>
        ''' <returns>A memory stream containing all texts as bytes in Unicode format</returns>
        Friend Shared Function WriteDataTableToCsvMemoryStream(ByVal dataTable As System.Data.DataTable, ByVal includesColumnHeaders As Boolean, ByVal encoding As System.Text.Encoding, ByVal cultureFormatProvider As System.Globalization.CultureInfo, ByVal columnSeparator As String, ByVal recognizeTextBy As Char, lineEncodings As CompuMaster.Data.Csv.WriteLineEncodings) As System.IO.MemoryStream

            Dim RequestedRowLineBreak As String = RowLineBreak(lineEncodings)

            If cultureFormatProvider Is Nothing Then
                cultureFormatProvider = System.Globalization.CultureInfo.InvariantCulture
            End If

            If columnSeparator = Nothing OrElse columnSeparator = vbNullChar Then
                columnSeparator = cultureFormatProvider.TextInfo.ListSeparator
            End If

            'Create stream writer
            Dim Result As New MemoryStream
            Dim writer As StreamWriter = Nothing
            Try
                writer = New StreamWriter(Result, encoding)

                'Column headers
                If includesColumnHeaders Then
                    For ColCounter As Integer = 0 To dataTable.Columns.Count - 1
                        If ColCounter <> 0 Then
                            writer.Write(columnSeparator)
                        End If
                        If recognizeTextBy <> Nothing Then writer.Write(recognizeTextBy)
                        writer.Write(CsvEncode(dataTable.Columns(ColCounter).ColumnName, recognizeTextBy, lineEncodings))
                        If recognizeTextBy <> Nothing Then writer.Write(recognizeTextBy)
                    Next
                    writer.Write(RequestedRowLineBreak)
                End If

                'Data values
                For RowCounter As Integer = 0 To dataTable.Rows.Count - 1
                    For ColCounter As Integer = 0 To dataTable.Columns.Count - 1
                        If ColCounter <> 0 Then
                            writer.Write(columnSeparator)
                        End If
                        WriteCellValue(dataTable.Columns(ColCounter).DataType, dataTable.Rows(RowCounter)(ColCounter), recognizeTextBy, columnSeparator, cultureFormatProvider, lineEncodings, writer, Nothing)
                    Next
                    writer.Write(RequestedRowLineBreak)
                Next

            Finally
                writer?.Close()
            End Try

            Return Result

        End Function

        ''' <summary>
        ''' The line break for rows
        ''' </summary>
        ''' <param name="lineEncoding"></param>
        ''' <returns></returns>
        Friend Shared Function RowLineBreak(lineEncoding As CompuMaster.Data.Csv.WriteLineEncodings) As String
            Select Case lineEncoding
                Case Csv.WriteLineEncodings.None, Csv.WriteLineEncodings.Auto
                    Return System.Environment.NewLine
                Case Csv.WriteLineEncodings.Default, Csv.WriteLineEncodings.RowBreakCrLf_CellLineBreakCr, Csv.WriteLineEncodings.RowBreakCrLf_CellLineBreakLf, Csv.WriteLineEncodings.RowBreakCrLf_CellLineBreakRemoved, Csv.WriteLineEncodings.RowBreakCrLf_CellLineBreakSpaceChar
                    Return ControlChars.CrLf
                Case Csv.WriteLineEncodings.RowBreakCr_CellLineBreakLf, Csv.WriteLineEncodings.RowBreakCr_CellLineBreakRemoved, Csv.WriteLineEncodings.RowBreakCr_CellLineBreakSpaceChar
                    Return ControlChars.Cr
                Case Csv.WriteLineEncodings.RowBreakLf_CellLineBreakCr, Csv.WriteLineEncodings.RowBreakLf_CellLineBreakRemoved, Csv.WriteLineEncodings.RowBreakLf_CellLineBreakSpaceChar
                    Return ControlChars.Lf
                Case Else
                    Throw New NotImplementedException("Invalid lineEncoding")
            End Select
        End Function

        ''' <summary>
        '''     Encode a string into CSV encoding
        ''' </summary>
        ''' <param name="value">The unencoded text</param>
        ''' <param name="recognizeTextBy">The character to identify a string in the CSV file</param>
        ''' <returns>The encoded writing style of the given text</returns>
        Friend Shared Function CsvEncode(ByVal value As String, ByVal recognizeTextBy As Char, lineEncoding As CompuMaster.Data.Csv.WriteLineEncodings) As String
            Dim Result As String
            If recognizeTextBy <> Nothing Then
                Result = Replace(value, recognizeTextBy, recognizeTextBy & recognizeTextBy)
            Else
                Result = value
            End If
            If lineEncoding = Csv.WriteLineEncodings.Auto Then
                Select Case System.Environment.NewLine
                    Case ControlChars.CrLf
                        'Windows platforms
                        lineEncoding = Csv.WriteLineEncodings.RowBreakCrLf_CellLineBreakLf
                    Case ControlChars.Lf
                        'Linux platforms
                        lineEncoding = Csv.WriteLineEncodings.RowBreakLf_CellLineBreakCr
                    Case ControlChars.Cr
                        'Mac platforms
                        lineEncoding = Csv.WriteLineEncodings.RowBreakCr_CellLineBreakLf
                    Case Else
                        Throw New NotImplementedException
                End Select
            End If
            Select Case lineEncoding
                Case Csv.WriteLineEncodings.None
                Case Csv.WriteLineEncodings.Default, Csv.WriteLineEncodings.RowBreakCrLf_CellLineBreakLf, Csv.WriteLineEncodings.RowBreakCr_CellLineBreakLf
                    Result = Replace(Result, ControlChars.CrLf, ControlChars.Lf)
                    Result = Replace(Result, ControlChars.Cr, ControlChars.Lf)
                Case Csv.WriteLineEncodings.RowBreakCrLf_CellLineBreakCr, Csv.WriteLineEncodings.RowBreakLf_CellLineBreakCr
                    Result = Replace(Result, ControlChars.CrLf, ControlChars.Cr)
                    Result = Replace(Result, ControlChars.Lf, ControlChars.Cr)
                Case Csv.WriteLineEncodings.RowBreakCrLf_CellLineBreakSpaceChar, Csv.WriteLineEncodings.RowBreakLf_CellLineBreakSpaceChar, Csv.WriteLineEncodings.RowBreakCr_CellLineBreakSpaceChar
                    Result = Replace(Result, ControlChars.CrLf, " ")
                    Result = Replace(Result, ControlChars.Lf, " "c)
                    Result = Replace(Result, ControlChars.Cr, " "c)
                Case Csv.WriteLineEncodings.RowBreakCrLf_CellLineBreakRemoved, Csv.WriteLineEncodings.RowBreakLf_CellLineBreakRemoved, Csv.WriteLineEncodings.RowBreakCr_CellLineBreakRemoved
                    Result = Replace(Result, ControlChars.CrLf, "")
                    Result = Replace(Result, ControlChars.Lf, "")
                    Result = Replace(Result, ControlChars.Cr, "")
                Case Else
                    Throw New NotSupportedException("Not supported/implemented: lineEncoding " & lineEncoding)
            End Select
            Return Result
        End Function

        Friend Shared Sub WriteDataViewToCsvFile(ByVal path As String, ByVal dataview As System.Data.DataView, lineEncodings As CompuMaster.Data.Csv.WriteLineEncodings)
            WriteDataViewToCsvFile(path, dataview, True, System.Globalization.CultureInfo.InvariantCulture, "UTF-8", vbNullChar, """"c, lineEncodings)
        End Sub

        Friend Shared Sub WriteDataViewToCsvFile(ByVal path As String, ByVal dataView As System.Data.DataView, ByVal includesColumnHeaders As Boolean, ByVal cultureFormatProvider As System.Globalization.CultureInfo, ByVal encoding As String, ByVal columnSeparator As String, ByVal recognizeTextBy As Char, lineEncodings As CompuMaster.Data.Csv.WriteLineEncodings)

            Dim DataTable As System.Data.DataTable = dataView.Table
            Dim RequestedRowLineBreak As String = RowLineBreak(lineEncodings)

            If cultureFormatProvider Is Nothing Then
                cultureFormatProvider = System.Globalization.CultureInfo.InvariantCulture
            End If

            If columnSeparator = Nothing OrElse columnSeparator = vbNullChar Then
                columnSeparator = cultureFormatProvider.TextInfo.ListSeparator
            End If

            'Create stream writer
            Dim writer As StreamWriter = Nothing
            Try
                writer = New StreamWriter(path, False, System.Text.Encoding.GetEncoding(encoding))

                'Column headers
                If includesColumnHeaders Then
                    For ColCounter As Integer = 0 To DataTable.Columns.Count - 1
                        If ColCounter <> 0 Then
                            writer.Write(columnSeparator)
                        End If
                        If recognizeTextBy <> Nothing Then writer.Write(recognizeTextBy)
                        writer.Write(CsvEncode(DataTable.Columns(ColCounter).ColumnName, recognizeTextBy, lineEncodings))
                        If recognizeTextBy <> Nothing Then writer.Write(recognizeTextBy)
                    Next
                    writer.Write(RequestedRowLineBreak)
                End If

                'Data values
                For RowCounter As Integer = 0 To dataView.Count - 1
                    For ColCounter As Integer = 0 To DataTable.Columns.Count - 1
                        If ColCounter <> 0 Then
                            writer.Write(columnSeparator)
                        End If
                        WriteCellValue(DataTable.Columns(ColCounter).DataType, dataView.Item(RowCounter).Row(ColCounter), recognizeTextBy, columnSeparator, cultureFormatProvider, lineEncodings, writer, Nothing)
                    Next
                Next

            Finally
                writer?.Close()
            End Try

        End Sub

        ''' <summary>
        '''     Write to a CSV file
        ''' </summary>
        ''' <param name="path">The path of the file</param>
        ''' <param name="dataView">A dataview object with the desired rows</param>
        ''' <param name="includesColumnHeaders">Indicates wether column headers are present</param>
        ''' <param name="encoding">The text encoding of the file</param>
        ''' <param name="columnSeparator">Choose the required character for splitting the columns. Set to null (Nothing in VisualBasic) to enable fixed column widths mode</param>
        ''' <param name="recognizeTextBy">A character indicating the start and end of text strings</param>
        ''' <param name="lineEncodings">Encoding style for linebreaks</param>
        Friend Shared Sub WriteDataViewToCsvFile(ByVal path As String, ByVal dataView As System.Data.DataView, ByVal includesColumnHeaders As Boolean, ByVal encoding As String, ByVal columnSeparator As String, ByVal recognizeTextBy As Char, ByVal decimalSeparator As Char, lineEncodings As CompuMaster.Data.Csv.WriteLineEncodings)

            Dim DataTable As System.Data.DataTable = dataView.Table
            Dim RequestedRowLineBreak As String = RowLineBreak(lineEncodings)

            Dim cultureFormatProvider As New System.Globalization.CultureInfo("")
            cultureFormatProvider.NumberFormat.CurrencyDecimalSeparator = decimalSeparator
            cultureFormatProvider.NumberFormat.NumberDecimalSeparator = decimalSeparator
            cultureFormatProvider.NumberFormat.PercentDecimalSeparator = decimalSeparator

            'Create stream writer
            Dim writer As StreamWriter = Nothing
            Try
                writer = New StreamWriter(path, False, System.Text.Encoding.GetEncoding(encoding))

                'Column headers
                If includesColumnHeaders Then
                    For ColCounter As Integer = 0 To DataTable.Columns.Count - 1
                        If ColCounter <> 0 Then
                            writer.Write(columnSeparator)
                        End If
                        If recognizeTextBy <> Nothing Then writer.Write(recognizeTextBy)
                        writer.Write(CsvEncode(DataTable.Columns(ColCounter).ColumnName, recognizeTextBy, lineEncodings))
                        If recognizeTextBy <> Nothing Then writer.Write(recognizeTextBy)
                    Next
                    writer.Write(RequestedRowLineBreak)
                End If

                'Data values
                For RowCounter As Integer = 0 To dataView.Count - 1
                    For ColCounter As Integer = 0 To DataTable.Columns.Count - 1
                        If ColCounter <> 0 Then
                            writer.Write(columnSeparator)
                        End If
                        WriteCellValue(DataTable.Columns(ColCounter).DataType, dataView.Item(RowCounter).Row(ColCounter), recognizeTextBy, columnSeparator, cultureFormatProvider, lineEncodings, writer, Nothing)
                    Next
                    writer.Write(RequestedRowLineBreak)
                Next

            Finally
                writer?.Close()
            End Try

        End Sub

        ''' <summary>
        ''' Test for provided writer object and use it to write back the value
        ''' </summary>
        ''' <param name="value"></param>
        ''' <param name="writerStream"></param>
        ''' <param name="writerStringBuilder"></param>
        Private Shared Sub WriteCellValueToWriter(value As String, writerStream As System.IO.StreamWriter, writerStringBuilder As System.Text.StringBuilder)
            If writerStream IsNot Nothing Then
                writerStream.Write(value)
            Else
                writerStringBuilder.Append(value)
            End If
        End Sub

        ''' <summary>
        ''' Write the cell value to the given writer object
        ''' </summary>
        ''' <param name="cellColumnDataType"></param>
        ''' <param name="cellValue"></param>
        ''' <param name="columnSeparator">Choose the required character for splitting the columns. Set to null (Nothing in VisualBasic) to enable fixed column widths mode</param>
        ''' <param name="recognizeTextBy">A character indicating the start and end of text strings</param>
        ''' <param name="cultureFormatProvider">A culture for all conversions</param>
        ''' <param name="lineEncodings">Encoding style for linebreaks</param>
        ''' <param name="writerStream"></param>
        ''' <param name="writerStringBuilder"></param>
        Private Shared Sub WriteCellValue(cellColumnDataType As Type, cellValue As Object, recognizeTextBy As Char, columnSeparator As String, cultureFormatProvider As System.Globalization.CultureInfo, lineEncodings As CompuMaster.Data.Csv.WriteLineEncodings,
                                          writerStream As System.IO.StreamWriter, writerStringBuilder As System.Text.StringBuilder)
            If cellColumnDataType Is GetType(String) Then
                'Strings
                If cellValue IsNot DBNull.Value Then
                    If recognizeTextBy <> Nothing Then WriteCellValueToWriter(recognizeTextBy, writerStream, writerStringBuilder)
                    WriteCellValueToWriter(CsvEncode(CType(cellValue, String), recognizeTextBy, lineEncodings), writerStream, writerStringBuilder)
                    If recognizeTextBy <> Nothing Then WriteCellValueToWriter(recognizeTextBy, writerStream, writerStringBuilder)
                End If
            ElseIf cellColumnDataType Is GetType(System.Double) Then
                'Doubles
                If cellValue IsNot DBNull.Value Then
                    'Other data types which do not require textual handling
                    Dim Value As String = CType(cellValue, Double).ToString(cultureFormatProvider)
                    If Value <> "" AndAlso Value.Contains(columnSeparator) Then WriteCellValueToWriter(recognizeTextBy, writerStream, writerStringBuilder)
                    WriteCellValueToWriter(Value, writerStream, writerStringBuilder)
                    If Value <> "" AndAlso Value.Contains(columnSeparator) Then WriteCellValueToWriter(recognizeTextBy, writerStream, writerStringBuilder)
                End If
            ElseIf cellColumnDataType Is GetType(System.Decimal) Then
                'Decimals
                If cellValue IsNot DBNull.Value Then
                    'Other data types which do not require textual handling
                    Dim Value As String = CType(cellValue, Decimal).ToString(cultureFormatProvider)
                    If Value <> "" AndAlso Value.Contains(columnSeparator) Then WriteCellValueToWriter(recognizeTextBy, writerStream, writerStringBuilder)
                    WriteCellValueToWriter(Value, writerStream, writerStringBuilder)
                    If Value <> "" AndAlso Value.Contains(columnSeparator) Then WriteCellValueToWriter(recognizeTextBy, writerStream, writerStringBuilder)
                End If
            ElseIf cellColumnDataType Is GetType(System.DateTime) Then
                'Datetime
                If cellValue IsNot DBNull.Value Then
                    'Other data types which do not require textual handling
                    Dim Value As String
                    If cultureFormatProvider Is Globalization.CultureInfo.InvariantCulture Then
                        Value = (CType(cellValue, DateTime).ToString("yyyy-MM-dd HH:mm:ss.fff", Threading.Thread.CurrentThread.CurrentCulture))
                    Else
                        Value = (CType(cellValue, DateTime).ToString(cultureFormatProvider))
                    End If
                    If Value <> "" AndAlso Value.Contains(columnSeparator) Then WriteCellValueToWriter(recognizeTextBy, writerStream, writerStringBuilder)
                    WriteCellValueToWriter(Value, writerStream, writerStringBuilder)
                    If Value <> "" AndAlso Value.Contains(columnSeparator) Then WriteCellValueToWriter(recognizeTextBy, writerStream, writerStringBuilder)
                End If
            Else
                'Other data types
                If cellValue IsNot DBNull.Value Then
                    'Other data types which do not require textual handling
                    Dim Value As String = CType(cellValue, String)
                    If Value <> "" AndAlso Value.Contains(columnSeparator) Then WriteCellValueToWriter(recognizeTextBy, writerStream, writerStringBuilder)
                    WriteCellValueToWriter(Value, writerStream, writerStringBuilder)
                    If Value <> "" AndAlso Value.Contains(columnSeparator) Then WriteCellValueToWriter(recognizeTextBy, writerStream, writerStringBuilder)
                End If
            End If
        End Sub

#End Region

    End Class

End Namespace