Option Explicit On 
Option Strict On

Namespace CompuMaster.Data

    ''' <summary>
    '''     Provides simplified access to CSV files
    ''' </summary>
    Public Class Csv

        ''' <summary>
        '''     Read from a CSV file
        ''' </summary>
        ''' <param name="path">The path of the file</param>
        ''' <param name="csvContainsColumnHeaders">Indicates wether column headers are present</param>
        ''' <param name="encoding">The text encoding of the file</param>
        ''' <param name="columnSeparator">Choose the required character for splitting the columns. Set to null (Nothing in VisualBasic) to enable fixed column widths mode</param>
        ''' <param name="recognizeTextBy">A character indicating the start and end of text strings</param>
        ''' <param name="recognizeMultipleColumnSeparatorCharsAsOne">Specifies whether multiple seperator characters should be recognized as one</param>
        ''' <param name="convertEmptyStringsToDBNull">Convert values with empty strings automatically to DbNull</param>
        ''' <returns></returns>
        ''' <remarks>
        ''' In case of duplicate column names, all additional occurances of the same column name will be modified to use a unique column name
        ''' </remarks>
        Public Shared Function ReadDataTableFromCsvFile(ByVal path As String, ByVal csvContainsColumnHeaders As Boolean, Optional ByVal encoding As String = "UTF-8", Optional ByVal columnSeparator As Char = ","c, Optional ByVal recognizeTextBy As Char = """"c, Optional ByVal recognizeMultipleColumnSeparatorCharsAsOne As Boolean = False, Optional ByVal convertEmptyStringsToDBNull As Boolean = False) As DataTable
            Return CompuMaster.Data.CsvTools.ReadDataTableFromCsvFile(path, csvContainsColumnHeaders, encoding, columnSeparator, recognizeTextBy, recognizeMultipleColumnSeparatorCharsAsOne, convertEmptyStringsToDBNull)
        End Function

        ''' <summary>
        '''     Read from a CSV file
        ''' </summary>
        ''' <param name="path">The path of the file</param>
        ''' <param name="csvContainsColumnHeaders">Indicates wether column headers are present</param>
        ''' <param name="encoding">The text encoding of the file</param>
        ''' <param name="cultureFormatProvider"></param>
        ''' <param name="recognizeTextBy">A character indicating the start and end of text strings</param>
        ''' <param name="recognizeDoubledColumnSeparatorCharAsOne">Currently without purpose</param>
        ''' <param name="convertEmptyStringsToDBNull">Convert values with empty strings automatically to DbNull</param>
        ''' <returns></returns>
        ''' <remarks>
        ''' In case of duplicate column names, all additional occurances of the same column name will be modified to use a unique column name
        ''' </remarks>
        Public Shared Function ReadDataTableFromCsvFile(ByVal path As String, ByVal csvContainsColumnHeaders As Boolean, ByVal encoding As System.Text.Encoding, ByVal cultureFormatProvider As System.Globalization.CultureInfo, Optional ByVal recognizeTextBy As Char = """"c, Optional ByVal recognizeDoubledColumnSeparatorCharAsOne As Boolean = True, Optional ByVal convertEmptyStringsToDBNull As Boolean = False) As DataTable
            Return CompuMaster.Data.CsvTools.ReadDataTableFromCsvFile(path, csvContainsColumnHeaders, encoding, cultureFormatProvider, recognizeTextBy, recognizeDoubledColumnSeparatorCharAsOne, convertEmptyStringsToDBNull)
        End Function

        ''' <summary>
        '''     Read from a CSV table
        ''' </summary>
        ''' <param name="data">The content of a CSV file</param>
        ''' <param name="csvContainsColumnHeaders">Indicates wether column headers are present</param>
        ''' <param name="cultureFormatProvider"></param>
        ''' <param name="recognizeTextBy">A character indicating the start and end of text strings</param>
        ''' <param name="recognizeDoubledColumnSeparatorCharAsOne">Currently without purpose</param>
        ''' <param name="convertEmptyStringsToDBNull">Convert values with empty strings automatically to DbNull</param>
        ''' <returns></returns>
        ''' <remarks>
        ''' In case of duplicate column names, all additional occurances of the same column name will be modified to use a unique column name
        ''' </remarks>
        Public Shared Function ReadDataTableFromCsvString(ByVal data As String, ByVal csvContainsColumnHeaders As Boolean, ByVal cultureFormatProvider As System.Globalization.CultureInfo, Optional ByVal recognizeTextBy As Char = """"c, Optional ByVal recognizeDoubledColumnSeparatorCharAsOne As Boolean = True, Optional ByVal convertEmptyStringsToDBNull As Boolean = False) As DataTable
            Return CompuMaster.Data.CsvTools.ReadDataTableFromCsvString(data, csvContainsColumnHeaders, cultureFormatProvider, recognizeTextBy, recognizeDoubledColumnSeparatorCharAsOne, convertEmptyStringsToDBNull)
        End Function

        ''' <summary>
        '''     Read from a CSV table
        ''' </summary>
        ''' <param name="data">The content of a CSV file</param>
        ''' <param name="csvContainsColumnHeaders">Indicates wether column headers are present</param>
        ''' <param name="columnSeparator">Choose the required character for splitting the columns. Set to null (Nothing in VisualBasic) to enable fixed column widths mode</param>
        ''' <param name="RecognizeTextBy">A character indicating the start and end of text strings</param>
        ''' <param name="recognizeDoubledColumnSeparatorCharAsOne">Currently without purpose</param>
        ''' <param name="convertEmptyStringsToDBNull">Convert values with empty strings automatically to DbNull</param>
        ''' <returns></returns>
        ''' <remarks>
        ''' In case of duplicate column names, all additional occurances of the same column name will be modified to use a unique column name
        ''' </remarks>
        Public Shared Function ReadDataTableFromCsvString(ByVal data As String, ByVal csvContainsColumnHeaders As Boolean, Optional ByVal columnSeparator As Char = ","c, Optional ByVal recognizeTextBy As Char = """"c, Optional ByVal recognizeDoubledColumnSeparatorCharAsOne As Boolean = True, Optional ByVal convertEmptyStringsToDBNull As Boolean = False) As DataTable
            Return CompuMaster.Data.CsvTools.ReadDataTableFromCsvString(data, csvContainsColumnHeaders, columnSeparator, recognizeTextBy, recognizeDoubledColumnSeparatorCharAsOne, convertEmptyStringsToDBNull)
        End Function

        ''' <summary>
        '''     Read from a CSV file
        ''' </summary>
        ''' <param name="path">The path of the file</param>
        ''' <param name="csvContainsColumnHeaders">Indicates wether column headers are present</param>
        ''' <param name="encoding">The text encoding of the file</param>
        ''' <param name="columnWidths">An array of integers specifying the fixed column widths</param>
        ''' <param name="convertEmptyStringsToDBNull">Convert values with empty strings automatically to DbNull</param>
        ''' <returns></returns>
        ''' <remarks>
        ''' In case of duplicate column names, all additional occurances of the same column name will be modified to use a unique column name
        ''' </remarks>
        Public Shared Function ReadDataTableFromCsvFile(ByVal path As String, ByVal csvContainsColumnHeaders As Boolean, ByVal columnWidths As Integer(), Optional ByVal encoding As String = "UTF-8", Optional ByVal convertEmptyStringsToDBNull As Boolean = False) As DataTable
            Return CompuMaster.Data.CsvTools.ReadDataTableFromCsvFile(path, csvContainsColumnHeaders, columnWidths, encoding, convertEmptyStringsToDBNull)
        End Function

        ''' <summary>
        '''     Read from a CSV file
        ''' </summary>
        ''' <param name="path">The path of the file</param>
        ''' <param name="csvContainsColumnHeaders">Indicates wether column headers are present</param>
        ''' <param name="columnWidths">An array of integers specifying the fixed column widths</param>
        ''' <param name="encoding">The text encoding of the file</param>
        ''' <param name="cultureFormatProvider"></param>
        ''' <param name="convertEmptyStringsToDBNull">Convert values with empty strings automatically to DbNull</param>
        ''' <returns></returns>
        ''' <remarks>
        ''' In case of duplicate column names, all additional occurances of the same column name will be modified to use a unique column name
        ''' </remarks>
        Public Shared Function ReadDataTableFromCsvFile(ByVal path As String, ByVal csvContainsColumnHeaders As Boolean, ByVal columnWidths As Integer(), ByVal encoding As System.Text.Encoding, ByVal cultureFormatProvider As System.Globalization.CultureInfo, Optional ByVal convertEmptyStringsToDBNull As Boolean = False) As DataTable
            Return CompuMaster.Data.CsvTools.ReadDataTableFromCsvFile(path, csvContainsColumnHeaders, columnWidths, encoding, cultureFormatProvider, convertEmptyStringsToDBNull)
        End Function

        ''' <summary>
        '''     Read from a CSV table
        ''' </summary>
        ''' <param name="data">The content of a CSV file</param>
        ''' <param name="csvContainsColumnHeaders">Indicates wether column headers are present</param>
        ''' <param name="cultureFormatProvider"></param>
        ''' <param name="columnWidths">An array of integers specifying the fixed column widths</param>
        ''' <param name="convertEmptyStringsToDBNull">Convert values with empty strings automatically to DbNull</param>
        ''' <returns></returns>
        ''' <remarks>
        ''' In case of duplicate column names, all additional occurances of the same column name will be modified to use a unique column name
        ''' </remarks>
        Public Shared Function ReadDataTableFromCsvString(ByVal data As String, ByVal csvContainsColumnHeaders As Boolean, ByVal cultureFormatProvider As System.Globalization.CultureInfo, ByVal columnWidths As Integer(), Optional ByVal convertEmptyStringsToDBNull As Boolean = False) As DataTable
            Return CompuMaster.Data.CsvTools.ReadDataTableFromCsvString(data, csvContainsColumnHeaders, cultureFormatProvider, columnWidths, convertEmptyStringsToDBNull)
        End Function

        ''' <summary>
        '''     Read from a CSV table
        ''' </summary>
        ''' <param name="data">The content of a CSV file</param>
        ''' <param name="csvContainsColumnHeaders">Indicates wether column headers are present</param>
        ''' <param name="columnWidths">An array of integers specifying the fixed column widths</param>
        ''' <param name="convertEmptyStringsToDBNull">Convert values with empty strings automatically to DbNull</param>
        ''' <returns></returns>
        ''' <remarks>
        ''' In case of duplicate column names, all additional occurances of the same column name will be modified to use a unique column name
        ''' </remarks>
        Public Shared Function ReadDataTableFromCsvString(ByVal data As String, ByVal csvContainsColumnHeaders As Boolean, ByVal columnWidths As Integer(), Optional ByVal convertEmptyStringsToDBNull As Boolean = False) As DataTable
            Return CompuMaster.Data.CsvTools.ReadDataTableFromCsvString(data, csvContainsColumnHeaders, columnWidths, convertEmptyStringsToDBNull)
        End Function

        ''' <summary>
        '''     Write to a CSV file
        ''' </summary>
        ''' <param name="path"></param>
        ''' <param name="dataTable"></param>
        Public Shared Sub WriteDataTableToCsvFile(ByVal path As String, ByVal dataTable As System.Data.DataTable)
            CompuMaster.Data.CsvTools.WriteDataTableToCsvFile(path, dataTable)
        End Sub

        ''' <summary>
        '''     Write to a CSV with fixed column widths
        ''' </summary>
        ''' <param name="path">The path of the CSV file</param>
        ''' <param name="dataTable">A datatable which shall be exported</param>
        ''' <param name="writeCsvColumnHeaders">Add a row with column headers on the top</param>
        ''' <param name="columnWidths">An array of integers specifying the fixed column widths</param>
        ''' <param name="cultureFormatProvider">A culture format provider which declares list and decimal separators, etc.</param>
        ''' <param name="encoding">A file encoding (default UTF-8)</param>
        Public Shared Sub WriteDataTableToCsvFile(ByVal path As String, ByVal dataTable As System.Data.DataTable, ByVal writeCsvColumnHeaders As Boolean, ByVal columnWidths As Integer(), ByVal cultureFormatProvider As System.Globalization.CultureInfo, Optional ByVal encoding As String = "UTF-8")
            CompuMaster.Data.CsvTools.WriteDataTableToCsvFile(path, dataTable, writeCsvColumnHeaders, columnWidths, cultureFormatProvider, encoding)
        End Sub

        ''' <summary>
        '''     Write to a CSV file
        ''' </summary>
        ''' <param name="path">The path of the CSV file</param>
        ''' <param name="dataTable">A datatable which shall be exported</param>
        ''' <param name="writeCsvColumnHeaders">Add a row with column headers on the top</param>
        ''' <param name="cultureFormatProvider">A culture format provider which declares list and decimal separators, etc.</param>
        ''' <param name="encoding">A file encoding (default UTF-8)</param>
        Public Shared Sub WriteDataTableToCsvFile(ByVal path As String, ByVal dataTable As System.Data.DataTable, ByVal writeCsvColumnHeaders As Boolean, ByVal cultureFormatProvider As System.Globalization.CultureInfo, Optional ByVal encoding As String = "UTF-8")
            WriteDataTableToCsvFile(path, dataTable, writeCsvColumnHeaders, cultureFormatProvider, encoding, Nothing)
        End Sub

        ''' <summary>
        '''     Write to a CSV file
        ''' </summary>
        ''' <param name="path">The path of the CSV file</param>
        ''' <param name="dataTable">A datatable which shall be exported</param>
        ''' <param name="writeCsvColumnHeaders">Add a row with column headers on the top</param>
        ''' <param name="cultureFormatProvider">A culture format provider which declares list and decimal separators, etc.</param>
        ''' <param name="encoding">A file encoding (default UTF-8)</param>
        ''' <param name="columnSeparator">A column separator (culture default if empty value)</param>
        ''' <param name="recognizeTextBy">Recognize text by this character (default: quotation marks)</param>
        Public Shared Sub WriteDataTableToCsvFile(ByVal path As String, ByVal dataTable As System.Data.DataTable, ByVal writeCsvColumnHeaders As Boolean, ByVal cultureFormatProvider As System.Globalization.CultureInfo, ByVal encoding As String, ByVal columnSeparator As String, Optional ByVal recognizeTextBy As Char = """"c)
            If encoding = Nothing Then
                encoding = "UTF-8"
            End If
            If columnSeparator = Nothing Then
                columnSeparator = cultureFormatProvider.TextInfo.ListSeparator
            End If
            CompuMaster.Data.CsvTools.WriteDataTableToCsvFile(path, dataTable, writeCsvColumnHeaders, cultureFormatProvider, encoding, columnSeparator, recognizeTextBy)
        End Sub

        ''' <summary>
        '''     Convert the datatable to a string based, comma-separated format
        ''' </summary>
        ''' <param name="dataTable">A datatable which shall be exported</param>
        ''' <param name="writeCsvColumnHeaders">Add a row with column headers on the top</param>
        ''' <param name="cultureFormatProvider">A culture format provider which declares list and decimal separators, etc.</param>
        ''' <param name="columnSeparator">A column separator (culture default if empty value)</param>
        ''' <param name="recognizeTextBy">Recognize text by this character (default: quotation marks)</param>
        ''' <returns>A formatted text output</returns>
        Public Shared Function ConvertDataTableToText(ByVal dataTable As System.Data.DataTable, ByVal writeCsvColumnHeaders As Boolean, ByVal cultureFormatProvider As System.Globalization.CultureInfo, Optional ByVal columnSeparator As String = Nothing, Optional ByVal recognizeTextBy As Char = """"c) As String
            If columnSeparator = Nothing Then
                columnSeparator = cultureFormatProvider.TextInfo.ListSeparator
            End If
            Return CompuMaster.Data.CsvTools.ConvertDataTableToCsv(dataTable, writeCsvColumnHeaders, cultureFormatProvider, columnSeparator, recognizeTextBy)
        End Function

        ''' <summary>
        '''     Convert the datatable to a string based, comma-separated format
        ''' </summary>
        ''' <param name="dataTable">A datatable which shall be exported</param>
        ''' <param name="writeCsvColumnHeaders">Add a row with column headers on the top</param>
        ''' <param name="cultureFormatProvider">A culture format provider which declares list and decimal separators, etc.</param>
        ''' <param name="columnWidths">An array of integers specifying the fixed column widths</param>
        ''' <returns>The table as text with comma-separated structure</returns>
        Public Shared Function ConvertDataTableToText(ByVal dataTable As System.Data.DataTable, ByVal writeCsvColumnHeaders As Boolean, ByVal cultureFormatProvider As System.Globalization.CultureInfo, ByVal columnWidths As Integer()) As String
            Return CompuMaster.Data.CsvTools.ConvertDataTableToCsv(dataTable, writeCsvColumnHeaders, cultureFormatProvider, columnWidths)
        End Function

        ''' <summary>
        '''     Write to a CSV file
        ''' </summary>
        ''' <param name="path">The path of the file</param>
        ''' <param name="dataTable">A datatable which shall be exported</param>
        ''' <param name="writeCsvColumnHeaders">Indicates wether column headers are present</param>
        ''' <param name="encoding">The text encoding of the file</param>
        ''' <param name="columnSeparator">Choose the required character for splitting the columns. Set to null (Nothing in VisualBasic) to enable fixed column widths mode</param>
        ''' <param name="recognizeTextBy">A character indicating the start and end of text strings</param>
        Public Shared Sub WriteDataTableToCsvFile(ByVal path As String, ByVal dataTable As System.Data.DataTable, ByVal writeCsvColumnHeaders As Boolean, Optional ByVal encoding As String = "UTF-8", Optional ByVal columnSeparator As String = ","c, Optional ByVal recognizeTextBy As Char = """"c, Optional ByVal decimalSeparator As Char = "."c)
            CompuMaster.Data.CsvTools.WriteDataTableToCsvFile(path, dataTable, writeCsvColumnHeaders, encoding, columnSeparator, recognizeTextBy, decimalSeparator)
        End Sub

        ''' <summary>
        '''     Create a CSV table (contains BOF signature for unicode encodings)
        ''' </summary>
        ''' <param name="dataTable"></param>
        ''' <param name="writeCsvColumnHeaders">Indicates wether column headers are present</param>
        ''' <param name="encoding">The text encoding of the file</param>
        ''' <param name="columnSeparator">Choose the required character for splitting the columns. Set to null (Nothing in VisualBasic) to enable fixed column widths mode</param>
        ''' <param name="recognizeTextBy">A character indicating the start and end of text strings</param>
        ''' <param name="decimalSeparator"></param>
        ''' <returns>A string containing the CSV table with integrated file encoding for writing with e.g. System.IO.File.WriteAllText()</returns>
        <Obsolete("Better use WriteDataTableToCsvFileStringWithTextEncoding() instead"), ComponentModel.EditorBrowsable(ComponentModel.EditorBrowsableState.Never)>
        Public Shared Function WriteDataTableToCsvString(ByVal dataTable As System.Data.DataTable, ByVal writeCsvColumnHeaders As Boolean, Optional ByVal encoding As String = "UTF-8", Optional ByVal columnSeparator As Char = ","c, Optional ByVal recognizeTextBy As Char = """"c, Optional ByVal decimalSeparator As Char = "."c) As String
            Return WriteDataTableToCsvFileStringWithTextEncoding(dataTable, writeCsvColumnHeaders, encoding, columnSeparator, recognizeTextBy, decimalSeparator)
        End Function

        ''' <summary>
        '''     Create a CSV table (contains BOF signature for unicode encodings)
        ''' </summary>
        ''' <param name="dataTable"></param>
        ''' <param name="writeCsvColumnHeaders">Indicates wether column headers are present</param>
        ''' <param name="encoding">The text encoding of the file</param>
        ''' <param name="columnSeparator">Choose the required character for splitting the columns. Set to null (Nothing in VisualBasic) to enable fixed column widths mode</param>
        ''' <param name="recognizeTextBy">A character indicating the start and end of text strings</param>
        ''' <param name="decimalSeparator"></param>
        ''' <returns>A string containing the CSV table with integrated file encoding for writing with e.g. System.IO.File.WriteAllText()</returns>
        Public Shared Function WriteDataTableToCsvFileStringWithTextEncoding(ByVal dataTable As System.Data.DataTable, ByVal writeCsvColumnHeaders As Boolean, Optional ByVal encoding As String = "UTF-8", Optional ByVal columnSeparator As String = ",", Optional ByVal recognizeTextBy As Char = """"c, Optional ByVal decimalSeparator As Char = "."c) As String
            Dim MyStream As System.IO.MemoryStream = WriteDataTableToCsvMemoryStream(dataTable, writeCsvColumnHeaders, System.Text.Encoding.Unicode.EncodingName, columnSeparator, recognizeTextBy, decimalSeparator)
            Return System.Text.Encoding.Unicode.GetString(MyStream.ToArray)
        End Function

        ''' <summary>
        '''     Create a CSV table
        ''' </summary>
        ''' <param name="dataTable"></param>
        ''' <param name="writeCsvColumnHeaders">Indicates wether column headers are present</param>
        ''' <param name="columnSeparator">Choose the required character for splitting the columns. Set to null (Nothing in VisualBasic) to enable fixed column widths mode</param>
        ''' <param name="recognizeTextBy">A character indicating the start and end of text strings</param>
        ''' <param name="decimalSeparator"></param>
        ''' <returns>A string containing the CSV table</returns>
        Public Shared Function WriteDataTableToCsvTextString(ByVal dataTable As System.Data.DataTable, ByVal writeCsvColumnHeaders As Boolean, Optional ByVal columnSeparator As String = ",", Optional ByVal recognizeTextBy As Char = """"c, Optional ByVal decimalSeparator As Char = "."c) As String
            Dim WrittenStream As System.IO.MemoryStream = WriteDataTableToCsvMemoryStream(dataTable, writeCsvColumnHeaders, System.Text.Encoding.Unicode.EncodingName, columnSeparator, recognizeTextBy, decimalSeparator)
            Dim ReaderStream As New System.IO.MemoryStream(WrittenStream.ToArray)
            WrittenStream.Dispose()
            Dim SR As New System.IO.StreamReader(ReaderStream)
            Dim Result As String = SR.ReadToEnd()
            SR.Close()
            SR.Dispose()
            ReaderStream.Close()
            ReaderStream.Dispose()
            Return Result
        End Function

        ''' <summary>
        '''     Create a CSV table
        ''' </summary>
        ''' <param name="dataTable"></param>
        ''' <param name="writeCsvColumnHeaders">Indicates wether column headers are present</param>
        ''' <param name="encoding">The text encoding of the file</param>
        ''' <param name="columnSeparator">Choose the required character for splitting the columns. Set to null (Nothing in VisualBasic) to enable fixed column widths mode</param>
        ''' <param name="recognizeTextBy">A character indicating the start and end of text strings</param>
        ''' <param name="decimalSeparator"></param>
        ''' <returns>A string containing the CSV table</returns>
        Public Shared Function WriteDataTableToCsvBytes(ByVal dataTable As System.Data.DataTable, ByVal writeCsvColumnHeaders As Boolean, Optional ByVal encoding As String = "UTF-8", Optional ByVal columnSeparator As Char = ","c, Optional ByVal recognizeTextBy As Char = """"c, Optional ByVal decimalSeparator As Char = "."c) As Byte()
            Return CompuMaster.Data.CsvTools.WriteDataTableToCsvBytes(dataTable, writeCsvColumnHeaders, encoding, columnSeparator, recognizeTextBy, decimalSeparator)
        End Function

        ''' <summary>
        '''     Create a CSV table
        ''' </summary>
        ''' <param name="dataTable"></param>
        ''' <param name="writeCsvColumnHeaders">Indicates wether column headers are present</param>
        ''' <param name="encoding">The text encoding of the file</param>
        ''' <param name="cultureFormatProvider">A globalization information object for the conversion of all data to strings</param>
        ''' <param name="columnSeparator">Choose the required character for splitting the columns. Set to null (Nothing in VisualBasic) to enable fixed column widths mode</param>
        ''' <param name="recognizeTextBy">A character indicating the start and end of text strings</param>
        ''' <returns>A string containing the CSV table</returns>
        Public Shared Function WriteDataTableToCsvBytes(ByVal dataTable As System.Data.DataTable, ByVal writeCsvColumnHeaders As Boolean, ByVal encoding As System.Text.Encoding, ByVal cultureFormatProvider As System.Globalization.CultureInfo, Optional ByVal columnSeparator As Char = Nothing, Optional ByVal recognizeTextBy As Char = """"c) As Byte()
            Dim charColumnSeparator As Char
            If columnSeparator = Nothing Then
                charColumnSeparator = CType(cultureFormatProvider.TextInfo.ListSeparator, Char)
            Else
                charColumnSeparator = columnSeparator
            End If
            Return CompuMaster.Data.CsvTools.WriteDataTableToCsvBytes(dataTable, writeCsvColumnHeaders, encoding, cultureFormatProvider, charColumnSeparator, recognizeTextBy)
        End Function

        ''' <summary>
        '''     Create a CSV table
        ''' </summary>
        ''' <param name="dataTable"></param>
        ''' <param name="writeCsvColumnHeaders">Indicates wether column headers are present</param>
        ''' <param name="encoding">The text encoding of the file</param>
        ''' <param name="columnSeparator">Choose the required character for splitting the columns. Set to null (Nothing in VisualBasic) to enable fixed column widths mode</param>
        ''' <param name="recognizeTextBy">A character indicating the start and end of text strings</param>
        ''' <param name="decimalSeparator"></param>
        ''' <returns>A memory stream containing all texts as bytes in Unicode format</returns>
        Public Shared Function WriteDataTableToCsvMemoryStream(ByVal dataTable As System.Data.DataTable, ByVal writeCsvColumnHeaders As Boolean, Optional ByVal encoding As String = "UTF-8", Optional ByVal columnSeparator As String = ","c, Optional ByVal recognizeTextBy As Char = """"c, Optional ByVal decimalSeparator As Char = "."c) As System.IO.MemoryStream
            Return CompuMaster.Data.CsvTools.WriteDataTableToCsvMemoryStream(dataTable, writeCsvColumnHeaders, encoding, columnSeparator, recognizeTextBy, decimalSeparator)
        End Function

        ''' <summary>
        '''     Write to a CSV file
        ''' </summary>
        ''' <param name="path"></param>
        ''' <param name="dataview"></param>
        Public Shared Sub WriteDataViewToCsvFile(ByVal path As String, ByVal dataview As System.Data.DataView)
            CompuMaster.Data.CsvTools.WriteDataViewToCsvFile(path, dataview)
        End Sub

        ''' <summary>
        '''     Write to a CSV file
        ''' </summary>
        ''' <param name="path"></param>
        ''' <param name="dataView"></param>
        ''' <param name="writeCsvColumnHeaders"></param>
        ''' <param name="cultureFormatProvider"></param>
        ''' <param name="encoding"></param>
        Public Shared Sub WriteDataViewToCsvFile(ByVal path As String, ByVal dataView As System.Data.DataView, ByVal writeCsvColumnHeaders As Boolean, ByVal cultureFormatProvider As System.Globalization.CultureInfo, Optional ByVal encoding As String = "UTF-8", Optional ByVal columnSeparator As String = Nothing, Optional ByVal recognizeTextBy As Char = """"c)
            If columnSeparator = Nothing Then
                columnSeparator = cultureFormatProvider.TextInfo.ListSeparator
            End If
            CompuMaster.Data.CsvTools.WriteDataViewToCsvFile(path, dataView, writeCsvColumnHeaders, cultureFormatProvider, encoding, columnSeparator, recognizeTextBy)
        End Sub

        ''' <summary>
        '''     Write to a CSV file
        ''' </summary>
        ''' <param name="path"></param>
        ''' <param name="dataView"></param>
        ''' <param name="writeCsvColumnHeaders"></param>
        ''' <param name="encoding"></param>
        ''' <param name="columnSeparator"></param>
        ''' <param name="recognizeTextBy"></param>
        ''' <param name="decimalSeparator"></param>
        Public Shared Sub WriteDataViewToCsvFile(ByVal path As String, ByVal dataView As System.Data.DataView, ByVal writeCsvColumnHeaders As Boolean, Optional ByVal encoding As String = "UTF-8", Optional ByVal columnSeparator As String = ",", Optional ByVal recognizeTextBy As Char = """"c, Optional ByVal decimalSeparator As Char = "."c)
            CompuMaster.Data.CsvTools.WriteDataViewToCsvFile(path, dataView, writeCsvColumnHeaders, encoding, columnSeparator, recognizeTextBy, decimalSeparator)
        End Sub

    End Class

End Namespace
