Option Explicit On 
Option Strict On
Imports System.Data

Namespace CompuMaster.Data

    ''' <summary>
    '''     Provides simplified access to CSV files
    ''' </summary>
    Public NotInheritable Class Csv

        Public Enum WriteLineEncodings As Byte
            ''' <summary>
            ''' Platform dependent NewLine encoding for row separation, but keep line breaks in cells unchanged (danger: cell line breaks might conflict with line break of current platform)
            ''' </summary>
            None = 0
            RowBreakCrLf_CellLineBreakLf = 1
            RowBreakCrLf_CellLineBreakCr = 2
            RowBreakCr_CellLineBreakLf = 3
            RowBreakLf_CellLineBreakCr = 4
            RowBreakCrLf_CellLineBreakSpaceChar = 10 'replace line breaks into space char
            RowBreakCr_CellLineBreakSpaceChar = 11 'replace line breaks into space char
            RowBreakLf_CellLineBreakSpaceChar = 12 'replace line breaks into space char
            RowBreakCrLf_CellLineBreakRemoved = 13 'remove all line breaks
            RowBreakCr_CellLineBreakRemoved = 14 'remove all line breaks
            RowBreakLf_CellLineBreakRemoved = 15 'remove all line breaks
            ''' <summary>
            ''' Rule as RowBreakCrLf_CellLineBreakLf
            ''' </summary>
            [Default] = 1
            'Windows = 1
            'Mac = 2
            'Linux = 3
            ''' <summary>
            ''' Platform dependent NewLine encoding for row separation, line breaks in cells LF (Windows+Mac) or CR (Linux+Unix)
            ''' </summary>
            Auto = 255
        End Enum

        ''' <summary>
        ''' Line encoding of CSV files 
        ''' </summary>
        <CodeAnalysis.SuppressMessage("Design", "CA1027:Mark enums with FlagsAttribute", Justification:="<Ausstehend>")>
        Public Enum ReadLineEncodings As Byte
            ''' <summary>
            ''' Force reading line break in cell value as row break
            ''' </summary>
            None = 0
            RowBreakCrLf_CellLineBreakLf = 1
            RowBreakCrLf_CellLineBreakCr = 2
            RowBreakCr_CellLineBreakLf = 3
            RowBreakLf_CellLineBreakCr = 4
            RowBreakCrLfOrCr_CellLineBreakLf = 5
            RowBreakCrLfOrLf_CellLineBreakCr = 6
            ''' <summary>
            ''' WARNING: FEATURE STILL BETA DUE TO DESIGN ISSUES: Read lines for rows and detect cell line breaks by incomplete column data per row
            ''' </summary>
            ''' <remarks>
            ''' CURRENT DESIGN ISSUE WITH TROUBLE: LineBreaks in first and last column can't be identified if its for the previous row or for the next row since this data is missing in CSV file
            ''' </remarks>
            RowBreakCrLfOrCrOrLf_CellLineBreakCrLfOrCrOrLf = 7
            ''' <summary>
            ''' Rule as RowBreakCrLfOrCr_CellLineBreakLf
            ''' </summary>
            [Default] = 5
            ''' <summary>
            ''' Platform dependent NewLine encoding for row separation, line breaks in cells LF (Windows+Mac) or CR (Linux+Unix)
            ''' </summary>
            Auto = 255
        End Enum

        ''' <summary>
        ''' Auto conversion of detected line breaks in cell to platform specific linebreak
        ''' </summary>
        Public Enum ReadLineEncodingAutoConversion As Byte
            NoAutoConversion = 0
            AutoConvertLineBreakToCrLf = 1
            AutoConvertLineBreakToCr = 2
            AutoConvertLineBreakToLf = 3
            AutoConvertLineBreakToSystemEnvironmentNewLine = 4
        End Enum

        ''' <summary>
        '''     Read from a CSV file
        ''' </summary>
        ''' <param name="path">The path of the file</param>
        ''' <param name="csvContainsColumnHeaders">Indicates wether column headers are present</param>
        ''' <param name="fileEncoding">The text encoding of the file</param>
        ''' <param name="columnSeparator">Choose the required character for splitting the columns. Set to null (Nothing in VisualBasic) to enable fixed column widths mode</param>
        ''' <param name="recognizeTextBy">A character indicating the start and end of text strings</param>
        ''' <param name="recognizeMultipleColumnSeparatorCharsAsOne">Specifies whether multiple seperator characters should be recognized as one</param>
        ''' <param name="convertEmptyStringsToDBNull">Convert values with empty strings automatically to DbNull</param>
        ''' <returns></returns>
        ''' <remarks>
        ''' In case of duplicate column names, all additional occurances of the same column name will be modified to use a unique column name
        ''' </remarks>
        Public Shared Function ReadDataTableFromCsvFile(ByVal path As String, ByVal csvContainsColumnHeaders As Boolean, Optional ByVal fileEncoding As String = "UTF-8", Optional ByVal columnSeparator As Char = ","c, Optional ByVal recognizeTextBy As Char = """"c, Optional ByVal recognizeMultipleColumnSeparatorCharsAsOne As Boolean = False, Optional ByVal convertEmptyStringsToDBNull As Boolean = False) As DataTable
            Return CompuMaster.Data.CsvTools.ReadDataTableFromCsvFile(path, csvContainsColumnHeaders, fileEncoding, columnSeparator, recognizeTextBy, recognizeMultipleColumnSeparatorCharsAsOne, convertEmptyStringsToDBNull, ReadLineEncodings.Default, ReadLineEncodingAutoConversion.AutoConvertLineBreakToSystemEnvironmentNewLine)
        End Function

        ''' <summary>
        '''     Read from a CSV file
        ''' </summary>
        ''' <param name="path">The path of the file</param>
        ''' <param name="csvContainsColumnHeaders">Indicates wether column headers are present</param>
        ''' <param name="fileEncoding">The text encoding of the file</param>
        ''' <param name="cultureFormatProvider">A culture for all conversions</param>
        ''' <param name="recognizeTextBy">A character indicating the start and end of text strings</param>
        ''' <param name="recognizeDoubledColumnSeparatorCharAsOne">Currently without purpose</param>
        ''' <param name="convertEmptyStringsToDBNull">Convert values with empty strings automatically to DbNull</param>
        ''' <returns></returns>
        ''' <remarks>
        ''' In case of duplicate column names, all additional occurances of the same column name will be modified to use a unique column name
        ''' </remarks>
        Public Shared Function ReadDataTableFromCsvFile(ByVal path As String, ByVal csvContainsColumnHeaders As Boolean, ByVal fileEncoding As System.Text.Encoding, ByVal cultureFormatProvider As System.Globalization.CultureInfo, Optional ByVal recognizeTextBy As Char = """"c, Optional ByVal recognizeDoubledColumnSeparatorCharAsOne As Boolean = False, Optional ByVal convertEmptyStringsToDBNull As Boolean = False) As DataTable
            Return CompuMaster.Data.CsvTools.ReadDataTableFromCsvFile(path, csvContainsColumnHeaders, fileEncoding, cultureFormatProvider, recognizeTextBy, recognizeDoubledColumnSeparatorCharAsOne, convertEmptyStringsToDBNull, ReadLineEncodings.Default, ReadLineEncodingAutoConversion.AutoConvertLineBreakToSystemEnvironmentNewLine)
        End Function

        ''' <summary>
        '''     Read from a CSV table
        ''' </summary>
        ''' <param name="data">The content of a CSV file</param>
        ''' <param name="csvContainsColumnHeaders">Indicates wether column headers are present</param>
        ''' <param name="cultureFormatProvider">A culture for all conversions</param>
        ''' <param name="recognizeTextBy">A character indicating the start and end of text strings</param>
        ''' <param name="recognizeDoubledColumnSeparatorCharAsOne">Currently without purpose</param>
        ''' <param name="convertEmptyStringsToDBNull">Convert values with empty strings automatically to DbNull</param>
        ''' <returns></returns>
        ''' <remarks>
        ''' In case of duplicate column names, all additional occurances of the same column name will be modified to use a unique column name
        ''' </remarks>
        Public Shared Function ReadDataTableFromCsvString(ByVal data As String, ByVal csvContainsColumnHeaders As Boolean, ByVal cultureFormatProvider As System.Globalization.CultureInfo, Optional ByVal recognizeTextBy As Char = """"c, Optional ByVal recognizeDoubledColumnSeparatorCharAsOne As Boolean = False, Optional ByVal convertEmptyStringsToDBNull As Boolean = False) As DataTable
            Return CompuMaster.Data.CsvTools.ReadDataTableFromCsvString(data, csvContainsColumnHeaders, cultureFormatProvider, recognizeTextBy, recognizeDoubledColumnSeparatorCharAsOne, convertEmptyStringsToDBNull, ReadLineEncodings.Default, ReadLineEncodingAutoConversion.AutoConvertLineBreakToSystemEnvironmentNewLine)
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
        Public Shared Function ReadDataTableFromCsvString(ByVal data As String, ByVal csvContainsColumnHeaders As Boolean, Optional ByVal columnSeparator As Char = ","c, Optional ByVal recognizeTextBy As Char = """"c, Optional ByVal recognizeDoubledColumnSeparatorCharAsOne As Boolean = False, Optional ByVal convertEmptyStringsToDBNull As Boolean = False) As DataTable
            Return CompuMaster.Data.CsvTools.ReadDataTableFromCsvString(data, csvContainsColumnHeaders, columnSeparator, recognizeTextBy, recognizeDoubledColumnSeparatorCharAsOne, convertEmptyStringsToDBNull, ReadLineEncodings.Default, ReadLineEncodingAutoConversion.AutoConvertLineBreakToSystemEnvironmentNewLine)
        End Function

        ''' <summary>
        '''     Read from a CSV file
        ''' </summary>
        ''' <param name="path">The path of the file</param>
        ''' <param name="csvContainsColumnHeaders">Indicates wether column headers are present</param>
        ''' <param name="fileEncoding">The text encoding of the file</param>
        ''' <param name="columnWidths">An array of integers specifying the fixed column widths</param>
        ''' <param name="convertEmptyStringsToDBNull">Convert values with empty strings automatically to DbNull</param>
        ''' <returns></returns>
        ''' <remarks>
        ''' In case of duplicate column names, all additional occurances of the same column name will be modified to use a unique column name
        ''' </remarks>
        Public Shared Function ReadDataTableFromCsvFile(ByVal path As String, ByVal csvContainsColumnHeaders As Boolean, ByVal columnWidths As Integer(), Optional ByVal fileEncoding As String = "UTF-8", Optional ByVal convertEmptyStringsToDBNull As Boolean = False) As DataTable
            Return CompuMaster.Data.CsvTools.ReadDataTableFromCsvFile(path, csvContainsColumnHeaders, columnWidths, fileEncoding, convertEmptyStringsToDBNull, ReadLineEncodings.Default, ReadLineEncodingAutoConversion.AutoConvertLineBreakToSystemEnvironmentNewLine)
        End Function

        ''' <summary>
        '''     Read from a CSV file
        ''' </summary>
        ''' <param name="path">The path of the file</param>
        ''' <param name="csvContainsColumnHeaders">Indicates wether column headers are present</param>
        ''' <param name="columnWidths">An array of integers specifying the fixed column widths</param>
        ''' <param name="fileEncoding">The text encoding of the file</param>
        ''' <param name="cultureFormatProvider">A culture for all conversions</param>
        ''' <param name="convertEmptyStringsToDBNull">Convert values with empty strings automatically to DbNull</param>
        ''' <returns></returns>
        ''' <remarks>
        ''' In case of duplicate column names, all additional occurances of the same column name will be modified to use a unique column name
        ''' </remarks>
        Public Shared Function ReadDataTableFromCsvFile(ByVal path As String, ByVal csvContainsColumnHeaders As Boolean, ByVal columnWidths As Integer(), ByVal fileEncoding As System.Text.Encoding, ByVal cultureFormatProvider As System.Globalization.CultureInfo, Optional ByVal convertEmptyStringsToDBNull As Boolean = False) As DataTable
            Return CompuMaster.Data.CsvTools.ReadDataTableFromCsvFile(path, csvContainsColumnHeaders, columnWidths, fileEncoding, cultureFormatProvider, convertEmptyStringsToDBNull, ReadLineEncodings.Default, ReadLineEncodingAutoConversion.AutoConvertLineBreakToSystemEnvironmentNewLine)
        End Function

        ''' <summary>
        '''     Read from a CSV table
        ''' </summary>
        ''' <param name="data">The content of a CSV file</param>
        ''' <param name="csvContainsColumnHeaders">Indicates wether column headers are present</param>
        ''' <param name="cultureFormatProvider">A culture for all conversions</param>
        ''' <param name="columnWidths">An array of integers specifying the fixed column widths</param>
        ''' <param name="convertEmptyStringsToDBNull">Convert values with empty strings automatically to DbNull</param>
        ''' <returns></returns>
        ''' <remarks>
        ''' In case of duplicate column names, all additional occurances of the same column name will be modified to use a unique column name
        ''' </remarks>
        Public Shared Function ReadDataTableFromCsvString(ByVal data As String, ByVal csvContainsColumnHeaders As Boolean, ByVal cultureFormatProvider As System.Globalization.CultureInfo, ByVal columnWidths As Integer(), Optional ByVal convertEmptyStringsToDBNull As Boolean = False) As DataTable
            Return CompuMaster.Data.CsvTools.ReadDataTableFromCsvString(data, csvContainsColumnHeaders, cultureFormatProvider, columnWidths, convertEmptyStringsToDBNull, ReadLineEncodings.Default, ReadLineEncodingAutoConversion.AutoConvertLineBreakToSystemEnvironmentNewLine)
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
            Return CompuMaster.Data.CsvTools.ReadDataTableFromCsvString(data, csvContainsColumnHeaders, columnWidths, convertEmptyStringsToDBNull, ReadLineEncodings.Default, ReadLineEncodingAutoConversion.AutoConvertLineBreakToSystemEnvironmentNewLine)
        End Function

        ''' <summary>
        '''     Read from a CSV file
        ''' </summary>
        ''' <param name="path">The path of the file</param>
        ''' <param name="csvContainsColumnHeaders">Indicates wether column headers are present</param>
        ''' <param name="fileEncoding">The text encoding of the file</param>
        ''' <param name="columnSeparator">Choose the required character for splitting the columns. Set to null (Nothing in VisualBasic) to enable fixed column widths mode</param>
        ''' <param name="recognizeTextBy">A character indicating the start and end of text strings</param>
        ''' <param name="recognizeMultipleColumnSeparatorCharsAsOne">Specifies whether multiple seperator characters should be recognized as one</param>
        ''' <param name="convertEmptyStringsToDBNull">Convert values with empty strings automatically to DbNull</param>
        ''' <param name="lineEncodings">Encoding style for linebreaks</param>
        ''' <param name="lineEncodingAutoConversions">Change linebreak encodings on reading</param>
        ''' <returns></returns>
        ''' <remarks>
        ''' In case of duplicate column names, all additional occurances of the same column name will be modified to use a unique column name
        ''' </remarks>
        Public Shared Function ReadDataTableFromCsvFile(ByVal path As String, ByVal csvContainsColumnHeaders As Boolean, lineEncodings As CompuMaster.Data.Csv.ReadLineEncodings, lineEncodingAutoConversions As CompuMaster.Data.Csv.ReadLineEncodingAutoConversion, Optional ByVal fileEncoding As String = "UTF-8", Optional ByVal columnSeparator As Char = ","c, Optional ByVal recognizeTextBy As Char = """"c, Optional ByVal recognizeMultipleColumnSeparatorCharsAsOne As Boolean = False, Optional ByVal convertEmptyStringsToDBNull As Boolean = False) As DataTable
            Return CompuMaster.Data.CsvTools.ReadDataTableFromCsvFile(path, csvContainsColumnHeaders, fileEncoding, columnSeparator, recognizeTextBy, recognizeMultipleColumnSeparatorCharsAsOne, convertEmptyStringsToDBNull, lineEncodings, lineEncodingAutoConversions)
        End Function

        ''' <summary>
        '''     Read from a CSV file
        ''' </summary>
        ''' <param name="path">The path of the file</param>
        ''' <param name="csvContainsColumnHeaders">Indicates wether column headers are present</param>
        ''' <param name="fileEncoding">The text encoding of the file</param>
        ''' <param name="cultureFormatProvider">A culture for all conversions</param>
        ''' <param name="recognizeTextBy">A character indicating the start and end of text strings</param>
        ''' <param name="recognizeDoubledColumnSeparatorCharAsOne">Currently without purpose</param>
        ''' <param name="convertEmptyStringsToDBNull">Convert values with empty strings automatically to DbNull</param>
        ''' <param name="lineEncodings">Encoding style for linebreaks</param>
        ''' <param name="lineEncodingAutoConversions">Change linebreak encodings on reading</param>
        ''' <returns></returns>
        ''' <remarks>
        ''' In case of duplicate column names, all additional occurances of the same column name will be modified to use a unique column name
        ''' </remarks>
        Public Shared Function ReadDataTableFromCsvFile(ByVal path As String, ByVal csvContainsColumnHeaders As Boolean, lineEncodings As CompuMaster.Data.Csv.ReadLineEncodings, lineEncodingAutoConversions As CompuMaster.Data.Csv.ReadLineEncodingAutoConversion, ByVal fileEncoding As System.Text.Encoding, ByVal cultureFormatProvider As System.Globalization.CultureInfo, Optional ByVal recognizeTextBy As Char = """"c, Optional ByVal recognizeDoubledColumnSeparatorCharAsOne As Boolean = False, Optional ByVal convertEmptyStringsToDBNull As Boolean = False) As DataTable
            Return CompuMaster.Data.CsvTools.ReadDataTableFromCsvFile(path, csvContainsColumnHeaders, fileEncoding, cultureFormatProvider, recognizeTextBy, recognizeDoubledColumnSeparatorCharAsOne, convertEmptyStringsToDBNull, lineEncodings, lineEncodingAutoConversions)
        End Function

        ''' <summary>
        '''     Read from a CSV table
        ''' </summary>
        ''' <param name="data">The content of a CSV file</param>
        ''' <param name="csvContainsColumnHeaders">Indicates wether column headers are present</param>
        ''' <param name="cultureFormatProvider">A culture for all conversions</param>
        ''' <param name="recognizeTextBy">A character indicating the start and end of text strings</param>
        ''' <param name="recognizeDoubledColumnSeparatorCharAsOne">Currently without purpose</param>
        ''' <param name="convertEmptyStringsToDBNull">Convert values with empty strings automatically to DbNull</param>
        ''' <param name="lineEncodings">Encoding style for linebreaks</param>
        ''' <param name="lineEncodingAutoConversions">Change linebreak encodings on reading</param>
        ''' <returns></returns>
        ''' <remarks>
        ''' In case of duplicate column names, all additional occurances of the same column name will be modified to use a unique column name
        ''' </remarks>
        Public Shared Function ReadDataTableFromCsvString(ByVal data As String, ByVal csvContainsColumnHeaders As Boolean, lineEncodings As CompuMaster.Data.Csv.ReadLineEncodings, lineEncodingAutoConversions As CompuMaster.Data.Csv.ReadLineEncodingAutoConversion, ByVal cultureFormatProvider As System.Globalization.CultureInfo, Optional ByVal recognizeTextBy As Char = """"c, Optional ByVal recognizeDoubledColumnSeparatorCharAsOne As Boolean = False, Optional ByVal convertEmptyStringsToDBNull As Boolean = False) As DataTable
            Return CompuMaster.Data.CsvTools.ReadDataTableFromCsvString(data, csvContainsColumnHeaders, cultureFormatProvider, recognizeTextBy, recognizeDoubledColumnSeparatorCharAsOne, convertEmptyStringsToDBNull, lineEncodings, lineEncodingAutoConversions)
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
        ''' <param name="lineEncodings">Encoding style for linebreaks</param>
        ''' <param name="lineEncodingAutoConversions">Change linebreak encodings on reading</param>
        ''' <returns></returns>
        ''' <remarks>
        ''' In case of duplicate column names, all additional occurances of the same column name will be modified to use a unique column name
        ''' </remarks>
        Public Shared Function ReadDataTableFromCsvString(ByVal data As String, ByVal csvContainsColumnHeaders As Boolean, lineEncodings As CompuMaster.Data.Csv.ReadLineEncodings, lineEncodingAutoConversions As CompuMaster.Data.Csv.ReadLineEncodingAutoConversion, Optional ByVal columnSeparator As Char = ","c, Optional ByVal recognizeTextBy As Char = """"c, Optional ByVal recognizeDoubledColumnSeparatorCharAsOne As Boolean = False, Optional ByVal convertEmptyStringsToDBNull As Boolean = False) As DataTable
            Return CompuMaster.Data.CsvTools.ReadDataTableFromCsvString(data, csvContainsColumnHeaders, columnSeparator, recognizeTextBy, recognizeDoubledColumnSeparatorCharAsOne, convertEmptyStringsToDBNull, lineEncodings, lineEncodingAutoConversions)
        End Function

        ''' <summary>
        '''     Read from a CSV file
        ''' </summary>
        ''' <param name="path">The path of the file</param>
        ''' <param name="csvContainsColumnHeaders">Indicates wether column headers are present</param>
        ''' <param name="fileEncoding">The text encoding of the file</param>
        ''' <param name="columnWidths">An array of integers specifying the fixed column widths</param>
        ''' <param name="convertEmptyStringsToDBNull">Convert values with empty strings automatically to DbNull</param>
        ''' <param name="lineEncodings">Encoding style for linebreaks</param>
        ''' <param name="lineEncodingAutoConversions">Change linebreak encodings on reading</param>
        ''' <returns></returns>
        ''' <remarks>
        ''' In case of duplicate column names, all additional occurances of the same column name will be modified to use a unique column name
        ''' </remarks>
        Public Shared Function ReadDataTableFromCsvFile(ByVal path As String, ByVal csvContainsColumnHeaders As Boolean, lineEncodings As CompuMaster.Data.Csv.ReadLineEncodings, lineEncodingAutoConversions As CompuMaster.Data.Csv.ReadLineEncodingAutoConversion, ByVal columnWidths As Integer(), Optional ByVal fileEncoding As String = "UTF-8", Optional ByVal convertEmptyStringsToDBNull As Boolean = False) As DataTable
            Return CompuMaster.Data.CsvTools.ReadDataTableFromCsvFile(path, csvContainsColumnHeaders, columnWidths, fileEncoding, convertEmptyStringsToDBNull, lineEncodings, lineEncodingAutoConversions)
        End Function

        ''' <summary>
        '''     Read from a CSV file
        ''' </summary>
        ''' <param name="path">The path of the file</param>
        ''' <param name="csvContainsColumnHeaders">Indicates wether column headers are present</param>
        ''' <param name="columnWidths">An array of integers specifying the fixed column widths</param>
        ''' <param name="fileEncoding">The text encoding of the file</param>
        ''' <param name="cultureFormatProvider">A culture for all conversions</param>
        ''' <param name="convertEmptyStringsToDBNull">Convert values with empty strings automatically to DbNull</param>
        ''' <param name="lineEncodings">Encoding style for linebreaks</param>
        ''' <param name="lineEncodingAutoConversions">Change linebreak encodings on reading</param>
        ''' <returns></returns>
        ''' <remarks>
        ''' In case of duplicate column names, all additional occurances of the same column name will be modified to use a unique column name
        ''' </remarks>
        Public Shared Function ReadDataTableFromCsvFile(ByVal path As String, ByVal csvContainsColumnHeaders As Boolean, lineEncodings As CompuMaster.Data.Csv.ReadLineEncodings, lineEncodingAutoConversions As CompuMaster.Data.Csv.ReadLineEncodingAutoConversion, ByVal columnWidths As Integer(), ByVal fileEncoding As System.Text.Encoding, ByVal cultureFormatProvider As System.Globalization.CultureInfo, Optional ByVal convertEmptyStringsToDBNull As Boolean = False) As DataTable
            Return CompuMaster.Data.CsvTools.ReadDataTableFromCsvFile(path, csvContainsColumnHeaders, columnWidths, fileEncoding, cultureFormatProvider, convertEmptyStringsToDBNull, lineEncodings, lineEncodingAutoConversions)
        End Function

        ''' <summary>
        '''     Read from a CSV table
        ''' </summary>
        ''' <param name="data">The content of a CSV file</param>
        ''' <param name="csvContainsColumnHeaders">Indicates wether column headers are present</param>
        ''' <param name="cultureFormatProvider">A culture for all conversions</param>
        ''' <param name="columnWidths">An array of integers specifying the fixed column widths</param>
        ''' <param name="convertEmptyStringsToDBNull">Convert values with empty strings automatically to DbNull</param>
        ''' <param name="lineEncodings">Encoding style for linebreaks</param>
        ''' <param name="lineEncodingAutoConversions">Change linebreak encodings on reading</param>
        ''' <returns></returns>
        ''' <remarks>
        ''' In case of duplicate column names, all additional occurances of the same column name will be modified to use a unique column name
        ''' </remarks>
        Public Shared Function ReadDataTableFromCsvString(ByVal data As String, ByVal csvContainsColumnHeaders As Boolean, lineEncodings As CompuMaster.Data.Csv.ReadLineEncodings, lineEncodingAutoConversions As CompuMaster.Data.Csv.ReadLineEncodingAutoConversion, ByVal cultureFormatProvider As System.Globalization.CultureInfo, ByVal columnWidths As Integer(), Optional ByVal convertEmptyStringsToDBNull As Boolean = False) As DataTable
            Return CompuMaster.Data.CsvTools.ReadDataTableFromCsvString(data, csvContainsColumnHeaders, cultureFormatProvider, columnWidths, convertEmptyStringsToDBNull, lineEncodings, lineEncodingAutoConversions)
        End Function

        ''' <summary>
        '''     Read from a CSV table
        ''' </summary>
        ''' <param name="data">The content of a CSV file</param>
        ''' <param name="csvContainsColumnHeaders">Indicates wether column headers are present</param>
        ''' <param name="columnWidths">An array of integers specifying the fixed column widths</param>
        ''' <param name="convertEmptyStringsToDBNull">Convert values with empty strings automatically to DbNull</param>
        ''' <param name="lineEncodings">Encoding style for linebreaks</param>
        ''' <param name="lineEncodingAutoConversions">Change linebreak encodings on reading</param>
        ''' <returns></returns>
        ''' <remarks>
        ''' In case of duplicate column names, all additional occurances of the same column name will be modified to use a unique column name
        ''' </remarks>
        Public Shared Function ReadDataTableFromCsvString(ByVal data As String, ByVal csvContainsColumnHeaders As Boolean, lineEncodings As CompuMaster.Data.Csv.ReadLineEncodings, lineEncodingAutoConversions As CompuMaster.Data.Csv.ReadLineEncodingAutoConversion, ByVal columnWidths As Integer(), Optional ByVal convertEmptyStringsToDBNull As Boolean = False) As DataTable
            Return CompuMaster.Data.CsvTools.ReadDataTableFromCsvString(data, csvContainsColumnHeaders, columnWidths, convertEmptyStringsToDBNull, lineEncodings, lineEncodingAutoConversions)
        End Function

        ''' <summary>
        '''     Write to a CSV file
        ''' </summary>
        ''' <param name="path">The path of the file</param>
        ''' <param name="dataTable">A datatable which shall be exported</param>
        Public Shared Sub WriteDataTableToCsvFile(ByVal path As String, ByVal dataTable As System.Data.DataTable)
            CompuMaster.Data.CsvTools.WriteDataTableToCsvFile(path, dataTable, WriteLineEncodings.Default)
        End Sub

        ''' <summary>
        '''     Write to a CSV file
        ''' </summary>
        ''' <param name="path">The path of the file</param>
        ''' <param name="dataTable">A datatable which shall be exported</param>
        ''' <param name="lineEncodings">Encoding style for linebreaks</param>
        Public Shared Sub WriteDataTableToCsvFile(ByVal path As String, ByVal dataTable As System.Data.DataTable, lineEncodings As CompuMaster.Data.Csv.WriteLineEncodings)
            CompuMaster.Data.CsvTools.WriteDataTableToCsvFile(path, dataTable, lineEncodings)
        End Sub

        ''' <summary>
        '''     Write to a CSV with fixed column widths
        ''' </summary>
        ''' <param name="path">The path of the CSV file</param>
        ''' <param name="dataTable">A datatable which shall be exported</param>
        ''' <param name="writeCsvColumnHeaders">Add a row with column headers on the top</param>
        ''' <param name="columnWidths">An array of integers specifying the fixed column widths</param>
        ''' <param name="cultureFormatProvider">A culture format provider which declares list and decimal separators, etc.</param>
        ''' <param name="fileEncoding">A file encoding (default UTF-8)</param>
        Public Shared Sub WriteDataTableToCsvFile(ByVal path As String, ByVal dataTable As System.Data.DataTable, ByVal writeCsvColumnHeaders As Boolean, ByVal columnWidths As Integer(), ByVal cultureFormatProvider As System.Globalization.CultureInfo, Optional ByVal fileEncoding As String = "UTF-8")
            CompuMaster.Data.CsvTools.WriteDataTableToCsvFile(path, dataTable, writeCsvColumnHeaders, columnWidths, cultureFormatProvider, fileEncoding, WriteLineEncodings.Default)
        End Sub

        ''' <summary>
        '''     Write to a CSV with fixed column widths
        ''' </summary>
        ''' <param name="path">The path of the CSV file</param>
        ''' <param name="dataTable">A datatable which shall be exported</param>
        ''' <param name="writeCsvColumnHeaders">Add a row with column headers on the top</param>
        ''' <param name="columnWidths">An array of integers specifying the fixed column widths</param>
        ''' <param name="cultureFormatProvider">A culture format provider which declares list and decimal separators, etc.</param>
        ''' <param name="fileEncoding">A file encoding (default UTF-8)</param>
        ''' <param name="lineEncodings">Encoding style for linebreaks</param>
        Public Shared Sub WriteDataTableToCsvFile(ByVal path As String, ByVal dataTable As System.Data.DataTable, ByVal writeCsvColumnHeaders As Boolean, lineEncodings As CompuMaster.Data.Csv.WriteLineEncodings, ByVal columnWidths As Integer(), ByVal cultureFormatProvider As System.Globalization.CultureInfo, Optional ByVal fileEncoding As String = "UTF-8")
            CompuMaster.Data.CsvTools.WriteDataTableToCsvFile(path, dataTable, writeCsvColumnHeaders, columnWidths, cultureFormatProvider, fileEncoding, lineEncodings)
        End Sub

        ''' <summary>
        '''     Write to a CSV file
        ''' </summary>
        ''' <param name="path">The path of the CSV file</param>
        ''' <param name="dataTable">A datatable which shall be exported</param>
        ''' <param name="writeCsvColumnHeaders">Add a row with column headers on the top</param>
        ''' <param name="cultureFormatProvider">A culture format provider which declares list and decimal separators, etc.</param>
        ''' <param name="fileEncoding">A file encoding (default UTF-8)</param>
        Public Shared Sub WriteDataTableToCsvFile(ByVal path As String, ByVal dataTable As System.Data.DataTable, ByVal writeCsvColumnHeaders As Boolean, ByVal cultureFormatProvider As System.Globalization.CultureInfo, Optional ByVal fileEncoding As String = "UTF-8")
            WriteDataTableToCsvFile(path, dataTable, writeCsvColumnHeaders, cultureFormatProvider, fileEncoding, Nothing)
        End Sub

        ''' <summary>
        '''     Write to a CSV file
        ''' </summary>
        ''' <param name="path">The path of the CSV file</param>
        ''' <param name="dataTable">A datatable which shall be exported</param>
        ''' <param name="writeCsvColumnHeaders">Add a row with column headers on the top</param>
        ''' <param name="cultureFormatProvider">A culture format provider which declares list and decimal separators, etc.</param>
        ''' <param name="fileEncoding">A file encoding (default UTF-8)</param>
        ''' <param name="lineEncodings">Encoding style for linebreaks</param>
        Public Shared Sub WriteDataTableToCsvFile(ByVal path As String, ByVal dataTable As System.Data.DataTable, ByVal writeCsvColumnHeaders As Boolean, lineEncodings As CompuMaster.Data.Csv.WriteLineEncodings, ByVal cultureFormatProvider As System.Globalization.CultureInfo, Optional ByVal fileEncoding As String = "UTF-8")
            WriteDataTableToCsvFile(path, dataTable, writeCsvColumnHeaders, lineEncodings, cultureFormatProvider, fileEncoding, Nothing)
        End Sub

        ''' <summary>
        '''     Write to a CSV file
        ''' </summary>
        ''' <param name="path">The path of the CSV file</param>
        ''' <param name="dataTable">A datatable which shall be exported</param>
        ''' <param name="writeCsvColumnHeaders">Add a row with column headers on the top</param>
        ''' <param name="cultureFormatProvider">A culture format provider which declares list and decimal separators, etc.</param>
        ''' <param name="fileEncoding">A file encoding (default UTF-8)</param>
        ''' <param name="columnSeparator">A column separator (culture default if empty value)</param>
        ''' <param name="recognizeTextBy">Recognize text by this character (default: quotation marks)</param>
        Public Shared Sub WriteDataTableToCsvFile(ByVal path As String, ByVal dataTable As System.Data.DataTable, ByVal writeCsvColumnHeaders As Boolean, ByVal cultureFormatProvider As System.Globalization.CultureInfo, ByVal fileEncoding As String, ByVal columnSeparator As String, Optional ByVal recognizeTextBy As Char = """"c)
            If fileEncoding = Nothing Then
                fileEncoding = "UTF-8"
            End If
            If columnSeparator = Nothing Then
                columnSeparator = cultureFormatProvider.TextInfo.ListSeparator
            End If
            CompuMaster.Data.CsvTools.WriteDataTableToCsvFile(path, dataTable, writeCsvColumnHeaders, cultureFormatProvider, fileEncoding, columnSeparator, recognizeTextBy, WriteLineEncodings.Default)
        End Sub

        ''' <summary>
        '''     Write to a CSV file
        ''' </summary>
        ''' <param name="path">The path of the CSV file</param>
        ''' <param name="dataTable">A datatable which shall be exported</param>
        ''' <param name="writeCsvColumnHeaders">Add a row with column headers on the top</param>
        ''' <param name="cultureFormatProvider">A culture format provider which declares list and decimal separators, etc.</param>
        ''' <param name="fileEncoding">A file encoding (default UTF-8)</param>
        ''' <param name="columnSeparator">A column separator (culture default if empty value)</param>
        ''' <param name="recognizeTextBy">Recognize text by this character (default: quotation marks)</param>
        ''' <param name="lineEncodings">Encoding style for linebreaks</param>
        Public Shared Sub WriteDataTableToCsvFile(ByVal path As String, ByVal dataTable As System.Data.DataTable, ByVal writeCsvColumnHeaders As Boolean, lineEncodings As CompuMaster.Data.Csv.WriteLineEncodings, ByVal cultureFormatProvider As System.Globalization.CultureInfo, ByVal fileEncoding As String, ByVal columnSeparator As String, Optional ByVal recognizeTextBy As Char = """"c)
            If fileEncoding = Nothing Then
                fileEncoding = "UTF-8"
            End If
            If columnSeparator = Nothing Then
                columnSeparator = cultureFormatProvider.TextInfo.ListSeparator
            End If
            CompuMaster.Data.CsvTools.WriteDataTableToCsvFile(path, dataTable, writeCsvColumnHeaders, cultureFormatProvider, fileEncoding, columnSeparator, recognizeTextBy, lineEncodings)
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
        Public Shared Function ConvertDataTableToTextAsStringBuilder(ByVal dataTable As System.Data.DataTable, ByVal writeCsvColumnHeaders As Boolean, ByVal cultureFormatProvider As System.Globalization.CultureInfo, Optional ByVal columnSeparator As String = Nothing, Optional ByVal recognizeTextBy As Char = """"c) As System.Text.StringBuilder
            If columnSeparator = Nothing Then
                columnSeparator = cultureFormatProvider.TextInfo.ListSeparator
            End If
            Return CompuMaster.Data.CsvTools.ConvertDataTableToCsv(dataTable, writeCsvColumnHeaders, cultureFormatProvider, columnSeparator, recognizeTextBy, WriteLineEncodings.Default)
        End Function

        ''' <summary>
        '''     Convert the datatable to a string based, comma-separated format
        ''' </summary>
        ''' <param name="dataTable">A datatable which shall be exported</param>
        ''' <param name="writeCsvColumnHeaders">Add a row with column headers on the top</param>
        ''' <param name="cultureFormatProvider">A culture format provider which declares list and decimal separators, etc.</param>
        ''' <param name="columnSeparator">A column separator (culture default if empty value)</param>
        ''' <param name="recognizeTextBy">Recognize text by this character (default: quotation marks)</param>
        ''' <param name="lineEncodings">Encoding style for linebreaks</param>
        ''' <returns>A formatted text output</returns>
        Public Shared Function ConvertDataTableToTextAsStringBuilder(ByVal dataTable As System.Data.DataTable, ByVal writeCsvColumnHeaders As Boolean, lineEncodings As CompuMaster.Data.Csv.WriteLineEncodings, ByVal cultureFormatProvider As System.Globalization.CultureInfo, Optional ByVal columnSeparator As String = Nothing, Optional ByVal recognizeTextBy As Char = """"c) As System.Text.StringBuilder
            If columnSeparator = Nothing Then
                columnSeparator = cultureFormatProvider.TextInfo.ListSeparator
            End If
            Return CompuMaster.Data.CsvTools.ConvertDataTableToCsv(dataTable, writeCsvColumnHeaders, cultureFormatProvider, columnSeparator, recognizeTextBy, lineEncodings)
        End Function

        ''' <summary>
        '''     Convert the datatable to a string based, comma-separated format
        ''' </summary>
        ''' <param name="dataTable">A datatable which shall be exported</param>
        ''' <param name="writeCsvColumnHeaders">Add a row with column headers on the top</param>
        ''' <param name="cultureFormatProvider">A culture format provider which declares list and decimal separators, etc.</param>
        ''' <param name="columnWidths">An array of integers specifying the fixed column widths</param>
        ''' <returns>The table as text with comma-separated structure</returns>
        Public Shared Function ConvertDataTableToTextAsStringBuilder(ByVal dataTable As System.Data.DataTable, ByVal writeCsvColumnHeaders As Boolean, ByVal cultureFormatProvider As System.Globalization.CultureInfo, ByVal columnWidths As Integer()) As System.Text.StringBuilder
            Return CompuMaster.Data.CsvTools.ConvertDataTableToCsv(dataTable, writeCsvColumnHeaders, cultureFormatProvider, columnWidths, WriteLineEncodings.Default)
        End Function

        ''' <summary>
        '''     Convert the datatable to a string based, comma-separated format
        ''' </summary>
        ''' <param name="dataTable">A datatable which shall be exported</param>
        ''' <param name="writeCsvColumnHeaders">Add a row with column headers on the top</param>
        ''' <param name="cultureFormatProvider">A culture format provider which declares list and decimal separators, etc.</param>
        ''' <param name="columnWidths">An array of integers specifying the fixed column widths</param>
        ''' <param name="lineEncodings">Encoding style for linebreaks</param>
        ''' <returns>The table as text with comma-separated structure</returns>
        Public Shared Function ConvertDataTableToTextAsStringBuilder(ByVal dataTable As System.Data.DataTable, ByVal writeCsvColumnHeaders As Boolean, lineEncodings As CompuMaster.Data.Csv.WriteLineEncodings, ByVal cultureFormatProvider As System.Globalization.CultureInfo, ByVal columnWidths As Integer()) As System.Text.StringBuilder
            Return CompuMaster.Data.CsvTools.ConvertDataTableToCsv(dataTable, writeCsvColumnHeaders, cultureFormatProvider, columnWidths, lineEncodings)
        End Function

        ''' <summary>
        '''     Convert the datatable to a string based, comma-separated format (for large tables, better use ConvertDataTableToTextAsStringBuilder to avoid System.OutOfMemoryExceptions)
        ''' </summary>
        ''' <param name="dataTable">A datatable which shall be exported</param>
        ''' <param name="writeCsvColumnHeaders">Add a row with column headers on the top</param>
        ''' <param name="cultureFormatProvider">A culture format provider which declares list and decimal separators, etc.</param>
        ''' <param name="columnSeparator">A column separator (culture default if empty value)</param>
        ''' <param name="recognizeTextBy">Recognize text by this character (default: quotation marks)</param>
        ''' <returns>A formatted text output</returns>
        <Obsolete("For large tables, better use ConvertDataTableToTextAsStringBuilder to avoid System.OutOfMemoryExceptions")> Public Shared Function ConvertDataTableToText(ByVal dataTable As System.Data.DataTable, ByVal writeCsvColumnHeaders As Boolean, ByVal cultureFormatProvider As System.Globalization.CultureInfo, Optional ByVal columnSeparator As String = Nothing, Optional ByVal recognizeTextBy As Char = """"c) As String
            If columnSeparator = Nothing Then
                columnSeparator = cultureFormatProvider.TextInfo.ListSeparator
            End If
            Return CompuMaster.Data.CsvTools.ConvertDataTableToCsv(dataTable, writeCsvColumnHeaders, cultureFormatProvider, columnSeparator, recognizeTextBy, WriteLineEncodings.Default).ToString
        End Function

        ''' <summary>
        '''     Convert the datatable to a string based, comma-separated format (for large tables, better use ConvertDataTableToTextAsStringBuilder to avoid System.OutOfMemoryExceptions)
        ''' </summary>
        ''' <param name="dataTable">A datatable which shall be exported</param>
        ''' <param name="writeCsvColumnHeaders">Add a row with column headers on the top</param>
        ''' <param name="cultureFormatProvider">A culture format provider which declares list and decimal separators, etc.</param>
        ''' <param name="columnWidths">An array of integers specifying the fixed column widths</param>
        ''' <returns>The table as text with comma-separated structure</returns>
        <Obsolete("For large tables, better use ConvertDataTableToTextAsStringBuilder to avoid System.OutOfMemoryExceptions")> Public Shared Function ConvertDataTableToText(ByVal dataTable As System.Data.DataTable, ByVal writeCsvColumnHeaders As Boolean, ByVal cultureFormatProvider As System.Globalization.CultureInfo, ByVal columnWidths As Integer()) As String
            Return CompuMaster.Data.CsvTools.ConvertDataTableToCsv(dataTable, writeCsvColumnHeaders, cultureFormatProvider, columnWidths, WriteLineEncodings.Default).ToString
        End Function

        ''' <summary>
        '''     Write to a CSV file
        ''' </summary>
        ''' <param name="path">The path of the file</param>
        ''' <param name="dataTable">A datatable which shall be exported</param>
        ''' <param name="writeCsvColumnHeaders">Indicates wether column headers are present</param>
        ''' <param name="fileEncoding">The text encoding of the file</param>
        ''' <param name="columnSeparator">Choose the required character for splitting the columns. Set to null (Nothing in VisualBasic) to enable fixed column widths mode</param>
        ''' <param name="recognizeTextBy">A character indicating the start and end of text strings</param>
        Public Shared Sub WriteDataTableToCsvFile(ByVal path As String, ByVal dataTable As System.Data.DataTable, ByVal writeCsvColumnHeaders As Boolean, Optional ByVal fileEncoding As String = "UTF-8", Optional ByVal columnSeparator As String = ","c, Optional ByVal recognizeTextBy As Char = """"c, Optional ByVal decimalSeparator As Char = "."c)
            CompuMaster.Data.CsvTools.WriteDataTableToCsvFile(path, dataTable, writeCsvColumnHeaders, fileEncoding, columnSeparator, recognizeTextBy, decimalSeparator, WriteLineEncodings.Default)
        End Sub

        ''' <summary>
        '''     Write to a CSV file
        ''' </summary>
        ''' <param name="path">The path of the file</param>
        ''' <param name="dataTable">A datatable which shall be exported</param>
        ''' <param name="writeCsvColumnHeaders">Indicates wether column headers are present</param>
        ''' <param name="fileEncoding">The text encoding of the file</param>
        ''' <param name="columnSeparator">Choose the required character for splitting the columns. Set to null (Nothing in VisualBasic) to enable fixed column widths mode</param>
        ''' <param name="recognizeTextBy">A character indicating the start and end of text strings</param>
        ''' <param name="lineEncodings">Encoding style for linebreaks</param>
        Public Shared Sub WriteDataTableToCsvFile(ByVal path As String, ByVal dataTable As System.Data.DataTable, ByVal writeCsvColumnHeaders As Boolean, lineEncodings As CompuMaster.Data.Csv.WriteLineEncodings, Optional ByVal fileEncoding As String = "UTF-8", Optional ByVal columnSeparator As String = ","c, Optional ByVal recognizeTextBy As Char = """"c, Optional ByVal decimalSeparator As Char = "."c)
            CompuMaster.Data.CsvTools.WriteDataTableToCsvFile(path, dataTable, writeCsvColumnHeaders, fileEncoding, columnSeparator, recognizeTextBy, decimalSeparator, lineEncodings)
        End Sub

        ''' <summary>
        '''     Create a CSV table (contains BOF signature for unicode encodings)
        ''' </summary>
        ''' <param name="dataTable">A datatable which shall be exported</param>
        ''' <param name="writeCsvColumnHeaders">Indicates wether column headers are present</param>
        ''' <param name="fileEncoding">The text encoding of the file</param>
        ''' <param name="columnSeparator">Choose the required character for splitting the columns. Set to null (Nothing in VisualBasic) to enable fixed column widths mode</param>
        ''' <param name="recognizeTextBy">A character indicating the start and end of text strings</param>
        ''' <param name="decimalSeparator">A character indicating the decimal separator in the text string</param>
        ''' <returns>A string containing the CSV table with integrated file encoding for writing with e.g. System.IO.File.WriteAllText()</returns>
        <Obsolete("Better use WriteDataTableToCsvFileStringWithTextEncoding() instead"), ComponentModel.EditorBrowsable(ComponentModel.EditorBrowsableState.Never)>
        Public Shared Function WriteDataTableToCsvString(ByVal dataTable As System.Data.DataTable, ByVal writeCsvColumnHeaders As Boolean, Optional ByVal fileEncoding As String = "UTF-8", Optional ByVal columnSeparator As Char = ","c, Optional ByVal recognizeTextBy As Char = """"c, Optional ByVal decimalSeparator As Char = "."c) As String
            Return WriteDataTableToCsvFileStringWithTextEncoding(dataTable, writeCsvColumnHeaders, WriteLineEncodings.Default, fileEncoding, columnSeparator, recognizeTextBy, decimalSeparator)
        End Function

        ''' <summary>
        '''     Create a CSV table (contains BOF signature for unicode encodings)
        ''' </summary>
        ''' <param name="dataTable">A datatable which shall be exported</param>
        ''' <param name="writeCsvColumnHeaders">Indicates wether column headers are present</param>
        ''' <param name="fileEncoding">The text encoding of the file</param>
        ''' <param name="columnSeparator">Choose the required character for splitting the columns. Set to null (Nothing in VisualBasic) to enable fixed column widths mode</param>
        ''' <param name="recognizeTextBy">A character indicating the start and end of text strings</param>
        ''' <param name="decimalSeparator">A character indicating the decimal separator in the text string</param>
        ''' <returns>A string containing the CSV table with integrated file encoding for writing with e.g. System.IO.File.WriteAllText()</returns>
        <Obsolete("Better use WriteDataTableToCsvTextString() instead since Strings don't support fileEncoding"), ComponentModel.EditorBrowsable(ComponentModel.EditorBrowsableState.Never)>
        Public Shared Function WriteDataTableToCsvFileStringWithTextEncoding(ByVal dataTable As System.Data.DataTable, ByVal writeCsvColumnHeaders As Boolean, Optional ByVal fileEncoding As String = "UTF-8", Optional ByVal columnSeparator As String = ",", Optional ByVal recognizeTextBy As Char = """"c, Optional ByVal decimalSeparator As Char = "."c) As String
            Dim MyStream As System.IO.MemoryStream = WriteDataTableToCsvMemoryStream(dataTable, writeCsvColumnHeaders, WriteLineEncodings.Default, System.Text.Encoding.Unicode.EncodingName, columnSeparator, recognizeTextBy, decimalSeparator)
            Return System.Text.Encoding.Unicode.GetString(MyStream.ToArray)
        End Function

        ''' <summary>
        '''     Create a CSV table (contains BOF signature for unicode encodings)
        ''' </summary>
        ''' <param name="dataTable">A datatable which shall be exported</param>
        ''' <param name="writeCsvColumnHeaders">Indicates wether column headers are present</param>
        ''' <param name="fileEncoding">The text encoding of the file</param>
        ''' <param name="columnSeparator">Choose the required character for splitting the columns. Set to null (Nothing in VisualBasic) to enable fixed column widths mode</param>
        ''' <param name="recognizeTextBy">A character indicating the start and end of text strings</param>
        ''' <param name="decimalSeparator">A character indicating the decimal separator in the text string</param>
        ''' <param name="lineEncodings">Encoding style for linebreaks</param>
        ''' <returns>A string containing the CSV table with integrated file encoding for writing with e.g. System.IO.File.WriteAllText()</returns>
        <Obsolete("Better use WriteDataTableToCsvTextString() instead since Strings don't support fileEncoding"), ComponentModel.EditorBrowsable(ComponentModel.EditorBrowsableState.Never)>
        Public Shared Function WriteDataTableToCsvFileStringWithTextEncoding(ByVal dataTable As System.Data.DataTable, ByVal writeCsvColumnHeaders As Boolean, lineEncodings As CompuMaster.Data.Csv.WriteLineEncodings, Optional ByVal fileEncoding As String = "UTF-8", Optional ByVal columnSeparator As String = ",", Optional ByVal recognizeTextBy As Char = """"c, Optional ByVal decimalSeparator As Char = "."c) As String
            Dim MyStream As System.IO.MemoryStream = WriteDataTableToCsvMemoryStream(dataTable, writeCsvColumnHeaders, lineEncodings, System.Text.Encoding.Unicode.EncodingName, columnSeparator, recognizeTextBy, decimalSeparator)
            Return System.Text.Encoding.Unicode.GetString(MyStream.ToArray)
        End Function

        ''' <summary>
        '''     Create a CSV table
        ''' </summary>
        ''' <param name="dataTable">A datatable which shall be exported</param>
        ''' <param name="writeCsvColumnHeaders">Indicates wether column headers are present</param>
        ''' <param name="columnSeparator">Choose the required character for splitting the columns. Set to null (Nothing in VisualBasic) to enable fixed column widths mode</param>
        ''' <param name="recognizeTextBy">A character indicating the start and end of text strings</param>
        ''' <param name="decimalSeparator">A character indicating the decimal separator in the text string</param>
        ''' <returns>A string containing the CSV table</returns>
        Public Shared Function WriteDataTableToCsvTextString(ByVal dataTable As System.Data.DataTable, ByVal writeCsvColumnHeaders As Boolean, Optional ByVal columnSeparator As String = ",", Optional ByVal recognizeTextBy As Char = """"c, Optional ByVal decimalSeparator As Char = "."c) As String
            Dim WrittenStream As System.IO.MemoryStream = WriteDataTableToCsvMemoryStream(dataTable, writeCsvColumnHeaders, WriteLineEncodings.Default, System.Text.Encoding.Unicode.EncodingName, columnSeparator, recognizeTextBy, decimalSeparator)
            Dim ReaderStream As New System.IO.MemoryStream(WrittenStream.ToArray)
            WrittenStream.Dispose()
            Dim SR As New System.IO.StreamReader(ReaderStream, System.Text.Encoding.Unicode)
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
        ''' <param name="dataTable">A datatable which shall be exported</param>
        ''' <param name="writeCsvColumnHeaders">Indicates wether column headers are present</param>
        ''' <param name="columnSeparator">Choose the required character for splitting the columns. Set to null (Nothing in VisualBasic) to enable fixed column widths mode</param>
        ''' <param name="recognizeTextBy">A character indicating the start and end of text strings</param>
        ''' <param name="decimalSeparator">A character indicating the decimal separator in the text string</param>
        ''' <param name="lineEncodings">Encoding style for linebreaks</param>
        ''' <returns>A string containing the CSV table</returns>
        Public Shared Function WriteDataTableToCsvTextString(ByVal dataTable As System.Data.DataTable, ByVal writeCsvColumnHeaders As Boolean, lineEncodings As CompuMaster.Data.Csv.WriteLineEncodings, Optional ByVal columnSeparator As String = ",", Optional ByVal recognizeTextBy As Char = """"c, Optional ByVal decimalSeparator As Char = "."c) As String
            Dim WrittenStream As System.IO.MemoryStream = WriteDataTableToCsvMemoryStream(dataTable, writeCsvColumnHeaders, lineEncodings, System.Text.Encoding.Unicode.EncodingName, columnSeparator, recognizeTextBy, decimalSeparator)
            Dim ReaderStream As New System.IO.MemoryStream(WrittenStream.ToArray)
            WrittenStream.Dispose()
            Dim SR As New System.IO.StreamReader(ReaderStream, System.Text.Encoding.Unicode)
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
        ''' <param name="dataTable">A datatable which shall be exported</param>
        ''' <param name="writeCsvColumnHeaders">Indicates wether column headers are present</param>
        ''' <param name="fileEncoding">The text encoding of the file</param>
        ''' <param name="columnSeparator">Choose the required character for splitting the columns. Set to null (Nothing in VisualBasic) to enable fixed column widths mode</param>
        ''' <param name="recognizeTextBy">A character indicating the start and end of text strings</param>
        ''' <param name="decimalSeparator">A character indicating the decimal separator in the text string</param>
        ''' <returns>A string containing the CSV table</returns>
        Public Shared Function WriteDataTableToCsvBytes(ByVal dataTable As System.Data.DataTable, ByVal writeCsvColumnHeaders As Boolean, Optional ByVal fileEncoding As String = "UTF-8", Optional ByVal columnSeparator As Char = ","c, Optional ByVal recognizeTextBy As Char = """"c, Optional ByVal decimalSeparator As Char = "."c) As Byte()
            Return CompuMaster.Data.CsvTools.WriteDataTableToCsvBytes(dataTable, writeCsvColumnHeaders, fileEncoding, columnSeparator, recognizeTextBy, decimalSeparator, WriteLineEncodings.Default)
        End Function

        ''' <summary>
        '''     Create a CSV table
        ''' </summary>
        ''' <param name="dataTable">A datatable which shall be exported</param>
        ''' <param name="writeCsvColumnHeaders">Indicates wether column headers are present</param>
        ''' <param name="fileEncoding">The text encoding of the file</param>
        ''' <param name="columnSeparator">Choose the required character for splitting the columns. Set to null (Nothing in VisualBasic) to enable fixed column widths mode</param>
        ''' <param name="recognizeTextBy">A character indicating the start and end of text strings</param>
        ''' <param name="decimalSeparator">A character indicating the decimal separator in the text string</param>
        ''' <param name="lineEncodings">Encoding style for linebreaks</param>
        ''' <returns>A string containing the CSV table</returns>
        Public Shared Function WriteDataTableToCsvBytes(ByVal dataTable As System.Data.DataTable, ByVal writeCsvColumnHeaders As Boolean, lineEncodings As CompuMaster.Data.Csv.WriteLineEncodings, Optional ByVal fileEncoding As String = "UTF-8", Optional ByVal columnSeparator As Char = ","c, Optional ByVal recognizeTextBy As Char = """"c, Optional ByVal decimalSeparator As Char = "."c) As Byte()
            Return CompuMaster.Data.CsvTools.WriteDataTableToCsvBytes(dataTable, writeCsvColumnHeaders, fileEncoding, columnSeparator, recognizeTextBy, decimalSeparator, lineEncodings)
        End Function

        ''' <summary>
        '''     Create a CSV table
        ''' </summary>
        ''' <param name="dataTable">A datatable which shall be exported</param>
        ''' <param name="writeCsvColumnHeaders">Indicates wether column headers are present</param>
        ''' <param name="fileEncoding">The text encoding of the file</param>
        ''' <param name="cultureFormatProvider">A globalization information object for the conversion of all data to strings</param>
        ''' <param name="columnSeparator">Choose the required character for splitting the columns. Set to null (Nothing in VisualBasic) to enable fixed column widths mode</param>
        ''' <param name="recognizeTextBy">A character indicating the start and end of text strings</param>
        ''' <returns>A string containing the CSV table</returns>
        Public Shared Function WriteDataTableToCsvBytes(ByVal dataTable As System.Data.DataTable, ByVal writeCsvColumnHeaders As Boolean, ByVal fileEncoding As System.Text.Encoding, ByVal cultureFormatProvider As System.Globalization.CultureInfo, Optional ByVal columnSeparator As Char = Nothing, Optional ByVal recognizeTextBy As Char = """"c) As Byte()
            Dim charColumnSeparator As Char
            If columnSeparator = Nothing Then
                charColumnSeparator = CType(cultureFormatProvider.TextInfo.ListSeparator, Char)
            Else
                charColumnSeparator = columnSeparator
            End If
            Return CompuMaster.Data.CsvTools.WriteDataTableToCsvBytes(dataTable, writeCsvColumnHeaders, fileEncoding, cultureFormatProvider, charColumnSeparator, recognizeTextBy, WriteLineEncodings.Default)
        End Function

        ''' <summary>
        '''     Create a CSV table
        ''' </summary>
        ''' <param name="dataTable">A datatable which shall be exported</param>
        ''' <param name="writeCsvColumnHeaders">Indicates wether column headers are present</param>
        ''' <param name="fileEncoding">The text encoding of the file</param>
        ''' <param name="cultureFormatProvider">A globalization information object for the conversion of all data to strings</param>
        ''' <param name="columnSeparator">Choose the required character for splitting the columns. Set to null (Nothing in VisualBasic) to enable fixed column widths mode</param>
        ''' <param name="recognizeTextBy">A character indicating the start and end of text strings</param>
        ''' <param name="lineEncodings">Encoding style for linebreaks</param>
        ''' <returns>A string containing the CSV table</returns>
        Public Shared Function WriteDataTableToCsvBytes(ByVal dataTable As System.Data.DataTable, ByVal writeCsvColumnHeaders As Boolean, lineEncodings As CompuMaster.Data.Csv.WriteLineEncodings, ByVal fileEncoding As System.Text.Encoding, ByVal cultureFormatProvider As System.Globalization.CultureInfo, Optional ByVal columnSeparator As Char = Nothing, Optional ByVal recognizeTextBy As Char = """"c) As Byte()
            Dim charColumnSeparator As Char
            If columnSeparator = Nothing Then
                charColumnSeparator = CType(cultureFormatProvider.TextInfo.ListSeparator, Char)
            Else
                charColumnSeparator = columnSeparator
            End If
            Return CompuMaster.Data.CsvTools.WriteDataTableToCsvBytes(dataTable, writeCsvColumnHeaders, fileEncoding, cultureFormatProvider, charColumnSeparator, recognizeTextBy, lineEncodings)
        End Function

        ''' <summary>
        '''     Create a CSV table
        ''' </summary>
        ''' <param name="dataTable">A datatable which shall be exported</param>
        ''' <param name="writeCsvColumnHeaders">Indicates wether column headers are present</param>
        ''' <param name="fileEncoding">The text encoding of the file</param>
        ''' <param name="columnSeparator">Choose the required character for splitting the columns. Set to null (Nothing in VisualBasic) to enable fixed column widths mode</param>
        ''' <param name="recognizeTextBy">A character indicating the start and end of text strings</param>
        ''' <param name="decimalSeparator">A character indicating the decimal separator in the text string</param>
        ''' <returns>A memory stream containing all texts as bytes in Unicode format</returns>
        Public Shared Function WriteDataTableToCsvMemoryStream(ByVal dataTable As System.Data.DataTable, ByVal writeCsvColumnHeaders As Boolean, Optional ByVal fileEncoding As String = "UTF-8", Optional ByVal columnSeparator As String = ","c, Optional ByVal recognizeTextBy As Char = """"c, Optional ByVal decimalSeparator As Char = "."c) As System.IO.MemoryStream
            Return CompuMaster.Data.CsvTools.WriteDataTableToCsvMemoryStream(dataTable, writeCsvColumnHeaders, fileEncoding, columnSeparator, recognizeTextBy, decimalSeparator, WriteLineEncodings.Default)
        End Function

        ''' <summary>
        '''     Create a CSV table
        ''' </summary>
        ''' <param name="dataTable">A datatable which shall be exported</param>
        ''' <param name="writeCsvColumnHeaders">Indicates wether column headers are present</param>
        ''' <param name="fileEncoding">The text encoding of the file</param>
        ''' <param name="columnSeparator">Choose the required character for splitting the columns. Set to null (Nothing in VisualBasic) to enable fixed column widths mode</param>
        ''' <param name="recognizeTextBy">A character indicating the start and end of text strings</param>
        ''' <param name="decimalSeparator">A character indicating the decimal separator in the text string</param>
        ''' <param name="lineEncodings">Encoding style for linebreaks</param>
        ''' <returns>A memory stream containing all texts as bytes in Unicode format</returns>
        Public Shared Function WriteDataTableToCsvMemoryStream(ByVal dataTable As System.Data.DataTable, ByVal writeCsvColumnHeaders As Boolean, lineEncodings As CompuMaster.Data.Csv.WriteLineEncodings, Optional ByVal fileEncoding As String = "UTF-8", Optional ByVal columnSeparator As String = ","c, Optional ByVal recognizeTextBy As Char = """"c, Optional ByVal decimalSeparator As Char = "."c) As System.IO.MemoryStream
            Return CompuMaster.Data.CsvTools.WriteDataTableToCsvMemoryStream(dataTable, writeCsvColumnHeaders, fileEncoding, columnSeparator, recognizeTextBy, decimalSeparator, lineEncodings)
        End Function

        ''' <summary>
        '''     Write to a CSV file
        ''' </summary>
        ''' <param name="path">The path of the file</param>
        ''' <param name="dataView">A dataview object with the desired rows</param>
        Public Shared Sub WriteDataViewToCsvFile(ByVal path As String, ByVal dataview As System.Data.DataView)
            CompuMaster.Data.CsvTools.WriteDataViewToCsvFile(path, dataview, WriteLineEncodings.Default)
        End Sub

        ''' <summary>
        '''     Write to a CSV file
        ''' </summary>
        ''' <param name="path">The path of the file</param>
        ''' <param name="dataView">A dataview object with the desired rows</param>
        ''' <param name="lineEncodings">Encoding style for linebreaks</param>
        Public Shared Sub WriteDataViewToCsvFile(ByVal path As String, ByVal dataview As System.Data.DataView, lineEncodings As CompuMaster.Data.Csv.WriteLineEncodings)
            CompuMaster.Data.CsvTools.WriteDataViewToCsvFile(path, dataview, lineEncodings)
        End Sub

        ''' <summary>
        '''     Write to a CSV file
        ''' </summary>
        ''' <param name="path">The path of the file</param>
        ''' <param name="dataView">A dataview object with the desired rows</param>
        ''' <param name="writeCsvColumnHeaders">Add a row with column headers on the top</param>
        ''' <param name="cultureFormatProvider">A culture for all conversions</param>
        ''' <param name="fileEncoding">The text encoding of the file</param>
        Public Shared Sub WriteDataViewToCsvFile(ByVal path As String, ByVal dataView As System.Data.DataView, ByVal writeCsvColumnHeaders As Boolean, ByVal cultureFormatProvider As System.Globalization.CultureInfo, Optional ByVal fileEncoding As String = "UTF-8", Optional ByVal columnSeparator As String = Nothing, Optional ByVal recognizeTextBy As Char = """"c)
            If columnSeparator = Nothing Then
                columnSeparator = cultureFormatProvider.TextInfo.ListSeparator
            End If
            CompuMaster.Data.CsvTools.WriteDataViewToCsvFile(path, dataView, writeCsvColumnHeaders, cultureFormatProvider, fileEncoding, columnSeparator, recognizeTextBy, WriteLineEncodings.Default)
        End Sub

        ''' <summary>
        '''     Write to a CSV file
        ''' </summary>
        ''' <param name="path">The path of the file</param>
        ''' <param name="dataView">A dataview object with the desired rows</param>
        ''' <param name="writeCsvColumnHeaders">Add a row with column headers on the top</param>
        ''' <param name="cultureFormatProvider">A culture for all conversions</param>
        ''' <param name="fileEncoding">The text encoding of the file</param>
        ''' <param name="lineEncodings">Encoding style for linebreaks</param>
        Public Shared Sub WriteDataViewToCsvFile(ByVal path As String, ByVal dataView As System.Data.DataView, ByVal writeCsvColumnHeaders As Boolean, lineEncodings As CompuMaster.Data.Csv.WriteLineEncodings, ByVal cultureFormatProvider As System.Globalization.CultureInfo, Optional ByVal fileEncoding As String = "UTF-8", Optional ByVal columnSeparator As String = Nothing, Optional ByVal recognizeTextBy As Char = """"c)
            If columnSeparator = Nothing Then
                columnSeparator = cultureFormatProvider.TextInfo.ListSeparator
            End If
            CompuMaster.Data.CsvTools.WriteDataViewToCsvFile(path, dataView, writeCsvColumnHeaders, cultureFormatProvider, fileEncoding, columnSeparator, recognizeTextBy, lineEncodings)
        End Sub

        ''' <summary>
        '''     Write to a CSV file
        ''' </summary>
        ''' <param name="path">The path of the file</param>
        ''' <param name="dataView">A dataview object with the desired rows</param>
        ''' <param name="writeCsvColumnHeaders">Add a row with column headers on the top</param>
        ''' <param name="fileEncoding">The text encoding of the file</param>
        ''' <param name="columnSeparator">Choose the required character for splitting the columns. Set to null (Nothing in VisualBasic) to enable fixed column widths mode</param>
        ''' <param name="recognizeTextBy">A character indicating the start and end of text strings</param>
        ''' <param name="decimalSeparator">A character indicating the decimal separator in the text string</param>
        Public Shared Sub WriteDataViewToCsvFile(ByVal path As String, ByVal dataView As System.Data.DataView, ByVal writeCsvColumnHeaders As Boolean, Optional ByVal fileEncoding As String = "UTF-8", Optional ByVal columnSeparator As String = ",", Optional ByVal recognizeTextBy As Char = """"c, Optional ByVal decimalSeparator As Char = "."c)
            CompuMaster.Data.CsvTools.WriteDataViewToCsvFile(path, dataView, writeCsvColumnHeaders, fileEncoding, columnSeparator, recognizeTextBy, decimalSeparator, WriteLineEncodings.Default)
        End Sub

        ''' <summary>
        '''     Write to a CSV file
        ''' </summary>
        ''' <param name="path">The path of the file</param>
        ''' <param name="dataView">A dataview object with the desired rows</param>
        ''' <param name="writeCsvColumnHeaders">Add a row with column headers on the top</param>
        ''' <param name="fileEncoding">The text encoding of the file</param>
        ''' <param name="columnSeparator">Choose the required character for splitting the columns. Set to null (Nothing in VisualBasic) to enable fixed column widths mode</param>
        ''' <param name="recognizeTextBy">A character indicating the start and end of text strings</param>
        ''' <param name="decimalSeparator">A character indicating the decimal separator in the text string</param>
        ''' <param name="lineEncodings">Encoding style for linebreaks</param>
        Public Shared Sub WriteDataViewToCsvFile(ByVal path As String, ByVal dataView As System.Data.DataView, ByVal writeCsvColumnHeaders As Boolean, lineEncodings As CompuMaster.Data.Csv.WriteLineEncodings, Optional ByVal fileEncoding As String = "UTF-8", Optional ByVal columnSeparator As String = ",", Optional ByVal recognizeTextBy As Char = """"c, Optional ByVal decimalSeparator As Char = "."c)
            CompuMaster.Data.CsvTools.WriteDataViewToCsvFile(path, dataView, writeCsvColumnHeaders, fileEncoding, columnSeparator, recognizeTextBy, decimalSeparator, lineEncodings)
        End Sub
    End Class

End Namespace
