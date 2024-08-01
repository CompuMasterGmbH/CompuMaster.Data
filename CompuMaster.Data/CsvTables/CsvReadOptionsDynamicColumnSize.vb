Option Explicit On
Option Strict On

Imports System.IO
Imports System.Data
Imports CompuMaster.Data.Strings
Imports System.Text

Namespace CompuMaster.Data.CsvTables

    ''' <summary>
    ''' Configuration options for reading CSV data from a file or stream
    ''' </summary>
    Public Class CsvReadOptionsDynamicColumnSize
        Inherits CsvReadBaseOptions

        ''' <summary>
        ''' Create a new instance of CsvReadOptionsDynamicColumnSize
        ''' </summary>
        Public Sub New()
        End Sub

        ''' <summary>
        ''' Create a new instance of CsvReadOptionsDynamicColumnSize
        ''' </summary>
        ''' <param name="csvContainsColumnHeaders">Indicates wether column headers are present</param>
        ''' <param name="columnSeparator">Choose the required character for splitting the columns. Set to null (Nothing in VisualBasic) to enable fixed column widths mode</param>
        ''' <param name="recognizeTextBy">A character indicating the start and end of text strings</param>
        ''' <param name="convertEmptyStringsToDBNull">Convert values with empty strings automatically to DbNull</param>
        Public Sub New(ByVal csvContainsColumnHeaders As Boolean, Optional ByVal columnSeparator As Char = ","c, Optional ByVal recognizeTextBy As Char = """"c, Optional ByVal convertEmptyStringsToDBNull As Boolean = False)
            Me.New(csvContainsColumnHeaders, 0, CompuMaster.Data.Csv.ReadLineEncodings.Default, CompuMaster.Data.Csv.ReadLineEncodingAutoConversion.AutoConvertLineBreakToSystemEnvironmentNewLine, columnSeparator, recognizeTextBy, convertEmptyStringsToDBNull)
        End Sub

        ''' <summary>
        ''' Create a new instance of CsvReadOptionsDynamicColumnSize
        ''' </summary>
        ''' <param name="csvContainsColumnHeaders">Indicates wether column headers are present</param>
        ''' <param name="startAtLineIndex">Start reading of table data at specified line index (most often 0 for very first line)</param>
        ''' <param name="columnSeparator">Choose the required character for splitting the columns. Set to null (Nothing in VisualBasic) to enable fixed column widths mode</param>
        ''' <param name="recognizeTextBy">A character indicating the start and end of text strings</param>
        ''' <param name="convertEmptyStringsToDBNull">Convert values with empty strings automatically to DbNull</param>
        Public Sub New(ByVal csvContainsColumnHeaders As Boolean, startAtLineIndex As Integer, Optional ByVal columnSeparator As Char = ","c, Optional ByVal recognizeTextBy As Char = """"c, Optional ByVal convertEmptyStringsToDBNull As Boolean = False)
            Me.New(csvContainsColumnHeaders, startAtLineIndex, CompuMaster.Data.Csv.ReadLineEncodings.Default, CompuMaster.Data.Csv.ReadLineEncodingAutoConversion.AutoConvertLineBreakToSystemEnvironmentNewLine, columnSeparator, recognizeTextBy, convertEmptyStringsToDBNull)
        End Sub

        ''' <summary>
        ''' Create a new instance of CsvReadOptionsDynamicColumnSize
        ''' </summary>
        ''' <param name="csvContainsColumnHeaders">Indicates wether column headers are present</param>
        ''' <param name="startAtLineIndex">Start reading of table data at specified line index (most often 0 for very first line)</param>
        ''' <param name="lineEncodings">Encoding style for linebreaks</param>
        ''' <param name="lineEncodingAutoConversions">Change linebreak encodings on reading</param>
        ''' <param name="columnSeparator">Choose the required character for splitting the columns. Set to null (Nothing in VisualBasic) to enable fixed column widths mode</param>
        ''' <param name="recognizeTextBy">A character indicating the start and end of text strings</param>
        ''' <param name="convertEmptyStringsToDBNull">Convert values with empty strings automatically to DbNull</param>
        Public Sub New(ByVal csvContainsColumnHeaders As Boolean, startAtLineIndex As Integer, lineEncodings As CompuMaster.Data.Csv.ReadLineEncodings, lineEncodingAutoConversions As CompuMaster.Data.Csv.ReadLineEncodingAutoConversion, Optional ByVal columnSeparator As Char = ","c, Optional ByVal recognizeTextBy As Char = """"c, Optional ByVal convertEmptyStringsToDBNull As Boolean = False)
            Me._CultureFormatProvider = System.Globalization.CultureInfo.CurrentCulture
            Me.CsvContainsColumnHeaders = csvContainsColumnHeaders
            Me.StartAtLineIndex = startAtLineIndex
            Me.LineEncodings = lineEncodings
            Me.LineEncodingAutoConversions = lineEncodingAutoConversions
            Me.ColumnSeparator = columnSeparator
            Me.RecognizeTextBy = recognizeTextBy
            Me.ConvertEmptyStringsToDBNull = convertEmptyStringsToDBNull
        End Sub

        ''' <summary>
        ''' Create a new instance of CsvReadOptionsDynamicColumnSize
        ''' </summary>
        ''' <param name="csvContainsColumnHeaders">Indicates wether column headers are present</param>
        ''' <param name="cultureFormatProvider">Choose the required character for splitting the columns. Set to null (Nothing in VisualBasic) to enable fixed column widths mode</param>
        ''' <param name="recognizeTextBy">A character indicating the start and end of text strings</param>
        ''' <param name="convertEmptyStringsToDBNull">Convert values with empty strings automatically to DbNull</param>
        Public Sub New(ByVal csvContainsColumnHeaders As Boolean, ByVal cultureFormatProvider As System.Globalization.CultureInfo, Optional ByVal recognizeTextBy As Char = """"c, Optional ByVal convertEmptyStringsToDBNull As Boolean = False)
            Me.New(csvContainsColumnHeaders, 0, CompuMaster.Data.Csv.ReadLineEncodings.Default, CompuMaster.Data.Csv.ReadLineEncodingAutoConversion.AutoConvertLineBreakToSystemEnvironmentNewLine, cultureFormatProvider, recognizeTextBy, convertEmptyStringsToDBNull)
        End Sub

        ''' <summary>
        ''' Create a new instance of CsvReadOptionsDynamicColumnSize
        ''' </summary>
        ''' <param name="csvContainsColumnHeaders">Indicates wether column headers are present</param>
        ''' <param name="startAtLineIndex">Start reading of table data at specified line index (most often 0 for very first line)</param>
        ''' <param name="cultureFormatProvider">Choose the required character for splitting the columns. Set to null (Nothing in VisualBasic) to enable fixed column widths mode</param>
        ''' <param name="recognizeTextBy">A character indicating the start and end of text strings</param>
        ''' <param name="convertEmptyStringsToDBNull">Convert values with empty strings automatically to DbNull</param>
        Public Sub New(ByVal csvContainsColumnHeaders As Boolean, startAtLineIndex As Integer, ByVal cultureFormatProvider As System.Globalization.CultureInfo, Optional ByVal recognizeTextBy As Char = """"c, Optional ByVal convertEmptyStringsToDBNull As Boolean = False)
            Me.New(csvContainsColumnHeaders, startAtLineIndex, CompuMaster.Data.Csv.ReadLineEncodings.Default, CompuMaster.Data.Csv.ReadLineEncodingAutoConversion.AutoConvertLineBreakToSystemEnvironmentNewLine, cultureFormatProvider, recognizeTextBy, convertEmptyStringsToDBNull)
        End Sub

        ''' <summary>
        ''' Create a new instance of CsvReadOptionsDynamicColumnSize
        ''' </summary>
        ''' <param name="csvContainsColumnHeaders">Indicates wether column headers are present</param>
        ''' <param name="startAtLineIndex">Start reading of table data at specified line index (most often 0 for very first line)</param>
        ''' <param name="lineEncodings">Encoding style for linebreaks</param>
        ''' <param name="lineEncodingAutoConversions">Change linebreak encodings on reading</param>
        ''' <param name="cultureFormatProvider">Choose the required character for splitting the columns. Set to null (Nothing in VisualBasic) to enable fixed column widths mode</param>
        ''' <param name="recognizeTextBy">A character indicating the start and end of text strings</param>
        ''' <param name="convertEmptyStringsToDBNull">Convert values with empty strings automatically to DbNull</param>
        Public Sub New(ByVal csvContainsColumnHeaders As Boolean, startAtLineIndex As Integer, lineEncodings As CompuMaster.Data.Csv.ReadLineEncodings, lineEncodingAutoConversions As CompuMaster.Data.Csv.ReadLineEncodingAutoConversion, ByVal cultureFormatProvider As System.Globalization.CultureInfo, Optional ByVal recognizeTextBy As Char = """"c, Optional ByVal convertEmptyStringsToDBNull As Boolean = False)
            Me._CultureFormatProvider = cultureFormatProvider
            Me.CsvContainsColumnHeaders = csvContainsColumnHeaders
            Me.StartAtLineIndex = startAtLineIndex
            Me.LineEncodings = lineEncodings
            Me.LineEncodingAutoConversions = lineEncodingAutoConversions
            Me.RecognizeTextBy = recognizeTextBy
            Me.ConvertEmptyStringsToDBNull = convertEmptyStringsToDBNull
        End Sub

        Private _CultureFormatProvider As System.Globalization.CultureInfo
        ''' <summary>
        ''' Culture settings for expected CSV structure (especially for list separators aka column separators)
        ''' </summary>
        ''' <returns></returns>
        Public Property CultureFormatProvider As System.Globalization.CultureInfo
            Get
                If Me._CultureFormatProvider Is Nothing Then
                    Return System.Globalization.CultureInfo.InvariantCulture
                Else
                    Return Me._CultureFormatProvider
                End If
            End Get
            Set(value As System.Globalization.CultureInfo)
                _CultureFormatProvider = value
                '_ColumnSeparator = value.TextInfo.ListSeparator.Chars(0)
            End Set
        End Property

        Public Property _ColumnSeparator As Char '= ","c
        ''' <summary>
        ''' Column separator character, default is comma (,)
        ''' </summary>
        ''' <returns></returns>
        Public Property ColumnSeparator As Char
            Get
                If _ColumnSeparator = Nothing OrElse _ColumnSeparator = vbNullChar Then
                    'ATTENTION: list separator is a string, but columnSeparator is implemented as char! Might be a bug in some special cultures
                    If Me.CultureFormatProvider.TextInfo.ListSeparator.Length > 1 Then
                        Throw New NotSupportedException("No column separator has been defined and the current culture declares a list separator with more than 1 character. Column separators with more than 1 characters are currenlty not supported.")
                    End If
                    Return Me.CultureFormatProvider.TextInfo.ListSeparator.Chars(0)
                Else
                    Return _ColumnSeparator
                End If
            End Get
            Set(value As Char)
                _ColumnSeparator = value
            End Set
        End Property

        ''' <summary>
        ''' Text recognition character, default is double quote (")
        ''' </summary>
        ''' <returns></returns>
        Public Property RecognizeTextBy As Char = """"c

        ''' <summary>
        ''' Recognize multiple column separator characters as one (feature inspired by MS Excel import wizard)
        ''' </summary>
        ''' <returns></returns>
        Public Property RecognizeMultipleColumnSeparatorCharsAsOne As Boolean

        <Obsolete("Use RecognizeMultipleColumnSeparatorCharsAsOne instead", True)>
        Private Property RecognizeDoubledColumnSeparatorCharAsOne As Boolean

        ''' <summary>
        ''' Decode backslash escape sequences like \" or \, or \\ in cell content
        ''' </summary>
        ''' <returns></returns>
        Public Property RecognizeBackslashEscapes As Boolean

    End Class

End Namespace