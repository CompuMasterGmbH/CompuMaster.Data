Imports System.Diagnostics.Eventing
Imports System.IO
Imports System.Text

Namespace CompuMaster.Data.CsvTables

    Public Class CsvReadOptionsWithFixedColumnWidths

        Public Sub New()

        End Sub

        Public Sub New(ByVal csvContainsColumnHeaders As Boolean, startAtLineIndex As Integer, lineEncodings As CompuMaster.Data.Csv.ReadLineEncodings, lineEncodingAutoConversions As CompuMaster.Data.Csv.ReadLineEncodingAutoConversion, Optional ByVal fileEncoding As String = "UTF-8", Optional ByVal columnSeparator As Char = ","c, Optional ByVal recognizeTextBy As Char = """"c, Optional ByVal recognizeMultipleColumnSeparatorCharsAsOne As Boolean = False, Optional ByVal convertEmptyStringsToDBNull As Boolean = False)
            Me.CsvContainsColumnHeaders = csvContainsColumnHeaders
            Me.StartAtLineIndex = startAtLineIndex
            Me.LineEncodings = lineEncodings
            Me.LineEncodingAutoConversions = lineEncodingAutoConversions

            If fileEncoding = "" Then
                Me.FileEncoding = System.Text.Encoding.Default
            Else
                Me.FileEncoding = System.Text.Encoding.GetEncoding(fileEncoding)
            End If

            Me.ColumnSeparator = columnSeparator
            Me.RecognizeMultipleColumnSeparatorCharsAsOne = recognizeMultipleColumnSeparatorCharsAsOne
            Me.RecognizeTextBy = recognizeTextBy
            Me.ConvertEmptyStringsToDBNull = convertEmptyStringsToDBNull

        End Sub

        Public Sub New(ByVal csvContainsColumnHeaders As Boolean, startAtLineIndex As Integer, lineEncodings As CompuMaster.Data.Csv.ReadLineEncodings, lineEncodingAutoConversions As CompuMaster.Data.Csv.ReadLineEncodingAutoConversion, ByVal fileEncoding As System.Text.Encoding, ByVal cultureFormatProvider As System.Globalization.CultureInfo, Optional ByVal recognizeTextBy As Char = """"c, Optional ByVal recognizeMultipleColumnSeparatorCharsAsOne As Boolean = False, Optional ByVal convertEmptyStringsToDBNull As Boolean = False)
            Me.CsvContainsColumnHeaders = csvContainsColumnHeaders
            Me.StartAtLineIndex = startAtLineIndex
            Me.LineEncodings = lineEncodings
            Me.LineEncodingAutoConversions = lineEncodingAutoConversions
            Me.FileEncoding = fileEncoding
            Me.CultureFormatProvider = cultureFormatProvider
            Me.RecognizeTextBy = recognizeTextBy
            Me.RecognizeMultipleColumnSeparatorCharsAsOne = recognizeMultipleColumnSeparatorCharsAsOne
            Me.ConvertEmptyStringsToDBNull = convertEmptyStringsToDBNull
        End Sub

        Public Property CsvContainsColumnHeaders As Boolean

        Public Property StartAtLineIndex As Integer

        Public Property LineEncodings As CompuMaster.Data.Csv.ReadLineEncodings

        Public Property LineEncodingAutoConversions As CompuMaster.Data.Csv.ReadLineEncodingAutoConversion

        Public Property FileEncoding As System.Text.Encoding

        Public Property CultureFormatProvider As System.Globalization.CultureInfo

        Public Property ColumnSeparator As Char = ","c

        Public Property RecognizeTextBy As Char = """"c

        Public Property RecognizeMultipleColumnSeparatorCharsAsOne As Boolean

        <Obsolete("Use RecognizeMultipleColumnSeparatorCharsAsOne instead", True)>
        Private Property RecognizeDoubledColumnSeparatorCharAsOne As Boolean

        Public Property ConvertEmptyStringsToDBNull As Boolean

        Public Property RecognizeBackslashEscapes As Boolean

    End Class

End Namespace