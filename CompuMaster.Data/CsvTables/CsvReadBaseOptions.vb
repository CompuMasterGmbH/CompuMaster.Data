Imports System.Diagnostics.Eventing
Imports System.IO
Imports System.Text

Namespace CompuMaster.Data.CsvTables

    Public MustInherit Class CsvReadBaseOptions

        Public Sub New()
        End Sub

        Private _CsvContainsColumnHeaders As Boolean
        ''' <summary>
        ''' CSV file contains column headers in first line
        ''' </summary>
        ''' <returns></returns>
        Public Property CsvContainsColumnHeaders As Boolean
            Get
                If _CsvContainsColumnHeadersLineCount > 1 Then
                    Throw New NotSupportedException("Use CsvContainsColumnHeadersLineCount instead")
                End If
                Return _CsvContainsColumnHeaders
            End Get
            Set(value As Boolean)
                _CsvContainsColumnHeaders = value
                _CsvContainsColumnHeadersLineCount = 1
            End Set
        End Property


        Private _CsvContainsColumnHeadersLineCount As Integer
        ''' <summary>
        ''' CSV file contains 0, 1 or more column headers lines on top of table
        ''' </summary>
        ''' <returns></returns>
        <Obsolete("TODO: Missing feature implementation")>
        Public Property CsvContainsColumnHeadersLineCount As Integer
            Get
                Return _CsvContainsColumnHeadersLineCount
            End Get
            Set(value As Integer)
                _CsvContainsColumnHeadersLineCount = value
                _CsvContainsColumnHeaders = (value <> 0)
            End Set
        End Property

        'TODO: maybe "line" is incorrect, should be "row"?
        ''' <summary>
        ''' Start reading at line index (0-based)
        ''' </summary>
        ''' <returns></returns>
        Public Property StartAtLineIndex As Integer

        ''' <summary>
        ''' Line encodings style in CSV table
        ''' </summary>
        ''' <returns></returns>
        Public Property LineEncodings As CompuMaster.Data.Csv.ReadLineEncodings

        ''' <summary>
        ''' Line encodings style in CSV table after auto-detection
        ''' </summary>
        ''' <returns></returns>
        Friend Function LineEncodingsAfterAutoDetection(checkFor_RowBreakCrLfOrCrOrLf_CellLineBreakCrLfOrCrOrLf_AND_ContainsHeadersLineCount_Equals_Zero As Boolean) As CompuMaster.Data.Csv.ReadLineEncodings
            'Provide buffered result for performance reasons (avoiding multiple checks of operation system for the same result)
            Static BufferedResult As CompuMaster.Data.Csv.ReadLineEncodings?
            Static BufferedResultBasedOnInput As CompuMaster.Data.Csv.ReadLineEncodings

            Dim Result As CompuMaster.Data.Csv.ReadLineEncodings
            If BufferedResult.HasValue = False OrElse BufferedResult.Value <> Me.LineEncodings Then
                '(Re-)detect line endings of current platform
                If LineEncodings = Csv.ReadLineEncodings.Auto Then
                    Select Case System.Environment.NewLine
                        Case ControlChars.CrLf
                            'Windows platforms
                            Result = Csv.ReadLineEncodings.RowBreakCrLf_CellLineBreakLf
                        Case ControlChars.Cr
                            'Mac platforms
                            Result = Csv.ReadLineEncodings.RowBreakCr_CellLineBreakLf
                        Case ControlChars.Lf
                            'Linux platforms
                            Result = Csv.ReadLineEncodings.RowBreakLf_CellLineBreakCr
                        Case Else
                            Throw New NotImplementedException
                    End Select
                End If
                BufferedResultBasedOnInput = Me.LineEncodings
                BufferedResult = Result
            Else
                'Re-use buffered result
                Result = BufferedResult.Value
            End If

            'Check for invalid settings
            If checkFor_RowBreakCrLfOrCrOrLf_CellLineBreakCrLfOrCrOrLf_AND_ContainsHeadersLineCount_Equals_Zero AndAlso Result = Csv.ReadLineEncodings.RowBreakCrLfOrCrOrLf_CellLineBreakCrLfOrCrOrLf AndAlso _CsvContainsColumnHeadersLineCount = 0 Then
                Throw New ArgumentException("Line endings setting RowBreakCrLfOrCrOrLf_CellLineBreakCrLfOrCrOrLf requires the CSV data to provide column headers")
            End If

            Return Result
        End Function

        ''' <summary>
        ''' Line encodings style in loaded CSV table
        ''' </summary>
        ''' <returns></returns>
        Public Property LineEncodingAutoConversions As CompuMaster.Data.Csv.ReadLineEncodingAutoConversion

        ''' <summary>
        ''' Convert empty strings to DBNull
        ''' </summary>
        ''' <returns></returns>
        Public Property ConvertEmptyStringsToDBNull As Boolean

    End Class

End Namespace