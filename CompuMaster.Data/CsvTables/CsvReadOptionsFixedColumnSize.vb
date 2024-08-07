﻿Option Explicit On
Option Strict On

Imports System.IO
Imports System.Data
Imports CompuMaster.Data.Strings
Imports System.Text

Namespace CompuMaster.Data.CsvTables

    Public Class CsvReadOptionsFixedColumnSize
        Inherits CsvReadBaseOptions

        ''' <summary>
        ''' Create a new instance of CsvReadOptionsFixedColumnSize
        ''' </summary>
        Public Sub New()
            'Me._CultureFormatProvider = System.Globalization.CultureInfo.InvariantCulture
            Me._CultureFormatProvider = System.Globalization.CultureInfo.CurrentCulture
        End Sub

        <Obsolete("Might be obsolete code since argument/property is without effect")>
        Friend Sub New(includesColumnHeaders As Boolean, startAtLineIndex As Integer, lineEncodings As CompuMaster.Data.Csv.ReadLineEncodings, lineEncodingAutoConversions As CompuMaster.Data.Csv.ReadLineEncodingAutoConversion, cultureFormatProvider As System.Globalization.CultureInfo, columnWidths As Integer(), convertEmptyStringsToDBNull As Boolean)
            Me.CsvContainsColumnHeaders = includesColumnHeaders
            Me.StartAtLineIndex = startAtLineIndex
            Me.LineEncodings = lineEncodings
            Me.LineEncodingAutoConversions = lineEncodingAutoConversions
            Me.ConvertEmptyStringsToDBNull = convertEmptyStringsToDBNull
            Me.ColumnWidths = columnWidths
            Me._CultureFormatProvider = cultureFormatProvider
        End Sub

        ''' <summary>
        ''' Create a new instance of CsvReadOptionsFixedColumnSize
        ''' </summary>
        ''' <param name="includesColumnHeaders">Indicates wether column headers are present</param>
        ''' <param name="lineEncodings">Encoding style for linebreaks</param>
        ''' <param name="lineEncodingAutoConversions">Change linebreak encodings on reading</param>
        ''' <param name="columnWidths">An array of column widths in their order</param>
        ''' <param name="convertEmptyStringsToDBNull">Convert values with empty strings automatically to DbNull</param>
        Public Sub New(includesColumnHeaders As Boolean, lineEncodings As CompuMaster.Data.Csv.ReadLineEncodings, lineEncodingAutoConversions As CompuMaster.Data.Csv.ReadLineEncodingAutoConversion, columnWidths As Integer(), convertEmptyStringsToDBNull As Boolean)
            Me.New(includesColumnHeaders, 0, lineEncodings, lineEncodingAutoConversions, columnWidths, convertEmptyStringsToDBNull)
        End Sub

        ''' <summary>
        ''' Create a new instance of CsvReadOptionsFixedColumnSize
        ''' </summary>
        ''' <param name="includesColumnHeaders">Indicates wether column headers are present</param>
        ''' <param name="startAtLineIndex">Start reading of table data at specified line index (most often 0 for very first line)</param>
        ''' <param name="lineEncodings">Encoding style for linebreaks</param>
        ''' <param name="lineEncodingAutoConversions">Change linebreak encodings on reading</param>
        ''' <param name="columnWidths">An array of column widths in their order</param>
        ''' <param name="convertEmptyStringsToDBNull">Convert values with empty strings automatically to DbNull</param>
        Public Sub New(includesColumnHeaders As Boolean, startAtLineIndex As Integer, lineEncodings As CompuMaster.Data.Csv.ReadLineEncodings, lineEncodingAutoConversions As CompuMaster.Data.Csv.ReadLineEncodingAutoConversion, columnWidths As Integer(), convertEmptyStringsToDBNull As Boolean)
            Me.CsvContainsColumnHeaders = includesColumnHeaders
            Me.StartAtLineIndex = startAtLineIndex
            Me.LineEncodings = lineEncodings
            Me.LineEncodingAutoConversions = lineEncodingAutoConversions
            Me.ConvertEmptyStringsToDBNull = convertEmptyStringsToDBNull
            Me.ColumnWidths = columnWidths
            Me._CultureFormatProvider = System.Globalization.CultureInfo.CurrentCulture
        End Sub

        Private _ColumnWidths As Integer()
        ''' <summary>
        ''' 
        ''' </summary>
        ''' <returns></returns>
        Public Property ColumnWidths As Integer()
            Get
                If _ColumnWidths Is Nothing Then
                    Return New Integer() {Integer.MaxValue}
                Else
                    Return _ColumnWidths
                End If
            End Get
            Private Set(value As Integer())
                _ColumnWidths = value
            End Set
        End Property

        Private _CultureFormatProvider As System.Globalization.CultureInfo
        <Obsolete("Might be obsolete code since argument/property is without effect")>
        Private Property CultureFormatProvider As System.Globalization.CultureInfo
            Get
                If Me._CultureFormatProvider Is Nothing Then
                    Return System.Globalization.CultureInfo.InvariantCulture
                Else
                    Return Me._CultureFormatProvider
                End If
            End Get
            Set(value As System.Globalization.CultureInfo)
                _CultureFormatProvider = value
            End Set
        End Property

    End Class

End Namespace