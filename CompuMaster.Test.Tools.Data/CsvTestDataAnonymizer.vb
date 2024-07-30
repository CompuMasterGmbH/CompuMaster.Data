Imports System
Imports System.IO
Imports System.Text
Imports NUnit.Framework

Namespace CompuMaster.Test.Data

    ''' <summary>
    ''' Helper class for developer to anonymize test data files
    ''' </summary>
    Public Class CsvTestDataAnonymizer

        <Test>
        <Ignore("Intended for manual file conversion by developer only")>
        <Explicit("Intended for manual file conversion by developer only")>
        Public Sub AnonymizeCsvTestFile()
            ' Beispielaufruf der Methode            
            AnonymizeCsv("D:\GitHub-OpenSource\CompuMaster.Data\CompuMaster.Test.Tools.Data\testfiles\lexoffice_sensitive.csv", "D:\GitHub-OpenSource\CompuMaster.Data\CompuMaster.Test.Tools.Data\testfiles\lexoffice.csv")
        End Sub

        ''' <summary>
        ''' Anonymize a CSV file by replacing all letters with random letters and all German umlauts with random umlauts. Numbers and special characters are kept as they are. File encodings and BOMs are preserved.
        ''' </summary>
        ''' <param name="inputFilePath"></param>
        ''' <param name="outputFilePath"></param>
        Public Shared Sub AnonymizeCsv(inputFilePath As String, outputFilePath As String)
            ' Datei binär öffnen
            Dim fileBytes As Byte() = File.ReadAllBytes(inputFilePath)

            ' Encoding der Datei ermitteln und BOM beibehalten
            Dim encoding As Encoding = DetectFileEncoding(fileBytes)

            ' BOM aus den Bytes entfernen, falls vorhanden
            Dim preamble As Byte() = encoding.GetPreamble()
            Dim contentBytes As Byte()
            If fileBytes.Take(preamble.Length).SequenceEqual(preamble) Then
                contentBytes = fileBytes.Skip(preamble.Length).ToArray()
            Else
                contentBytes = fileBytes
            End If

            ' Inhalt der Datei als String einlesen
            Dim fileContent As String = encoding.GetString(contentBytes)

            ' Anonymisierte Inhalte erstellen
            Dim anonymizedContent As New StringBuilder()

            Dim random As New Random()

            For Each ch As Char In fileContent
                If Char.IsLetter(ch) OrElse "äöüÄÖÜß".Contains(ch) Then
                    anonymizedContent.Append(RandomizeChar(ch, random))
                Else
                    anonymizedContent.Append(ch)
                End If
            Next

            ' Anonymisierte Inhalte zurück in Bytes konvertieren
            Dim anonymizedBytes As Byte() = encoding.GetBytes(anonymizedContent.ToString())

            ' BOM hinzufügen, falls vorhanden
            Dim finalBytes As Byte() = If(preamble.Length > 0, preamble.Concat(anonymizedBytes).ToArray(), anonymizedBytes)

            ' Anonymisierte Datei speichern
            File.WriteAllBytes(outputFilePath, finalBytes)
        End Sub

        Private Shared Function DetectFileEncoding(fileBytes As Byte()) As Encoding
            ' Überprüfen auf BOMs für UTF-8, UTF-16 und UTF-32
            If fileBytes.Length >= 3 AndAlso fileBytes(0) = &HEF AndAlso fileBytes(1) = &HBB AndAlso fileBytes(2) = &HBF Then
                Return New UTF8Encoding(True)
            ElseIf fileBytes.Length >= 2 AndAlso fileBytes(0) = &HFF AndAlso fileBytes(1) = &HFE Then
                Return Encoding.Unicode ' UTF-16 LE
            ElseIf fileBytes.Length >= 2 AndAlso fileBytes(0) = &HFE AndAlso fileBytes(1) = &HFF Then
                Return Encoding.BigEndianUnicode ' UTF-16 BE
            ElseIf fileBytes.Length >= 4 AndAlso fileBytes(0) = &HFF AndAlso fileBytes(1) = &HFE AndAlso fileBytes(2) = &H0 AndAlso fileBytes(3) = &H0 Then
                Return Encoding.UTF32 ' UTF-32 LE
            ElseIf fileBytes.Length >= 4 AndAlso fileBytes(0) = &H0 AndAlso fileBytes(1) = &H0 AndAlso fileBytes(2) = &HFE AndAlso fileBytes(3) = &HFF Then
                Return New UTF32Encoding(True, True) ' UTF-32 BE
            End If

            ' Standard Encoding
            Return Encoding.Default
        End Function

        Private Shared Function RandomizeChar(ch As Char, random As Random) As Char
            Dim offset As Integer
            If Char.IsLower(ch) Then
                ' Kleinbuchstaben
                offset = AscW("a"c)
                Return ChrW(offset + random.Next(0, 26))
            ElseIf Char.IsUpper(ch) Then
                ' Großbuchstaben
                offset = AscW("A"c)
                Return ChrW(offset + random.Next(0, 26))
            Else
                ' Deutsche Umlaute und ß
                Dim umlauts As String = "äöüÄÖÜß"
                Dim randomUmlaut As Char = umlauts(random.Next(0, umlauts.Length))
                Return randomUmlaut
            End If
        End Function

    End Class

End Namespace