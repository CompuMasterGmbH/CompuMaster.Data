Namespace CompuMaster.Data.Web

    Friend Class Utils

        ''' <summary>
        '''     Converts all line breaks into HTML line breaks (&quot;&lt;br&gt;&quot;)
        ''' </summary>
        ''' <param name="text">A text string which might contain line breaks of any platform type</param>
        ''' <returns>The text string with encoded line breaks to &quot;&lt;br&gt;&quot;</returns>
        ''' <remarks>
        '''     Supported line breaks are linebreaks of Windows, MacOS as well as Linux/Unix.
        ''' </remarks>
        Public Shared Function HTMLEncodeLineBreaks(ByVal text As String) As String
            If text = Nothing Then
                Return text
            Else
                Return text.Replace(ControlChars.CrLf, "<br>").Replace(ControlChars.Cr, "<br>").Replace(ControlChars.Lf, "<br>")
            End If
        End Function

        ''' <summary>
        '''     Checks for DBNull and returns the second value alternatively
        ''' </summary>
        ''' <param name="CheckValueIfDBNull">The value to be checked</param>
        ''' <param name="ReplaceWithThis">The alternative value, null (Nothing in VisualBasic) if not defined</param>
        ''' <returns>A value which is not DBNull</returns>
        Public Shared Function Nz(ByVal CheckValueIfDBNull As Object, ByVal ReplaceWithThis As String) As String
            If IsDBNull(CheckValueIfDBNull) Then
                Return (ReplaceWithThis)
            Else
                Return CType(CheckValueIfDBNull, String)
            End If
        End Function

        ''' <summary>
        ''' Build string from given objects using stringbuilder
        ''' </summary>
        ''' <returns>Builded string</returns>
        ''' <history>
        ''' [clemens] 05.06.2013
        ''' </history>
        Public Shared Function BuildString(ByVal ParamArray str As Object()) As String
			Dim sb As Text.StringBuilder = New Text.StringBuilder()

			For counter As Integer = 0 To str.Length - 1
				sb.Append(str(counter))
			Next

			Return sb.ToString()
		End Function
	End Class

End Namespace