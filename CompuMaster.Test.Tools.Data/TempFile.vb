Imports System

Namespace CompuMaster.Test.Data

    Friend Class TemporaryFile
        Implements IDisposable

        ''' <summary>
        ''' Create a temp file with auto-cleanup of file
        ''' </summary>
        ''' <param name="fileNameExtension">An extension inclusive the leading dot, e.g. .txt</param>
        Public Sub New(fileNameExtension As String)
            Me.Path = System.IO.Path.GetTempFileName() & fileNameExtension
        End Sub

        Public ReadOnly Property Path As String

        Public Function FileSize() As Long
            Dim File As New System.IO.FileInfo(Me.Path)
            Return File.Length
        End Function

#Region "IDisposable Support"
        Private disposedValue As Boolean ' Dient zur Erkennung redundanter Aufrufe.

        <CodeAnalysis.SuppressMessage("Design", "CA1031:Do not catch general exception types", Justification:="<Ausstehend>")>
        <CodeAnalysis.SuppressMessage("Major Code Smell", "S1066:Collapsible ""if"" statements should be merged", Justification:="<Ausstehend>")>
        Protected Overridable Sub Dispose(disposing As Boolean)
            If Not disposedValue Then
                If disposing Then
                    Try
                        If System.IO.File.Exists(Me.Path) Then System.IO.File.Delete(Me.Path)
                    Catch
                        'Ignore exceptions and leave the file on disk
                    End Try
                End If
            End If
            disposedValue = True
        End Sub

        Public Sub Dispose() Implements IDisposable.Dispose
            Dispose(True)
        End Sub
#End Region


    End Class

End Namespace