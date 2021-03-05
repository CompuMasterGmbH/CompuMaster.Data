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

        ' IDisposable
        <CodeAnalysis.SuppressMessage("Design", "CA1031:Do not catch general exception types", Justification:="<Ausstehend>")>
        Protected Overridable Sub Dispose(disposing As Boolean)
            If Not disposedValue Then
                If disposing Then
                    Try
                        If System.IO.File.Exists(Me.Path) Then System.IO.File.Delete(Me.Path)
                    Catch
                    End Try
                    ' TODO: verwalteten Zustand (verwaltete Objekte) entsorgen.
                End If

                ' TODO: nicht verwaltete Ressourcen (nicht verwaltete Objekte) freigeben und Finalize() weiter unten überschreiben.
                ' TODO: große Felder auf Null setzen.
            End If
            disposedValue = True
        End Sub

        ' TODO: Finalize() nur überschreiben, wenn Dispose(disposing As Boolean) weiter oben Code zur Bereinigung nicht verwalteter Ressourcen enthält.
        'Protected Overrides Sub Finalize()
        '    ' Ändern Sie diesen Code nicht. Fügen Sie Bereinigungscode in Dispose(disposing As Boolean) weiter oben ein.
        '    Dispose(False)
        '    MyBase.Finalize()
        'End Sub

        ' Dieser Code wird von Visual Basic hinzugefügt, um das Dispose-Muster richtig zu implementieren.
        Public Sub Dispose() Implements IDisposable.Dispose
            ' Ändern Sie diesen Code nicht. Fügen Sie Bereinigungscode in Dispose(disposing As Boolean) weiter oben ein.
            Dispose(True)
            ' TODO: Auskommentierung der folgenden Zeile aufheben, wenn Finalize() oben überschrieben wird.
            ' GC.SuppressFinalize(Me)
        End Sub
#End Region


    End Class

End Namespace