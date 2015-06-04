Imports System.IO
Imports System.Windows.Forms

Public Class Log

    Public Shared Sub ErrorLog(ByVal mensaje As String)
        Try

            Dim directorio As String = AppDomain.CurrentDomain.BaseDirectory & "Log"
            Dim archivo As String = directorio & "\Log.txt"

            If Not Directory.Exists(directorio) Then
                Directory.CreateDirectory(directorio)
            End If

            Dim streamWriter As StreamWriter = New StreamWriter(archivo, True)

            streamWriter.WriteLine(Date.Now.ToShortDateString() & " " & Date.Now.ToShortTimeString() & " " & mensaje)
            streamWriter.Close()

        Catch ex As Exception

            ErrorLog(mensaje)

        End Try
    End Sub
End Class
