Imports System.IO
Imports System.Windows.Forms

Public Class Log

    Public Shared Sub Registrar_Log(ByVal pIntentos As Integer, ByVal pServicio As String, _
                             ByVal ParamArray pParametros As String())

        If pIntentos = 0 Then
            Exit Sub
        Else
            pIntentos -= 1
        End If

        Try

            Dim pPatch_Log As String = String.Empty
            Dim pFile_log As String = String.Empty
            Dim strLog As String = "<" & Date.Now.ToShortTimeString() & ">"

            pPatch_Log = AppDomain.CurrentDomain.BaseDirectory & "Log"

            pFile_log = pPatch_Log & "\Log_" & pServicio & +Date.Now.Year.ToString("0000") & Date.Now.Month.ToString("00") & Date.Now.Day.ToString("00") + ".txt"

            If Not Directory.Exists(pPatch_Log) Then
                Directory.CreateDirectory(pPatch_Log)
            End If

            Dim sw As StreamWriter = New StreamWriter(pFile_log, True)

            For Each Item As String In pParametros

                strLog += "<" & Item & ">"

            Next

            sw.WriteLine(strLog)
            sw.Close()

        Catch ex As Exception

            Registrar_Log(pIntentos, pServicio, pParametros)

        End Try

    End Sub

    Public Shared Sub ErrorLog(ByVal mensaje As String)
        Try

            Dim directorio As String = AppDomain.CurrentDomain.BaseDirectory & "Log"
            Dim archivo As String = directorio & "\log_" & +Date.Now.Year.ToString("0000") & Date.Now.Month.ToString("00") & Date.Now.Day.ToString("00") + ".txt"

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
