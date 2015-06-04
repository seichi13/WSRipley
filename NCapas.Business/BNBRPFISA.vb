Imports NCapas.Entity
Imports NCapas.Data
Imports NCapas.Utility
Imports NCapas.Utility.Log

Public Class BNBRPFISA
    Inherits Singleton(Of BNBRPFISA)

    Public Function ObtenerEmailFISA(ByVal numeroDocumento As String) As String
        Dim respuesta As String = String.Empty
        Try
            respuesta = DABRPFISA.Instancia.ObtenerEmailFISA(numeroDocumento)
        Catch ex As Exception
            ErrorLog(ex.Message)
            respuesta = String.Empty
        End Try
        Return respuesta
    End Function

    Public Function ObtenerEmailPersonal(ByVal tipoDocumento As String, ByVal numeroDocumento As String) As String
        Dim respuesta As String = String.Empty
        Try
            respuesta = DABRPFISA.Instancia.ObtenerEmailPersonal(tipoDocumento, numeroDocumento)
        Catch ex As Exception
            ErrorLog(ex.Message)
            respuesta = String.Empty
        End Try
        Return respuesta
    End Function

    Public Function ObtenerEmailLaboral(ByVal tipoDocumento As String, ByVal numeroDocumento As String) As String
        Dim respuesta As String = String.Empty
        Try
            respuesta = DABRPFISA.Instancia.ObtenerEmailLaboral(tipoDocumento, numeroDocumento)
        Catch ex As Exception
            ErrorLog(ex.Message)
            respuesta = String.Empty
        End Try
        Return respuesta
    End Function

    Public Function ObtenerOtrosEmail(ByVal tipoDocumento As String, ByVal numeroDocumento As String) As Cliente
        Dim oCliente As New Cliente
        Try
            oCliente = DABRPFISA.Instancia.ObtenerOtrosEmail(tipoDocumento, numeroDocumento)
        Catch ex As Exception
            ErrorLog(ex.Message)
        End Try
        Return oCliente
    End Function

End Class
