Imports NCapas.Entity
Imports NCapas.Data
Imports NCapas.Utility

Public Class BNMetodoTipoDocumento
    Inherits Singleton(Of BNMetodoTipoDocumento)

    Public Function BuscarMetodoPorDocumentoYTarjeta(ByVal tipoDocumento As Integer, ByVal idTipoTarjeta As Integer) As MetodoTipoDocumento
        Dim metodoTipoDocumento As New MetodoTipoDocumento
        Try
            metodoTipoDocumento = DAMetodoTipoDocumento.Instancia.BuscarMetodoPorDocumentoYTarjeta(tipoDocumento, idTipoTarjeta)
        Catch ex As Exception
            metodoTipoDocumento = New MetodoTipoDocumento
        End Try
        Return metodoTipoDocumento
    End Function
End Class
