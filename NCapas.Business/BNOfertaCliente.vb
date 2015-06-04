Imports NCapas.Entity
Imports NCapas.Data
Imports NCapas.Utility
Imports NCapas.Utility.Log

Public Class BNOfertaCliente
    Inherits Singleton(Of BNOfertaCliente)

    Public Function VerificarReenganche(ByVal tipoProducto As String, ByVal tipoDocumento As String, ByVal numeroDocumento As String) As OfertaCliente
        Dim oferta As New OfertaCliente
        Try
            oferta = DAOfertaCliente.Instancia.VerificarReenganche(tipoProducto, tipoDocumento, numeroDocumento)
        Catch ex As Exception
            oferta = New OfertaCliente
        End Try
        Return oferta
    End Function

    Public Function VerificarDesembolso(ByVal tipoProducto As String, ByVal tipoDocumento As String, ByVal numeroDocumento As String) As OfertaCliente
        Dim oferta As New OfertaCliente
        Try
            oferta = DAOfertaCliente.Instancia.VerificarDesembolso(tipoProducto, tipoDocumento, numeroDocumento)
        Catch ex As Exception
            oferta = New OfertaCliente
        End Try
        Return oferta
    End Function

    Public Function ActualizarOfertaPorSilumacion(ByVal tipoProducto As String, ByVal tipoDocumento As String, ByVal numeroDocumento As String, ByVal monto As String, ByVal plazoMeses As String, ByVal tasa As String, ByVal cuota As String, ByVal nombreUsuario As String, ByVal codigoSucursal As String) As OfertaCliente
        Dim oferta As New OfertaCliente
        Try
            oferta = DAOfertaCliente.Instancia.ActualizarOfertaPorSilumacion(tipoProducto, tipoDocumento, numeroDocumento, monto, plazoMeses, tasa, cuota, nombreUsuario, codigoSucursal)
        Catch ex As Exception
            oferta = New OfertaCliente
        End Try
        Return oferta
    End Function

    Public Function ObtenerOferta(ByVal tipoProducto As String, ByVal tipoDocumento As String, ByVal numeroDocumento As String) As OfertaCliente
        Dim oferta As New OfertaCliente
        Try
            oferta = DAOfertaCliente.Instancia.ObtenerOferta(tipoProducto, tipoDocumento, numeroDocumento)
        Catch ex As Exception
            oferta = New OfertaCliente
        End Try
        Return oferta
    End Function
End Class
