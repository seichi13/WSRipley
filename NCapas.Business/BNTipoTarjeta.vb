Imports NCapas.Entity
Imports NCapas.Data
Imports NCapas.Utility

Public Class BNTipoTarjeta
    Inherits Singleton(Of BNTipoTarjeta)
    Public Function ObtenerTarjetasParaDocumento() As List(Of TipoTarjeta)
        Dim tiposTarjeta As New List(Of TipoTarjeta)
        Try
            tiposTarjeta = DATipoTarjeta.Instancia.ObtenerTarjetasParaDocumento()
        Catch ex As Exception
            tiposTarjeta = New List(Of TipoTarjeta)
        End Try
        Return tiposTarjeta
    End Function
End Class
