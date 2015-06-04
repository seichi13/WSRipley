Imports NCapas.Entity
Imports NCapas.Data
Imports NCapas.Utility
Imports NCapas.Utility.Log

Public Class BNTarjetaBin
    Inherits Singleton(Of BNTarjetaBin)

    Public Function ObtenerTarjetaBin(ByVal bin As String) As TarjetaBin
        Dim tarjetaBin As New TarjetaBin
        Try
            tarjetaBin = DATarjetaBin.Instancia.ObtenerTarjetaBin(bin)
        Catch ex As Exception
            ErrorLog(ex.Message.Trim())
        End Try
        Return tarjetaBin
    End Function

End Class
