Imports NCapas.Entity
Imports NCapas.Data
Imports NCapas.Utility
Public Class BNConfiguracionKiosco
    Inherits Singleton(Of BNConfiguracionKiosco)

    Public Function BuscarConfiguracionKioskoPorCodigoKiosco(ByVal sCodigoKiosko As String) As ConfiguracionKiosko
        Return DAConfiguracionKiosco.Instancia.BuscarConfiguracionKioskoPorCodigoKiosco(sCodigoKiosko)
    End Function

    Public Function BuscarConfiguracionKioskoPorIP(ByVal sIP As String) As ConfiguracionKiosko
        Dim config = New ConfiguracionKiosko
        Try
            config = DAConfiguracionKiosco.Instancia.BuscarConfiguracionKioskoPorIP(sIP)
        Catch ex As Exception
            config = New ConfiguracionKiosko
        End Try
        Return config
    End Function

End Class
