Imports NCapas.Entity
Imports NCapas.Data
Imports NCapas.Utility
Imports NCapas.Utility.Log

Public Class BNPin4TarjetaBloqueada
    Inherits Singleton(Of BNPin4TarjetaBloqueada)

    Public Function GetPin4TarjetaBloqueadaByNroTarjeta(ByVal nroTarjeta As String) As Pin4TarjetaBloqueada
        Return DAPin4TarjetaBloqueada.Instancia.GetPin4TarjetaBloqueadaByNroTarjeta(nroTarjeta)
    End Function

    Public Sub InsertPin4TarjetaBloqueadaByNroTarjeta(ByVal oPin4TarjetaBloqueada As Pin4TarjetaBloqueada)
        DAPin4TarjetaBloqueada.Instancia.InsertPin4TarjetaBloqueadaByNroTarjeta(oPin4TarjetaBloqueada)
    End Sub

    Public Sub UpdatePin4TarjetaBloqueadaByNroTarjeta(ByVal oPin4TarjetaBloqueada As Pin4TarjetaBloqueada)
        DAPin4TarjetaBloqueada.Instancia.UpdatePin4TarjetaBloqueadaByNroTarjeta(oPin4TarjetaBloqueada)
    End Sub
End Class
