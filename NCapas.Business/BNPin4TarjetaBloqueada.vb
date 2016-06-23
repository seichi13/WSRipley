Imports NCapas.Entity
Imports NCapas.Data
Imports NCapas.Utility
Imports NCapas.Utility.Log

Public Class BNPin4TarjetaBloqueada
    Inherits Singleton(Of BNPin4TarjetaBloqueada)

    Public Function GetByNroTarjeta(ByVal nroTarjeta As String) As Pin4TarjetaBloqueada
        Return DAPin4TarjetaBloqueada.Instancia.GetByNroTarjeta(nroTarjeta)
    End Function

    Public Sub InsertByNroTarjeta(ByVal oPin4TarjetaBloqueada As Pin4TarjetaBloqueada)
        DAPin4TarjetaBloqueada.Instancia.InsertByNroTarjeta(oPin4TarjetaBloqueada)
    End Sub

    Public Sub UpdateByNroTarjeta(ByVal oPin4TarjetaBloqueada As Pin4TarjetaBloqueada)
        DAPin4TarjetaBloqueada.Instancia.UpdateByNroTarjeta(oPin4TarjetaBloqueada)
    End Sub

    Public Function ValidateCreateNewInstance(ByVal oPin4TarjetaBloqueada As Pin4TarjetaBloqueada, ByVal oConfiguracionKiosko As ConfiguracionKiosko) As Boolean
        Dim oResult As Boolean
        If IsNothing(oPin4TarjetaBloqueada) Then
            oResult = True
        ElseIf Not oPin4TarjetaBloqueada.EstaBloqueada Then
            oResult = False
        ElseIf oConfiguracionKiosko.Pin4Intentos = 0 Then
            oResult = True
        ElseIf oPin4TarjetaBloqueada.FechaBloqueo.AddHours(oConfiguracionKiosko.Pin4HorasBloqueo) < oPin4TarjetaBloqueada.FechaActual Then
            oResult = True
        Else
            oResult = False
        End If
        Return oResult
    End Function
End Class
