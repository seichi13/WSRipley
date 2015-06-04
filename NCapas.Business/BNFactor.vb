Imports NCapas.Entity
Imports NCapas.Data
Imports NCapas.Utility
Imports NCapas.Utility.Log

Public Class BNFactor
    Inherits Singleton(Of BNFactor)

    Public Function ObtenerFactorPorPlazo(ByVal plazo As String) As Factor
        Dim factor As New Factor
        Try
            factor = DAFactor.Instancia.ObtenerFactorPorPlazo(plazo)
        Catch ex As Exception
            factor = New Factor
        End Try
        Return factor
    End Function
End Class
