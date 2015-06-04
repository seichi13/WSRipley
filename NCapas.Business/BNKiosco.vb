Imports NCapas.Data
Imports NCapas.Entity
Public Class BNKiosco
    Inherits Singleton(Of BNKiosco)

    Public Function BuscarKioscoPorNombre(ByVal sNombreKiosco As String) As Kiosco
        Dim kiosco As Kiosco
        Try
            kiosco = DAKiosco.Instancia.BuscarKioscoPorNombre(sNombreKiosco)
        Catch ex As Exception
            kiosco = Nothing
        End Try
        Return kiosco
    End Function
End Class
