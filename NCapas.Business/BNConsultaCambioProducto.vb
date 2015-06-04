Imports NCapas.Entity
Imports NCapas.Data
Imports NCapas.Utility

Public Class BNConsultaCambioProducto
    Inherits Singleton(Of BNConsultaCambioProducto)

    Public Function InsertConsultaCambioProducto(ByVal consultaCambioProducto As ConsultaCambioProducto) As Boolean
        Dim resultado As Boolean
        Try
            resultado = DAConsultaCambioProducto.Instancia.InsertConsultaCambioProducto(consultaCambioProducto)
        Catch ex As Exception
            resultado = False
        End Try
        Return resultado
    End Function
End Class
