Imports NCapas.Entity
Imports NCapas.Data
Imports NCapas.Utility
Public Class BNConexion
    Inherits Singleton(Of BNConexion)
    Public Function PRUEBA_CONEXION_SQL() As String
        Dim result As String = ""
        Try
            Dim rpta As Boolean = DAConexion.Instancia.PRUEBA_CONEXION_SQL()
            If rpta Then
                result = "OK"
            Else
                result = Constantes.MSGErrorConexion
            End If
        Catch ex As Exception
            result = Constantes.MSGErrorConexion
        End Try
        Return result
    End Function

    Public Function OBTENER_VARIABLES_TIMEOUT() As String
        Dim result As String = ""
        Try
            result = DAConexion.Instancia.OBTENER_VARIABLES_TIMEOUT()
        Catch ex As Exception
            result = Constantes.MSGErrorConexion
        End Try
        Return result
    End Function
End Class
