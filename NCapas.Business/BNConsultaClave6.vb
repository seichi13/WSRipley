Imports NCapas.Entity
Imports NCapas.Data
Imports NCapas.Utility
Imports NCapas.Utility.Log

Public Class BNConsultaClave6
    Inherits Singleton(Of BNConsultaClave6)

    Public Function InsertConsultaClave6(ByVal consultaClave6 As ConsultaClave6) As Boolean
        Dim resultado As Boolean
        Try
            resultado = DAConsultaClave6.Instancia.InsertConsultaClave6(consultaClave6)
        Catch ex As Exception
            ErrorLog("Error en InsertConsultaClave6 = " & ex.Message)
            resultado = False
        End Try
        Return resultado
    End Function

    Public Function UpdateConsultaClave6(ByVal consultaClave6 As ConsultaClave6) As Boolean
        Dim resultado As Boolean
        Try
            resultado = DAConsultaClave6.Instancia.UpdateConsultaClave6(consultaClave6)
        Catch ex As Exception
            ErrorLog("Error en UpdateConsultaClave6 = " & ex.Message)
            resultado = False
        End Try
        Return resultado
    End Function

    Public Function GetTerminosCondiciones() As TerminosCondiciones
        Dim terminos As TerminosCondiciones = Nothing
        Try
            terminos = DAConsultaClave6.Instancia.GetTerminosCondiciones()
        Catch ex As Exception
            ErrorLog("Error en GetTerminosCondiciones = " & ex.Message)
        End Try
        Return terminos
    End Function

    Public Function ValidarClave(ByVal clave As String) As Integer
        Dim respuesta As Integer = 0
        Try
            respuesta = DAConsultaClave6.Instancia.ValidarClave(clave)
        Catch ex As Exception
            ErrorLog("Error en ValidarClave = " & ex.Message)
        End Try
        Return respuesta
    End Function
End Class
