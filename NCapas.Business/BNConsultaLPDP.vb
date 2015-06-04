Imports NCapas.Entity
Imports NCapas.Data
Imports NCapas.Utility
Imports NCapas.Utility.Log

Public Class BNConsultaLPDP
    Inherits Singleton(Of BNConsultaLPDP)

    Public Function InsertConsultaLPDP(ByVal consultaLPDP As ConsultaLPDP) As Boolean
        Dim resultado As Boolean
        Try
            resultado = DAConsultaLPDP.Instancia.InsertConsultaLPDP(ConsultaLPDP)
        Catch ex As Exception
            ErrorLog("Error InsertConsultaLPDP = " & ex.Message)
            resultado = False
        End Try
        Return resultado
    End Function

    Public Function GetTerminosCondicionesLPDP() As TerminosCondicionesLPDP
        Dim terminos As TerminosCondicionesLPDP = Nothing
        Try
            terminos = DAConsultaLPDP.Instancia.GetTerminosCondicionesLPDP()
        Catch ex As Exception
            ErrorLog("Error en GetTerminosCondicionesLPDP = " & ex.Message)
        End Try
        Return terminos
    End Function

    Public Function GetLoadLPDP(ByVal nroDocumento As String) As LoadLPDP
        Dim loadLPDP As LoadLPDP = Nothing
        Try
            loadLPDP = DAConsultaLPDP.Instancia.GetLoadLPDP(nroDocumento)
        Catch ex As Exception
            ErrorLog("Error en GetLoadLPDP = " & ex.Message)
        End Try
        Return loadLPDP
    End Function

End Class
