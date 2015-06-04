Imports NCapas.Entity
Imports System.Data.SqlClient
Imports NCapas.Data.DAConexion
Imports NCapas.Utility.Log
Imports System.Data.OracleClient

Public Class DAFactor
    Inherits Singleton(Of DAFactor)

    Public Function ObtenerFactorPorPlazo(ByVal plazo As String) As Factor
        Dim factor As Factor = New Factor
        Dim sMensajeError_SQL As String = ""
        Using oConexion = New OracleConnection(DAConexion.Instancia.Get_CadenaConexionOracleFINX7Q2())
            If Not oConexion Is Nothing Then
                oConexion.Open()
                If oConexion.State = ConnectionState.Open Then
                    Using cmd As OracleCommand = oConexion.CreateCommand()
                        cmd.CommandTimeout = 600
                        cmd.CommandType = CommandType.StoredProcedure
                        cmd.CommandText = "PKG_RIPLEYMATICO_FACTOR.PRC_OBTENER_FACTOR_SEF"

                        Dim parametro1 As New OracleParameter
                        Dim parametro2 As New OracleParameter

                        parametro1.ParameterName = "I_TERM"
                        parametro1.OracleType = OracleType.Double
                        parametro1.Direction = ParameterDirection.Input
                        parametro1.Value = CDbl(plazo.Trim)
                        cmd.Parameters.Add(parametro1)

                        parametro2.ParameterName = "O_FACTOR"
                        parametro2.OracleType = OracleType.Double
                        parametro2.Direction = ParameterDirection.Output
                        cmd.Parameters.Add(parametro2)

                        cmd.ExecuteNonQuery()
                        factor.FactorPlazo = cmd.Parameters("O_FACTOR").Value.ToString()
                    End Using
                End If
            End If
        End Using
        Return factor
    End Function

End Class
