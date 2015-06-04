Imports NCapas.Entity
Imports System.Data.SqlClient
Imports NCapas.Data.DAConexion

Public Class DATipoTarjeta
    Inherits Singleton(Of DATipoTarjeta)

    Public Function ObtenerTarjetasParaDocumento() As List(Of TipoTarjeta)
        Dim tiposTarjeta As List(Of TipoTarjeta) = Nothing
        Dim sMensajeError_SQL As String = ""
        Using oConexion As SqlConnection = DAConexion.Instancia.SQL_ConnectionOpen(DAConexion.Instancia.Get_CadenaConexion(), sMensajeError_SQL)
            If oConexion.State = ConnectionState.Open Then
                Using cmd As SqlClient.SqlCommand = oConexion.CreateCommand
                    cmd.CommandTimeout = 600
                    cmd.CommandType = CommandType.StoredProcedure
                    cmd.CommandText = "USP_GET_TIPOTARJETASDOCUMENTO"
                    Using reader As SqlClient.SqlDataReader = cmd.ExecuteReader
                        While reader.Read = True
                            Dim tipoTarjeta As New TipoTarjeta
                            tipoTarjeta.Id = reader.GetInt32(reader.GetOrdinal("ID"))
                            tipoTarjeta.NombreTarjeta = reader.GetInt32(reader.GetOrdinal("NOMBRE_TARJETA"))
                            tiposTarjeta.Add(tipoTarjeta)
                        End While
                    End Using
                End Using
            End If
        End Using
        Return tiposTarjeta
    End Function

End Class
