Imports NCapas.Entity
Imports System.Data.SqlClient
Imports NCapas.Data.DAConexion
Imports NCapas.Utility.Log

Public Class DATarjetaBin
    Inherits Singleton(Of DATarjetaBin)

    Public Function ObtenerTarjetaBin(ByVal bin As String) As TarjetaBin
        Dim tarjetaBin As New TarjetaBin

        Try
            Dim sMensajeError_SQL As String = ""
            Using oConexion As SqlConnection = DAConexion.Instancia.SQL_ConnectionOpen(DAConexion.Instancia.Get_CadenaConexion(), sMensajeError_SQL)
                If oConexion.State = ConnectionState.Open Then
                    Using cmd As SqlClient.SqlCommand = oConexion.CreateCommand
                        cmd.CommandTimeout = 600
                        cmd.CommandType = CommandType.StoredProcedure
                        cmd.CommandText = "USP_TARJETABIN_GET"
                        cmd.Parameters.AddWithValue("@Bin", bin)
                        Using reader As SqlClient.SqlDataReader = cmd.ExecuteReader
                            If reader.Read = True Then
                                tarjetaBin.Id = IIf(IsDBNull(reader.GetOrdinal("ID")), 0, reader.GetInt32(reader.GetOrdinal("ID")))
                                tarjetaBin.Bin = IIf(IsDBNull(reader.GetOrdinal("BIN")), String.Empty, reader.GetString(reader.GetOrdinal("BIN")))
                                tarjetaBin.Tipo = IIf(IsDBNull(reader.GetOrdinal("TIPO")), String.Empty, reader.GetString(reader.GetOrdinal("TIPO")))
                                tarjetaBin.Nombre = IIf(IsDBNull(reader.GetOrdinal("NOMBRE_TARJETA")), String.Empty, reader.GetString(reader.GetOrdinal("NOMBRE_TARJETA")))
                                tarjetaBin.Metodo = IIf(IsDBNull(reader.GetOrdinal("METODO_WS_CALL")), String.Empty, reader.GetString(reader.GetOrdinal("METODO_WS_CALL")))
                                tarjetaBin.TipoPasado = IIf(IsDBNull(reader.GetOrdinal("TIPO_TARJETA_PASADA")), String.Empty, reader.GetString(reader.GetOrdinal("TIPO_TARJETA_PASADA")))
                                tarjetaBin.TipoSEF = IIf(IsDBNull(reader.GetOrdinal("TIPO_SEF")), String.Empty, reader.GetString(reader.GetOrdinal("TIPO_SEF")))
                                tarjetaBin.TipoSimuladorSEF = IIf(IsDBNull(reader.GetOrdinal("TIPO_SIMULADOR_SEF")), String.Empty, reader.GetString(reader.GetOrdinal("TIPO_SIMULADOR_SEF")))
                                tarjetaBin.TipoAbiertoRsat = IIf(IsDBNull(reader.GetOrdinal("TIPO_ABIERTO_RSAT")), 0, reader.GetInt32(reader.GetOrdinal("TIPO_ABIERTO_RSAT")))
                            End If
                        End Using
                    End Using
                End If
            End Using
        Catch ex As Exception
            ErrorLog(ex.Message.Trim())
        End Try

        Return tarjetaBin
    End Function
End Class
