Imports NCapas.Entity
Imports System.Data.SqlClient
Imports NCapas.Data.DAConexion
Imports NCapas.Utility.Log

Public Class DAPin4TarjetaBloqueada
    Inherits Singleton(Of DAPin4TarjetaBloqueada)

    Public Function GetByNroTarjeta(ByVal nroTarjeta As String) As Pin4TarjetaBloqueada
        Dim oPin4TarjetaBloqueada As Pin4TarjetaBloqueada = Nothing
        Dim sMensajeError_SQL As String = ""

        Using oConexion As SqlConnection = DAConexion.Instancia.SQL_ConnectionOpen(DAConexion.Instancia.Get_CadenaConexion(), sMensajeError_SQL)
            If Not String.IsNullOrEmpty(sMensajeError_SQL) Then
                Throw New Exception(sMensajeError_SQL)
            End If
            If oConexion.State <> ConnectionState.Open Then
                Throw New Exception("No se pudo abrir la conexion a la base de datos")
            End If

            Using cmd As SqlClient.SqlCommand = oConexion.CreateCommand
                cmd.CommandTimeout = 600
                cmd.CommandType = CommandType.StoredProcedure
                cmd.CommandText = "Usp_Get_Pin4_Tarjetas_Bloqueadas_ByNroTarjeta"
                cmd.Parameters.AddWithValue("@NroTarjeta", nroTarjeta)

                Using reader As SqlClient.SqlDataReader = cmd.ExecuteReader
                    If reader.Read = True Then
                        oPin4TarjetaBloqueada = New Pin4TarjetaBloqueada
                        oPin4TarjetaBloqueada.Id = reader.GetInt32(reader.GetOrdinal("Id"))
                        oPin4TarjetaBloqueada.NroTarjeta = reader.GetString(reader.GetOrdinal("NroTarjeta"))
                        oPin4TarjetaBloqueada.EstaBloqueada = reader.GetBoolean(reader.GetOrdinal("EstaBloqueada"))
                        oPin4TarjetaBloqueada.NroIntentos = reader.GetInt16(reader.GetOrdinal("NroIntentos"))

                        If reader.IsDBNull(reader.GetOrdinal("FechaBloqueo")) Then
                            oPin4TarjetaBloqueada.FechaBloqueo = Nothing
                        Else
                            oPin4TarjetaBloqueada.FechaBloqueo = reader.GetDateTime(reader.GetOrdinal("FechaBloqueo"))
                        End If

                        oPin4TarjetaBloqueada.FechaActual = reader.GetDateTime(reader.GetOrdinal("FechaActual"))
                    End If
                End Using
            End Using
        End Using

        Return oPin4TarjetaBloqueada
    End Function

    Public Sub InsertByNroTarjeta(ByVal oPin4TarjetaBloqueada As Pin4TarjetaBloqueada)
        Dim sMensajeError_SQL As String = ""
        Using oConexion As SqlConnection = DAConexion.Instancia.SQL_ConnectionOpen(DAConexion.Instancia.Get_CadenaConexion(), sMensajeError_SQL)
            If Not String.IsNullOrEmpty(sMensajeError_SQL) Then
                Throw New Exception(sMensajeError_SQL)
            End If
            If oConexion.State <> ConnectionState.Open Then
                Throw New Exception("No se pudo abrir la conexion a la base de datos")
            End If

            Using cmd As SqlClient.SqlCommand = oConexion.CreateCommand
                cmd.CommandTimeout = 600
                cmd.CommandType = CommandType.StoredProcedure
                cmd.CommandText = "Usp_Insert_Pin4_Tarjetas_Bloqueadas_ByNroTarjeta"
                cmd.Parameters.AddWithValue("@NroTarjeta", oPin4TarjetaBloqueada.NroTarjeta)
                cmd.Parameters.AddWithValue("@EstaBloqueada", oPin4TarjetaBloqueada.EstaBloqueada)
                cmd.Parameters.AddWithValue("@NroIntentos", oPin4TarjetaBloqueada.NroIntentos)

                cmd.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    Public Sub UpdateByNroTarjeta(ByVal oPin4TarjetaBloqueada As Pin4TarjetaBloqueada)
        Dim sMensajeError_SQL As String = ""
        Using oConexion As SqlConnection = DAConexion.Instancia.SQL_ConnectionOpen(DAConexion.Instancia.Get_CadenaConexion(), sMensajeError_SQL)
            If Not String.IsNullOrEmpty(sMensajeError_SQL) Then
                Throw New Exception(sMensajeError_SQL)
            End If
            If oConexion.State <> ConnectionState.Open Then
                Throw New Exception("No se pudo abrir la conexion a la base de datos")
            End If

            Using cmd As SqlClient.SqlCommand = oConexion.CreateCommand
                cmd.CommandTimeout = 600
                cmd.CommandType = CommandType.StoredProcedure
                cmd.CommandText = "Usp_Update_Pin4_Tarjetas_Bloqueadas_ByNroTarjeta"
                cmd.Parameters.AddWithValue("@NroTarjeta", oPin4TarjetaBloqueada.NroTarjeta)
                cmd.Parameters.AddWithValue("@EstaBloqueada", oPin4TarjetaBloqueada.EstaBloqueada)
                cmd.Parameters.AddWithValue("@NroIntentos", oPin4TarjetaBloqueada.NroIntentos)

                cmd.ExecuteNonQuery()
            End Using
        End Using
    End Sub
End Class
