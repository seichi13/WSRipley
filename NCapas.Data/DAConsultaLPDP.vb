Imports NCapas.Entity
Imports System.Data.SqlClient
Imports NCapas.Data.DAConexion
Imports NCapas.Utility.Log

Public Class DAConsultaLPDP
    Inherits Singleton(Of DAConsultaLPDP)

    Public Function InsertConsultaLPDP(ByVal consultaLPDP As ConsultaLPDP) As Boolean

        Dim sMensajeError_SQL As String = ""
        Using oConexion As SqlConnection = DAConexion.Instancia.SQL_ConnectionOpen(DAConexion.Instancia.Get_CadenaConexion(), sMensajeError_SQL)
            If Not oConexion Is Nothing Then
                If oConexion.State = ConnectionState.Open Then
                    Using comando As SqlCommand = oConexion.CreateCommand()
                        comando.CommandTimeout = 600
                        comando.CommandType = CommandType.StoredProcedure
                        comando.CommandText = "Usp_Insert_ConsultaLPDP"

                        comando.Parameters.AddWithValue("@COD_SUCURSAL_BAN", consultaLPDP.Cod_Sucursal_Ban)
                        comando.Parameters.AddWithValue("@NOMBRE_KIOSCO", consultaLPDP.Nombre_Kiosko)
                        comando.Parameters.AddWithValue("@TIPO_DOC", consultaLPDP.Tipo_Doc)
                        comando.Parameters.AddWithValue("@NRO_DOC", consultaLPDP.Nro_Doc)
                        comando.Parameters.AddWithValue("@OPCION", consultaLPDP.Opcion)
                        comando.Parameters.AddWithValue("@NUMERO_TARJETA", consultaLPDP.Numero_Tarjeta)
                        comando.Parameters.AddWithValue("@NRO_CUENTA", consultaLPDP.Numero_Cuenta)
                        comando.Parameters.AddWithValue("@ID_KIOSKO", consultaLPDP.Id_Sucursal)

                        comando.ExecuteNonQuery()
                    End Using
                End If
            End If
        End Using
        Return True
    End Function

    Public Function GetTerminosCondicionesLPDP() As TerminosCondicionesLPDP
        Dim terminos As TerminosCondicionesLPDP = Nothing
        Dim sMensajeError_SQL As String = ""
        Using oConexion As SqlConnection = DAConexion.Instancia.SQL_ConnectionOpen(DAConexion.Instancia.Get_CadenaConexion(), sMensajeError_SQL)
            If oConexion.State = ConnectionState.Open Then
                Using cmd As SqlClient.SqlCommand = oConexion.CreateCommand
                    cmd.CommandTimeout = 600
                    cmd.CommandType = CommandType.StoredProcedure
                    cmd.CommandText = "USP_GET_TYC_LPDP"
                    Using reader As SqlClient.SqlDataReader = cmd.ExecuteReader
                        If reader.Read = True Then
                            terminos = New TerminosCondicionesLPDP
                            terminos.Id = reader.GetInt32(reader.GetOrdinal("Id"))
                            terminos.Titulo = reader.GetString(reader.GetOrdinal("Titulo"))
                            terminos.Cabecera = reader.GetString(reader.GetOrdinal("Cabecera"))
                            terminos.Condiciones = reader.GetString(reader.GetOrdinal("Condiciones"))
                            terminos.Confidencialidad = reader.GetString(reader.GetOrdinal("Confidencialidad"))
                            terminos.Caducidad = reader.GetString(reader.GetOrdinal("Caducidad"))
                        End If
                    End Using
                End Using
            End If
        End Using
        Return terminos
    End Function

    Public Function GetLoadLPDP(ByVal nroDocumento As String) As LoadLPDP
        Dim loadLPDP As LoadLPDP = Nothing
        Dim sMensajeError_SQL As String = ""
        Using oConexion As SqlConnection = DAConexion.Instancia.SQL_ConnectionOpen(DAConexion.Instancia.Get_CadenaConexion(), sMensajeError_SQL)
            If oConexion.State = ConnectionState.Open Then
                Using cmd As SqlClient.SqlCommand = oConexion.CreateCommand
                    cmd.CommandTimeout = 600
                    cmd.CommandType = CommandType.StoredProcedure
                    cmd.CommandText = "USP_GET_LOAD_LPDP"
                    cmd.Parameters.AddWithValue("@NRO_DOC", nroDocumento)
                    Using reader As SqlClient.SqlDataReader = cmd.ExecuteReader
                        If reader.Read = True Then
                            loadLPDP = New LoadLPDP
                            loadLPDP.Id = reader.GetInt32(reader.GetOrdinal("Id"))
                            loadLPDP.NroDocumento = reader.GetString(reader.GetOrdinal("NroDocumento"))
                        End If
                    End Using
                End Using
            End If
        End Using
        Return loadLPDP
    End Function

End Class
