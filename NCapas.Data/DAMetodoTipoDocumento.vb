Imports NCapas.Entity
Imports System.Data.SqlClient
Imports NCapas.Data.DAConexion
Imports NCapas.Utility.Log

Public Class DAMetodoTipoDocumento
    Inherits Singleton(Of DAMetodoTipoDocumento)

    Public Function BuscarMetodoPorDocumentoYTarjeta(ByVal tipoDocumento As Integer, ByVal idTipoTarjeta As Integer) As MetodoTipoDocumento
        Dim metodoTipoDocumento As MetodoTipoDocumento = Nothing
        Dim sMensajeError_SQL As String = ""
        Using oConexion As SqlConnection = DAConexion.Instancia.SQL_ConnectionOpen(DAConexion.Instancia.Get_CadenaConexion(), sMensajeError_SQL)
            If oConexion.State = ConnectionState.Open Then
                Using cmd As SqlClient.SqlCommand = oConexion.CreateCommand
                    cmd.CommandTimeout = 600
                    cmd.CommandType = CommandType.StoredProcedure
                    cmd.CommandText = "USP_GET_METODOTIPODOCUMENTOXTIPOTARJETA"
                    cmd.Parameters.AddWithValue("@TipoDocumento", tipoDocumento)
                    cmd.Parameters.AddWithValue("@IdTipoTarjeta", idTipoTarjeta)
                    Using reader As SqlClient.SqlDataReader = cmd.ExecuteReader
                        If reader.Read = True Then
                            metodoTipoDocumento = New MetodoTipoDocumento
                            metodoTipoDocumento.Id = reader.GetInt32(reader.GetOrdinal("ID"))
                            metodoTipoDocumento.TipoDocumento = reader.GetInt32(reader.GetOrdinal("TIPO_DOCUMENTO"))
                            metodoTipoDocumento.IdTipoTarjeta = reader.GetInt32(reader.GetOrdinal("ID_TIPO_TARJETA"))
                            metodoTipoDocumento.Metodo = reader.GetString(reader.GetOrdinal("METODO"))
                            metodoTipoDocumento.Estado = reader.GetInt32(reader.GetOrdinal("ESTADO"))
                        End If
                    End Using
                End Using
            End If
        End Using
        Return metodoTipoDocumento
    End Function

End Class
