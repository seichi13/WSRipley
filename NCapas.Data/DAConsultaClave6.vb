Imports NCapas.Entity
Imports System.Data.SqlClient
Imports NCapas.Data.DAConexion
Imports NCapas.Utility.Log

Public Class DAConsultaClave6
    Inherits Singleton(Of DAConsultaClave6)

    Public Function InsertConsultaClave6(ByVal consultaClave6 As ConsultaClave6) As Boolean
        ErrorLog("DA  Inicio InsertConsultaClave6")
        Dim sMensajeError_SQL As String = ""
        Using oConexion As SqlConnection = DAConexion.Instancia.SQL_ConnectionOpen(DAConexion.Instancia.Get_CadenaConexion(), sMensajeError_SQL)
            If oConexion.State = ConnectionState.Open Then
                Using comando As SqlCommand = oConexion.CreateCommand()
                    comando.CommandTimeout = 600
                    comando.CommandType = CommandType.StoredProcedure
                    comando.CommandText = "Usp_Insert_ConsultaClave6"
                    ErrorLog("DA  InsertConsultaClave6")
                    comando.Parameters.AddWithValue("@COD_SUCURSAL_BAN", consultaClave6.Cod_Sucursal_Ban)
                    comando.Parameters.AddWithValue("@NOMBRE_KIOSCO", consultaClave6.Nombre_Kiosko)
                    comando.Parameters.AddWithValue("@ID_KIOSKO", consultaClave6.Id_Kiosko)
                    comando.Parameters.AddWithValue("@TIPO_DOC", consultaClave6.Tipo_Doc)
                    comando.Parameters.AddWithValue("@NRO_DOC", consultaClave6.Nro_Doc)
                    comando.Parameters.AddWithValue("@NUMERO_TARJETA", consultaClave6.Numero_Tarjeta)
                    comando.Parameters.AddWithValue("@OPCION", consultaClave6.Opcion)
                    comando.Parameters.AddWithValue("@EMAIL", consultaClave6.Email)
                    comando.Parameters.AddWithValue("@MENSAJE_ERROR", consultaClave6.Mensaje_Error)
                    ErrorLog("Antes de ejecutar Usp_Insert_ConsultaClave6")
                    comando.ExecuteNonQuery()
                    ErrorLog("Después de ejecutar Usp_Insert_ConsultaClave6")
                End Using
            End If
        End Using
        Return True
    End Function

    Public Function UpdateConsultaClave6(ByVal consultaClave6 As ConsultaClave6) As Boolean

        Dim sMensajeError_SQL As String = ""
        Using oConexion As SqlConnection = DAConexion.Instancia.SQL_ConnectionOpen(DAConexion.Instancia.Get_CadenaConexion(), sMensajeError_SQL)
            If oConexion.State = ConnectionState.Open Then
                Using comando As SqlCommand = oConexion.CreateCommand()
                    comando.CommandTimeout = 600
                    comando.CommandType = CommandType.StoredProcedure
                    comando.CommandText = "Usp_Update_ConsultaClave6"

                    comando.Parameters.AddWithValue("@NRO_DOC", consultaClave6.Nro_Doc)
                    comando.Parameters.AddWithValue("@EMAIL", consultaClave6.Email)
                    comando.Parameters.AddWithValue("@ENVIO", consultaClave6.Envio)
                    comando.Parameters.AddWithValue("@CONTADOR", consultaClave6.Contador)
                    ErrorLog("Antes de ejecutar UpdateConsultaClave6")
                    comando.ExecuteNonQuery()
                    ErrorLog("Despues de ejecutar UpdateConsultaClave6")
                End Using
            End If
        End Using
        Return True
    End Function

    Public Function GetTerminosCondiciones() As TerminosCondiciones
        Dim terminos As TerminosCondiciones = Nothing
        Dim sMensajeError_SQL As String = ""
        Using oConexion As SqlConnection = DAConexion.Instancia.SQL_ConnectionOpen(DAConexion.Instancia.Get_CadenaConexion(), sMensajeError_SQL)
            If oConexion.State = ConnectionState.Open Then
                Using cmd As SqlClient.SqlCommand = oConexion.CreateCommand
                    cmd.CommandTimeout = 600
                    cmd.CommandType = CommandType.StoredProcedure
                    cmd.CommandText = "USP_GET_TYC"
                    Using reader As SqlClient.SqlDataReader = cmd.ExecuteReader
                        If reader.Read = True Then
                            terminos = New TerminosCondiciones
                            terminos.Id = reader.GetInt32(reader.GetOrdinal("Id"))
                            terminos.Titulo = reader.GetString(reader.GetOrdinal("Titulo"))
                            terminos.Cabecera = reader.GetString(reader.GetOrdinal("Cabecera"))
                            terminos.Condiciones = reader.GetString(reader.GetOrdinal("Condiciones"))
                            terminos.Confidencialidad = reader.GetString(reader.GetOrdinal("Confidencialidad"))
                            terminos.ResponsabilidadInf = reader.GetString(reader.GetOrdinal("ResponsabilidadInf"))
                            terminos.Acceso = reader.GetString(reader.GetOrdinal("Acceso"))
                            terminos.Fallas = reader.GetString(reader.GetOrdinal("Fallas"))
                            terminos.Propiedad = reader.GetString(reader.GetOrdinal("Propiedad"))
                            terminos.Responsabilidad = reader.GetString(reader.GetOrdinal("Responsabilidad"))
                            terminos.Otros = reader.GetString(reader.GetOrdinal("Otros"))
                            terminos.Anexo = reader.GetString(reader.GetOrdinal("Anexo"))
                        End If
                    End Using
                End Using
            End If
        End Using
        Return terminos
    End Function

    Public Function ValidarClave(ByVal clave As String) As Integer
        Dim respuesta As Integer = 0
        Dim sMensajeError_SQL As String = ""
        Using oConexion As SqlConnection = DAConexion.Instancia.SQL_ConnectionOpen(DAConexion.Instancia.Get_CadenaConexion(), sMensajeError_SQL)
            If oConexion.State = ConnectionState.Open Then
                Using comando As SqlCommand = oConexion.CreateCommand()
                    comando.CommandTimeout = 600
                    comando.CommandType = CommandType.StoredProcedure
                    comando.CommandText = "USP_VALIDAR_CLAVE"
                    comando.Parameters.AddWithValue("@Clave", clave)
                    Using reader As SqlClient.SqlDataReader = comando.ExecuteReader
                        If reader.Read = True Then
                            respuesta = reader.GetInt32(reader.GetOrdinal("Estado"))
                        End If
                    End Using
                End Using
            End If
        End Using
        Return respuesta
    End Function

    Public Function CountClientesByNroTarjeta(ByVal nroTarjeta As String) As Integer
        Dim respuesta As Integer = 0
        Dim sMensajeError_SQL As String = ""
        Using oConexion As SqlConnection = DAConexion.Instancia.SQL_ConnectionOpen(DAConexion.Instancia.Get_CadenaConexion(), sMensajeError_SQL)
            If oConexion.State = ConnectionState.Open Then
                Using comando As SqlCommand = oConexion.CreateCommand()
                    comando.CommandTimeout = 600
                    comando.CommandType = CommandType.StoredProcedure
                    comando.CommandText = "USP_Count_KIO_CLAVE6_CLIENTES_BY_NroTarjeta"
                    comando.Parameters.AddWithValue("@NroTarjeta", nroTarjeta)
                    Using reader As SqlClient.SqlDataReader = comando.ExecuteReader
                        If reader.Read = True Then
                            respuesta = reader.GetInt32(reader.GetOrdinal("COUNT"))
                        End If
                    End Using
                End Using
            End If
        End Using
        Return respuesta
    End Function
End Class
