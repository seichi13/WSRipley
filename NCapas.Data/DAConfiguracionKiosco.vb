Imports NCapas.Entity

Public Class DAConfiguracionKiosco
    Inherits Singleton(Of DAConfiguracionKiosco)

    Public Function BuscarConfiguracionKioskoPorCodigoKiosco(ByVal sCodigoKiosko As String) As ConfiguracionKiosko
        Dim config As New ConfiguracionKiosko
        Dim sMensajeError_SQL As String = ""
        Using oConexion = DAConexion.Instancia.SQL_ConnectionOpen(DAConexion.Instancia.Get_CadenaConexion(), sMensajeError_SQL)
            If sMensajeError_SQL <> "" Then
                config = New ConfiguracionKiosko
                config.ID = 100
                config.Nombre = sCodigoKiosko
            Else
                If oConexion.State = ConnectionState.Open Then
                    Dim m_ssql As String = "USP_GET_CONFIGURACION_POR_CODIGOKIOSKO"
                    Using cmd As SqlClient.SqlCommand = oConexion.CreateCommand
                        Dim lTotal As Double = 0
                        cmd.CommandTimeout = 600
                        cmd.CommandType = CommandType.StoredProcedure
                        cmd.CommandText = m_ssql
                        cmd.Parameters.AddWithValue("@COD_KIOSKO", sCodigoKiosko)
                        config = Nothing
                        Using rd_cod As SqlClient.SqlDataReader = cmd.ExecuteReader
                            If rd_cod.Read = True Then
                                config = New ConfiguracionKiosko

                                config.ID = rd_cod.GetInt32(rd_cod.GetOrdinal("ID"))
                                config.Nombre = rd_cod.GetString(rd_cod.GetOrdinal("NOMBRE"))
                                config.Server = rd_cod.GetString(rd_cod.GetOrdinal("SERVER"))
                                config.Server_Simulador = rd_cod.GetString(rd_cod.GetOrdinal("SERVER_SIMULADOR"))
                                config.Server_Com = rd_cod.GetString(rd_cod.GetOrdinal("SERVER_COM"))
                                config.Server_Print = rd_cod.GetString(rd_cod.GetOrdinal("SERVER_PRINT"))

                                config.Codigo_Kiosko = rd_cod.GetString(rd_cod.GetOrdinal("Codigo_Kiosko"))
                                config.Codigo_Sucursal = rd_cod.GetInt32(rd_cod.GetOrdinal("Codigo_Sucursal")).ToString()
                                config.ID_Kiosko = rd_cod.GetInt32(rd_cod.GetOrdinal("ID_Kiosko")).ToString()

                                config.Bin1 = rd_cod.GetString(rd_cod.GetOrdinal("BIN1"))
                                config.Longitud_Tarjeta_Bin1 = rd_cod.GetInt32(rd_cod.GetOrdinal("LONGITUD_TARJETA_BIN1"))
                                config.Posini_Bin1 = rd_cod.GetInt32(rd_cod.GetOrdinal("POSINI_BIN1"))
                                config.Longitud_Bin_Bin1 = rd_cod.GetInt32(rd_cod.GetOrdinal("LONGITUD_BIN_BIN1"))

                                config.Bin2 = rd_cod.GetString(rd_cod.GetOrdinal("BIN2"))
                                config.Longitud_Tarjeta_Bin2 = rd_cod.GetInt32(rd_cod.GetOrdinal("LONGITUD_TARJETA_BIN2"))
                                config.Posini_Bin2 = rd_cod.GetInt32(rd_cod.GetOrdinal("POSINI_BIN2"))
                                config.Longitud_Bin_Bin2 = rd_cod.GetInt32(rd_cod.GetOrdinal("LONGITUD_BIN_BIN2"))

                                config.Pin4Intentos = rd_cod.GetInt16(rd_cod.GetOrdinal("PIN4_INTENTOS"))
                                config.Pin4HorasBloqueo = rd_cod.GetInt16(rd_cod.GetOrdinal("PIN4_HORAS_BLOQUEO"))
                            End If
                        End Using
                    End Using
                Else
                    config = Nothing
                End If
            End If
        End Using
        Return config
    End Function

    Public Function BuscarConfiguracionKioskoPorIP(ByVal sIP As String) As ConfiguracionKiosko
        Dim config = New ConfiguracionKiosko
        Dim sMensajeError_SQL As String = ""
        Using oConexion = DAConexion.Instancia.SQL_ConnectionOpen(DAConexion.Instancia.Get_CadenaConexion(), sMensajeError_SQL)
            If sMensajeError_SQL <> "" Then
                config = Nothing
            Else
                If oConexion.State = ConnectionState.Open Then
                    Dim m_ssql As String = "USP_GET_CONFIGURACION_POR_IP"
                    Using cmd As SqlClient.SqlCommand = oConexion.CreateCommand
                        Dim lTotal As Double = 0
                        cmd.CommandTimeout = 600
                        cmd.CommandType = CommandType.StoredProcedure
                        cmd.CommandText = m_ssql
                        cmd.Parameters.AddWithValue("@IP", sIP)
                        Using rd_cod As SqlClient.SqlDataReader = cmd.ExecuteReader
                            If rd_cod.Read = True Then
                                config = New ConfiguracionKiosko
                                config.ID = rd_cod.GetInt32(rd_cod.GetOrdinal("ID"))
                                config.Nombre = rd_cod.GetString(rd_cod.GetOrdinal("NOMBRE"))
                                config.Server = rd_cod.GetString(rd_cod.GetOrdinal("SERVER"))
                                config.Server_Simulador = rd_cod.GetString(rd_cod.GetOrdinal("SERVER_SIMULADOR"))
                                config.Server_Com = rd_cod.GetString(rd_cod.GetOrdinal("SERVER_COM"))
                                config.Server_Print = rd_cod.GetString(rd_cod.GetOrdinal("SERVER_PRINT"))

                                config.Codigo_Kiosko = rd_cod.GetString(rd_cod.GetOrdinal("Codigo_Kiosko"))
                                config.Codigo_Sucursal = rd_cod.GetInt32(rd_cod.GetOrdinal("Codigo_Sucursal")).ToString()
                                config.ID_Kiosko = rd_cod.GetInt32(rd_cod.GetOrdinal("ID_Kiosko")).ToString()

                                config.Bin1 = rd_cod.GetString(rd_cod.GetOrdinal("BIN1"))
                                config.Longitud_Tarjeta_Bin1 = rd_cod.GetInt32(rd_cod.GetOrdinal("LONGITUD_TARJETA_BIN1"))
                                config.Posini_Bin1 = rd_cod.GetInt32(rd_cod.GetOrdinal("POSINI_BIN1"))
                                config.Longitud_Bin_Bin1 = rd_cod.GetInt32(rd_cod.GetOrdinal("LONGITUD_BIN_BIN1"))

                                config.Bin2 = rd_cod.GetString(rd_cod.GetOrdinal("BIN2"))
                                config.Longitud_Tarjeta_Bin2 = rd_cod.GetInt32(rd_cod.GetOrdinal("LONGITUD_TARJETA_BIN2"))
                                config.Posini_Bin2 = rd_cod.GetInt32(rd_cod.GetOrdinal("POSINI_BIN2"))
                                config.Longitud_Bin_Bin2 = rd_cod.GetInt32(rd_cod.GetOrdinal("LONGITUD_BIN_BIN2"))
                            End If
                        End Using
                    End Using
                Else
                    config = Nothing
                End If
            End If
        End Using
        Return config
    End Function

End Class
