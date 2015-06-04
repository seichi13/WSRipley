Imports NCapas.Entity
Imports System.Data.SqlClient
Imports NCapas.Data.DAConexion
Imports NCapas.Utility.Log

Public Class DASucursal
    Inherits Singleton(Of DASucursal)

    Public Function BuscarSucursalPorCodigo(ByVal sCodigoKiosko As String) As Sucursal
        Dim sucursal As New Sucursal
        'Realizar Conexion a la base de datos
        Dim sMensajeError_SQL As String = ""

        Using oConexion As SqlConnection = DAConexion.Instancia.SQL_ConnectionOpen(DAConexion.Instancia.Get_CadenaConexion(), sMensajeError_SQL)
            If sMensajeError_SQL <> "" Then
                sucursal = Nothing
            Else
                If oConexion.State = ConnectionState.Open Then
                    Dim m_ssql As String = "USP_GET_SUCURSALPORKIOSKO"
                    Using cmd As SqlClient.SqlCommand = oConexion.CreateCommand
                        Dim lTotal As Double = 0
                        cmd.CommandTimeout = 600
                        cmd.CommandType = CommandType.StoredProcedure
                        cmd.CommandText = m_ssql
                        cmd.Parameters.AddWithValue("@COD_KIOSKO", sCodigoKiosko)
                        Using rd_cod As SqlClient.SqlDataReader = cmd.ExecuteReader
                            sucursal = Nothing
                            If rd_cod.Read = True Then
                                sucursal = New Sucursal

                                sucursal.ID = rd_cod.GetInt32(rd_cod.GetOrdinal("ID"))
                                sucursal.Ciudad = rd_cod.GetString(rd_cod.GetOrdinal("CIUDAD"))
                                sucursal.Sucursal = rd_cod.GetString(rd_cod.GetOrdinal("SUCURSAL"))
                                sucursal.IdUbigeo = rd_cod.GetInt32(rd_cod.GetOrdinal("IdUbigeo"))
                                sucursal.EstTienda = rd_cod.GetInt32(rd_cod.GetOrdinal("esttienda"))
                                sucursal.Direccion = rd_cod.GetString(rd_cod.GetOrdinal("direccion"))
                                sucursal.Hini_com1 = rd_cod.GetString(rd_cod.GetOrdinal("hini_com1"))
                                sucursal.Hfin_com1 = rd_cod.GetString(rd_cod.GetOrdinal("hfin_com1"))
                                sucursal.Hini_com2 = rd_cod.GetString(rd_cod.GetOrdinal("hini_com2"))
                                sucursal.Hfin_com2 = rd_cod.GetString(rd_cod.GetOrdinal("hfin_com2"))
                                sucursal.Cod_Sucursal_Banco = rd_cod.GetString(rd_cod.GetOrdinal("cod_sucursal_banco"))
                                If rd_cod.IsDBNull(rd_cod.GetOrdinal("hini_cli")) Then
                                    sucursal.Hini_Cli = ""
                                Else
                                    sucursal.Hini_Cli = rd_cod.GetString(rd_cod.GetOrdinal("hini_cli"))
                                End If
                                If rd_cod.IsDBNull(rd_cod.GetOrdinal("hfin_cli")) Then
                                    sucursal.Hfin_Cli = ""
                                Else
                                    sucursal.Hfin_Cli = rd_cod.GetString(rd_cod.GetOrdinal("hfin_cli"))
                                End If
                            End If
                        End Using
                    End Using
                Else
                    sucursal = Nothing
                End If
            End If
        End Using
        Return sucursal
    End Function

    Public Function BuscarSucursalPorId(ByVal sIDSucursalKiosco As String) As Sucursal
        Dim sucursal As New Sucursal
            'Realizar Conexion a la base de datos
            Dim sMensajeError_SQL As String = ""

            Using oConexion As SqlConnection = DAConexion.Instancia.SQL_ConnectionOpen(DAConexion.Instancia.Get_CadenaConexion(), sMensajeError_SQL)
                If sMensajeError_SQL <> "" Then
                    sucursal = Nothing
                Else
                    If oConexion.State = ConnectionState.Open Then
                        Dim m_ssql As String = "SELECT ID,CIUDAD,SUCURSAL,IdUbigeo,esttienda,direccion,hini_com1,hfin_com1,hini_com2,hfin_com2,cod_sucursal_banco,hini_cli,hfin_cli FROM dbo.KIO_SUCURSAL WHERE ID='" & sIDSucursalKiosco.Trim & "'"
                        Using cmd As SqlClient.SqlCommand = oConexion.CreateCommand
                            Dim lTotal As Double = 0
                            cmd.CommandTimeout = 600
                            cmd.CommandType = CommandType.Text
                            cmd.CommandText = m_ssql
                            Using rd_cod As SqlClient.SqlDataReader = cmd.ExecuteReader
                                sucursal = Nothing
                                If rd_cod.Read = True Then
                                    sucursal = New Sucursal

                                    sucursal.ID = rd_cod.GetInt32(rd_cod.GetOrdinal("ID"))
                                    sucursal.Ciudad = rd_cod.GetString(rd_cod.GetOrdinal("CIUDAD"))
                                    sucursal.Sucursal = rd_cod.GetString(rd_cod.GetOrdinal("SUCURSAL"))
                                    sucursal.IdUbigeo = rd_cod.GetInt32(rd_cod.GetOrdinal("IdUbigeo"))
                                    sucursal.EstTienda = rd_cod.GetInt32(rd_cod.GetOrdinal("esttienda"))
                                    sucursal.Direccion = rd_cod.GetString(rd_cod.GetOrdinal("direccion"))
                                    sucursal.Hini_com1 = rd_cod.GetString(rd_cod.GetOrdinal("hini_com1"))
                                    sucursal.Hfin_com1 = rd_cod.GetString(rd_cod.GetOrdinal("hfin_com1"))
                                    sucursal.Hini_com2 = rd_cod.GetString(rd_cod.GetOrdinal("hini_com2"))
                                    sucursal.Hfin_com2 = rd_cod.GetString(rd_cod.GetOrdinal("hfin_com2"))
                                    sucursal.Cod_Sucursal_Banco = rd_cod.GetString(rd_cod.GetOrdinal("cod_sucursal_banco"))
                                    If rd_cod.IsDBNull(rd_cod.GetOrdinal("hini_cli")) Then
                                        sucursal.Hini_Cli = ""
                                    Else
                                        sucursal.Hini_Cli = rd_cod.GetString(rd_cod.GetOrdinal("hini_cli"))
                                    End If
                                    If rd_cod.IsDBNull(rd_cod.GetOrdinal("hfin_cli")) Then
                                        sucursal.Hfin_Cli = ""
                                    Else
                                        sucursal.Hfin_Cli = rd_cod.GetString(rd_cod.GetOrdinal("hfin_cli"))
                                    End If
                                End If
                            End Using
                        End Using
                    Else
                        sucursal = Nothing
                    End If
                End If
            End Using
        Return sucursal
    End Function

    Public Function ObtenerSucursalBanco(ByVal codigoKiosko As String) As Sucursal
        Dim sucursal As Sucursal = New Sucursal
        Dim sMensajeError_SQL As String = ""
        Using oConexion As SqlConnection = DAConexion.Instancia.SQL_ConnectionOpen(DAConexion.Instancia.Get_CadenaConexion(), sMensajeError_SQL)
            If oConexion.State = ConnectionState.Open Then
                Using cmd As SqlClient.SqlCommand = oConexion.CreateCommand
                    cmd.CommandTimeout = 600
                    cmd.CommandType = CommandType.StoredProcedure
                    cmd.CommandText = "USP_GET_SUCURSALBANCO"
                    cmd.Parameters.AddWithValue("@COD_KIOSKO", codigoKiosko)
                    ErrorLog("Entro al sp USP_GET_SUCURSALBANCO(" & codigoKiosko & ")")
                    Using reader As SqlClient.SqlDataReader = cmd.ExecuteReader
                        If reader.Read = True Then
                            ErrorLog("Si hay registros")
                            sucursal.ID = reader.GetInt32(reader.GetOrdinal("IdKiosko"))
                            ErrorLog("sucursal.ID " & sucursal.ID)
                            sucursal.Banco = reader.GetString(reader.GetOrdinal("Banco"))
                            ErrorLog("sucursal.ID " & sucursal.Banco)
                        End If
                    End Using
                End Using
            End If
        End Using
        Return sucursal
    End Function
End Class
