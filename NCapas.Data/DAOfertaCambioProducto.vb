Imports NCapas.Entity
Imports System.Data.SqlClient
Imports NCapas.Data.DAConexion
Imports NCapas.Utility.Log
Imports System.Data.OracleClient

Public Class DAOfertaCambioProducto
    Inherits Singleton(Of DAOfertaCambioProducto)

    Public Function ObtenerOferta(ByVal contrato As String) As OfertaCambioProducto
        Dim oferta As OfertaCambioProducto = New OfertaCambioProducto

        Try
            Using oConexion = New OracleConnection(DAConexion.Instancia.Get_CadenaConexionOracle())
                If Not oConexion Is Nothing Then
                    oConexion.Open()
                    If oConexion.State = ConnectionState.Open Then
                        Using comando As OracleCommand = oConexion.CreateCommand()
                            comando.CommandTimeout = 600
                            comando.CommandType = CommandType.StoredProcedure
                            comando.CommandText = "PKG_VEX_OFERTAS_CAMPROD.PRC_VEX_SELECT_OFERTA_TICKET"
                            Dim parametro1 As New OracleParameter
                            Dim parametro2 As New OracleParameter

                            parametro1.ParameterName = "I_CONTRATO"
                            parametro1.OracleType = OracleType.VarChar
                            parametro1.Direction = ParameterDirection.Input
                            parametro1.Value = contrato.Trim()
                            comando.Parameters.Add(parametro1)

                            parametro2.ParameterName = "POCUR_OFERTA"
                            parametro2.OracleType = OracleType.Cursor
                            parametro2.Direction = ParameterDirection.Output
                            parametro2.Value = vbNull
                            comando.Parameters.Add(parametro2)
                            ErrorLog("Antes de ejecutar " & comando.CommandText)
                            Using reader As OracleDataReader = comando.ExecuteReader
                                If reader.Read = True Then
                                    ErrorLog("Si hay registro en Oracle Cambio Producto")
                                    oferta.ContratoTarjeta = IIf(IsDBNull(reader.Item("CONTRATO_TAR")), "", reader.Item("CONTRATO_TAR").ToString())
                                    oferta.ContratoSEF = IIf(IsDBNull(reader.Item("CONTRATO_SEF")), "", reader.Item("CONTRATO_SEF").ToString())
                                    oferta.CodCambioTarjeta = IIf(IsDBNull(reader.Item("COD_CAMBIO_TAR")), "", reader.Item("COD_CAMBIO_TAR").ToString())
                                    oferta.CodCambioSEF = IIf(IsDBNull(reader.Item("COD_CAMBIO_SEF")), "", reader.Item("COD_CAMBIO_SEF").ToString())
                                    oferta.TipoDocumento = IIf(IsDBNull(reader.Item("TIPO_DOC")), "", reader.Item("TIPO_DOC").ToString())
                                    oferta.NumeroDocumento = IIf(IsDBNull(reader.Item("NUM_DOC")), "", reader.Item("NUM_DOC").ToString())
                                    oferta.NombreCliente = IIf(IsDBNull(reader.Item("NOM_CLIE")), "", reader.Item("NOM_CLIE").ToString())
                                    oferta.FechaInicioVigencia = IIf(IsDBNull(reader.Item("FEC_INI_VIG")), "", reader.Item("FEC_INI_VIG").ToString())
                                    oferta.FechaFinVigencia = IIf(IsDBNull(reader.Item("FEC_FIN_VIG")), "", reader.Item("FEC_FIN_VIG").ToString())
                                    oferta.Estado = IIf(IsDBNull(reader.Item("ESTADO")), "", reader.Item("ESTADO").ToString())
                                    oferta.CodSucursalBanco = IIf(IsDBNull(reader.Item("COD_SUC_TCK")), "0", reader.Item("COD_SUC_TCK").ToString())
                                    oferta.NumeroCaja = IIf(IsDBNull(reader.Item("NUM_CAJ_TCK")), "", reader.Item("NUM_CAJ_TCK").ToString())
                                    oferta.FechaImpresion = IIf(IsDBNull(reader.Item("FEC_TCK")), "", reader.Item("FEC_TCK").ToString())
                                    oferta.DatosTarjeta = IIf(IsDBNull(reader.Item("DATOS_TAR")), "", reader.Item("DATOS_TAR").ToString())
                                    oferta.DatosSEF = IIf(IsDBNull(reader.Item("DATOS_SEF")), "", reader.Item("DATOS_SEF").ToString())
                                    oferta.OfertaVigente = IIf(IsDBNull(reader.Item("OFERTA_VIGENTE")), "", reader.Item("OFERTA_VIGENTE").ToString())
                                End If
                            End Using

                        End Using
                    Else
                        oferta.MensajeError = "No se pudo conectar al servidor de base de datos Oracle."
                    End If
                End If
            End Using
        Catch ex As Exception
            ErrorLog(ex.Message.Trim())
            oferta.MensajeError = ex.Message.Trim()
        End Try
        Return oferta
    End Function

    Public Function InsertarLogIncidenciaCambioProducto(ByVal contrato As String, _
                                                    ByVal codOficina As String, _
                                                    ByVal codUsuario As String, _
                                                    ByVal sistema As String, _
                                                    ByVal subSistema As String, _
                                                    ByVal codRespuesta As String, _
                                                    ByVal desRespuesta As String) As String

        Dim estado As String = "0"

        Try
            Using oConexion = New OracleConnection(DAConexion.Instancia.Get_CadenaConexionOracle())
                If Not oConexion Is Nothing Then
                    oConexion.Open()
                    If oConexion.State = ConnectionState.Open Then
                        Using comando As OracleCommand = oConexion.CreateCommand()
                            comando.CommandTimeout = 600
                            comando.CommandType = CommandType.StoredProcedure
                            comando.CommandText = "PKG_VEX_OFERTAS_CAMPROD.PRC_VEX_REGISTRA_OFERTA_LOG"
                            Dim parametro1 As New OracleParameter
                            Dim parametro2 As New OracleParameter
                            Dim parametro3 As New OracleParameter
                            Dim parametro4 As New OracleParameter
                            Dim parametro5 As New OracleParameter
                            Dim parametro6 As New OracleParameter
                            Dim parametro7 As New OracleParameter
                            Dim parametro8 As New OracleParameter

                            parametro1.ParameterName = "I_CONTRATO_TAR"
                            parametro1.OracleType = OracleType.VarChar
                            parametro1.Direction = ParameterDirection.Input
                            parametro1.Value = contrato.Trim()
                            comando.Parameters.Add(parametro1)

                            parametro2.ParameterName = "I_COD_SUC"
                            parametro2.OracleType = OracleType.Number
                            parametro2.Direction = ParameterDirection.Input
                            parametro2.Value = Convert.ToInt32(codOficina.Trim())
                            comando.Parameters.Add(parametro2)

                            parametro3.ParameterName = "I_TERMINAL"
                            parametro3.OracleType = OracleType.VarChar
                            parametro3.Direction = ParameterDirection.Input
                            parametro3.Value = codUsuario.Trim()
                            comando.Parameters.Add(parametro3)

                            parametro4.ParameterName = "I_SISTEMA"
                            parametro4.OracleType = OracleType.VarChar
                            parametro4.Direction = ParameterDirection.Input
                            parametro4.Value = sistema.Trim()
                            comando.Parameters.Add(parametro4)

                            parametro5.ParameterName = "I_SUB_SISTEMA"
                            parametro5.OracleType = OracleType.VarChar
                            parametro5.Direction = ParameterDirection.Input
                            parametro5.Value = subSistema.Trim()
                            comando.Parameters.Add(parametro5)

                            parametro6.ParameterName = "I_COD_RESPUESTA"
                            parametro6.OracleType = OracleType.VarChar
                            parametro6.Direction = ParameterDirection.Input
                            parametro6.Value = codRespuesta.Trim()
                            comando.Parameters.Add(parametro6)

                            parametro7.ParameterName = "I_DES_RESPUESTA"
                            parametro7.OracleType = OracleType.VarChar
                            parametro7.Direction = ParameterDirection.Input
                            parametro7.Value = desRespuesta.Trim()
                            comando.Parameters.Add(parametro7)

                            parametro8.ParameterName = "O_OK"
                            parametro8.OracleType = OracleType.Float
                            parametro8.Direction = ParameterDirection.Output
                            comando.Parameters.Add(parametro8)

                            ErrorLog("Antes de ejecutar " & comando.CommandText)
                            comando.ExecuteNonQuery()
                            estado = comando.Parameters("O_OK").Value.ToString()

                        End Using
                    Else
                        estado = "0"
                    End If
                End If
            End Using
        Catch ex As Exception
            ErrorLog(ex.Message.Trim())
            estado = "0"
        End Try
        Return estado
    End Function

    Public Function ActualizarCambioProducto(ByVal contrato As String, _
                                        ByVal codVendedor As Integer, _
                                        ByVal codSucursal As Integer, _
                                        ByVal numeroCaja As String, _
                                        ByVal numeroTicket As Integer) As String
        ErrorLog("ActualizarCambioProducto(" & contrato & "," & codVendedor & "," & codSucursal & "," & numeroCaja & "," & numeroTicket & ",)")

        Dim estado As String = "0"

        Try
            Using oConexion = New OracleConnection(DAConexion.Instancia.Get_CadenaConexionOracle())
                If Not oConexion Is Nothing Then
                    oConexion.Open()
                    If oConexion.State = ConnectionState.Open Then
                        Using comando As OracleCommand = oConexion.CreateCommand()
                            comando.CommandTimeout = 600
                            comando.CommandType = CommandType.StoredProcedure
                            comando.CommandText = "PKG_VEX_OFERTAS_CAMPROD.PRC_VEX_IMPRIME_OFERTA_TICKET"
                            Dim parametro1 As New OracleParameter
                            Dim parametro2 As New OracleParameter
                            Dim parametro3 As New OracleParameter
                            Dim parametro4 As New OracleParameter
                            Dim parametro5 As New OracleParameter
                            Dim parametro6 As New OracleParameter

                            parametro1.ParameterName = "I_CONTRATO"
                            parametro1.OracleType = OracleType.VarChar
                            parametro1.Direction = ParameterDirection.Input
                            parametro1.Value = contrato.Trim()
                            comando.Parameters.Add(parametro1)

                            parametro2.ParameterName = "I_COD_VENDEDOR"
                            parametro2.OracleType = OracleType.Number
                            parametro2.Direction = ParameterDirection.Input
                            parametro2.Value = codVendedor
                            comando.Parameters.Add(parametro2)

                            parametro3.ParameterName = "I_COD_SUC_TCK"
                            parametro3.OracleType = OracleType.Number
                            parametro3.Direction = ParameterDirection.Input
                            parametro3.Value = codSucursal
                            comando.Parameters.Add(parametro3)

                            parametro4.ParameterName = "I_NUM_CAJ_TCK"
                            parametro4.OracleType = OracleType.VarChar
                            parametro4.Direction = ParameterDirection.Input
                            parametro4.Value = numeroCaja
                            comando.Parameters.Add(parametro4)

                            parametro5.ParameterName = "I_NUM_TCK"
                            parametro5.OracleType = OracleType.Number
                            parametro5.Direction = ParameterDirection.Input
                            parametro5.Value = numeroTicket
                            comando.Parameters.Add(parametro5)

                            parametro6.ParameterName = "O_OK"
                            parametro6.OracleType = OracleType.Float
                            parametro6.Direction = ParameterDirection.Output
                            comando.Parameters.Add(parametro6)

                            ErrorLog("Antes de ejecutar " & comando.CommandText)
                            comando.ExecuteNonQuery()
                            estado = comando.Parameters("O_OK").Value.ToString()

                        End Using
                    Else
                        estado = "0"
                    End If
                End If
            End Using
        Catch ex As Exception
            ErrorLog(ex.Message.Trim())
            estado = "0"
        End Try
        Return estado
    End Function


End Class
