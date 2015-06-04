Imports NCapas.Entity
Imports System.Data.SqlClient
Imports NCapas.Data.DAConexion
Imports NCapas.Utility.Log
Imports System.Data.OracleClient

Public Class DAOfertaCliente
    Inherits Singleton(Of DAOfertaCliente)

    ''' <summary>
    ''' Verifica si el cliente tiene reenganche
    ''' </summary>
    ''' <param name="tipoProducto">Tipo de Oferta</param>
    ''' <param name="tipoDocumento">Tipo de Documento de Identidad del Cliente</param>
    ''' <param name="numeroDocumento">Número de documento del Cliente</param>
    ''' <returns>Retorna 1 si tiene reenganche, caso contrario retorna 0</returns>
    ''' <remarks>Retorna 1 si tiene reenganche, caso contrario retorna 0</remarks>
    Public Function VerificarReenganche(ByVal tipoProducto As String, ByVal tipoDocumento As String, ByVal numeroDocumento As String) As OfertaCliente
        Dim oferta As OfertaCliente = New OfertaCliente
        Dim sMensajeError_SQL As String = ""
        Using oConexion = New OracleConnection(DAConexion.Instancia.Get_CadenaConexionOracle())
            If Not oConexion Is Nothing Then
                oConexion.Open()
                If oConexion.State = ConnectionState.Open Then
                    Using cmd As OracleCommand = oConexion.CreateCommand()
                        cmd.CommandTimeout = 600
                        cmd.CommandType = CommandType.StoredProcedure
                        cmd.CommandText = "PKG_OFERTAS_PRODUCTOS.PRC_VERIFICAR_REENGANCHE_SEF"

                        Dim parametro1 As New OracleParameter
                        Dim parametro2 As New OracleParameter
                        Dim parametro3 As New OracleParameter
                        Dim parametro4 As New OracleParameter

                        parametro1.ParameterName = "I_TIP_PROD"
                        parametro1.OracleType = OracleType.Double
                        parametro1.Direction = ParameterDirection.Input
                        parametro1.Value = CDbl(tipoProducto.Trim)
                        cmd.Parameters.Add(parametro1)

                        parametro2.ParameterName = "I_TIP_DOC"
                        parametro2.OracleType = OracleType.Double
                        parametro2.Direction = ParameterDirection.Input
                        parametro2.Value = CDbl(tipoDocumento.Trim)
                        cmd.Parameters.Add(parametro2)

                        parametro3.ParameterName = "I_NUM_DOC"
                        parametro3.OracleType = OracleType.VarChar
                        parametro3.Size = 15
                        parametro3.Direction = ParameterDirection.Input
                        parametro3.Value = numeroDocumento.Trim
                        cmd.Parameters.Add(parametro3)

                        parametro4.ParameterName = "O_REENGANCHE"
                        parametro4.OracleType = OracleType.Double
                        parametro4.Direction = ParameterDirection.Output
                        cmd.Parameters.Add(parametro4)

                        cmd.ExecuteNonQuery()
                        oferta.Reenganche = cmd.Parameters("O_REENGANCHE").Value.ToString()
                    End Using
                End If
            End If
        End Using
        Return oferta
    End Function

    ''' <summary>
    ''' Verifica si el cliente ha desembolsado su Oferta SEF
    ''' </summary>
    ''' <param name="tipoProducto">Tipo de Oferta</param>
    ''' <param name="tipoDocumento">Tipo de Documento de Identidad del Cliente</param>
    ''' <param name="numeroDocumento">Número de documento del Cliente</param>
    ''' <returns>Retorna 1 si ha desembolsado, caso contrario retorna 0</returns>
    ''' <remarks>Retorna 1 si ha desembolsado, caso contrario retorna 0</remarks>
    Public Function VerificarDesembolso(ByVal tipoProducto As String, ByVal tipoDocumento As String, ByVal numeroDocumento As String) As OfertaCliente
        Dim oferta As OfertaCliente = New OfertaCliente
        Dim sMensajeError_SQL As String = ""
        Using oConexion = New OracleConnection(DAConexion.Instancia.Get_CadenaConexionOracle())
            If Not oConexion Is Nothing Then
                oConexion.Open()
                If oConexion.State = ConnectionState.Open Then
                    Using cmd As OracleCommand = oConexion.CreateCommand()
                        cmd.CommandTimeout = 600
                        cmd.CommandType = CommandType.StoredProcedure
                        cmd.CommandText = "PKG_OFERTAS_PRODUCTOS.PRC_VERIFICAR_DESEMBOLSO_SEF"
                        Dim parametro1 As New OracleParameter
                        Dim parametro2 As New OracleParameter
                        Dim parametro3 As New OracleParameter
                        Dim parametro4 As New OracleParameter

                        parametro1.ParameterName = "I_TIP_PROD"
                        parametro1.OracleType = OracleType.Double
                        parametro1.Direction = ParameterDirection.Input
                        parametro1.Value = CDbl(tipoProducto.Trim)
                        cmd.Parameters.Add(parametro1)

                        parametro2.ParameterName = "I_TIP_DOC"
                        parametro2.OracleType = OracleType.Double
                        parametro2.Direction = ParameterDirection.Input
                        parametro2.Value = CDbl(tipoDocumento.Trim)
                        cmd.Parameters.Add(parametro2)

                        parametro3.ParameterName = "I_NUM_DOC"
                        parametro3.OracleType = OracleType.VarChar
                        parametro3.Size = 15
                        parametro3.Direction = ParameterDirection.Input
                        parametro3.Value = numeroDocumento.Trim
                        cmd.Parameters.Add(parametro3)

                        parametro4.ParameterName = "O_ESTADO"
                        parametro4.OracleType = OracleType.VarChar
                        parametro4.Size = 1
                        parametro4.Direction = ParameterDirection.Output
                        cmd.Parameters.Add(parametro4)

                        cmd.ExecuteNonQuery()
                        oferta.Estado = cmd.Parameters("O_ESTADO").Value.ToString()
                    End Using
                End If
            End If
        End Using
        Return oferta
    End Function

    ''' <summary>
    ''' Actualiza la oferta, con los campos que maneja el simulador. Tales como Monto, Plazo, Tasa y Cuota
    ''' </summary>
    ''' <param name="tipoProducto">Tipo de Oferta</param>
    ''' <param name="tipoDocumento">Tipo de Documento de Identidad del Cliente</param>
    ''' <param name="numeroDocumento">Número de documento del Cliente</param>
    ''' <param name="monto">Monto de la simulación</param>
    ''' <param name="plazoMeses">Plazo en meses de la simulacion</param>
    ''' <param name="tasa">Tasa que calcula la simulación</param>
    ''' <param name="cuota">Cuota que calcula la simulación</param>
    ''' <param name="nombreUsuario">Nombre del Usuario</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ActualizarOfertaPorSilumacion(ByVal tipoProducto As String, ByVal tipoDocumento As String, ByVal numeroDocumento As String, ByVal monto As String, ByVal plazoMeses As String, ByVal tasa As String, ByVal cuota As String, ByVal nombreUsuario As String, ByVal codigoSucursal As String) As OfertaCliente
        Dim oferta As OfertaCliente = New OfertaCliente
        Dim sMensajeError_SQL As String = ""
        Dim sFechaTransaccion As String = DateTime.Now.Day.ToString("00").Trim + DateTime.Now.Month.ToString("00").Trim + DateTime.Now.Year.ToString("0000").Trim
        Dim sHoraTransaccion As String = DateTime.Now.Hour.ToString("00").Trim + DateTime.Now.Minute.ToString("00").Trim + DateTime.Now.Second.ToString("00").Trim

        Try
            Using oConexion = New OracleConnection(DAConexion.Instancia.Get_CadenaConexionOracle())
                If Not oConexion Is Nothing Then
                    oConexion.Open()
                    If oConexion.State = ConnectionState.Open Then
                        Using cmd As OracleCommand = oConexion.CreateCommand()
                            cmd.CommandTimeout = 600
                            cmd.CommandType = CommandType.StoredProcedure
                            cmd.CommandText = "PKG_OFERTAS_PRODUCTOS.PRC_UPDATE_OFERTA_SEF"
                            Dim parametro1 As New OracleParameter
                            Dim parametro2 As New OracleParameter
                            Dim parametro3 As New OracleParameter
                            Dim parametro4 As New OracleParameter
                            Dim parametro5 As New OracleParameter
                            Dim parametro6 As New OracleParameter
                            Dim parametro7 As New OracleParameter
                            Dim parametro8 As New OracleParameter
                            Dim parametro9 As New OracleParameter
                            Dim parametro10 As New OracleParameter
                            Dim parametro11 As New OracleParameter
                            Dim parametro12 As New OracleParameter

                            parametro1.ParameterName = "I_TIP_PROD"
                            parametro1.OracleType = OracleType.Double
                            parametro1.Direction = ParameterDirection.Input
                            parametro1.Value = CDbl(tipoProducto.Trim)
                            cmd.Parameters.Add(parametro1)

                            parametro2.ParameterName = "I_TIP_DOC"
                            parametro2.OracleType = OracleType.Double
                            parametro2.Direction = ParameterDirection.Input
                            parametro2.Value = CDbl(tipoDocumento.Trim)
                            cmd.Parameters.Add(parametro2)

                            parametro3.ParameterName = "I_NUM_DOC"
                            parametro3.OracleType = OracleType.VarChar
                            parametro3.Size = 15
                            parametro3.Direction = ParameterDirection.Input
                            parametro3.Value = numeroDocumento.Trim
                            cmd.Parameters.Add(parametro3)

                            parametro4.ParameterName = "I_DATO_NUM2"
                            parametro4.OracleType = OracleType.Double
                            parametro4.Direction = ParameterDirection.Input
                            parametro4.Value = CDbl(monto.Trim)
                            cmd.Parameters.Add(parametro4)

                            parametro5.ParameterName = "I_DATO_NUM3"
                            parametro5.OracleType = OracleType.Double
                            parametro5.Direction = ParameterDirection.Input
                            parametro5.Value = CDbl(plazoMeses.Trim)
                            cmd.Parameters.Add(parametro5)

                            parametro6.ParameterName = "I_DATO_NUM4"
                            parametro6.OracleType = OracleType.Double
                            parametro6.Direction = ParameterDirection.Input
                            parametro6.Value = CDbl(tasa.Trim)
                            cmd.Parameters.Add(parametro6)

                            parametro7.ParameterName = "I_DATO_NUM5"
                            parametro7.OracleType = OracleType.Double
                            parametro7.Direction = ParameterDirection.Input
                            parametro7.Value = CDbl(cuota.Trim)
                            cmd.Parameters.Add(parametro7)

                            parametro8.ParameterName = "I_FEC_TRX"
                            parametro8.OracleType = OracleType.Double
                            parametro8.Direction = ParameterDirection.Input
                            parametro8.Value = sFechaTransaccion.Trim
                            cmd.Parameters.Add(parametro8)

                            parametro9.ParameterName = "I_HOR_TRX"
                            parametro9.OracleType = OracleType.Double
                            parametro9.Direction = ParameterDirection.Input
                            parametro9.Value = sHoraTransaccion.Trim
                            cmd.Parameters.Add(parametro9)

                            parametro10.ParameterName = "I_NOM_USR"
                            parametro10.OracleType = OracleType.VarChar
                            parametro10.Size = 30
                            parametro10.Direction = ParameterDirection.Input
                            parametro10.Value = nombreUsuario.Trim
                            cmd.Parameters.Add(parametro10)

                            parametro11.ParameterName = "I_COD_SUC"
                            parametro11.OracleType = OracleType.Double
                            parametro11.Direction = ParameterDirection.Input
                            parametro11.Value = codigoSucursal.Trim
                            cmd.Parameters.Add(parametro11)

                            parametro12.ParameterName = "O_OK"
                            parametro12.OracleType = OracleType.Double
                            parametro12.Direction = ParameterDirection.Output
                            cmd.Parameters.Add(parametro12)

                            cmd.ExecuteNonQuery()
                            oferta.Estado = cmd.Parameters("O_OK").Value.ToString()

                            If oferta.Estado = "1" Then 'OK
                                oferta.Estado = ""
                            Else
                                oferta.Estado = "No se actualizaron los valores de la simulación"
                            End If

                        End Using
                    Else
                        oferta.Estado = "No se pudo conectar al servidor de base de datos Oracle."
                    End If
                End If
            End Using
        Catch ex As Exception
            oferta.Estado = ex.Message.Trim()
        End Try
        Return oferta
    End Function

    ''' <summary>
    ''' Obtiene la Oferta por Cliente
    ''' </summary>
    ''' <param name="tipoProducto">Tipo de Oferta</param>
    ''' <param name="tipoDocumento">Tipo de Documento de Identidad del Cliente</param>
    ''' <param name="numeroDocumento">Número de documento del Cliente</param>
    ''' <returns>Retorna el objeto Oferta</returns>
    ''' <remarks></remarks>
    Public Function ObtenerOferta(ByVal tipoProducto As String, ByVal tipoDocumento As String, ByVal numeroDocumento As String) As OfertaCliente
        Dim oferta As OfertaCliente = New OfertaCliente
        Dim sMensajeError_SQL As String = ""

        Try
            Using oConexion = New OracleConnection(DAConexion.Instancia.Get_CadenaConexionOracle())
                If Not oConexion Is Nothing Then
                    oConexion.Open()
                    If oConexion.State = ConnectionState.Open Then
                        Using cmd As OracleCommand = oConexion.CreateCommand()
                            cmd.CommandTimeout = 600
                            cmd.CommandType = CommandType.StoredProcedure
                            cmd.CommandText = "PKG_OFERTAS_PRODUCTOS.PRC_CONS_OFERTA_SEF_INI"
                            Dim parametro1 As New OracleParameter
                            Dim parametro2 As New OracleParameter
                            Dim parametro3 As New OracleParameter
                            Dim parametro4 As New OracleParameter
                            Dim parametro5 As New OracleParameter
                            Dim parametro6 As New OracleParameter
                            Dim parametro7 As New OracleParameter
                            Dim parametro8 As New OracleParameter
                            Dim parametro9 As New OracleParameter
                            Dim parametro10 As New OracleParameter
                            Dim parametro11 As New OracleParameter
                            Dim parametro12 As New OracleParameter
                            Dim parametro13 As New OracleParameter
                            Dim parametro14 As New OracleParameter
                            Dim parametro15 As New OracleParameter
                            Dim parametro16 As New OracleParameter
                            Dim parametro17 As New OracleParameter
                            Dim parametro18 As New OracleParameter
                            Dim parametro19 As New OracleParameter
                            Dim parametro20 As New OracleParameter
                            Dim parametro21 As New OracleParameter

                            parametro1.ParameterName = "I_TIP_PROD"
                            parametro1.OracleType = OracleType.Double
                            parametro1.Direction = ParameterDirection.Input
                            parametro1.Value = CDbl(tipoProducto.Trim)
                            cmd.Parameters.Add(parametro1)

                            parametro2.ParameterName = "I_TIP_DOC"
                            parametro2.OracleType = OracleType.Double
                            parametro2.Direction = ParameterDirection.Input
                            parametro2.Value = CDbl(tipoDocumento.Trim)
                            cmd.Parameters.Add(parametro2)

                            parametro3.ParameterName = "I_NUM_DOC"
                            parametro3.OracleType = OracleType.VarChar
                            parametro3.Size = 15
                            parametro3.Direction = ParameterDirection.Input
                            parametro3.Value = numeroDocumento.Trim
                            cmd.Parameters.Add(parametro3)

                            parametro4.ParameterName = "O_EXISTE"
                            parametro4.OracleType = OracleType.Double
                            parametro4.Direction = ParameterDirection.Output
                            cmd.Parameters.Add(parametro4)

                            parametro5.ParameterName = "O_STIP_PROD"
                            parametro5.OracleType = OracleType.VarChar
                            parametro5.Size = 50
                            parametro5.Direction = ParameterDirection.Output
                            cmd.Parameters.Add(parametro5)

                            parametro6.ParameterName = "O_NOM_CLIE"
                            parametro6.OracleType = OracleType.VarChar
                            parametro6.Size = 80
                            parametro6.Direction = ParameterDirection.Output
                            cmd.Parameters.Add(parametro6)

                            parametro7.ParameterName = "O_NUM_CUENTA"
                            parametro7.OracleType = OracleType.VarChar
                            parametro7.Size = 50
                            parametro7.Direction = ParameterDirection.Output
                            cmd.Parameters.Add(parametro7)

                            parametro8.ParameterName = "O_LINEA_OFE"
                            parametro8.OracleType = OracleType.Double
                            parametro8.Direction = ParameterDirection.Output
                            cmd.Parameters.Add(parametro8)

                            parametro9.ParameterName = "O_PLAZO"
                            parametro9.OracleType = OracleType.Double
                            parametro9.Direction = ParameterDirection.Output
                            cmd.Parameters.Add(parametro9)

                            parametro10.ParameterName = "O_TASA"
                            parametro10.OracleType = OracleType.Double
                            parametro10.Direction = ParameterDirection.Output
                            cmd.Parameters.Add(parametro10)

                            parametro11.ParameterName = "O_CUOTA"
                            parametro11.OracleType = OracleType.Double
                            parametro11.Direction = ParameterDirection.Output
                            cmd.Parameters.Add(parametro11)

                            parametro12.ParameterName = "O_TEM"
                            parametro12.OracleType = OracleType.Double
                            parametro12.Direction = ParameterDirection.Output
                            cmd.Parameters.Add(parametro12)

                            parametro13.ParameterName = "O_TEA"
                            parametro13.OracleType = OracleType.Double
                            parametro13.Direction = ParameterDirection.Output
                            cmd.Parameters.Add(parametro13)

                            parametro14.ParameterName = "O_FEC_INI_VIG"
                            parametro14.OracleType = OracleType.VarChar
                            parametro14.Size = 10
                            parametro14.Direction = ParameterDirection.Output
                            cmd.Parameters.Add(parametro14)

                            parametro15.ParameterName = "O_FEC_FIN_VIG"
                            parametro15.OracleType = OracleType.VarChar
                            parametro15.Size = 10
                            parametro15.Direction = ParameterDirection.Output
                            cmd.Parameters.Add(parametro15)

                            parametro16.ParameterName = "O_LINEA_OFE_INI"
                            parametro16.OracleType = OracleType.Double
                            parametro16.Direction = ParameterDirection.Output
                            cmd.Parameters.Add(parametro16)

                            parametro17.ParameterName = "O_PLAZO_OFE_INI"
                            parametro17.OracleType = OracleType.Double
                            parametro17.Direction = ParameterDirection.Output
                            cmd.Parameters.Add(parametro17)

                            parametro18.ParameterName = "O_TASA_OFE_INI"
                            parametro18.OracleType = OracleType.Double
                            parametro18.Direction = ParameterDirection.Output
                            cmd.Parameters.Add(parametro18)

                            parametro19.ParameterName = "O_CUOTA_OFE_INI"
                            parametro19.OracleType = OracleType.Double
                            parametro19.Direction = ParameterDirection.Output
                            cmd.Parameters.Add(parametro19)

                            parametro20.ParameterName = "O_COD_OFERTA"
                            parametro20.OracleType = OracleType.VarChar
                            parametro20.Size = 30
                            parametro20.Direction = ParameterDirection.Output
                            cmd.Parameters.Add(parametro20)

                            parametro21.ParameterName = "O_OK"
                            parametro21.OracleType = OracleType.Double
                            parametro21.Direction = ParameterDirection.Output
                            cmd.Parameters.Add(parametro21)
                            ErrorLog("Antes de llamar a PKG_OFERTAS_PRODUCTOS.PRC_CONS_OFERTA_SEF_INI")
                            cmd.ExecuteNonQuery()
                            ErrorLog(cmd.Parameters("O_EXISTE").Value.ToString())
                            oferta.Existe = cmd.Parameters("O_EXISTE").Value.ToString()
                            oferta.TipoProducto = cmd.Parameters("O_STIP_PROD").Value.ToString()
                            oferta.NombreCliente = cmd.Parameters("O_NOM_CLIE").Value.ToString()
                            oferta.NumeroCuenta = cmd.Parameters("O_NUM_CUENTA").Value.ToString()
                            oferta.LineaOferta = cmd.Parameters("O_LINEA_OFE").Value.ToString()
                            oferta.Plazo = cmd.Parameters("O_PLAZO").Value.ToString()
                            oferta.Tasa = cmd.Parameters("O_TASA").Value.ToString()
                            oferta.Cuota = cmd.Parameters("O_CUOTA").Value.ToString()
                            oferta.Tem = cmd.Parameters("O_TEM").Value.ToString()
                            oferta.Tea = cmd.Parameters("O_TEA").Value.ToString()
                            oferta.FechaInicialVigencia = cmd.Parameters("O_FEC_INI_VIG").Value.ToString().Trim()
                            oferta.FechaFinVigencia = cmd.Parameters("O_FEC_FIN_VIG").Value.ToString().Trim()
                            oferta.LineaOfertaInicial = cmd.Parameters("O_LINEA_OFE_INI").Value.ToString()
                            oferta.PlazoOfertaInicial = cmd.Parameters("O_PLAZO_OFE_INI").Value.ToString()
                            oferta.TasaOfertaInicial = cmd.Parameters("O_TASA_OFE_INI").Value.ToString()
                            oferta.CuotaOfertaInicial = cmd.Parameters("O_CUOTA_OFE_INI").Value.ToString()
                            oferta.CodigoOferta = cmd.Parameters("O_COD_OFERTA").Value.ToString()
                            oferta.Estado = cmd.Parameters("O_OK").Value.ToString()

                            If oferta.Estado <> "-1" Then 'OK
                                oferta.Estado = ""
                            Else
                                oferta.Estado = "No se pudo realizar la consulta"
                            End If

                        End Using
                    Else
                        oferta.Estado = "No se pudo conectar al servidor de base de datos Oracle."
                    End If
                End If
            End Using
        Catch ex As Exception
            ErrorLog(ex.Message.Trim())
            oferta.Estado = ex.Message.Trim()
        End Try
        Return oferta
    End Function

End Class
