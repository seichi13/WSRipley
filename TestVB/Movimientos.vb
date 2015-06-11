Imports MQCOMLib
Imports System.Configuration
Imports System.Globalization
Imports System.Threading

Public Enum TServidor
    SICRON = 1
    RSAT = 2
End Enum

Public Class Movimiento
    Private _Concepto As String
    Public Property Concepto() As String
        Get
            Return _Concepto
        End Get
        Set(ByVal value As String)
            _Concepto = value
        End Set
    End Property

    Private _Ticket As String
    Public Property Ticket() As String
        Get
            Return _Ticket
        End Get
        Set(ByVal value As String)
            _Ticket = value
        End Set
    End Property

    Private _Fecha_Consumo As String
    Public Property Fecha_Consumo() As String
        Get
            Return _Fecha_Consumo
        End Get
        Set(ByVal value As String)
            _Fecha_Consumo = value
        End Set
    End Property

    Private _Fecha_Proceso As String
    Public Property Fecha_Proceso() As String
        Get
            Return _Fecha_Proceso
        End Get
        Set(ByVal value As String)
            _Fecha_Proceso = value
        End Set
    End Property

    Private _Importe As String
    Public Property Importe() As String
        Get
            Return _Importe
        End Get
        Set(ByVal value As String)
            _Importe = value
        End Set
    End Property

    Private _Sucursal As String
    Public Property Sucursal() As String
        Get
            Return _Sucursal
        End Get
        Set(ByVal value As String)
            _Sucursal = value
        End Set
    End Property

    Private _Codigo_Desc As String
    Public Property Codigo_Desc() As String
        Get
            Return _Codigo_Desc
        End Get
        Set(ByVal value As String)
            _Codigo_Desc = value
        End Set
    End Property
End Class

Public Class Movimientos
    Private Shared _Instancia As Movimientos = New Movimientos

    Private Shared c_TipoFactura As String(,)
    Private _Detalle_Movimientos As New List(Of Movimiento)

    Public Shared ReadOnly Property Instancia() As Movimientos
        Get
            Return _Instancia
        End Get
    End Property

    Private Movimientos

    Public Function MOVIMIENTOS_CLASICA(ByVal sNroTarjeta As String, ByVal sNroCuenta As String, ByVal sDATA_MONITOR_KIOSCO As String, ByVal Servidor As TServidor, ByVal sMigrado As String) As String
        Dim Respuesta As String
        Respuesta = String.Empty

        Select Case Servidor
            'Case TServidor.SICRON
            '    Respuesta = MOVIMIENTOS_CLASICA_SICRON(sNroCuenta, sDATA_MONITOR_KIOSCO)
            Case TServidor.RSAT
                Respuesta = MOVIMIENTOS_CLASICA_RSAT(sNroTarjeta, sNroCuenta, sMigrado)
            Case Else
                Respuesta = "ERROR:Servidor no especificado"
        End Select

        If Respuesta.Substring(0, 5) <> "ERROR" And sDATA_MONITOR_KIOSCO.Length > 0 Then
            Dim ADATA_MONITOR As Array
            ADATA_MONITOR = Split(sDATA_MONITOR_KIOSCO, "|\t|", , CompareMethod.Text)

            Dim sNroTarjetax As String = ADATA_MONITOR(1)
            Dim sIDSucursal As String = ADATA_MONITOR(6)
            Dim sIDTerminal As String = ADATA_MONITOR(7)

            'GUADAR LOG CONSULTAS
            'SAVE_LOG_CONSULTAS(sIDSucursal.Trim, sIDTerminal.Trim, "34", FUN_BUSCAR_TIPO_TARJETA(sNroTarjetax.Trim))
        End If

        Return Respuesta
    End Function

    Public Function MOVIMIENTOS_ASOCIADA(ByVal sNroTarjeta As String, ByVal sDATA_MONITOR_KIOSCO As String) As String
        Dim sGetTramaMC_MOV As String = ""
        Dim sParametros_MC As String = ""
        Dim sMensajeErrorUsuario As String = ""
        Dim sXML As String = ""

        'Return "09/01/2014|\t|20:31:10|\t|001398|¿**?|09/01/2014|\t|CARGOS POR PROCESAR|\t||\t||\t|36.80|\n|06/01/2014|\t|WONG SAN MIGUEL T009|\t||\t||\t|33.26|\n|02/01/2014|\t|PAGOS EN TIENDAS RIP|\t||\t||\t|103.00|\n|27/12/2013|\t|BANCO CREDITO/PAGO A|\t||\t||\t|2445.59|\n|26/12/2013|\t|OPENENGLISH.COM|\t||\t||\t|202.50|\n|26/12/2013|\t|FRUTIX|\t||\t||\t|23.70|\n|24/12/2013|\t|RIPLEY MIRAFLORES|\t||\t||\t|252.20|\n|24/12/2013|\t|RIPLEY MIRAFLORES|\t||\t||\t|140.00|\n|19/12/2013|\t|RIPLEY SAN MIGUEL|\t||\t||\t|234.32|\n|16/12/2013|\t|REST JHONNY Y JENNIF|\t||\t||\t|28.00|\n|15/12/2013|\t|RIPLEY LIMA NORTE|\t||\t||\t|254.20|\n|14/12/2013|\t|BOTICAS INKAFARMA|\t||\t||\t|69.63|\n|13/12/2013|\t|RIPLEY SAN MIGUEL|\t||\t||\t|136.88|\n|13/12/2013|\t|RIPLEY|\t||\t||\t|427.83|\n|12/12/2013|\t|PARDOS CHICKEN|\t||\t||\t|53.40|\n|09/12/2013|\t|CLINICA STELLA MARIS|\t||\t||\t|40.00|\n|09/12/2013|\t|WONG SAN MIGUEL T009|\t||\t||\t|61.49|\n|07/12/2013|\t|MONTALVO HAIR PERU|\t||\t||\t|45.00|\n|05/12/2013|\t|REST LA CARAVANA|\t||\t||\t|30.50|\n|01/12/2013|\t|BANCO CREDITO/PAGO A|\t||\t||\t|2794.89"

        Try
            If sNroTarjeta.Trim.Length = 16 Then
                'Dim obSendWAS_ As New WSCONSULTAS_MC_VS.defaultService
                'sParametros_MC = "HQ000001073RMDQCM" & sNroTarjeta.Trim & Strings.StrDup(16, " ") & "0000000000" & "0" & "000000" & "000073" & "@"
                'sGetTramaMC_MOV = obSendWAS_.execute(sParametros_MC.Trim)
                'sGetTramaMC_MOV = "        0000000000SFSCANT0109                                       0000064995254350038841756      00010001000005465174M03 0000000000000000042800PAGO TARJETA RIPLEY DESDE OTR2015-03-312015-03-31700000000000045CCE                         0000000000010000 000000000033 18010PAGO DE SERVICIO INTERNET    2015-03-312015-03-31700000000000044Home Banking Ripley         0000000000010000 000000000046 10300PAGO TARJETA RIPLEY DESDE OTR2015-03-312015-03-31700000000000045                            0000000000000400 000000000033 00000                                                                                            0000000000000000 000000000000 00000                                                                                            0000000000000000 000000000000 00000                                                                                            0000000000000000 000000000000 00000                                                                                            0000000000000000 000000000000 00000"
                sGetTramaMC_MOV = "5254350038841756      00010001000005465174M03 0000000000000000042800PAGO TARJETA RIPLEY DESDE OTR2015-03-312015-03-31700000000000045CCE                         0000000000010000 000000000033 18010PAGO DE SERVICIO INTERNET    2015-03-312015-03-31700000000000044Home Banking Ripley         0000000000010000 000000000046 10300PAGO TARJETA RIPLEY DESDE OTR2015-03-312015-03-31700000000000045                            0000000000000400 000000000033 00000                                                                                            0000000000000000 000000000000 00000                                                                                            0000000000000000 000000000000 00000                                                                                            0000000000000000 000000000000 00000                                                                                            0000000000000000 000000000000 00000"

                'Evaluar Respuesta si es ERROR
                If sGetTramaMC_MOV.Trim.Length > 0 Then
                    If Left(sGetTramaMC_MOV.Trim, 2) = "HE" Then
                        sGetTramaMC_MOV = "ERROR:" & Trim(Mid(sGetTramaMC_MOV, 16, 30)) 'CORTAR MENSAJE DE ERROR
                    Else

                        'VARIABLES PARA MOV
                        Dim sBufferHeaderMovimientos As String = ""
                        Dim sBufferMovimientos As String = ""
                        Dim lLongitudBuffer As Long = 0
                        Dim sNTarjeta As String = ""
                        Dim sCliente As String = ""
                        Dim sTotalRegistros As String = ""
                        Dim lTotalCadenaDetalleMov As Long = 0


                        'lLongitudBuffer = Val(Trim(Mid(sGetTramaMC_MOV, 16, 6))).ToString
                        lLongitudBuffer = 1

                        If lLongitudBuffer > 0 Then
                            'EXTRAER EL ENCABEZADO Y DETALLE DE MOVIMIENTOS
                            'sBufferHeaderMovimientos = Mid(sGetTramaMC_MOV, 22, lLongitudBuffer)
                            sBufferHeaderMovimientos = sGetTramaMC_MOV
                            sNTarjeta = Trim(Mid(sBufferHeaderMovimientos, 1, 16))
                            sTotalRegistros = Trim(Mid(sBufferHeaderMovimientos, 109, 2))

                            'EXTRAER EL DETALLE DE MOVIMIENTOS
                            If sTotalRegistros.Trim = "" Then
                                sTotalRegistros = "0"
                            End If

                            'Cada registro del movimiento tiene 62 caracteres
                            lTotalCadenaDetalleMov = Val(sTotalRegistros.Trim) * 62

                            If lTotalCadenaDetalleMov > 0 Then
                                sBufferMovimientos = Mid(sBufferHeaderMovimientos, 111, lTotalCadenaDetalleMov)

                                'Llamar a la funcion cortar Movimientos para formatear las columnas de mov

                                sXML = Cortar_Movimientos_Asociada(sBufferMovimientos, Val(sTotalRegistros))

                                If sXML.Trim.Length > 0 Then

                                    Dim sFechaKiosco As String = DateTime.Now.Day.ToString("00").Trim + "/" + DateTime.Now.Month.ToString("00").Trim + "/" + DateTime.Now.Year.ToString("0000").Trim
                                    Dim sHoraKiosco As String = DateTime.Now.ToString("HH:mm:ss")
                                    Dim sNOperacion As String = GET_NUMERO_OPERACION()

                                    Dim DAT_NUM_OPERACION As String = sFechaKiosco & "|\t|" & sHoraKiosco & "|\t|" & sNOperacion
                                    Dim sDataXML_MOV As String = Left(sXML, sXML.Trim.Length - 4)


                                    sXML = ""
                                    sXML = DAT_NUM_OPERACION & "|¿**?|" & sDataXML_MOV

                                Else
                                    sXML = "ERROR:NODATA"
                                End If


                            Else
                                sXML = "ERROR:NODATA"
                            End If



                        Else
                            sXML = "ERROR:NODATA"
                        End If


                        sGetTramaMC_MOV = sXML

                    End If

                Else 'Sino devuelve nada
                    sGetTramaMC_MOV = "ERROR:CONSULTA NO DISPONIBLE"
                End If


            Else
                'Mostrar Mensaje de Error
                sGetTramaMC_MOV = "ERROR:Tarjeta no valida."
            End If

        Catch ex As Exception
            'save log error

            sGetTramaMC_MOV = "ERROR:" & ex.Message.Trim

        End Try


        'Save Monitor
        Dim sDataMonitor As String = ""
        Dim FECHA_TX As String = DateTime.Now.Year.ToString("0000").Trim & DateTime.Now.Month.ToString("00").Trim + DateTime.Now.Day.ToString("00").Trim
        Dim HORA_TX As String = DateTime.Now.Hour.ToString("00").Trim & DateTime.Now.Minute.ToString("00").Trim & DateTime.Now.Second.ToString("00").Trim
        Dim sNroCuentax As String = ""
        Dim sNroTarjetax As String = ""
        Dim sNombreCliente As String = ""
        Dim sModoEntrada As String = ""
        Dim sSaldoDisponible As String = ""
        Dim sTotalDeuda As String = ""
        Dim sIDSucursal As String = ""
        Dim sIDTerminal As String = ""
        Dim sCodigoTransaccion As String = ""

        Dim sRespuestaServidor As String = ""
        Dim sCanalAtencion As String = "01" 'RipleyMatico
        Dim sEstadoCuenta As String = ""

        If sDATA_MONITOR_KIOSCO.Length > 0 Then
            'RECORRER ARREGLO
            Dim ADATA_MONITOR As Array
            ADATA_MONITOR = Split(sDATA_MONITOR_KIOSCO, "|\t|", , CompareMethod.Text)


            sNroCuentax = ADATA_MONITOR(0)
            sNroTarjetax = ADATA_MONITOR(1)
            sNombreCliente = ADATA_MONITOR(2)
            sModoEntrada = ADATA_MONITOR(3)
            sSaldoDisponible = ADATA_MONITOR(4)
            sTotalDeuda = ADATA_MONITOR(5)
            sIDSucursal = ADATA_MONITOR(6)
            sIDTerminal = ADATA_MONITOR(7)
            sCodigoTransaccion = ADATA_MONITOR(8)

            'GUADAR LOG CONSULTAS
            'SAVE_LOG_CONSULTAS(sIDSucursal.Trim, sIDTerminal.Trim, "34", FUN_BUSCAR_TIPO_TARJETA(sNroTarjetax.Trim))

            sNroCuentax = Right("          " & sNroCuentax.Trim, 10)
            sNroTarjetax = Right("                    " & sNroTarjetax.Trim, 20)

            sNombreCliente = Right(Strings.StrDup(26, " ") & sNombreCliente.Trim, 26)

            sSaldoDisponible = Right("      " & sSaldoDisponible.Trim, 6)
            sTotalDeuda = Right("      " & sTotalDeuda.Trim, 6)


            'armar variable desde flash con: numero de cuenta, numero de tarjeta, nombre del cliente, modo de entrada, canal de atencion, id de sucursal,codigo kiosco y codigo de transaccion (opcion)
            If Left(sGetTramaMC_MOV.Trim, 5) <> "ERROR" Then
                'VALIDACION CORRECTA

                sEstadoCuenta = "01" 'Estado de la cuenta
                sRespuestaServidor = "01" 'Atendido

                'ENVIAR_MONITOR
                sDataMonitor = ""
                sDataMonitor = FECHA_TX & HORA_TX & sNroCuentax & sNroTarjetax & sNombreCliente & sModoEntrada & sCanalAtencion & sSaldoDisponible & sTotalDeuda & sIDSucursal & sIDTerminal & sEstadoCuenta & sRespuestaServidor & sCodigoTransaccion

                'ENVIAR_MONITOR(sDataMonitor)

            Else
                'ENVIAR_MONITOR

                sEstadoCuenta = "01" 'Estado de la cuenta
                sRespuestaServidor = "02" 'Rechazado

                sDataMonitor = ""
                sDataMonitor = FECHA_TX & HORA_TX & sNroCuentax & sNroTarjetax & sNombreCliente & sModoEntrada & sCanalAtencion & sSaldoDisponible & sTotalDeuda & sIDSucursal & sIDTerminal & sEstadoCuenta & sRespuestaServidor & sCodigoTransaccion

                'ENVIAR_MONITOR(sDataMonitor)
            End If
        End If

        Return sGetTramaMC_MOV
    End Function

    Private Function MOVIMIENTOS_CLASICA_RSAT(ByVal nrotarjeta As String, ByVal nrocuenta As String, ByVal sMigrado As String) As String
        'Dim objMQ As New MQ
        Dim strServicio As String = String.Empty
        Dim strMensaje As String = String.Empty
        Dim strParametros As String = String.Empty
        Dim strRespuesta As String = String.Empty
        Dim strPan As String = String.Empty
        Dim strNroPan As String = String.Empty
        Dim lngLargo As Long = 0
        Dim i As Integer = 0
        Dim g_Posicion As Integer = 0

        nrocuenta = nrocuenta.Trim
        nrotarjeta = nrotarjeta.Trim

        If (nrotarjeta.Substring(0, 6) = "542020" Or nrotarjeta.Substring(0, 6) = "525474" Or _
            nrotarjeta.Substring(0, 6) = "450034" Or nrotarjeta.Substring(0, 6) = "450035") And sMigrado = "SI" Then
            nrocuenta = "000000000000"
        End If


        'strServicio = "SFSCANT0109"
        'strMensaje = "0000000000SFSCANT0109                                       000000043" + nrotarjeta + "      00010001" + nrocuenta + "M"
        'strServicio = ReadAppConfig("SFSCAN_MOVIMIENTOS_CLASICA_RSAT")
        'strMensaje = "0000000000" & strServicio & "                                       000000043" + nrotarjeta + "      00010001" + nrocuenta + "M"
        'objMQ.Service = strServicio
        'objMQ.Message = strMensaje
        'objMQ.Execute()

        'If objMQ.ReasonMd <> 0 Then
        '    strRespuesta = "ERROR"
        'End If
        'If objMQ.ReasonApp = 0 Then
        '    strRespuesta = objMQ.Response
        'End If
        strRespuesta = "0000000000SFSCANT0109                                       0000064995254350038841756      00010001000005465174M03 0000000000000000042800PAGO TARJETA RIPLEY DESDE OTR2015-03-312015-03-31700000000000045CCE                         0000000000010000 000000000033 18010PAGO DE SERVICIO INTERNET    2015-03-312015-03-31700000000000044Home Banking Ripley         0000000000010000 000000000046 10300PAGO TARJETA RIPLEY DESDE OTR2015-03-312015-03-31700000000000045                            0000000000000400 000000000033 00000                                                                                            0000000000000000 000000000000 00000                                                                                            0000000000000000 000000000000 00000                                                                                            0000000000000000 000000000000 00000                                                                                            0000000000000000 000000000000 00000"

        'CAMPOS DE MOVIMIENTOS
        Dim Fecha_mov As String = ""
        Dim Descripcion_Mov As String = ""
        Dim Establecimiento As String = ""
        Dim Ticket As String = ""
        Dim Monto As String = ""
        Dim sDataMovEstablec As String = ""
        Dim sCodLinea As String = ""
        Dim sCodArticulo As String = ""
        Dim sCodigoEstablece As String = ""

        'Data Retornar Movimientos
        Dim sDATA_MOV_RETORNO As String = ""
        Dim g_TotMovimientos, g_TotRevol As String

        g_TotMovimientos = 0

        If Val(Mid(strRespuesta, 1, 10)) = 0 Then
            g_TotMovimientos = Val(Mid(strRespuesta, 113, 2))                               'CANTIDAD COMPRAS
            g_TotRevol = Val(Mid(strRespuesta, 115, 18))                                    'TOTAL REV


            'If CInt(g_TotMovimientos) > 15 Then
            '    g_TotMovimientos = 15
            'End If

            Dim TFacturaMOV As String

            g_Posicion = 115 + 18
            Dim x As Integer
            Dim oMovimiento As Movimiento
            Dim Cod_RSAT As String
            Dim Descripcion As String
            Dim curCulture As CultureInfo = Thread.CurrentThread.CurrentCulture
            Dim tInfo As TextInfo = curCulture.TextInfo()

            For x = 1 To g_TotMovimientos
                oMovimiento = New Movimiento
                TFacturaMOV = Trim(g_ExtraeData(strRespuesta, 4, g_Posicion))

                If Not ValidarTipFacturaMOV(TFacturaMOV, c_TipoFactura) Then
                    'Descripcion_Mov = Trim(g_ExtraeData(strRespuesta, 30, g_Posicion))
                    oMovimiento.Concepto = Trim(g_ExtraeData(strRespuesta, 30, g_Posicion)).ToLower()
                    Cod_RSAT = TFacturaMOV & oMovimiento.Concepto.Substring(0, 1)
                    'Descripcion = Buscar_Desc_Mov_RSAT(Cod_RSAT)
                    Descripcion = "ERROR"
                    oMovimiento.Concepto = IIf(Descripcion.Substring(0, 5) = "ERROR", tInfo.ToTitleCase(oMovimiento.Concepto.Substring(1)), Descripcion)
                    'Fecha_mov = Trim(g_ExtraeData(strRespuesta, 10, g_Posicion))
                    oMovimiento.Fecha_Consumo = Trim(g_ExtraeData(strRespuesta, 10, g_Posicion))
                    g_Posicion = g_Posicion + 25

                    'Establecimiento = Trim(g_ExtraeData(strRespuesta, 27, g_Posicion))
                    oMovimiento.Sucursal = Trim(g_ExtraeData(strRespuesta, 27, g_Posicion)).ToLower
                    oMovimiento.Sucursal = tInfo.ToTitleCase(oMovimiento.Sucursal)
                    'Monto = Format(CDbl(Trim(g_ExtraeData(strRespuesta, 17, g_Posicion))) / 100, "###,#0.00")
                    oMovimiento.Importe = Format(CDbl(Trim(g_ExtraeData(strRespuesta, 17, g_Posicion))) / 100, "###,#0.00")
                    g_Posicion = g_Posicion + 1
                    'Ticket = Trim(g_ExtraeData(strRespuesta, 12, g_Posicion))
                    oMovimiento.Ticket = Trim(g_ExtraeData(strRespuesta, 12, g_Posicion))
                    g_Posicion = g_Posicion + 1
                    'sDATA_MOV_RETORNO = sDATA_MOV_RETORNO & Fecha_mov.Trim & "|\t|" & Descripcion_Mov.Trim & "|\t|" & Establecimiento.Trim & "|\t|" & Ticket.Trim & "|\t|" & Monto.Trim & "|\n|"

                    If oMovimiento.Fecha_Consumo.Length > 0 Then
                        Agregar_Movimiento(oMovimiento)
                    End If
                Else
                    g_ExtraeData(strRespuesta, 30, g_Posicion)
                    g_ExtraeData(strRespuesta, 10, g_Posicion)
                    g_Posicion = g_Posicion + 25
                    g_ExtraeData(strRespuesta, 27, g_Posicion)
                    g_ExtraeData(strRespuesta, 17, g_Posicion)
                    g_Posicion = g_Posicion + 1
                    g_ExtraeData(strRespuesta, 12, g_Posicion)
                    g_Posicion = g_Posicion + 1
                End If
            Next

            If g_TotMovimientos > 0 Then
                _Detalle_Movimientos = OrdenaMovimientosxfecha(_Detalle_Movimientos)
            End If

            Dim oResultas As New List(Of Movimiento)

            If CInt(_Detalle_Movimientos.Count) > 15 Then
                g_TotMovimientos = 15
            Else
                g_TotMovimientos = _Detalle_Movimientos.Count
            End If

            If g_TotMovimientos > 0 Then
                For x = 0 To g_TotMovimientos - 1
                    oResultas.Add(_Detalle_Movimientos(x))
                Next
            End If

            Dim maxDESC As Integer = 0
            Dim maxSUC As Integer = 0
            Dim Concepto As String = String.Empty
            Dim Sucursal As String = String.Empty


            If g_TotMovimientos < 15 And sMigrado = "SI" Then

                Dim sTramaMov_MC As String
                Dim iNroMov_MC As Integer
                iNroMov_MC = 15 - g_TotMovimientos
                sTramaMov_MC = MOVIMIENTOS_ASOCIADA(nrotarjeta, "")

                Dim aTrama1 As Array
                Dim aTrama2 As Array
                Dim aTrama3 As Array

                aTrama1 = Split(sTramaMov_MC, "|¿**?|", , CompareMethod.Text)
                aTrama2 = Split(aTrama1(1), "|\n|", , CompareMethod.Text)

                If aTrama2.Length < iNroMov_MC Then
                    iNroMov_MC = aTrama2.Length
                End If
                g_TotMovimientos = g_TotMovimientos + iNroMov_MC

                For z As Integer = (iNroMov_MC - 1) To 0 Step -1

                    aTrama3 = Split(aTrama2(z), "|\t|", , CompareMethod.Text)
                    sDATA_MOV_RETORNO = sDATA_MOV_RETORNO & aTrama3(0) & "|\t|" & aTrama3(1) & "|\t|" & aTrama3(2) & "|\t|" & aTrama3(3) & "|\t|" & aTrama3(4) & "|\n|"
                Next

            End If

            For s As Integer = oResultas.Count - 1 To 0 Step -1
                maxDESC = oResultas(s).Concepto.Length
                maxSUC = oResultas(s).Sucursal.Length

                If maxDESC > 0 Then
                    Concepto = oResultas(s).Concepto.Substring(0, IIf(maxDESC > 17, 17, maxDESC))
                Else
                    Concepto = ""
                End If

                Sucursal = oResultas(s).Sucursal

                If sMigrado = "SI" Then
                    oResultas(s).Ticket = "(POR PROCESAR)"
                End If

                sDATA_MOV_RETORNO = sDATA_MOV_RETORNO & oResultas(s).Fecha_Consumo & "|\t|" _
                & Concepto & "|\t|" _
                & Sucursal & "|\t|" _
                & oResultas(s).Ticket & "|\t|" & _
                oResultas(s).Importe & "|\n|"
            Next

            Dim sFechaKiosco As String = DateTime.Now.Day.ToString("00").Trim + "/" + DateTime.Now.Month.ToString("00").Trim + "/" + DateTime.Now.Year.ToString("0000").Trim
            Dim sHoraKiosco As String = DateTime.Now.ToString("HH:mm:ss")
            Dim sNOperacion As String = GET_NUMERO_OPERACION()

            Dim DAT_NUM_OPERACION As String = sFechaKiosco & "|\t|" & sHoraKiosco & "|\t|" & sNOperacion
            Dim sDataOrdenada As String = String.Empty

            If g_TotMovimientos > 0 Then
                'CALL FUNCION PARA ODERNAR LA DATA
                sDataOrdenada = Left(sDATA_MOV_RETORNO, sDATA_MOV_RETORNO.Trim.Length - 4)
                sDataOrdenada = fun_OrdenarMovimientos(sDataOrdenada)
                strRespuesta = DAT_NUM_OPERACION & "|¿**?|" & sDataOrdenada
            Else
                strRespuesta = "ERROR:NODATA"
            End If
        End If

        Return strRespuesta
    End Function

    Private Function Cortar_Movimientos_Asociada(ByVal sDataTramaMOV As String, ByVal lTotalRegistros As Long) As String
        'CAMPOS DE MOVIMIENTOS
        Dim Fecha_mov_mc As String = ""
        Dim Descripcion_Mov_Estable_mc As String = ""
        Dim sMonto_mc As String = ""
        Dim dMonto_mc As Double = 0

        'Data Retornar Movimientos
        Dim sDATA_MOV_RETORNO_MC As String = ""
        Dim lFila As Long = 0
        Dim lLongitudRegistro As Long = 62
        Dim lCon1 As Long = 1 'Pos Fecha
        Dim lCon2 As Long = 9 'Pos Descripcion
        Dim lCon3 As Long = 49 'Pos Monto

        For lFila = 1 To lTotalRegistros
            'Limpiar Variables.
            Fecha_mov_mc = ""
            Descripcion_Mov_Estable_mc = ""
            sMonto_mc = ""

            'LLEGA YYYYMMDD CONVERTIRLO A DD/MM/YYYY
            Fecha_mov_mc = Trim(Mid(sDataTramaMOV, lCon1, 8))

            Descripcion_Mov_Estable_mc = Trim(Mid(sDataTramaMOV, lCon2, 20))

            sMonto_mc = Trim(Mid(sDataTramaMOV, lCon3, 14))

            lCon1 = lCon1 + 62
            lCon2 = lCon2 + 62
            lCon3 = lCon3 + 62

            'SI TIENE FECHA DE MOV
            If Fecha_mov_mc.Trim <> "00000000" Then
                'FORMATEAR FECHA DE 20111105 A 05/11/2011
                Fecha_mov_mc = Mid(Fecha_mov_mc, 7, 2) & "/" & Mid(Fecha_mov_mc, 5, 2) & "/" & Mid(Fecha_mov_mc, 1, 4) 'DD/MM/YYYY

                sMonto_mc = CLng(sMonto_mc).ToString
                sMonto_mc = Format(Val(SET_MONTO_DECIMAL(sMonto_mc)), "######0.00")

                sDATA_MOV_RETORNO_MC = sDATA_MOV_RETORNO_MC & Fecha_mov_mc.Trim & "|\t|" & Descripcion_Mov_Estable_mc.Trim & "|\t||\t||\t|" & sMonto_mc.Trim & "|\n|"
            End If
        Next
        Return sDATA_MOV_RETORNO_MC.Trim
    End Function

    Private Function GET_NUMERO_OPERACION() As String
        Return "001346"
        'Try
        '    'Realizar Conexion a la base de datos
        '    sMensajeError_SQL = ""

        '    oConexion = SQL_ConnectionOpen(Get_CadenaConexion(), sMensajeError_SQL)

        '    If sMensajeError_SQL <> "" Then
        '        sResultado_SQL = ""
        '    Else

        '        If oConexion.State = ConnectionState.Open Then

        '            m_ssql = "SP_NUM_OPERACION"

        '            Dim cmd As SqlClient.SqlCommand = oConexion.CreateCommand
        '            Dim lTotal As Double = 0

        '            cmd.CommandTimeout = 900
        '            cmd.CommandType = CommandType.StoredProcedure
        '            cmd.CommandText = m_ssql



        '            Dim rd_GET As SqlClient.SqlDataReader = cmd.ExecuteReader
        '            sXML_SQL = ""

        '            If rd_GET.Read = True Then
        '                sXML_SQL = rd_GET.GetValue(0)
        '            End If

        '            sResultado_SQL = sXML_SQL

        '            rd_GET.Close()
        '            rd_GET = Nothing
        '            cmd.Dispose()
        '            cmd = Nothing
        '            oConexion.Close()
        '            oConexion = Nothing
        '        Else
        '            sResultado_SQL = ""
        '            oConexion = Nothing
        '        End If
        '    End If


        'Catch ex As Exception
        '    sResultado_SQL = ""
        '    oConexion = Nothing
        'End Try

        'Return sResultado_SQL
    End Function

    Private Function SET_MONTO_DECIMAL(ByVal sMonto As String) As String
        Dim sData1 As String = ""
        Dim sData2 As String = ""
        Dim sDecimal As String = ""
        Dim sRepFinal As String = ""

        If Left(sMonto.Trim, 1) = "0" Then
            sRepFinal = "0.00"
        Else
            If sMonto.Trim.Length = 1 Or sMonto.Trim.Length = 2 Then
                sMonto = Right("00" & sMonto.Trim, 3)
            End If

            sData1 = sMonto.Trim
            'Extraer el monto sin considerar los dos ultimos valores
            sData2 = Left(sData1.Trim, sData1.Trim.Length - 2)
            'Extraer por la derecha dos caracteres
            sDecimal = Right(sData1.Trim, 2)

            sRepFinal = sData2 & "." & sDecimal
        End If

        Return sRepFinal
    End Function

    Private Function ValidarTipFacturaMOV(ByVal sCodigoTipFac As String, ByVal vg_ArrTipFacError As String(,)) As Boolean
        Dim ind As Byte
        Dim bTFEncontrada As Boolean
        bTFEncontrada = False
        If Not vg_ArrTipFacError Is Nothing Then
            If vg_ArrTipFacError.Length > 0 Then
                For ind = 0 To vg_ArrTipFacError.GetLength(0) - 1
                    If sCodigoTipFac = Mid(vg_ArrTipFacError(ind, 0), 5, 4) Then
                        bTFEncontrada = True
                        Exit For
                    End If
                Next
            End If
        End If

        Return bTFEncontrada
    End Function

    Private Function g_ExtraeData(ByVal p_data As String, ByVal p_Longitud As Long, ByRef g_Posicion As Long) As String
        g_ExtraeData = Mid(p_data, g_Posicion, p_Longitud)
        g_Posicion = g_Posicion + p_Longitud
        Return g_ExtraeData
    End Function

    Private Sub Agregar_Movimiento(ByVal oMovimiento As Movimiento)
        _Detalle_Movimientos.Add(oMovimiento)
    End Sub

    Public Function OrdenaMovimientosxfecha(ByVal g_ArrMovSAT As List(Of Movimiento)) As List(Of Movimiento)
        Dim I As Integer
        Dim j As Integer
        Dim strFecha_Consumo As String
        Dim strFecha_Proceso As String
        Dim strConcepto As String
        Dim strSucursal As String
        Dim strImporte As String
        Dim strTicket As String

        Dim Fecha1 As Integer
        Dim Fecha2 As Integer

        Dim g_TotMovimientos As Integer = g_ArrMovSAT.Count

        For I = 0 To g_TotMovimientos - 2
            For j = I + 1 To g_TotMovimientos - 1
                If g_ArrMovSAT(I).Fecha_Consumo = "  /  /  " Then
                    g_ArrMovSAT(I).Fecha_Consumo = "01/01/01"
                End If
                If g_ArrMovSAT(j).Fecha_Consumo = "  /  /  " Then
                    g_ArrMovSAT(j).Fecha_Consumo = "01/01/01"
                End If

                Try
                    Fecha1 = CInt(g_ArrMovSAT(I).Fecha_Consumo.Replace("-", ""))
                Catch ex As Exception
                    Fecha1 = 0
                End Try

                Try
                    Fecha2 = CInt(g_ArrMovSAT(j).Fecha_Consumo.Replace("-", ""))
                Catch ex As Exception
                    Fecha2 = 0
                End Try

                If Fecha1 < Fecha2 And (Fecha1 <> 0 And Fecha2 <> 0) Then
                    strFecha_Consumo = g_ArrMovSAT(I).Fecha_Consumo
                    strFecha_Proceso = g_ArrMovSAT(I).Fecha_Proceso
                    strConcepto = g_ArrMovSAT(I).Concepto
                    strImporte = g_ArrMovSAT(I).Importe
                    strSucursal = g_ArrMovSAT(I).Sucursal
                    strTicket = g_ArrMovSAT(I).Ticket

                    g_ArrMovSAT(I).Fecha_Consumo = g_ArrMovSAT(j).Fecha_Consumo
                    g_ArrMovSAT(I).Fecha_Proceso = g_ArrMovSAT(j).Fecha_Proceso
                    g_ArrMovSAT(I).Concepto = g_ArrMovSAT(j).Concepto
                    g_ArrMovSAT(I).Importe = g_ArrMovSAT(j).Importe
                    g_ArrMovSAT(I).Sucursal = g_ArrMovSAT(j).Sucursal
                    g_ArrMovSAT(I).Ticket = g_ArrMovSAT(j).Ticket

                    g_ArrMovSAT(j).Fecha_Consumo = strFecha_Consumo
                    g_ArrMovSAT(j).Fecha_Proceso = strFecha_Proceso
                    g_ArrMovSAT(j).Concepto = strConcepto
                    g_ArrMovSAT(j).Importe = strImporte
                    g_ArrMovSAT(j).Sucursal = strSucursal
                    g_ArrMovSAT(j).Ticket = strTicket

                End If

            Next j
        Next I

        For x As Integer = 0 To g_ArrMovSAT.Count - 1

            If g_ArrMovSAT(x).Fecha_Consumo.Length > 0 Then
                g_ArrMovSAT(x).Fecha_Consumo = IIf(g_ArrMovSAT(x).Fecha_Consumo.Length = 10, g_ArrMovSAT(x).Fecha_Consumo.Substring(8, 2) & "/" & g_ArrMovSAT(x).Fecha_Consumo.Substring(5, 2) & "/" & g_ArrMovSAT(x).Fecha_Consumo.Substring(0, 4), "")
            End If
        Next

        Return g_ArrMovSAT
    End Function

    Private Function fun_OrdenarMovimientos(ByVal sDataMovimientos As String) As String

        Dim sNUEVA_DATA_MOV As String = ""

        Try
            If sDataMovimientos.Trim.Length > 0 Then

                Dim ADATA_MOV As Array
                Dim ADATA_REGISTROS As Array
                Dim SDATA_REGISTROS As String = ""
                Dim nRegistros As Long = 0
                Dim lIndice As Long = 0


                ADATA_MOV = Split(sDataMovimientos, "|\n|", , CompareMethod.Text)

                lIndice = ADATA_MOV.Length - 1
                For lIndice = ADATA_MOV.Length - 1 To 0 Step -1
                    SDATA_REGISTROS = ADATA_MOV(lIndice)
                    ADATA_REGISTROS = Split(SDATA_REGISTROS, "|\t|", , CompareMethod.Text)  'Separar en columnas de array

                    sNUEVA_DATA_MOV = sNUEVA_DATA_MOV & ADATA_REGISTROS(0) & "|\t|" & ADATA_REGISTROS(1) & "|\t|" & ADATA_REGISTROS(2) & "|\t|" & ADATA_REGISTROS(3) & "|\t|" & ADATA_REGISTROS(4) & "|\n|"
                Next

                If sNUEVA_DATA_MOV.Trim.Length > 0 Then
                    sNUEVA_DATA_MOV = Left(sNUEVA_DATA_MOV, sNUEVA_DATA_MOV.Trim.Length - 4)
                End If

            End If

        Catch ex As Exception
            sNUEVA_DATA_MOV = ""
        End Try

        Return sNUEVA_DATA_MOV

    End Function

    Private Function ReadAppConfig(ByVal llave As String) As String
        Dim reader As New AppSettingsReader()
        Return reader.GetValue(llave, Type.GetType("System.String")).ToString
    End Function
End Class
