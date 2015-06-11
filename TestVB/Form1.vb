Imports System
Imports Microsoft.VisualBasic
Imports TestVB.Log
Imports TestVB.Funciones
Imports TestVB.Tipos

Public Class Form1

    Public Shared c_TablaBloqueo As String(,)
    Private TIME_OUT_SERVER As Long = CLng(40) 'TIME OUT
    Private SERVIDOR_MIRROR_DESTINO As String = "BRVMSO" 'destination
    Private SERVIDOR_MIRROR_NODE As String = "BRVMSO" 'Mirror Node
    Private PUERTO As String = "2000" 'PUERTO
    Private PRIORIDAD_S As String = "02" 'PUERTO
    Private parameterOut As New VTAAUTO5001ParameterOUT

    Private Sub btnCalcular_Click(sender As System.Object, e As System.EventArgs) Handles btnCalcular.Click

        Dim fecha As Date = DateTime.Now
        Dim lol23 As String = Format(fecha, "g")

        Dim sRespuesta As String = ""
        'VARIABLES RESUMEN DE ESTADO DE CUENTA
        Dim sLineaCredito As String = ""
        Dim sLineaCreditoTemporal As String = ""
        Dim sLineaSuperEfectivo As String = ""
        Dim sLineaCreditoTotal As String = ""
        Dim sPeriodoFacturacion As String = ""
        Dim sUltimoDiaPago As String = ""
        Dim sCreditoUtilizado As String = ""
        Dim sComisionCargos As String = ""
        Dim sDeudaTotal As String = ""
        Dim sDeudaVencida As String = ""
        Dim sPagoMinimoMes As String = ""
        Dim sPagoTotalMes As String = ""
        Dim sDispCompras24Cuotas As String = ""
        Dim sDispCompras36Cuotas As String = ""
        Dim sDispSuperEfectivo As String = ""
        Dim sDisponibleCompras As String = ""
        Dim sDispEfectivoExpress As String = ""

        'CALCULO DE PAGO TOTAL DEL MES
        Dim sSaldoAnterior As String = ""
        Dim sConsumoMes As String = ""
        Dim sCuotaMes As String = ""
        Dim sInteres As String = ""
        Dim sComisionCargos_Mes As String = ""
        Dim sPagosAbonos As String = ""
        Dim sPagoTotal_Mes As String = ""

        'CALCULO DE PAGO MINIMO DEL MES
        Dim sSaldoFavor As String = ""
        Dim sMontoMinimo As String = ""
        Dim sCuotaMes_Min As String = ""
        Dim sInteres_Min As String = ""
        Dim sComisionCargos_Min As String = ""
        Dim sPagoMinimoMes_Min As String = ""

        'PLAN DE CUOTAS DE LOS PROXIMOS 3 MESES
        Dim sMes1 As String = ""
        Dim sMonto1 As String = ""
        Dim sMes2 As String = ""
        Dim sMonto2 As String = ""
        Dim sMes3 As String = ""
        Dim sMonto3 As String = ""

        Dim sPeriodoFinal As String = ""
        Dim sPeriodo1 As String = ""
        Dim sPeriodo2 As String = ""
        Dim sPeriodo3 As String = ""
        Dim vFechaHoy As Date
        Dim fFechaHoyAux1 As Date
        Dim lEstadoLlamadaEECC As Long = 9 'Variable para iniciar llamadas al EECC 3 veces

        Dim sTrama As String = ""
        Dim sXMLMOV As String = ""
        Dim sXMLMOV_FINAL As String = ""
        Dim sXMLCAB As String = ""
        Dim sXMLPIE As String = ""
        Dim sDataMov As String = ""
        Dim sDataMovAux As String = ""

        Dim lFila As Long = 0
        Dim Incrementa As Long = 1
        Dim tamanioFilaDetalle As Long = 94
        Dim tamanioFilaDetalleAnt As Long = 87

        Dim outpputBuff As String = "        AU            008CABALLERO ZEÑA, DANIEL                            JR.SANTA MARIA 338 PS.04 INT.B                                                                      SAN MARTIN DE PORRES    LIMA            LIMA      52543-500328093-61                   1,500.00      0.00      0.00  1,500.0001/FEB/2015-31/MAR/201525/ABR/2015  2,136.81     58.77  2,195.69    186.16    392.47    392.47      0.00      0.00      0.00      0.00      0.00                                                                392.47                                                                                                                                                                                                                        186.16      5.90    101.48     59.91     39.00      0.00    392.47     11.80    101.48     59.91     39.00                                                            SALDO INICIAL                                               186.1620/FEB/201520/FEB/2015004125RET. EFEC CUOTAS TDAT  1000.05 43.74%02/24   29.10  30.91    60.0121/FEB/201521/FEB/2015004131RET. EFEC CUOTAS TDAT  1000.05 41.25%02/12   72.38  28.10   100.4831/MAR/201531/MAR/2015000100SEGURO DESGRAVAMEN  T     5.90       01/01                    5.90                            INTERES PERIODO FACT                                          0.31                            INTERES COMPENSATORI                                          0.59                            PENALIDAD PAGO VCDO.                                         39.00                                                                                                                                                                                                                                                                                                                                                                                                                                                                                          U    "
        outpputBuff = "        AU            006APARICIO SORIA, RONALD                            AV.CAMINO REAL 348 PISO 12                                                                          1401SAN ISIDRO          0015LIMA        LIMA      96041-005000244-91                   7,100.00      0.00      0.00  7,100.0026/NOV/2014-26/DIC/201420/ENE/2015    400.86      6.90    407.78      0.00     36.90    407.78  6,692.24      0.00      0.00      0.00  6,692.24    188.83    712.03      0.00      0.00      6.90    500.00    407.780000000000     30.00      0.00      0.00      6.90     36.90                                                                                                                                                                                                                                                                                                                                  SALDO INICIAL                                        188.8308/DIC/201409/DIC/2014370162MAKRO SUPERMAYORISTAT   706.1301/01                  706.1319/DIC/201419/DIC/2014049057PAGO TIENDA  RIPLEY T                               -500.0026/DIC/201426/DIC/2014000100SEGURO DESGRAVAMEN  T     5.9001/01                    5.90                            ENVIO EE.CC. MENSUAL                                   6.90                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                         U    "
        sPeriodoFinal = "201412"
        outpputBuff = RTrim(outpputBuff)
        sTrama = "0"
        If sTrama = "0" Then 'EXITO
            If outpputBuff.Length > 0 Then
                sTrama = outpputBuff.Substring(8, outpputBuff.Length - 8)
            End If

        End If





        MessageBox.Show(sTrama.Length)

        'VARIABLES RESUMEN DE ESTADO DE CUENTA
        sLineaCredito = Mid(sTrama, 253, 10)
        sDispEfectivoExpress = Mid(sTrama, 427, 10)
        sDisponibleCompras = Mid(sTrama, 387, 10)
        sPeriodoFacturacion = Mid(sTrama, 293, 23)
        sUltimoDiaPago = Mid(sTrama, 316, 11)

        sCreditoUtilizado = Mid(sTrama, 327, 10)
        sComisionCargos = Mid(sTrama, 337, 10)
        sDeudaTotal = Mid(sTrama, 347, 10)
        sDeudaVencida = Mid(sTrama, 357, 10)
        sPagoMinimoMes = Mid(sTrama, 367, 10)
        sPagoTotalMes = Mid(sTrama, 377, 10)


        'CALCULO DE PAGO TOTAL DEL MES
        sSaldoAnterior = Mid(sTrama, 437, 10)
        sConsumoMes = Mid(sTrama, 447, 10)
        sCuotaMes = Mid(sTrama, 457, 10)
        sInteres = Mid(sTrama, 467, 10)
        sComisionCargos_Mes = Mid(sTrama, 477, 10)
        sPagosAbonos = Mid(sTrama, 487, 10)
        sPagoTotal_Mes = Mid(sTrama, 497, 10)

        'CALCULO DE PAGO MINIMO DEL MES
        sSaldoFavor = Mid(sTrama, 507, 10)
        sMontoMinimo = Mid(sTrama, 517, 10)
        sCuotaMes_Min = Mid(sTrama, 527, 10)
        sInteres_Min = Mid(sTrama, 537, 10)
        sComisionCargos_Min = Mid(sTrama, 547, 10)
        sPagoMinimoMes_Min = Mid(sTrama, 557, 10)


        'PLAN DE CUOTAS DE LOS PROXIMOS 3 MESES
        sMes1 = Mid(sTrama, 615, 8)
        sMonto1 = Mid(sTrama, 623, 10)
        sMes2 = Mid(sTrama, 633, 8)
        sMonto2 = Mid(sTrama, 641, 10)
        sMes3 = Mid(sTrama, 651, 8)
        sMonto3 = Mid(sTrama, 659, 10)

        'MOVIMIENTOS
        sDataMovAux = ""
        sDataMov = ""
        sDataMov = Mid(sTrama, 861, sTrama.Length)


        Incrementa = 1
        If sDataMov.Length > 0 Then

            If sPeriodoFinal >= Constantes.PeriodoInclusionTEA Then
                For lFila = 1 To 13

                    If Trim(Mid(sDataMov, Incrementa, tamanioFilaDetalle)) <> "" Then
                        sDataMovAux = sDataMovAux & Mid(sDataMov, Incrementa, tamanioFilaDetalle) & "|\n|"
                    End If

                    Incrementa = Incrementa + tamanioFilaDetalle

                Next
            Else
                For lFila = 1 To 13

                    If Trim(Mid(sDataMov, Incrementa, tamanioFilaDetalleAnt)) <> "" Then
                        sDataMovAux = sDataMovAux & Mid(sDataMov, Incrementa, tamanioFilaDetalleAnt) & "|\n|"
                    End If

                    Incrementa = Incrementa + tamanioFilaDetalleAnt

                Next
            End If


        End If

        sXMLCAB = ""
        sXMLCAB = sLineaCredito.Trim & "|\t|" & sDispEfectivoExpress.Trim & "|\t|" & sDisponibleCompras.Trim
        sXMLCAB = sXMLCAB & "|\t|" & sPeriodoFacturacion.Trim & "|\t|" & sUltimoDiaPago.Trim
        sXMLCAB = sXMLCAB & "|\t|" & sCreditoUtilizado.Trim & "|\t|" & sComisionCargos.Trim & "|\t|" & sDeudaTotal.Trim
        sXMLCAB = sXMLCAB & "|\t|" & sDeudaVencida.Trim & "|\t|" & sPagoMinimoMes.Trim & "|\t|" & sPagoTotalMes.Trim


        sXMLPIE = ""
        sXMLPIE = sSaldoAnterior.Trim & "|\t|" & sConsumoMes.Trim & "|\t|" & sCuotaMes.Trim
        sXMLPIE = sXMLPIE & "|\t|" & sInteres.Trim & "|\t|" & sComisionCargos_Mes.Trim & "|\t|" & sPagosAbonos.Trim
        sXMLPIE = sXMLPIE & "|\t|" & sPagoTotal_Mes.Trim & "|\t|" & sSaldoFavor.Trim & "|\t|" & sMontoMinimo.Trim
        sXMLPIE = sXMLPIE & "|\t|" & sCuotaMes_Min.Trim & "|\t|" & sInteres_Min.Trim & "|\t|" & sComisionCargos_Min.Trim
        sXMLPIE = sXMLPIE & "|\t|" & sPagoMinimoMes_Min.Trim & "|\t|" & sMes1.Trim & "|\t|" & sMonto1.Trim
        sXMLPIE = sXMLPIE & "|\t|" & sMes2.Trim & "|\t|" & sMonto2.Trim & "|\t|" & sMes3.Trim & "|\t|" & sMonto3.Trim

        Dim lol As String = Microsoft.VisualBasic.Strings.Right(sTrama, 1)
        If Microsoft.VisualBasic.Strings.Right(sTrama, 1) = "U" Then
            'If Left(sTrama, 2) = "AU" Then 'FIN DE LA SOLICITUD

            If sDataMovAux.Length > 0 Then
                sDataMovAux = Microsoft.VisualBasic.Strings.Left(sDataMovAux, sDataMovAux.Length - 4)
                sXMLMOV = sXMLMOV & sDataMovAux

            Else
                sDataMovAux = "" 'NO HAY MOVIMIENTOS
                sXMLMOV = sXMLMOV & sDataMovAux


            End If

            If sXMLMOV.Length > 0 Then
                'Armar las columnas de los movimientos de EECC
                sXMLMOV_FINAL = EstadosCuenta.Instancia.FUN_DETALLE_EECC_TEA(sXMLMOV, sPeriodoFinal)

            End If

        End If

        'CADENA FINAL CON LOS DATOS FINALES
        sRespuesta = sXMLCAB.Trim & "*$¿*" & sXMLPIE.Trim & "*$¿*" & sXMLMOV_FINAL.Trim
        lblRespuesta.Text = sRespuesta
    End Sub

    Public Function ObtenerFechaFacturacion(ByVal ultimoDiaPago As String, ByVal mes As String, ByVal anio As String) As String
        Dim diasMes As Integer = 0
        diasMes = Date.DaysInMonth(CInt(anio), CInt(mes))
        Dim respuesta As String = String.Empty
        Try

            Select Case ultimoDiaPago
                Case 1
                    Select Case diasMes
                        Case 28
                            respuesta = "04"
                        Case 29
                            respuesta = "05"
                        Case 30
                            respuesta = "06"
                        Case 31
                            respuesta = "07"
                        Case Else
                            respuesta = "XX"
                    End Select
                Case 5
                    Select Case diasMes
                        Case 28
                            respuesta = "08"
                        Case 29
                            respuesta = "09"
                        Case 30
                            respuesta = "10"
                        Case 31
                            respuesta = "11"
                        Case Else
                            respuesta = "XX"
                    End Select
                Case 10
                    Select Case diasMes
                        Case 28
                            respuesta = "13"
                        Case 29
                            respuesta = "14"
                        Case 30
                            respuesta = "15"
                        Case 31
                            respuesta = "16"
                        Case Else
                            respuesta = "XX"
                    End Select
                Case 15
                    Select Case diasMes
                        Case 28
                            respuesta = "18"
                        Case 29
                            respuesta = "19"
                        Case 30
                            respuesta = "20"
                        Case 31
                            respuesta = "21"
                        Case Else
                            respuesta = "XX"
                    End Select
                Case 20
                    Select Case diasMes
                        Case 28
                            respuesta = "23"
                        Case 29
                            respuesta = "24"
                        Case 30
                            respuesta = "25"
                        Case 31
                            respuesta = "26"
                        Case Else
                            respuesta = "XX"
                    End Select
                Case 25
                    Select Case diasMes
                        Case 28
                            respuesta = "28"
                        Case 29
                            respuesta = "29"
                        Case 30
                            respuesta = "30"
                        Case 31
                            respuesta = "31"
                        Case Else
                            respuesta = "XX"
                    End Select
                Case Else
                    respuesta = "XX"

            End Select

            If respuesta.Contains("XX") Then
                respuesta = "ERROR:NODATA-Último día de pago es incorrecto."
            Else
                respuesta = anio & mes & respuesta
            End If

        Catch ex As Exception
            respuesta = "ERROR:NODATA-Hubo error en llamada a método ObtenerFechaFacturacion " & ex.Message
        End Try

        Return respuesta
    End Function

    Enum TServidor

        SICRON = 1
        RSAT = 2

    End Enum

    Private Sub ObtenerTramaEECC_Click(sender As System.Object, e As System.EventArgs) Handles ObtenerTramaEECC.Click
        'Dim respuesta As String = ESTADO_CUENTA_CLASICA_RSAT("5420700010596464", "05174142", "", TServidor.RSAT)
        Dim respuesta As String = EstadosCuenta.Instancia.ESTADO_CUENTA_CLASICA_RSAT("5254350068889667", "000005278239", "000005278239|\t|5254350068889667|\t|CASTILLO ALFARO, MIJAIL ALEXANDER|\t|00|\t|      |\t|      |\t|2|\t|SIRIP-99|\t|13", TServidor.RSAT)
        'Dim respuesta As String = Movimientos.Instancia.MOVIMIENTOS_CLASICA(
        '    "5254350068889667",
        '    "000005278239",
        '    "000005278239|\t|5254350068889667|\t|CASTILLO ALFARO, MIJAIL ALEXANDER|\t|00|\t|      |\t|      |\t|2|\t|SIRIP-99|\t|12",
        '    TServidor.RSAT,
        '    Nothing
        ')

        ErrorLog("respuesta = " & respuesta)
        txtSalidaEECC.Text = respuesta
    End Sub

    Public Function ObtenerOfertaCambioProducto(ByVal contrato As String) As Response
        Dim respuesta As New Response
        respuesta.Success = False
        respuesta.Warning = False
        Dim oferta As New OfertaCambioProducto
        Try
            ErrorLog("Entro al método ObtenerOfertaCambioProducto(" & contrato & ")")
            'oferta = BNOfertaCambioProducto.Instancia.ObtenerOferta(contrato)
            oferta.ContratoTarjeta = contrato
            oferta.DatosTarjeta = "014002TARJETA PLATINUM VISA SAT                         TARJETA CLASICA SAT                               "
            oferta.DatosSEF = "01"
            If oferta.ContratoTarjeta <> String.Empty Then
                Dim validaTarjeta As String
                Dim productoOrigenTarjeta As String
                Dim productoDestinoTarjeta As String
                Dim descripcionProductoOrigenTarjeta As String
                Dim descripcionProductoDestinoTarjeta As String

                Dim validaSEF As String
                Dim productoOrigenSEF As String
                Dim productoDestinoSEF As String
                Dim descripcionProductoOrigenSEF As String
                Dim descripcionProductoDestinoSEF As String

                If oferta.DatosTarjeta.Length = 106 Then
                    validaTarjeta = oferta.DatosTarjeta.Substring(0, 2)
                    productoOrigenTarjeta = oferta.DatosTarjeta.Substring(2, 2)
                    productoDestinoTarjeta = oferta.DatosTarjeta.Substring(4, 2)
                    descripcionProductoOrigenTarjeta = oferta.DatosTarjeta.Substring(6, 50)
                    descripcionProductoDestinoTarjeta = oferta.DatosTarjeta.Substring(56, 50)
                ElseIf oferta.DatosTarjeta.Length = 2 Then
                    validaTarjeta = oferta.DatosTarjeta.Substring(0, 2)
                End If

                If oferta.DatosSEF.Length = 106 Then
                    validaSEF = oferta.DatosSEF.Substring(0, 2)
                    productoOrigenSEF = oferta.DatosSEF.Substring(2, 2)
                    productoDestinoSEF = oferta.DatosSEF.Substring(4, 2)
                    descripcionProductoOrigenSEF = oferta.DatosSEF.Substring(6, 50)
                    descripcionProductoDestinoSEF = oferta.DatosSEF.Substring(56, 50)
                ElseIf oferta.DatosSEF.Length = 2 Then
                    validaSEF = oferta.DatosSEF.Substring(0, 2)
                End If
            Else
                respuesta.Warning = True
                respuesta.Message = "No tiene oferta de cambio de producto"
            End If

            respuesta.Success = True
        Catch ex As Exception
            respuesta.Message = "En estos momentos no podemos atenderle. Por favor inténtelo más tarde"
            ErrorLog("ObtenerOfertaCambioProducto Error =" & ex.Message)
        End Try

        Return respuesta
    End Function

    Private Sub btnCambioProducto_Click(sender As System.Object, e As System.EventArgs) Handles btnCambioProducto.Click

        AddOUTParameters(New VTAAUTO5001ParameterOUT)
        Dim lo232 As String = parameterOut.ToString()

        Dim tramaRespuesta As String = "0000000000VTAAUTO5001                                       0000010240001000004SIRIP-990188VACP2015-02-11RMPES                                                                                                                                                                                                                                                                                                                                                                                                                 000 ES                                                                                        00010001000005351359C 10085380    2SAT     188         SIRIP-99    10.25.4.6      04801100000000000000000000000200EL CONTRATO CUMPLE CON LAS CONDICIONES DE CAMBIO DE PRODUCTO. CODIGO DE CAMBIO DE PRODUCYO (0480) ENCONTRADO"
        Dim tananio As Integer = tramaRespuesta.Length
        Dim variante As Integer = 92
        Dim CuentaENC As String = tramaRespuesta.Substring(639, 1)
        Dim CodCambioENC As String = tramaRespuesta.Substring(640, 1)
        Dim TieneSEF As String = tramaRespuesta.Substring(641, 1)
        Dim ContratoSEF As String = tramaRespuesta.Substring(695 - variante, 20)
        Dim CambioOK As String = tramaRespuesta.Substring(715 - variante, 1)

        Dim Paso As String = tramaRespuesta.Substring(716 - variante, 2)
        Dim CodRpta As String = tramaRespuesta.Substring(718 - variante, 2)
        Dim DesRpta As String = tramaRespuesta.Substring(720 - variante, 350)

        Dim fechaActual As String = Date.Now.ToShortDateString()
        FormatearFecha(fechaActual)
        ObtenerOfertaCambioProducto("00010001000000011111")
    End Sub

    Private Function FormatearFecha(ByVal sFecha As String) As String
        Dim sResul As String = ""
        Dim sDia As String = ""
        Dim sMes As String = ""
        Dim sAnio As String = ""

        If sFecha.Trim.Length > 0 Then
            sDia = Microsoft.VisualBasic.Strings.Left(sFecha.Trim, 2)
            sMes = Mid(sFecha.Trim, 3, 2)
            sAnio = Microsoft.VisualBasic.Strings.Right(sFecha.Trim, 4)

            sResul = sDia.Trim & "/" & sMes.Trim & "/" & sAnio.Trim

        End If


        Return sResul.Trim

    End Function

    Sub AddOUTParameters(ByVal parametrosOUT As VTAAUTO5001ParameterOUT)

        parameterOut.CuentaENC = fx_Completar_Campo("0", 1, parametrosOUT.CuentaENC, TYPE_ALINEAR.DERECHA)
        parameterOut.ProductoO = fx_Completar_Campo("0", 2, parametrosOUT.ProductoO, TYPE_ALINEAR.DERECHA)
        parameterOut.CodCambioENC = fx_Completar_Campo("0", 1, parametrosOUT.CodCambioENC, TYPE_ALINEAR.DERECHA)
        parameterOut.CuentaPlataformaEnc = fx_Completar_Campo("0", 1, parametrosOUT.CuentaPlataformaEnc, TYPE_ALINEAR.DERECHA)
        parameterOut.TieneSEF = fx_Completar_Campo("0", 1, parametrosOUT.TieneSEF, TYPE_ALINEAR.DERECHA)
        parameterOut.ContratoSEF = fx_Completar_Campo("0", 20, parametrosOUT.ContratoSEF, TYPE_ALINEAR.DERECHA)
        parameterOut.CambioOK = fx_Completar_Campo("0", 1, parametrosOUT.CambioOK, TYPE_ALINEAR.DERECHA)
        parameterOut.Paso = fx_Completar_Campo("0", 2, parametrosOUT.Paso, TYPE_ALINEAR.DERECHA)
        parameterOut.CodRpta = fx_Completar_Campo("0", 2, parametrosOUT.CodRpta, TYPE_ALINEAR.DERECHA)

    End Sub

    Private Sub Form1_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Dim lol As String = "27/MAR/201527/MAR/2015004230COMPRA EN CUOTAS                                                                        1400.00"
        Dim rpta As String = Validar()
    End Sub

    Public Function Validar() As String
        Dim TipProducto As String = "2"
        Dim TipDocumento As String = "C"
        Dim NroDocumento As String = "44780645"
        Dim sRespuesta As String = "0000000000SFSCANC0040                                       000002000PE00010000020027X8  0       C44780645    1                                           00APAZA               CAHUANA             ANTHONY SILVERIO    026756840200010001000002615323050001-01-0101TI0000000000024500340005509259      46                              3000010001000002615323050001-01-0101TI0000000000014500340005509267      00                              30                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                      "
        Dim sNroTarjeta As String = String.Empty
        Dim sNroDocumento As String = "44780645"
        Dim sApellidoParterno As String = String.Empty
        Dim sApellidoMaterno As String = String.Empty
        Dim sNombres As String = String.Empty
        Dim sCliente As String = String.Empty
        Dim sFechaNac As String = String.Empty
        Dim sTipoProducto As String = String.Empty 'Segun el bin de la tarjeta 6 primeros numeros de la tarjeta
        Dim sTipoDocumento As String = String.Empty
        Dim sNroCuenta As String = String.Empty
        Dim sXMLRespuesta As String = String.Empty
        Dim Nro_Contrato As String
        Dim Tipo_Tarjeta_Ofertas As String

        If sRespuesta.Trim.Length > 0 Then
            If Microsoft.VisualBasic.Strings.Left(sRespuesta.Trim, 5) = "ERROR" Then
                'Save Error
                sRespuesta = "ERROR:Servicio no disponible."

            Else
                ErrorLog("Inicio de Corte")
                '---- INICIO CORTE ----'
                Dim fnEstMigra As String
                fnEstMigra = Mid(sRespuesta, 155, 2)
                If fnEstMigra = "00" Then
                    ErrorLog("fnEstMigra = 00")
                    sNombres = Mid$(sRespuesta, 197, 20)            '-> Nombres
                    sApellidoParterno = Mid$(sRespuesta, 157, 20)   '-> Apellido Paterno
                    sApellidoMaterno = Mid$(sRespuesta, 177, 20)    '-> Apellido Materno
                    sCliente = sApellidoParterno.Trim & " " & sApellidoMaterno.Trim & ", " & sNombres.Trim
                    Nro_Contrato = Mid$(sRespuesta, 227, 20)
                    Tipo_Tarjeta_Ofertas = "RSAT"
                    Dim int_cont, iIndEle, iLonBucle As Integer
                    Dim bCondCli As Boolean

                    Dim vg_RSAT_Sitarj, vg_RSAT_CodBloqueo, vg_RSAT_FecBaj, vg_RSAT_CodTitularAdicional As String

                    int_cont = Val(Mid(sRespuesta, 225, 2))
                    iLonBucle = 104
                    bCondCli = False

                    For iIndEle = 1 To int_cont
                        sNroCuenta = Mid(sRespuesta, 235 + iLonBucle * (iIndEle - 1), 12)
                        vg_RSAT_Sitarj = Mid(sRespuesta, 247 + iLonBucle * (iIndEle - 1), 2) 'SITUACION TARJETA
                        vg_RSAT_FecBaj = Mid(sRespuesta, 249 + iLonBucle * (iIndEle - 1), 10)
                        vg_RSAT_CodTitularAdicional = Mid(sRespuesta, 261 + iLonBucle * (iIndEle - 1), 2)
                        sNroTarjeta = Mid(sRespuesta, 275 + iLonBucle * (iIndEle - 1), 22)
                        vg_RSAT_CodBloqueo = Mid(sRespuesta, 297 + iLonBucle * (iIndEle - 1), 2)

                        If sNroCuenta.Length = 0 Then
                            sRespuesta = "ERROR: DNI no encontrado"
                            Exit For
                        End If

                        If VALIDAR_BLOQUEDO_TARJETA_RSAT(vg_RSAT_CodBloqueo) = False And vg_RSAT_Sitarj = "05" And vg_RSAT_FecBaj = "0001-01-01" And vg_RSAT_CodTitularAdicional = "TI" Then
                            bCondCli = True
                        Else
                            If VALIDAR_BLOQUEDO_TARJETA_RSAT(vg_RSAT_CodBloqueo) = False And vg_RSAT_Sitarj = "05" And vg_RSAT_FecBaj = "0001-01-01" And vg_RSAT_CodTitularAdicional = "BE" Then
                                bCondCli = True
                            End If
                        End If

                        '<INI TCK-563699-01 DHERRERA 20-03-2014>
                        'If bCondCli Then
                        If bCondCli = True And sNroTarjeta.Substring(6, 3) <> "761" Then
                            '<FIN TCK-563699-01 DHERRERA 20-03-2014>
                            If getTipProducto_AbiertaRSAT(sNroTarjeta.Substring(0, 6)).ToString = TipProducto Then

                                If TipDocumento = "C" Then 'DNI
                                    sTipoDocumento = "1"
                                Else
                                    sTipoDocumento = "2" 'Carnet de Extranjeria
                                End If

                                sTipoProducto = Mid(sNroTarjeta.Trim, 1, 6)
                                sRespuesta = sNroTarjeta.Trim & "|\t|" & NroDocumento.Trim & "|\t|" & sCliente.Trim & "|\t|" & sFechaNac.Trim & "|\t|" & sTipoProducto.Trim & "|\t||\t|RSAT"
                                Return sRespuesta
                            Else
                                sXMLRespuesta = "ERROR: DNI no encontrado"
                            End If

                            If iIndEle = int_cont Then
                                Return sXMLRespuesta
                            End If

                        Else

                            If iIndEle = int_cont Then
                                Return "ERROR:Servicio no disponible."
                            End If

                        End If


                    Next iIndEle



                Else
                    sRespuesta = "ERROR: DNI no encontrado"
                End If
            End If

        Else
            'Sino Encontro DNI en RSAT
            sRespuesta = "ERROR: DNI no encontrado"
        End If

        Return sRespuesta
    End Function

    Private Function VALIDAR_BLOQUEDO_TARJETA_RSAT(ByVal cod_bloquedo As String) As Boolean

        Dim ind As Byte
        Dim bTFEncontrada As Boolean
        bTFEncontrada = False

        Try
            For ind = 0 To c_TablaBloqueo.GetLength(0) - 1
                If cod_bloquedo = "00" Then
                    Exit For
                ElseIf cod_bloquedo = c_TablaBloqueo(ind, 0) And c_TablaBloqueo(ind, 3) = "S" Then
                    bTFEncontrada = True
                    Exit For
                End If
            Next
        Catch ex As Exception
            bTFEncontrada = False
        End Try

        Return bTFEncontrada

    End Function

    Private Function getTipProducto_AbiertaRSAT(ByVal BINN As String) As Integer

        Dim codigo As Integer

        Select Case BINN

            Case "525435"
                codigo = 2
            Case "542070"
                codigo = 1
            Case "450000"
                codigo = 5
            Case "450007"
                codigo = 4
                '<INI TCK-563699-01 DHERRERA 20-03-2014>
            Case "542020"
                codigo = 1
            Case "525474"
                codigo = 2
                '<FIN TCK-563699-01 DHERRERA 20-03-2014>
        End Select

        Return codigo

    End Function

    Private Sub BtnTitular_Click(sender As System.Object, e As System.EventArgs) Handles BtnTitular.Click
        txtDia.Text = Mid(txtSalidaEECC.Text, 261, 2)
    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Dim fechaActual As String = Date.Now.ToShortDateString()
        Dim dia As String = String.Empty
        Dim mes As String = String.Empty
        Dim anio As String = String.Empty
        Dim fechaResultante As String = String.Empty

        If fechaActual.Trim.Length > 0 Then
            dia = Microsoft.VisualBasic.Strings.Left(fechaActual.Trim, 2)
            mes = Mid(fechaActual.Trim, 4, 2)
            anio = Microsoft.VisualBasic.Strings.Right(fechaActual.Trim, 4)

            fechaResultante = anio.Trim() & "/" & mes.Trim() & "/" & dia.Trim()

        End If
    End Sub
End Class

