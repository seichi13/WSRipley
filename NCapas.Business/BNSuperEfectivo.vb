Imports NCapas.Entity.SuperEfectivo
Imports NCapas.RSAT.Servicio
Imports NCapas.RSAT
Imports NCapas.Utility.Log

Public Class BNSuperEfectivo

    Private oRSAT_SuperEfectivo As Servicio

    Public Function Aprobar_Oferta(ByVal params As Parametros) As Salida
        Dim oAprobacion As New Salida
        Dim oResultado As New PARAMETROS_OUT
        Registrar_Log(3, "", "Antes de llamar a ValidarParametrosEntradaSEF(params)")
        ValidarParametrosEntradaSEF(params)
        Registrar_Log(3, "", "Inicializar Service()")
        oRSAT_SuperEfectivo = New Servicio("SFSRPMT0100", Utility.Constantes.Long_Mensaje_SFSRPMT0100, params.sCodUsuario, params.sCodigoOficina, params.sCodigoCanal, params.sTerminalFisico)
        oRSAT_SuperEfectivo.Add_Parametros(params)
        Registrar_Log(3, "", "llamar a oRSAT_SuperEfectivo.Ejecutar_Servicio()")
        oResultado = oRSAT_SuperEfectivo.Ejecutar_Servicio()
        Registrar_Log(3, "", "oResultado.sWSMessage=" & oResultado.sWSMessage)
        If oResultado.sWSMessage = "OK" Then
            oAprobacion.nWSResult = oResultado.nWSResult
            oAprobacion.sNumeroCuenta = oResultado.sNumeroCuenta
            oAprobacion.sNumeroTarjeta = oResultado.sNumeroTarjeta
        Else
            oAprobacion.nWSResult = 0
            oAprobacion.sWSMessage = "Error"
        End If
        Return oAprobacion
    End Function

    Public Function Extraer_Oferta(ByVal params As Parametros) As String
        Dim oAprobacion As String = ""
        Try
            ValidarParametrosEntradaSEF(params)
            oRSAT_SuperEfectivo = New Servicio("SFSRPMT0100", Utility.Constantes.Long_Mensaje_SFSRPMT0100, params.sCodUsuario, params.sCodigoOficina, params.sCodigoCanal, params.sTerminalFisico)
            oRSAT_SuperEfectivo.Add_Parametros(params)
            oAprobacion = oRSAT_SuperEfectivo.Extraer_Trama()
        Catch ex As Exception
            oAprobacion = "Extraer_Oferta " + ex.Message
        End Try

        Return oAprobacion
    End Function

    Public Sub ValidarParametrosEntradaSEF(ByRef params As Parametros)
        If params.sTipoDocumento.Trim = "1" Then
            params.sTipoDocumento = "C"
        End If
        Registrar_Log(3, "", "Antes de llamar fx_Preparar_Decimal " & params.nMontoAceptado)
        params.nMontoAceptado = Utility.Funciones.fx_Preparar_Decimal(params.nMontoAceptado)
        Registrar_Log(3, "", "Antes de llamar fx_Preparar_Decimal " & params.nCuotaMes)
        params.nCuotaMes = Utility.Funciones.fx_Preparar_Decimal(params.nCuotaMes)
        Registrar_Log(3, "", "Antes de llamar fx_Preparar_Decimal " & params.nTEA)
        params.nTEA = Utility.Funciones.fx_Preparar_Decimal(params.nTEA)
        Registrar_Log(3, "", "Antes de llamar fx_Preparar_Decimal " & params.nTEM)
        params.nTEM = Utility.Funciones.fx_Preparar_Decimal(params.nTEM)
    End Sub
End Class
