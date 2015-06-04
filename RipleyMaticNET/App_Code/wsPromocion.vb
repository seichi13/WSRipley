Imports System.Web
Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports NCapas.Entity
Imports NCapas.Business
Imports NCapas.Utility.Log

' Para permitir que se llame a este servicio Web desde un script, usando ASP.NET AJAX, quite la marca de comentario de la siguiente línea.
' <System.Web.Script.Services.ScriptService()> _
<WebService(Namespace:="http://tempuri.org/")> _
<WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Public Class wsPromocion
    Inherits System.Web.Services.WebService

    <WebMethod()> _
    Public Function HelloWorld() As String
        Return "Hola a todos"
    End Function

    <WebMethod()> _
    Public Function txSEF() As String
        Dim valor As String = ""

        Return valor
    End Function

    <WebMethod(Description:="Actualizar oferta SEF por objeto", MessageName:="Funcion de Objeto")> _
    Public Function ACTUALIZAR_OFERTA_SEF(ByVal sDataParametros As SuperEfectivo.Parametros) As Salida
        Dim resultado As New Salida
        Try
            Dim oBN_SuperEfectivo As New BNSuperEfectivo
            Dim oResultado As New SuperEfectivo.Salida
            Registrar_Log(3, "", "Ingresó a Servicio ACTUALIZAR_OFERTA_SEF")
            oResultado = oBN_SuperEfectivo.Aprobar_Oferta(sDataParametros)
            Registrar_Log(3, "", "Terminó la llamada a oBN_SuperEfectivo.Aprobar_Oferta(sDataParametros)")
            If oResultado.nWSResult = 1 Then
                resultado.Code = NCapas.Utility.Constantes.CODE_EXITO
                'resultado.Mensaje = exitoRSAT
                'resultado.MensajeImpresion = exitoImpresionRSAT
            Else
                resultado.Code = 2
                'resultado.Mensaje = errorRSAT
            End If
        Catch ex As Exception
            Registrar_Log(3, "", "Catch Exception " & ex.Message)
            resultado.Code = 5
            resultado.Mensaje = ex.Message
        End Try
        Return resultado
    End Function

    <WebMethod(Description:="Actualizar oferta SEF por parametros", MessageName:="Funcion de parametros")> _
    Public Function ACTUALIZAR_OFERTA_SEF2(ByVal psCodigoOficina As String, _
                                          ByVal psTerminalFisico As String, _
                                          ByVal psCodigoCanal As String, _
                                          ByVal psNombreCliente As String, _
                                          ByVal psTipoDocumento As String, _
                                          ByVal psNumeroDocumento As String, _
                                          ByVal psTipoProducto As String, _
                                          ByVal psCodigoEntidad As String, _
                                          ByVal psCentroAlta As String, _
                                          ByVal psCuenta As String, _
                                          ByVal pnMontoAceptado As String, _
                                          ByVal pnPlazoMes As String, _
                                          ByVal pnCuotaMes As String, _
                                          ByVal pnTEA As String, _
                                          ByVal pnTEM As String, _
                                          ByVal psMoneda As String, _
                                          ByVal psFechaVencimiento As String, _
                                          ByVal psCanal As String, _
                                          ByVal psSucursal As String, _
                                          ByVal psCodigoVendedor As String, _
                                          ByVal psCodigoUsuario As String, _
                                          ByVal psCodigoAtencion As String, _
                                          ByVal psAplicaITF As String, _
                                          ByVal psIndicadorCommit As String) As Salida
        Dim resultado As New Salida
        Dim sDataParametros As New SuperEfectivo.Parametros
        sDataParametros.sCodigoOficina = psCodigoOficina
        sDataParametros.sTerminalFisico = psTerminalFisico
        sDataParametros.sCodigoCanal = psCodigoCanal
        sDataParametros.sNombreCliente = psNombreCliente
        sDataParametros.sTipoDocumento = psTipoDocumento
        sDataParametros.sNumeroDocumento = psNumeroDocumento
        sDataParametros.sTipoProducto = psTipoProducto
        sDataParametros.sCodigoEntidad = psCodigoEntidad
        sDataParametros.sCentroAlta = psCentroAlta
        sDataParametros.sCuenta = psCuenta
        sDataParametros.nMontoAceptado = pnMontoAceptado
        sDataParametros.nPlazoMes = pnPlazoMes
        sDataParametros.nCuotaMes = pnCuotaMes
        sDataParametros.nTEA = pnTEA
        sDataParametros.nTEM = pnTEM
        sDataParametros.sMoneda = psMoneda
        sDataParametros.sFechaVencimiento = psFechaVencimiento
        sDataParametros.sCanal = psCanal
        sDataParametros.sSucursal = psSucursal
        sDataParametros.sCodigoVendedor = psCodigoVendedor
        sDataParametros.sCodigoUsuario = psCodigoUsuario
        sDataParametros.sCodigoAtencion = psCodigoAtencion
        sDataParametros.sAplicaITF = psAplicaITF
        sDataParametros.sIndicadorCommit = psIndicadorCommit

        resultado = ACTUALIZAR_OFERTA_SEF(sDataParametros)
        Return resultado
    End Function

    <WebMethod(Description:="Actualizar oferta de incremento de linea.")> _
    Public Function Extraer_trama_OFERTA_SEF(ByVal sDataParametros As SuperEfectivo.Parametros) As String
        Dim resultado As String
        Try
            Dim oBN_Incremento_Linea As New BNSuperEfectivo

            resultado = oBN_Incremento_Linea.Extraer_Oferta(sDataParametros)
        Catch ex As Exception
            resultado = "Extraer_trama_OFERTA_SEF: " + ex.Message
        End Try

        Return resultado
    End Function

End Class