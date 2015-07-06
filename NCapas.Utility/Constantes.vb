Public Class Constantes
    Public Const CODE_EXITO As String = "OK"
    Public Const CODE_ERROR As String = "ERROR"
    Public Const CODE_NULL As String = "NULL"

    Public Const PROMOCION_INC As String = "INCREMENTO_LINEA"

    Public Const TIPO_USUARIO As String = "R"
    Public Const CODIGO_VENDEDOR As String = "0"

    Public Const DECIMALES_SALIDA As Integer = 2
    Public Const FORMATO_FECHAS As String = "dd/MM/yyy"
    Public Const FORMATO_FECHASDESEMBOLSO As String = "dd/MM/yyyy"

    Public Const BINN_VALIDADO_NO As String = "sBinValidadoNO"

    Public Const COD_APLICACION As String = "RM"
    Public Const MONEDA_SOLES As String = "604"
    Public Const Long_Mensaje_VTAAUTO0004 As String = "000002048"
    Public Const Long_Mensaje_SFSRPMT0100 As String = "000000374"

    Public Const ServicioSFSCANT0013CODRETORNO As String = "0000000000"
    Public Const ServicioSFSCANT0013NOMBRE As String = "SFSCANT0013"
    Public Const ServicioSFSCANT0013LONGMENSAJE As String = "000004182"
    Public Const ServicioSFSCANT0013ESTADOEXITOSO As String = "0"
    Public Const ServicioSFSCANT0013ESTADOEXITOSOVACIO As String = " "
    Public Const ServicioSFSCANT0013ESTADOERROR As String = "9"

    Public Const PeriodoInclusionTEA As String = "201501"
    Public Const ServidorRSAT As String = "S"
    Public Const TramaDoce As String = "S"
    Public Const TramaTrece As String = "T"

    Public Const PeriodoAgrandarGlosa As String = "201506"

#Region "Mensajes de Email"
    Public Const MensajeEmailClave6 As String = "<br><p style='font-weight: bold;'>¡YA TIENES TU CLAVE INTERNET!</p><p>Felicitaciones, tu Clave Internet (6 d&iacute;gitos) se gener&oacute; con &eacute;xito.</p><p>Desde ahora podr&aacute;s realizar tus operaciones bancarias m&aacute;s r&aacute;pido y f&aacute;cil, ingresando a nuestra Banca por Internet desde nuestra p&aacute;gina web <a href='http://www.bancoripley.com.pe/'>www.bancoripley.com.pe</a>.</p><br><p>Atentamente,</p><p>Banco Ripley</p>"
    'Public Const MensajeEmailClave6 As String = "<p style='text-align:center;'><b>GENERACIÓN DE CLAVE WEB</b></p><p><b>Te confirmamos la Generaci&oacute;n de tu clave web en forma exitosa, ahora podr&aacute;s disfrutar de las opciones de nuestro Banca por Internet.</b></p>El presente correo electr&oacute;nico por el Banco Ripley, tiene car&aacute;cter confidencial y s&oacute;lo puede ser utilizado por la persona a quien ha sido dirigida. Su divulgaci&oacute;n, copia y/u otro no autorizado est&aacute; estrictamente prohibida.<br/>Si usted no es el destinatario esperado, por favor cont&aacute;ctese con el remitente y elimine el mensaje. Esta comunicaci&oacute;n es s&oacute;lo para prop&oacute;sitos de informaci&oacute;n y no genera obligaci&oacute;n contractual alguna a cargo de RIPLEY.<br/>"
#End Region


#Region "Mensajes a mostrar"
    Public Const MSGErrorConexion As String = "ERROR:No se pudo conectar al servidor de base de datos."
#End Region

#Region "Cambio Producto"
    Public Const TipoTransaccion As String = "T"
    Public Const TipoConsulta As String = "C"
    Public Const OfertaVigente As String = "S"
    Public Const TarjetaSEFValido As String = "01"
    Public Const TipoDocumentoDNI As String = "C"
    Public Const TipoSistema As String = "SAT"
    Public Const ModuleId As String = "00"
    Public Const STipoTerminal As String = "000004"
    Public Const SistemaRM As String = "RM"
    Public Const ErrorWS As String = "01"
    Public Const ErrorRM As String = "02"
    Public Const NoTieneContratoSEF As String = "0"
    Public Const TieneContratoSEF As String = "1"

    Public Const ServicioVTAAUTO5001CODRETORNO As String = "0000000000"
    Public Const ServicioVTAAUTO5001NOMBRE As String = "VTAAUTO5000"
    Public Const ServicioVTAAUTO5001LONGMENSAJE As String = "000002600"
    Public Const ServicioVTAAUTO5001ESTADOEXITOSO As String = "0"
    Public Const ServicioVTAAUTO5001ESTADOEXITOSOVACIO As String = " "
    Public Const ServicioVTAAUTO5001ESTADOERROR As String = "9"
    Public Const ServicioVTAAUTO5001CODENTIDAD As String = "0001"
    Public Const ServicioVTAAUTO5001TIPOTERMINAL As String = "000004"
    Public Const ServicioVTAAUTO5001TERMINALFISICO As String = "VACP"
    Public Const ServicioVTAAUTO5001PAIS As String = "PE"
    Public Const ServicioVTAAUTO5001COMMIT As String = "S"
    Public Const ServicioVTAAUTO5001PANTALLA As String = "000"
    Public Const ServicioVTAAUTO5001OTROSDATOS As String = "ES"
    Public Const ServicioVTAAUTO5001PASO00 As String = "00"
    Public Const ServicioVTAAUTO5001PASO01 As String = "01"
    Public Const ServicioVTAAUTO5001PASO02 As String = "02"
    Public Const ServicioVTAAUTO5001PASO03 As String = "03"
    Public Const ServicioVTAAUTO5001PASO04 As String = "04"
    Public Const ServicioVTAAUTO5000PASO01 As String = "05"
    Public Const ServicioVTAAUTO5000PASO02 As String = "X"

    Public Const ServicioVTAAUTO5001RPTA00 As String = "00"
    Public Const ServicioVTAAUTO5001RPTA01 As String = "01"
    Public Const ServicioVTAAUTO5000RPTA As String = "01"

    Public Const ServicioVTAAUTO5001TIPOCONTRATOTARJETA As String = "1"
    Public Const ServicioVTAAUTO5001TIPOCONTRATOSEF As String = "2"

    Public Const ErrorValidarContratoSEF = "Error al Validar Contrato SEF"
    Public Const ErrorContratosDiferentes = "El Producto del Contrato TARJETA = {0} es diferente a el Contrato SEF {1}"
    Public Const ErrorContratoSEFOfertaDifentePlataforma = "Contrato SEF de la OFerta es diferente al Contrato SEF de Plataforma Comercial"
    Public Const ErrorOfertaConSEFPlataformaSinSEF = "Oferta tiene Contrato SEF, sin embargo Plataforma no tiene un SEF.)"
    Public Const ErrorValidarContratoTarjeta = "Error al Validar Contrato Tarjeta"
    Public Const ErrorCodigoTarjetaInvalido = "Código de Tarjeta Inválido. "
    Public Const ErrorCodigoSEFInvalido = "Código SEF Inválido. "
    Public Const ErrorPromocionNoVigente = "Promoción no está vigente."
    Public Const SinPromocion = "No tiene oferta de cambio de producto"
    Public Const ErrorServidor = "En estos momentos no podemos atenderle. Por favor inténtelo más tarde"

    Public Const PLAMC = "PLAMC"
    Public Const CLASICA = "CLASICA"
    Public Const GOLDMC = "GOLDMC"
    Public Const PLAVI = "PLAVI"
    Public Const SILVERMC = "SILVERMC"
    Public Const GOLDMCR = "GOLDMCR"
    Public Const PLAVISAR = "PLAVISAR"
    Public Const SILVERMCR = "SILVERMCR"
    Public Const SILVERVISA = "SILVERVISA"
    Public Const SILVERVISAR = "SILVERVISAR"

#End Region

#Region "Mensajes"
    Public Const MSG_NO_DATOS_CLIENTE As String = "No se encontraron los datos del Cliente."
    Public Const MSG_TARJETA_NO_VALIDA As String = "La tarjeta que desea consultar, no es una tarjeta válida."
    Public Const MSG_ERROR_CODIGO As String = "Ha ocurrido un error al procesar su consulta, intentelo más tarde."
    Public Const MSG_PROBLEMA_CONECCION As String = "Ha ocurrido un error de conección, comuniquelo al encargado."

#End Region

#Region "Titularidad"
    Public Const Titular As String = "TI"
    Public Const Beneficiario As String = "BE"
    Public Const OperacionTarjeta As String = "1"
    Public Const OperacionDNI As String = "2"
    Public Const FechaCaducidadPopUpInvasivo As String = "20150430"
#End Region
End Class