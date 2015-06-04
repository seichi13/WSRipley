Imports NCapas.Utility.Funciones
Imports NCapas.Utility.Tipos
Imports NCapas.Utility.Log
Imports NCapas.Entity.SuperEfectivo
Imports MQCOMLib

Public Class Servicio

    Private HEADER_ As New HEADER
    Private PARAMETROS_IN_ As New PARAMETROS_IN
    Private PARAMETROS_OUT_ As Salida
    Private service As String = ""

    Sub New(ByVal nombreServicio As String,
            ByVal largoMensaje As String,
           ByVal pCod_Usuario As String,
            ByVal pCod_Oficina As String,
            ByVal pCod_Canal As String,
            ByVal pTerm_Fisico As String)

        service = nombreServicio
        HEADER_.COD_RESPUESTA = "0000000000"
        HEADER_.NOMBRE_SERVICIO = fx_Completar_Campo(" ", 50, nombreServicio, TYPE_ALINEAR.IZQUIERDA)
        HEADER_.LARGO_MENSAJE = fx_Completar_Campo("0", 9, largoMensaje, TYPE_ALINEAR.DERECHA) '"000002048"

        HEADER_.COD_PAIS = "PE"
        HEADER_.COD_ENTIDAD = "0001"
        HEADER_.COD_CANAL = pCod_Canal '"000006" = Ripleymatico
        HEADER_.COD_OFICINA = fx_Completar_Campo("0", 4, pCod_Oficina, TYPE_ALINEAR.DERECHA) ' 0001

        HEADER_.TERM_FISICO = pTerm_Fisico '"RPMT" = Ripleymatico
        HEADER_.COD_USUARIO = fx_Completar_Campo("0", 8, pCod_Usuario, TYPE_ALINEAR.DERECHA) ' RPLMATIC = Ripleymatico

    End Sub

    Sub Add_Parametros(ByVal params As Parametros)

        PARAMETROS_IN_.sNombreCliente = fx_Completar_Campo(" ", 70, params.sNombreCliente, TYPE_ALINEAR.IZQUIERDA)
        PARAMETROS_IN_.sTipoDocumento = fx_Completar_Campo(" ", 2, params.sTipoDocumento, TYPE_ALINEAR.IZQUIERDA)
        PARAMETROS_IN_.sNumeroDocumento = fx_Completar_Campo(" ", 12, params.sNumeroDocumento, TYPE_ALINEAR.IZQUIERDA)
        PARAMETROS_IN_.sTipoProducto = fx_Completar_Campo("0", 2, params.sTipoProducto, TYPE_ALINEAR.DERECHA)
        PARAMETROS_IN_.sCuenta = fx_Completar_Campo(" ", 12, params.sCuenta, TYPE_ALINEAR.IZQUIERDA)
        PARAMETROS_IN_.nMontoAceptado = fx_Completar_Campo("0", 17, params.nMontoAceptado, TYPE_ALINEAR.DERECHA)
        PARAMETROS_IN_.nPlazoMes = fx_Completar_Campo("0", 2, params.nPlazoMes, TYPE_ALINEAR.DERECHA)
        PARAMETROS_IN_.nCuotaMes = fx_Completar_Campo("0", 17, params.nCuotaMes, TYPE_ALINEAR.DERECHA)
        PARAMETROS_IN_.nTEA = fx_Completar_Campo("0", 5, params.nTEA, TYPE_ALINEAR.DERECHA)
        PARAMETROS_IN_.nTEM = fx_Completar_Campo("0", 5, params.nTEM, TYPE_ALINEAR.DERECHA)
        PARAMETROS_IN_.sMoneda = Utility.Constantes.MONEDA_SOLES
        PARAMETROS_IN_.sFechaVencimiento = fx_Completar_Campo(" ", 10, params.sFechaVencimiento, TYPE_ALINEAR.IZQUIERDA)
        PARAMETROS_IN_.sCanal = fx_Completar_Campo(" ", 2, params.sCanal, TYPE_ALINEAR.IZQUIERDA) 'RM = Ripleymatico
        PARAMETROS_IN_.sSucursal = fx_Completar_Campo("0", 2, params.sSucursal, TYPE_ALINEAR.DERECHA)
        PARAMETROS_IN_.sCodigoVendedor = fx_Completar_Campo("0", 12, params.sCodigoVendedor, TYPE_ALINEAR.DERECHA)
        PARAMETROS_IN_.sCodigoUsuario = fx_Completar_Campo(" ", 12, params.sCodigoUsuario, TYPE_ALINEAR.IZQUIERDA)
        PARAMETROS_IN_.sCodigoAtencion = fx_Completar_Campo(" ", 40, params.sCodigoAtencion, TYPE_ALINEAR.IZQUIERDA)
        PARAMETROS_IN_.sAplicaITF = fx_Completar_Campo(" ", 1, params.sAplicaITF, TYPE_ALINEAR.IZQUIERDA)
        PARAMETROS_IN_.sIndicadorCommit = fx_Completar_Campo(" ", 1, params.sIndicadorCommit, TYPE_ALINEAR.IZQUIERDA) '"S" = tx desde el servicio
        PARAMETROS_IN_.sCentroAlta = fx_Completar_Campo("0", 4, params.sCentroAlta, TYPE_ALINEAR.DERECHA)
        PARAMETROS_IN_.sCodigoEntidad = fx_Completar_Campo("0", 4, params.sCodigoEntidad, TYPE_ALINEAR.DERECHA)
    End Sub

    Public Function Ejecutar_Servicio() As PARAMETROS_OUT

        Dim oPARAMETROS_OUT_ As PARAMETROS_OUT
        Dim oMQCOM As New MQ
        Dim Trama_Respuesta As String = String.Empty
        Dim Trama_Envio As String = String.Empty
        Dim pFecha_TRX As Double = CDbl(DateTime.Now.ToString("ddMMyyyy"))
        Dim pHora_TRX As Double = CDbl(DateTime.Now.ToString("HHmmss"))

        Try

            Trama_Envio = HEADER_.fx_Trama_Header() & PARAMETROS_IN_.fx_Trama_Parametros_IN()
            Registrar_Log(3, HEADER_.TERM_FISICO, "02-01", service.Trim, PARAMETROS_IN_.sCodigoEntidad & PARAMETROS_IN_.sCentroAlta & PARAMETROS_IN_.sCuenta, "IN:" & Trama_Envio)

            Registrar_Log(3, HEADER_.TERM_FISICO, "Antes de asignar service =  " & service)
            oMQCOM.Service = service
            Registrar_Log(3, HEADER_.TERM_FISICO, "Antes de asignar Trama_Envio =  " & Trama_Envio)
            oMQCOM.Message = Trama_Envio

            Registrar_Log(3, HEADER_.TERM_FISICO, "Antes de ejecutar oMQCOM.Execute() ")
            oMQCOM.Execute()
            Registrar_Log(3, HEADER_.TERM_FISICO, "Antes de asignar oMQCOM.Response " & oMQCOM.Response)
            Trama_Respuesta = oMQCOM.Response
            Registrar_Log(3, HEADER_.TERM_FISICO, "Trama_Respuesta " & Trama_Respuesta)
            If oMQCOM.ReasonMd = 0 Then
                Registrar_Log(3, HEADER_.TERM_FISICO, "oMQCOM.ReasonMd = 0 " & oMQCOM.ReasonMd)
                'OK
                Registrar_Log(3, HEADER_.TERM_FISICO, "02-01", service.Trim, PARAMETROS_IN_.sCodigoEntidad & PARAMETROS_IN_.sCentroAlta & PARAMETROS_IN_.sCuenta, "OUT:" & Trama_Respuesta)

                Trama_Respuesta = Trama_Respuesta.Substring(332)
                oPARAMETROS_OUT_ = New PARAMETROS_OUT(Trama_Respuesta)
                oPARAMETROS_OUT_.nWSResult = 1
                oPARAMETROS_OUT_.sWSMessage = "OK"
            Else
                Registrar_Log(3, HEADER_.TERM_FISICO, "oMQCOM.ReasonMd != 0 " & oMQCOM.ReasonMd)
                'ERROR
                oPARAMETROS_OUT_ = New PARAMETROS_OUT()
                Registrar_Log(3, "", "Salir de constructor por defecto")

                'Registrar_Log(3, HEADER_.TERM_FISICO, "02-01", service.Trim, PARAMETROS_IN_.sCodigoEntidad & PARAMETROS_IN_.sCentroAlta & PARAMETROS_IN_.sCuenta, "OUT:" & oPARAMETROS_OUT_.nWSResult.ToString & "," & oPARAMETROS_OUT_.sWSMessage.Trim)

            End If

        Catch ex As Exception
            Registrar_Log(3, HEADER_.TERM_FISICO, "Exception:" & ex.Message)
            oPARAMETROS_OUT_ = New PARAMETROS_OUT()
            'oPARAMETROS_OUT_.COD_RESPUESTA = "-2"
            'oPARAMETROS_OUT_.MSJ_RESPUESTA = ex.Message

            'Registrar_Log(3, HEADER_.TERM_FISICO, "02-01", service.Trim, PARAMETROS_IN_.sCodigoEntidad & PARAMETROS_IN_.sCentroAlta & PARAMETROS_IN_.sCuenta, "OUT:" & oPARAMETROS_OUT_.nWSResult.ToString & "," & oPARAMETROS_OUT_.sWSMessage.Trim)

        End Try

        Return oPARAMETROS_OUT_

    End Function


    Public Class HEADER

#Region "-------------- CONECTOR --------------"

        Private COD_RESPUESTA_ As String
        Private NOMBRE_SERVICIO_ As String
        Private LARGO_MENSAJE_ As String

        Public Property COD_RESPUESTA As String
            Get
                Return COD_RESPUESTA_
            End Get
            Set(ByVal value As String)
                COD_RESPUESTA_ = value
            End Set
        End Property

        Public Property NOMBRE_SERVICIO As String
            Get
                Return NOMBRE_SERVICIO_
            End Get
            Set(ByVal value As String)
                NOMBRE_SERVICIO_ = value
            End Set
        End Property

        Public Property LARGO_MENSAJE As String
            Get
                Return LARGO_MENSAJE_
            End Get
            Set(ByVal value As String)
                LARGO_MENSAJE_ = value
            End Set
        End Property

#End Region

#Region "-------------- CABECERA --------------"

        Private COD_ENTIDAD_ As String
        Private COD_CANAL_ As String
        Private COD_USUARIO_ As String
        Private COD_OFICINA_ As String
        Private TERM_FISICO_ As String
        Private COD_PAIS_ As String

        Public Property COD_ENTIDAD As String
            Get
                Return COD_ENTIDAD_
            End Get
            Set(ByVal value As String)
                COD_ENTIDAD_ = value
            End Set
        End Property

        Public Property COD_CANAL As String
            Get
                Return COD_CANAL_
            End Get
            Set(ByVal value As String)
                COD_CANAL_ = value
            End Set
        End Property

        Public Property COD_USUARIO As String
            Get
                Return COD_USUARIO_
            End Get
            Set(ByVal value As String)
                COD_USUARIO_ = value
            End Set
        End Property

        Public Property COD_OFICINA As String
            Get
                Return COD_OFICINA_
            End Get
            Set(ByVal value As String)
                COD_OFICINA_ = value
            End Set
        End Property

        Public Property TERM_FISICO As String
            Get
                Return TERM_FISICO_
            End Get
            Set(ByVal value As String)
                TERM_FISICO_ = value
            End Set
        End Property

        Public Property COD_PAIS As String
            Get
                Return COD_PAIS_
            End Get
            Set(ByVal value As String)
                COD_PAIS_ = value
            End Set
        End Property

#End Region

#Region "-------------- PAGINACION --------------"

        Private IND_PAGINACION_ As String
        Private SIGUIENTE_ As String
        Private ANTERIOR_ As String
        Private UNITARIA_ As String
        Private RESTO_ As String
        Private CLAVE_INICIO_ As String
        Private CLAVE_FIN_ As String
        Private PANTALLA_PAG_ As String
        Private IND_MAS_DATOS_ As String
        Private OTROS_DATOS_ As String

        Public Property IND_PAGINACION As TYPE_IND_PAGINACION
            Get
                Return IND_PAGINACION_
            End Get
            Set(ByVal value As TYPE_IND_PAGINACION)

                IND_PAGINACION_ = value

            End Set
        End Property

        Public Property CLAVE_INICIO As String
            Get
                Return CLAVE_INICIO_
            End Get
            Set(ByVal value As String)
                CLAVE_INICIO_ = value
            End Set
        End Property

        Public Property CLAVE_FIN As String
            Get
                Return CLAVE_FIN_
            End Get
            Set(ByVal value As String)
                CLAVE_FIN_ = value
            End Set
        End Property

        Public Property PANTALLA_PAG As String
            Get
                Return PANTALLA_PAG_
            End Get
            Set(ByVal value As String)
                PANTALLA_PAG_ = value
            End Set
        End Property

        Public Property IND_MAS_DATOS As String
            Get
                Return IND_MAS_DATOS_
            End Get
            Set(ByVal value As String)
                IND_MAS_DATOS_ = value
            End Set
        End Property

        Public Property OTROS_DATOS As String
            Get
                Return OTROS_DATOS_
            End Get
            Set(ByVal value As String)
                OTROS_DATOS_ = value
            End Set
        End Property


#End Region

        Public Function fx_Trama_Header() As String

            Dim Trama_Header As String = String.Empty

            'CONECTOR
            Trama_Header = COD_RESPUESTA
            Trama_Header = Trama_Header & NOMBRE_SERVICIO
            Trama_Header = Trama_Header & LARGO_MENSAJE

            'CABECERA
            Trama_Header = Trama_Header & COD_PAIS
            Trama_Header = Trama_Header & COD_ENTIDAD
            Trama_Header = Trama_Header & COD_CANAL
            Trama_Header = Trama_Header & COD_OFICINA
            Trama_Header = Trama_Header & TERM_FISICO
            Trama_Header = Trama_Header & COD_USUARIO

            ''PAGINACION
            'Dim strIND_PAGINACION_ As String = " "
            'Select Case IND_PAGINACION

            '    Case TYPE_IND_PAGINACION.ANTERIOR
            '        strIND_PAGINACION_ = "A"
            '    Case TYPE_IND_PAGINACION.SIGUIENTE
            '        strIND_PAGINACION_ = "S"
            '    Case TYPE_IND_PAGINACION.UNITARIA
            '        strIND_PAGINACION_ = "U"
            '    Case TYPE_IND_PAGINACION.RESTO
            '        strIND_PAGINACION_ = " "
            'End Select


            'Trama_Header = Trama_Header & strIND_PAGINACION_
            'Trama_Header = Trama_Header & CLAVE_INICIO
            'Trama_Header = Trama_Header & CLAVE_FIN
            'Trama_Header = Trama_Header & PANTALLA_PAG
            'Trama_Header = Trama_Header & IND_MAS_DATOS
            'Trama_Header = Trama_Header & OTROS_DATOS

            Return Trama_Header


        End Function


    End Class

    Public Class PARAMETROS_IN

        Private _sNombreCliente As String
        Private _sTipoDocumento As String
        Private _sNumeroDocumento As String
        Private _sTipoProducto As String
        Private _sCodigoEntidad As String
        Private _sCentroAlta As String
        Private _sCuenta As String
        Private _nMontoAceptado As String
        Private _nPlazoMes As String
        Private _nCuotaMes As String
        Private _nTEA As String
        Private _nTEM As String
        Private _sMoneda As String
        Private _sFechaVencimiento As String
        Private _sCanal As String
        Private _sSucural As String
        Private _sCodigoVendedor As String
        Private _sCodigoUsuario As String
        Private _sIndicadorCommit As String
        Private _sAplicaITF As String
        Private _sCodigoAtencion As String

        Public Property sNombreCliente As String
            Get
                Return _sNombreCliente
            End Get
            Set(ByVal value As String)
                _sNombreCliente = value
            End Set
        End Property
        Public Property sTipoDocumento As String
            Get
                Return _sTipoDocumento
            End Get
            Set(ByVal value As String)
                _sTipoDocumento = value
            End Set
        End Property
        Public Property sNumeroDocumento As String
            Get
                Return _sNumeroDocumento
            End Get
            Set(ByVal value As String)
                _sNumeroDocumento = value
            End Set
        End Property
        Public Property sTipoProducto() As String
            Get
                Return _sTipoProducto
            End Get
            Set(ByVal value As String)
                _sTipoProducto = value
            End Set
        End Property
        Public Property sCodigoEntidad() As String
            Get
                Return _sCodigoEntidad
            End Get
            Set(ByVal value As String)
                _sCodigoEntidad = value
            End Set
        End Property
        Public Property sCentroAlta() As String
            Get
                Return _sCentroAlta
            End Get
            Set(ByVal value As String)
                _sCentroAlta = value
            End Set
        End Property
        Public Property sCuenta() As String
            Get
                Return _sCuenta
            End Get
            Set(ByVal value As String)
                _sCuenta = value
            End Set
        End Property
        Public Property nMontoAceptado() As String
            Get
                Return _nMontoAceptado
            End Get
            Set(ByVal value As String)
                _nMontoAceptado = value
            End Set
        End Property
        Public Property nPlazoMes() As String
            Get
                Return _nPlazoMes
            End Get
            Set(ByVal value As String)
                _nPlazoMes = value
            End Set
        End Property
        Public Property nCuotaMes() As String
            Get
                Return _nCuotaMes
            End Get
            Set(ByVal value As String)
                _nCuotaMes = value
            End Set
        End Property
        Public Property nTEA() As String
            Get
                Return _nTEA
            End Get
            Set(ByVal value As String)
                _nTEA = value
            End Set
        End Property
        Public Property nTEM() As String
            Get
                Return _nTEM
            End Get
            Set(ByVal value As String)
                _nTEM = value
            End Set
        End Property
        Public Property sMoneda() As String
            Get
                Return _sMoneda
            End Get
            Set(ByVal value As String)
                _sMoneda = value
            End Set
        End Property
        Public Property sFechaVencimiento() As String
            Get
                Return _sFechaVencimiento
            End Get
            Set(ByVal value As String)
                _sFechaVencimiento = value
            End Set
        End Property
        Public Property sCanal() As String
            Get
                Return _sCanal
            End Get
            Set(ByVal value As String)
                _sCanal = value
            End Set
        End Property
        Public Property sSucursal() As String
            Get
                Return _sSucural
            End Get
            Set(ByVal value As String)
                _sSucural = value
            End Set
        End Property
        Public Property sCodigoVendedor() As String
            Get
                Return _sCodigoVendedor
            End Get
            Set(ByVal value As String)
                _sCodigoVendedor = value
            End Set
        End Property
        Public Property sCodigoUsuario() As String
            Get
                Return _sCodigoUsuario
            End Get
            Set(ByVal value As String)
                _sCodigoUsuario = value
            End Set
        End Property
        Public Property sCodigoAtencion() As String
            Get
                Return _sCodigoAtencion
            End Get
            Set(ByVal value As String)
                _sCodigoAtencion = value
            End Set
        End Property
        Public Property sAplicaITF() As String
            Get
                Return _sAplicaITF
            End Get
            Set(ByVal value As String)
                _sAplicaITF = value
            End Set
        End Property
        Public Property sIndicadorCommit() As String
            Get
                Return _sIndicadorCommit
            End Get
            Set(ByVal value As String)
                _sIndicadorCommit = value
            End Set
        End Property

        Public Function fx_Trama_Parametros_IN() As String

            Dim Trama_Parametros_IN As String = String.Empty

            Trama_Parametros_IN = sNombreCliente
            Trama_Parametros_IN = Trama_Parametros_IN & sTipoDocumento
            Trama_Parametros_IN = Trama_Parametros_IN & sNumeroDocumento
            Trama_Parametros_IN = Trama_Parametros_IN & sTipoProducto
            Trama_Parametros_IN = Trama_Parametros_IN & sCodigoEntidad
            Trama_Parametros_IN = Trama_Parametros_IN & sCentroAlta
            Trama_Parametros_IN = Trama_Parametros_IN & sCuenta
            Trama_Parametros_IN = Trama_Parametros_IN & nMontoAceptado
            Trama_Parametros_IN = Trama_Parametros_IN & nPlazoMes
            Trama_Parametros_IN = Trama_Parametros_IN & nCuotaMes
            Trama_Parametros_IN = Trama_Parametros_IN & nTEA
            Trama_Parametros_IN = Trama_Parametros_IN & nTEM
            Trama_Parametros_IN = Trama_Parametros_IN & sMoneda
            Trama_Parametros_IN = Trama_Parametros_IN & sFechaVencimiento
            Trama_Parametros_IN = Trama_Parametros_IN & sCanal
            Trama_Parametros_IN = Trama_Parametros_IN & sSucursal
            Trama_Parametros_IN = Trama_Parametros_IN & sCodigoVendedor
            Trama_Parametros_IN = Trama_Parametros_IN & sCodigoUsuario
            Trama_Parametros_IN = Trama_Parametros_IN & sCodigoAtencion
            Trama_Parametros_IN = Trama_Parametros_IN & sAplicaITF
            Trama_Parametros_IN = Trama_Parametros_IN & sIndicadorCommit

            Return Trama_Parametros_IN

        End Function
    End Class

    Public Class PARAMETROS_OUT


        Sub New(ByVal pTrama_Respuesta As String)

            'nWSResult = pTrama_Respuesta.Substring(0, 1)
            'sWSMessage = pTrama_Respuesta.Substring(1, 2)
            sNumeroCuenta = pTrama_Respuesta.Substring(0, 20)
            sNumeroTarjeta = pTrama_Respuesta.Substring(21, 22)
        End Sub

        Sub New()
            Registrar_Log(3, "", "Entro a constructor por defecto")
        End Sub


        Private _nWSResult As Integer
        Public Property nWSResult() As Integer
            Get
                Return _nWSResult
            End Get
            Set(ByVal value As Integer)
                _nWSResult = value
            End Set
        End Property
        Private _sWSMessage As String
        Public Property sWSMessage() As String
            Get
                Return _sWSMessage
            End Get
            Set(ByVal value As String)
                _sWSMessage = value
            End Set
        End Property
        Private _sNumeroCuenta As String
        Public Property sNumeroCuenta() As String
            Get
                Return _sNumeroCuenta
            End Get
            Set(ByVal value As String)
                _sNumeroCuenta = value
            End Set
        End Property
        Private _sNumeroTarjeta As String
        Public Property sNumeroTarjeta() As String
            Get
                Return _sNumeroTarjeta
            End Get
            Set(ByVal value As String)
                _sNumeroTarjeta = value
            End Set
        End Property

    End Class


    Public Function Extraer_Trama() As String

        Dim Trama_Envio As String = String.Empty
        Try
            Trama_Envio = HEADER_.fx_Trama_Header() & PARAMETROS_IN_.fx_Trama_Parametros_IN()
        Catch ex As Exception
            Trama_Envio = "Extraer_Trama: " + ex.Message
        End Try

        Return Trama_Envio

    End Function

End Class
