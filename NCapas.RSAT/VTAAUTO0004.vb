Imports NCapas.Utility.Funciones
Imports NCapas.Utility.Tipos
Imports NCapas.Utility.Log
Imports MQCOMLib

Public Class VTAAUTO0004

    Private HEADER_ As New HEADER
    Private PARAMETROS_IN_ As New PARAMETROS_IN
    Private PARAMETROS_OUT_ As PARAMETROS_OUT
    Private Const SERVICIO As String = "VTAAUTO0004"

    Sub New(ByVal pCod_Usuario As String,
            ByVal pCod_Oficina As String,
            ByVal pFecha_Conta As Date)

        HEADER_.COD_RESPUESTA = "0000000000"
        HEADER_.NOMBRE_SERVICIO = fx_Completar_Campo(" ", 50, SERVICIO, TYPE_ALINEAR.IZQUIERDA)
        HEADER_.LARGO_MENSAJE = "000002048"

        HEADER_.COD_ENTIDAD = "0001"
        HEADER_.TIPO_TERMINAL = "000006"
        HEADER_.COD_USUARIO = fx_Completar_Campo("0", 8, pCod_Usuario, TYPE_ALINEAR.DERECHA)
        HEADER_.COD_OFICINA = fx_Completar_Campo("0", 4, pCod_Oficina, TYPE_ALINEAR.DERECHA)
        'RMIL = RIPLEYMATICO INCREMENTO LINEA
        HEADER_.TERM_FISICO = "RMIL"
        HEADER_.FECHA_CONTA = pFecha_Conta.Year.ToString("0000") & "-" & pFecha_Conta.Month.ToString("00") & "-" & pFecha_Conta.Day.ToString("00")
        HEADER_.COD_APLICACION = "RM"
        HEADER_.COD_PAIS = "PE"
        HEADER_.IND_COMMIT = "S"

        HEADER_.IND_PAGINACION = TYPE_IND_PAGINACION.RESTO
        HEADER_.CLAVE_INICIO = New String(" ", 200)
        HEADER_.CLAVE_FIN = New String(" ", 200)
        HEADER_.PANTALLA_PAG = "000"
        HEADER_.IND_MAS_DATOS = New String(" ", 1)
        HEADER_.OTROS_DATOS = fx_Completar_Campo(" ", 90, "ES", TYPE_ALINEAR.IZQUIERDA)

    End Sub

    Sub Add_Parametros(ByVal pNRO_CONTRATO As String, _
                       ByVal pLINEA_EGM As String, _
                       ByVal pLINEA_1 As String, _
                       ByVal pLINEA_2 As String, _
                       ByVal pCOD_SUC As String, _
                       ByVal pNUM_CAJ As String, _
                       ByVal pNUM_TRX As String)

        pNRO_CONTRATO = pNRO_CONTRATO.Trim

        pLINEA_EGM = pLINEA_EGM.Trim & "00"
        pLINEA_1 = pLINEA_1.Trim & "00"
        pLINEA_2 = pLINEA_2.Trim & "00"

        PARAMETROS_IN_.CODENT = pNRO_CONTRATO.Substring(0, 4)
        PARAMETROS_IN_.CENTALTA = pNRO_CONTRATO.Substring(4, 4)
        PARAMETROS_IN_.CUENTA = pNRO_CONTRATO.Substring(8)

        PARAMETROS_IN_.CLAMON = "604"
        PARAMETROS_IN_.LIMCRECTA = fx_Completar_Campo("0", 17, pLINEA_EGM, TYPE_ALINEAR.DERECHA)
        PARAMETROS_IN_.LIMCRELNA_1 = fx_Completar_Campo("0", 17, pLINEA_1, TYPE_ALINEAR.DERECHA)
        PARAMETROS_IN_.LIMCRELNA_2 = fx_Completar_Campo("0", 17, pLINEA_2, TYPE_ALINEAR.DERECHA)

        PARAMETROS_IN_.COD_SUC = fx_Completar_Campo("0", 2, pCOD_SUC, TYPE_ALINEAR.DERECHA)
        PARAMETROS_IN_.NUM_CAJ = fx_Completar_Campo(" ", 20, pNUM_CAJ, TYPE_ALINEAR.IZQUIERDA)
        PARAMETROS_IN_.NUM_TRX = fx_Completar_Campo("0", 8, pNUM_TRX, TYPE_ALINEAR.DERECHA)

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
            Registrar_Log(3, "RMIL", "02-01", SERVICIO.Trim, PARAMETROS_IN_.CODENT & PARAMETROS_IN_.CENTALTA & PARAMETROS_IN_.CUENTA, "IN:" & Trama_Envio)

            oMQCOM.Service = SERVICIO
            oMQCOM.Message = Trama_Envio

            oMQCOM.Execute()
            Trama_Respuesta = oMQCOM.Response

            '<INI DATA TEST>
            'oMQCOM.ReasonMd = 0
            'Select Case pOferta.NUM_CONTRATO

            '    Case "00010001000000268926"
            '        Trama_Respuesta = "0000000000VTAAUTO0004                                       000002048000100000200000000000700X82014-06-10PCPES                                                                                                                                                                                                                                                                                                                                                                                                                 000 ES                                                                                        000100010000002689266040000000000082000000000000000820000000000000008200000200010000000000063000000000000000630000000000000006300001111112014-05-27-03.03.12.0109922014-04-29-07.57.25.0286452014-04-29-07.57.25.0286450500SE ACTUALIZARON LAS SIGUIENTES LINEAS: 00-LINEA_TOTAL. 01-CONSUMO. 02-EFECTIVO.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                         "
            '    Case "00010001000005153852"
            '        Trama_Respuesta = "0000000000VTAAUTO0004                                       000002048000100000200000000002000X82014-06-10PCPES                                                                                                                                                                                                                                                                                                                                                                                                                 000 ES                                                                                        000100010000051538526040000000000020000000000000000200000000000000000000001000010000000000015000000000000000150000000000000000750001111102014-06-10-01.16.31.0045912014-06-10-01.16.31.004591                          0500SE ACTUALIZARON LAS SIGUIENTES LINEAS: 00-LINEA_TOTAL. 01-CONSUMO.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                      "
            '    Case Else
            '        Trama_Respuesta = "0000000000VTAAUTO0004                                       000002048000100000200000000000700X82014-06-10PCPES                                                                                                                                                                                                                                                                                                                                                                                                                 000 ES                                                                                        000100010000000902496040000000000052000000000000000520000000000000005200000200010000000000052000000000000000520000000000000005200001110002014-06-10-18.59.20.068360                                                    0298ERR: NO SE PUDIERON INCREMENTAR LAS LINEAS DE CREDITO; YA QUE ESTAS NO SON SUPERIORES A LAS ACTUALES                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                    "
            'End Select
            '<FIN DATA TEST>

            If oMQCOM.ReasonMd = 0 Then

                'OK
                Registrar_Log(3, "RMIL", "02-01", SERVICIO.Trim, PARAMETROS_IN_.CODENT & PARAMETROS_IN_.CENTALTA & PARAMETROS_IN_.CUENTA, "OUT:" & Trama_Respuesta)

                Trama_Respuesta = Trama_Respuesta.Substring(709)
                oPARAMETROS_OUT_ = New PARAMETROS_OUT(Trama_Respuesta)
                oPARAMETROS_OUT_.COD_RESPUESTA = oPARAMETROS_OUT_.CODRPTA
                oPARAMETROS_OUT_.MSJ_RESPUESTA = oPARAMETROS_OUT_.DESRPTA

            Else
                'ERROR
                oPARAMETROS_OUT_ = New PARAMETROS_OUT()
                oPARAMETROS_OUT_.COD_RESPUESTA = oMQCOM.ReasonMd
                oPARAMETROS_OUT_.MSJ_RESPUESTA = oMQCOM.DescReasonMd

                Registrar_Log(3, "RMIL", "02-01", SERVICIO.Trim, PARAMETROS_IN_.CODENT & PARAMETROS_IN_.CENTALTA & PARAMETROS_IN_.CUENTA, "OUT:" & oPARAMETROS_OUT_.COD_RESPUESTA.Trim & "," & oPARAMETROS_OUT_.MSJ_RESPUESTA.Trim)

            End If

        Catch ex As Exception
            oPARAMETROS_OUT_ = New PARAMETROS_OUT()
            oPARAMETROS_OUT_.COD_RESPUESTA = "-2"
            oPARAMETROS_OUT_.MSJ_RESPUESTA = ex.Message

            Registrar_Log(3, "RMIL", "02-01", SERVICIO.Trim, PARAMETROS_IN_.CODENT & PARAMETROS_IN_.CENTALTA & PARAMETROS_IN_.CUENTA, "OUT:" & oPARAMETROS_OUT_.COD_RESPUESTA.Trim & "," & oPARAMETROS_OUT_.MSJ_RESPUESTA.Trim)

        End Try

        Return oPARAMETROS_OUT_

    End Function


    Class HEADER

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
        Private TIPO_TERMINAL_ As String
        Private COD_USUARIO_ As String
        Private COD_OFICINA_ As String
        Private TERM_FISICO_ As String
        Private FECHA_CONTA_ As String
        Private COD_APLICACION_ As String
        Private COD_PAIS_ As String
        Private IND_COMMIT_ As String

        Public Property COD_ENTIDAD As String
            Get
                Return COD_ENTIDAD_
            End Get
            Set(ByVal value As String)
                COD_ENTIDAD_ = value
            End Set
        End Property

        Public Property TIPO_TERMINAL As String
            Get
                Return TIPO_TERMINAL_
            End Get
            Set(ByVal value As String)
                TIPO_TERMINAL_ = value
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

        Public Property FECHA_CONTA As String
            Get
                Return FECHA_CONTA_
            End Get
            Set(ByVal value As String)
                FECHA_CONTA_ = value
            End Set
        End Property

        Public Property COD_APLICACION As String
            Get
                Return COD_APLICACION_
            End Get
            Set(ByVal value As String)
                COD_APLICACION_ = value
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

        Public Property IND_COMMIT As String
            Get
                Return IND_COMMIT_
            End Get
            Set(ByVal value As String)
                IND_COMMIT_ = value
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
            Trama_Header = Trama_Header & COD_ENTIDAD
            Trama_Header = Trama_Header & TIPO_TERMINAL
            Trama_Header = Trama_Header & COD_USUARIO
            Trama_Header = Trama_Header & COD_OFICINA
            Trama_Header = Trama_Header & TERM_FISICO
            Trama_Header = Trama_Header & FECHA_CONTA
            Trama_Header = Trama_Header & COD_APLICACION
            Trama_Header = Trama_Header & COD_PAIS
            Trama_Header = Trama_Header & IND_COMMIT

            'PAGINACION
            Dim strIND_PAGINACION_ As String = " "
            Select Case IND_PAGINACION

                Case TYPE_IND_PAGINACION.ANTERIOR
                    strIND_PAGINACION_ = "A"
                Case TYPE_IND_PAGINACION.SIGUIENTE
                    strIND_PAGINACION_ = "S"
                Case TYPE_IND_PAGINACION.UNITARIA
                    strIND_PAGINACION_ = "U"
                Case TYPE_IND_PAGINACION.RESTO
                    strIND_PAGINACION_ = " "
            End Select


            Trama_Header = Trama_Header & strIND_PAGINACION_
            Trama_Header = Trama_Header & CLAVE_INICIO
            Trama_Header = Trama_Header & CLAVE_FIN
            Trama_Header = Trama_Header & PANTALLA_PAG
            Trama_Header = Trama_Header & IND_MAS_DATOS
            Trama_Header = Trama_Header & OTROS_DATOS

            Return Trama_Header


        End Function


    End Class

    Class PARAMETROS_IN

        Private CODENT_ As String
        Private CENTALTA_ As String
        Private CUENTA_ As String
        Private CLAMON_ As String
        Private LIMCRECTA_ As String
        Private LIMCRELNA_1_ As String
        Private LIMCRELNA_2_ As String

        Private COD_SUC_ As String
        Private NUM_CAJ_ As String
        Private NUM_TRX_ As String

        Public Property CODENT As String
            Get
                Return CODENT_
            End Get
            Set(ByVal value As String)
                CODENT_ = value
            End Set
        End Property

        Public Property CENTALTA As String
            Get
                Return CENTALTA_
            End Get
            Set(ByVal value As String)
                CENTALTA_ = value
            End Set
        End Property

        Public Property CUENTA As String
            Get
                Return CUENTA_
            End Get
            Set(ByVal value As String)
                CUENTA_ = value
            End Set
        End Property

        Public Property CLAMON As String
            Get
                Return CLAMON_
            End Get
            Set(ByVal value As String)
                CLAMON_ = value
            End Set
        End Property

        Public Property LIMCRECTA As String
            Get
                Return LIMCRECTA_
            End Get
            Set(ByVal value As String)
                LIMCRECTA_ = value
            End Set
        End Property

        Public Property LIMCRELNA_1 As String
            Get
                Return LIMCRELNA_1_
            End Get
            Set(ByVal value As String)
                LIMCRELNA_1_ = value
            End Set
        End Property

        Public Property LIMCRELNA_2 As String
            Get
                Return LIMCRELNA_2_
            End Get
            Set(ByVal value As String)
                LIMCRELNA_2_ = value
            End Set
        End Property

        Public Function fx_Trama_Parametros_IN() As String

            Dim Trama_Parametros_IN As String = String.Empty

            Trama_Parametros_IN = CODENT_
            Trama_Parametros_IN = Trama_Parametros_IN & CENTALTA_
            Trama_Parametros_IN = Trama_Parametros_IN & CUENTA_
            Trama_Parametros_IN = Trama_Parametros_IN & CLAMON_
            Trama_Parametros_IN = Trama_Parametros_IN & LIMCRECTA_
            Trama_Parametros_IN = Trama_Parametros_IN & LIMCRELNA_1_
            Trama_Parametros_IN = Trama_Parametros_IN & LIMCRELNA_2_

            Trama_Parametros_IN = Trama_Parametros_IN & COD_SUC_
            Trama_Parametros_IN = Trama_Parametros_IN & NUM_CAJ_
            Trama_Parametros_IN = Trama_Parametros_IN & NUM_TRX_

            Return Trama_Parametros_IN

        End Function

        Public Property COD_SUC As String
            Get
                Return COD_SUC_
            End Get
            Set(ByVal value As String)
                COD_SUC_ = value
            End Set
        End Property

        Public Property NUM_CAJ As String
            Get
                Return NUM_CAJ_
            End Get
            Set(ByVal value As String)
                NUM_CAJ_ = value
            End Set
        End Property

        Public Property NUM_TRX As String
            Get
                Return NUM_TRX_
            End Get
            Set(ByVal value As String)
                NUM_TRX_ = value
            End Set
        End Property

    End Class

    Class PARAMETROS_OUT


        Sub New(ByVal pTrama_Respuesta As String)

            PRODUCTO_ = pTrama_Respuesta.Substring(0, 2)
            SUBPRODU_ = pTrama_Respuesta.Substring(2, 4)
            LIMCRECTA_OLD_ = pTrama_Respuesta.Substring(6, 17)
            LIMCRELNA_OLD_1_ = pTrama_Respuesta.Substring(23, 17)
            LIMCRELNA_OLD_2_ = pTrama_Respuesta.Substring(40, 17)

            LIMCRECTA_EXI_ = pTrama_Respuesta.Substring(57, 1)
            LIMCRELNA_EXI_1_ = pTrama_Respuesta.Substring(58, 1)
            LIMCRELNA_EXI_2_ = pTrama_Respuesta.Substring(59, 1)

            UPD_LIM_OK_ = pTrama_Respuesta.Substring(60, 1)
            UPD_LIN_OK_1_ = pTrama_Respuesta.Substring(61, 1)
            UPD_LIN_OK_2_ = pTrama_Respuesta.Substring(62, 1)

            CONTCUR_ = pTrama_Respuesta.Substring(63, 26)
            CONTCUR_1_ = pTrama_Respuesta.Substring(89, 26)
            CONTCUR_2_ = pTrama_Respuesta.Substring(115, 26)

            PASO_ = pTrama_Respuesta.Substring(141, 2)
            CODRPTA = pTrama_Respuesta.Substring(143, 2)
            DESRPTA_ = pTrama_Respuesta.Substring(145).Trim

        End Sub

        Sub New()

        End Sub


        Public Property PRODUCTO_ As String
        Private SUBPRODU_ As String
        Private LIMCRECTA_OLD_ As String
        Private LIMCRELNA_OLD_1_ As String
        Private LIMCRELNA_OLD_2_ As String
        Private LIMCRECTA_EXI_ As String
        Private LIMCRELNA_EXI_1_ As String
        Private LIMCRELNA_EXI_2_ As String
        Private UPD_LIM_OK_ As String
        Private UPD_LIN_OK_1_ As String
        Private UPD_LIN_OK_2_ As String
        Private CONTCUR_ As String
        Private CONTCUR_1_ As String
        Private CONTCUR_2_ As String
        Private PASO_ As String
        Private CODRPTA_ As String
        Private DESRPTA_ As String

        Public Property PRODUCTO As String
            Get
                Return PRODUCTO_
            End Get
            Set(ByVal value As String)
                PRODUCTO_ = value
            End Set
        End Property

        Public Property SUBPRODU As String
            Get
                Return SUBPRODU_
            End Get
            Set(ByVal value As String)
                SUBPRODU_ = value
            End Set
        End Property

        Public Property LIMCRECTA_OLD As String
            Get
                Return LIMCRECTA_OLD_
            End Get
            Set(ByVal value As String)
                LIMCRECTA_OLD_ = value
            End Set
        End Property

        Public Property LIMCRELNA_OLD_1 As String
            Get
                Return LIMCRELNA_OLD_1_
            End Get
            Set(ByVal value As String)
                LIMCRELNA_OLD_1_ = value
            End Set
        End Property

        Public Property LIMCRELNA_OLD_2 As String
            Get
                Return LIMCRELNA_OLD_2_
            End Get
            Set(ByVal value As String)
                LIMCRELNA_OLD_2_ = value
            End Set
        End Property

        Public Property LIMCRECTA_EXI As String
            Get
                Return LIMCRECTA_EXI_
            End Get
            Set(ByVal value As String)
                LIMCRECTA_EXI_ = value
            End Set
        End Property

        Public Property LIMCRELNA_EXI_1 As String
            Get
                Return LIMCRELNA_EXI_1_
            End Get
            Set(ByVal value As String)
                LIMCRELNA_EXI_1_ = value
            End Set
        End Property

        Public Property LIMCRELNA_EXI_2 As String
            Get
                Return LIMCRELNA_EXI_2_
            End Get
            Set(ByVal value As String)
                LIMCRELNA_EXI_2_ = value
            End Set
        End Property

        Public Property UPD_LIM_OK As String
            Get
                Return UPD_LIM_OK_
            End Get
            Set(ByVal value As String)
                UPD_LIM_OK_ = value
            End Set
        End Property

        Public Property UPD_LIN_OK_1 As String
            Get
                Return LIMCRELNA_EXI_1_
            End Get
            Set(ByVal value As String)
                LIMCRELNA_EXI_1_ = value
            End Set
        End Property

        Public Property UPD_LIN_OK_2 As String
            Get
                Return LIMCRELNA_EXI_2_
            End Get
            Set(ByVal value As String)
                LIMCRELNA_EXI_2_ = value
            End Set
        End Property

        Public Property CONTCUR As String
            Get
                Return CONTCUR_
            End Get
            Set(ByVal value As String)
                CONTCUR_ = value
            End Set
        End Property

        Public Property CONTCUR_1 As String
            Get
                Return CONTCUR_1_
            End Get
            Set(ByVal value As String)
                CONTCUR_1_ = value
            End Set
        End Property

        Public Property CONTCUR_2 As String
            Get
                Return CONTCUR_2_
            End Get
            Set(ByVal value As String)
                CONTCUR_2_ = value
            End Set
        End Property

        Public Property PASO As String
            Get
                Return PASO_
            End Get
            Set(ByVal value As String)
                PASO_ = value
            End Set
        End Property

        Public Property CODRPTA As String
            Get
                Return CODRPTA_
            End Get
            Set(ByVal value As String)
                CODRPTA_ = value
            End Set
        End Property

        Public Property DESRPTA As String
            Get
                Return DESRPTA_
            End Get
            Set(ByVal value As String)
                DESRPTA_ = value
            End Set
        End Property

        Private COD_RESPUESTA_ As String = "-1"
        Private MSJ_RESPUESTA_ As String = "SERVICIO NO INVOCADO"

        Public Property COD_RESPUESTA As String
            Get
                Return COD_RESPUESTA_
            End Get
            Set(ByVal value As String)
                COD_RESPUESTA_ = value
            End Set
        End Property

        Public Property MSJ_RESPUESTA As String
            Get
                Return MSJ_RESPUESTA_
            End Get
            Set(ByVal value As String)
                MSJ_RESPUESTA_ = value
            End Set
        End Property


    End Class

End Class
