Imports System.Web
Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.Configuration
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OracleClient
Imports System.Xml
Imports System.Net.Sockets
Imports System.Text
Imports MIRRCOM
Imports Channel
Imports Microsoft.VisualBasic

'<INI>
Imports MQCOMLib
Imports System.Collections.Generic
Imports System.Globalization
Imports System.Threading
Imports System

'<FIN>

'<INI PRY-0794-01 - Danny Herrera 01/08/2014
'Imports BusinessEntities
'Imports BusinessLogic
'<FIN PRY-0794-01 - Danny Herrera 01/08/2014

'<INI PRY-0794-01 - Danny Herrera 01/08/2014
Imports NCapas.Business
Imports NCapas.Entity
Imports NCapas.Entity.IncrementoLinea
Imports NCapas.Utility
Imports NCapas.Utility.Log
Imports NCapas.Utility.Tipos
Imports NCapas.Utility.Funciones
Imports System.Net.Mail

'<FIN PRY-0794-01 - Danny Herrera 01/08/2014

<WebService(Namespace:="http://tempuri.org/")> _
<WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Public Class Service
    Inherits System.Web.Services.WebService

    Private OpLog_Registro As String = ConfigurationSettings.AppSettings.Get("LOG_REGISTRO")
    Private SERVIDOR_MIRROR As String = ConfigurationSettings.AppSettings.Get("SERVIDOR_MIRROR")
    Private NOMBRE_APLICACION As String = ConfigurationSettings.AppSettings.Get("NOMBRE_APLICACION")
    Private USUARIO As String = ConfigurationSettings.AppSettings.Get("USUARIO")
    Private PROTOCOLO_TCP As String = ConfigurationSettings.AppSettings.Get("PROTOCOLO_TCP")
    Private TERMINAL As String = ConfigurationSettings.AppSettings.Get("TERMINAL")
    Private OpLog_Consulta As String = ConfigurationSettings.AppSettings.Get("LOG_CONSULTAS")

    Private TIME_OUT_SERVER As Long = CLng(ConfigurationSettings.AppSettings.Get("TIME_OUT_SERVER")) 'TIME OUT
    Private SERVIDOR_MIRROR_DESTINO As String = ConfigurationSettings.AppSettings.Get("SERVIDOR_MIRROR_DESTINO") 'destination
    Private SERVIDOR_MIRROR_NODE As String = ConfigurationSettings.AppSettings.Get("SERVIDOR_MIRROR_NODE") 'Mirror Node
    Private PUERTO As String = ConfigurationSettings.AppSettings.Get("PUERTO") 'PUERTO
    Private PRIORIDAD_S As String = ConfigurationSettings.AppSettings.Get("PRIORIDAD") 'PUERTO
    Private sXML_SQL As String = ""
    Private m_ssql As String = ""
    Private sResultado_SQL As String = ""
    Private sMensajeError_SQL As String = ""
    Private oConexion As SqlClient.SqlConnection

    'VARIABLES ORACLE FISA
    Private mCadenaConexion_ORA As String = "User ID=" & ConfigurationManager.AppSettings("USER_NAME_ORA") & ";Data Source=" + ConfigurationManager.AppSettings("SERVICE_NAME_ORA") & ";Password=" & ConfigurationManager.AppSettings("PASSWORD_ORA")
    'CONSULTAR TIPO DOC Y NRO DOC DE MC Y VISA
    Private mCadenaConexion_ORA_X As String = "User ID=" & ConfigurationManager.AppSettings("USER_NAME_ORAX") & ";Data Source=" + ConfigurationManager.AppSettings("SERVICE_NAME_ORAX") & ";Password=" & ConfigurationManager.AppSettings("PASSWORD_ORAX")
    'CONSULTAR DESCRIPCION DE ESTABLECIMIENTO ORA Y CONSULTA DE RIPLEY PUNTOS SOLICITAR USUARIO Y PASS CON TODOS LOS PERMISOS A LOS PAQUETES Y PROCEDIMIENTOS.
    Private mCadenaConexion_ORA_XX As String = "User ID=" & ConfigurationManager.AppSettings("USER_NAME_ORAXX") & ";Data Source=" + ConfigurationManager.AppSettings("SERVICE_NAME_ORAXX") & ";Password=" & ConfigurationManager.AppSettings("PASSWORD_ORAXX")
    'CONSULTA DE PUNTOS
    Private mCadenaConexion_ORA_PUNTOS As String = "User ID=" & ConfigurationManager.AppSettings("USER_NAME_ORA_PUNTOS") & ";Data Source=" + ConfigurationManager.AppSettings("SERVICE_NAME_ORA_PUNTOS") & ";Password=" & ConfigurationManager.AppSettings("PASSWORD_ORA_PUNTOS")
    'CONSULTA DE SUPEREFECTIVO SEF
    Private mCadenaConexion_ORA_SEF As String = "User ID=" & ConfigurationManager.AppSettings("USER_NAME_ORA_SEF") & ";Data Source=" + ConfigurationManager.AppSettings("SERVICE_NAME_ORA_SEF") & ";Password=" & ConfigurationManager.AppSettings("PASSWORD_ORA_SEF")
    'CONSULTA DE INCREMENTO DE LINEA
    Private cadenaConexionOraInc As String = "User ID=" & ConfigurationManager.AppSettings("USER_NAME_ORA_INC") & ";Data Source=" + ConfigurationManager.AppSettings("SERVICE_NAME_ORA_INC") & ";Password=" & ConfigurationManager.AppSettings("PASSWORD_ORA_INC")

    Private sXML_ORA As String = ""
    Private m_SQL_ORA As String = ""
    Private sResultado_ORA As String = ""
    Private sMensajeError_ORA As String = ""

    Public Shared p_clasica As String
    Public Shared p_asosiada As String

    Public Shared c_TablaBloqueo As String(,)
    Public Shared c_TipoFactura As String(,)
    Public Shared c_OrdenCore As String(,)


    <WebMethod(Description:="Prueba de conexión.")> _
    Public Function PRUEBA_CONEXION_SQL() As String
        Dim result As String = String.Empty
        result = BNConexion.Instancia.PRUEBA_CONEXION_SQL()
        Return result
    End Function

    'OBTENER LAS VARIABLES DE TIEMPO DE SESSION
    <WebMethod(Description:="Obtener variables de timeout.")> _
    Public Function OBTENER_VARIABLES_TIMEOUT() As String
        Dim result As String = String.Empty
        result = BNConexion.Instancia.OBTENER_VARIABLES_TIMEOUT()
        Return result
        'Try
        '    'Realizar Conexion a la base de datos
        '    sMensajeError_SQL = ""
        '    oConexion = SQL_ConnectionOpen(Get_CadenaConexion(), sMensajeError_SQL)
        '    If sMensajeError_SQL <> "" Then
        '        sResultado_SQL = "ERROR:No se pudo conectar al servidor de base de datos."
        '    Else
        '        If oConexion.State = ConnectionState.Open Then
        '            m_ssql = "SELECT TIEMPO_DOC,TIEMPO_OPCIONES,NRO_ERROR_TARJETA FROM TIEMPO"
        '            Dim cmdx As SqlClient.SqlCommand = oConexion.CreateCommand
        '            cmdx.CommandTimeout = 900
        '            cmdx.CommandType = CommandType.Text
        '            cmdx.CommandText = m_ssql
        '            Dim rd_time As SqlClient.SqlDataReader = cmdx.ExecuteReader
        '            sXML_SQL = ""
        '            If rd_time.Read = True Then
        '                'ARMAR CADENA
        '                sXML_SQL = IIf(IsDBNull(rd_time.Item("TIEMPO_DOC")), "0", rd_time.Item("TIEMPO_DOC").ToString) & "|\t|" & IIf(IsDBNull(rd_time.Item("TIEMPO_OPCIONES")), "0", rd_time.Item("TIEMPO_OPCIONES").ToString) & "|\t|" & IIf(IsDBNull(rd_time.Item("NRO_ERROR_TARJETA")), "0", rd_time.Item("NRO_ERROR_TARJETA").ToString)
        '            Else
        '                sXML_SQL = "5" & "|\t|" & "10" & "|\t|" & "5"
        '            End If
        '            sResultado_SQL = sXML_SQL
        '            rd_time.Close()
        '            rd_time = Nothing
        '            cmdx.Dispose()
        '            cmdx = Nothing
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

    'OBTENER EL CODIGO DE SUCURSAL DE BANCO
    <WebMethod(Description:="Buscar el codigo de sucursal de banco")> _
    Public Function BUSCAR_COD_SUCURSAL_BANCO(ByVal sIDSucursalKiosco As String) As String
        Dim codigo As String = String.Empty
        Dim sucursal As Sucursal
        sucursal = BNSucursal.Instancia.BuscarSucursalPorId(sIDSucursalKiosco)
        If sucursal Is Nothing Then
            codigo = ""
        Else
            codigo = sucursal.Cod_Sucursal_Banco
        End If
        Return codigo
        'Try
        '    'Realizar Conexion a la base de datos
        '    sMensajeError_SQL = ""
        '    oConexion = SQL_ConnectionOpen(Get_CadenaConexion(), sMensajeError_SQL)
        '    If sMensajeError_SQL <> "" Then
        '        sResultado_SQL = "ERROR:No se pudo conectar al servidor de base de datos."
        '    Else
        '        If oConexion.State = ConnectionState.Open Then
        '            m_ssql = "SELECT COD_SUCURSAL_BANCO FROM dbo.KIO_SUCURSAL WHERE ID='" & sIDSucursalKiosco.Trim & "'"
        '            Dim cmd_suc As SqlClient.SqlCommand = oConexion.CreateCommand
        '            cmd_suc.CommandTimeout = 600
        '            cmd_suc.CommandType = CommandType.Text
        '            cmd_suc.CommandText = m_ssql.Trim
        '            Dim rd_cod As SqlClient.SqlDataReader = cmd_suc.ExecuteReader
        '            sXML_SQL = ""
        '            If rd_cod.Read = True Then
        '                'ARMAR CADENA
        '                sXML_SQL = IIf(IsDBNull(rd_cod.Item("COD_SUCURSAL_BANCO")), "", rd_cod.Item("COD_SUCURSAL_BANCO").ToString)
        '            End If
        '            sResultado_SQL = sXML_SQL
        '            rd_cod.Close()
        '            rd_cod = Nothing
        '            cmd_suc.Dispose()
        '            cmd_suc = Nothing
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


    'OBTENER EL CODIGO DEL KIOSCO
    <WebMethod(Description:="Buscar el codigo del Kiosco")> _
    Public Function BUSCAR_ID_KIOSCO(ByVal sNombreKiosco As String) As String
        Dim result As String = String.Empty
        Dim kiosco As Kiosco
        kiosco = BNKiosco.Instancia.BuscarKioscoPorNombre(sNombreKiosco)
        If kiosco Is Nothing Then
            result = "0"
        Else
            result = kiosco.IdKiosco.ToString()
        End If
        Return result
        'Try
        '    sMensajeError_SQL = ""
        '    oConexion = SQL_ConnectionOpen(Get_CadenaConexion(), sMensajeError_SQL)
        '    If sMensajeError_SQL <> "" Then
        '        sResultado_SQL = "ERROR:No se pudo conectar al servidor de base de datos."
        '    Else
        '        If oConexion.State = ConnectionState.Open Then
        '            m_ssql = "SELECT IDKIOSKO FROM KIO_NOMBRE_KIOSKO WHERE NOMBRE_KIOSKO='" & sNombreKiosco.Trim & "'"
        '            Dim cmd_kio As SqlClient.SqlCommand = oConexion.CreateCommand
        '            cmd_kio.CommandTimeout = 600
        '            cmd_kio.CommandType = CommandType.Text
        '            cmd_kio.CommandText = m_ssql.Trim
        '            Dim rd_cod As SqlClient.SqlDataReader = cmd_kio.ExecuteReader
        '            sXML_SQL = ""
        '            If rd_cod.Read = True Then
        '                sXML_SQL = IIf(IsDBNull(rd_cod.Item("IDKIOSKO")), "0", rd_cod.Item("IDKIOSKO").ToString)
        '            Else
        '                sXML_SQL = "0"
        '            End If
        '            sResultado_SQL = sXML_SQL
        '            rd_cod.Close()
        '            rd_cod = Nothing
        '            cmd_kio.Dispose()
        '            cmd_kio = Nothing
        '            oConexion.Close()
        '            oConexion = Nothing
        '        Else
        '            sResultado_SQL = "0"
        '            oConexion = Nothing
        '        End If
        '    End If
        'Catch ex As Exception
        '    sResultado_SQL = "0"
        '    oConexion = Nothing
        'End Try
        'Return sResultado_SQL
    End Function


    'BUSCAR LAS TASAS PARA EL SIMULADOR OFERTA SUPER EFECTIVO
    '12|\t|36|\t|6.90|\t|1.70|\t|4.00|\t|1.99|\t|26.68|\t|1500
    <WebMethod(Description:="Consultar los Datos para el Calculo de la Oferta SuperEfectivo ")> _
    Public Function BUSCAR_DATOS_SIMULADOR_OFERTA_SEF(ByVal sTasaTEM As String, ByVal sNroTarjeta As String) As String

        Try
            'Realizar Conexion a la base de datos
            sMensajeError_SQL = ""
            Dim sTipoTarjetaSimulador As String = ""

            sTipoTarjetaSimulador = FUN_BUSCAR_TIPO_TARJETA_SIMULADOR_OFERTA(sNroTarjeta.Trim)

            oConexion = SQL_ConnectionOpen(Get_CadenaConexion(), sMensajeError_SQL)

            If sMensajeError_SQL <> "" Then
                sResultado_SQL = ""
            Else

                If oConexion.State = ConnectionState.Open Then

                    m_ssql = "SELECT C.PLAZO_MIN,C.PLAZO_MAX,C.ENVIO_EECC,C.SEG_DESG,C.SEG_PROT,C.TEM,D.TEA,C.MONTO_MIN FROM KIO_CAB_TEM_SIMULADOR C "
                    m_ssql = m_ssql & " INNER JOIN KIO_DET_TCEA_SIMULADOR D ON C.IDCAB=D.IDCAB WHERE C.TEM=" & sTasaTEM.Trim & " AND D.TIPO_TARJETA=" & sTipoTarjetaSimulador.Trim


                    Dim cmd_simu As SqlClient.SqlCommand = oConexion.CreateCommand

                    cmd_simu.CommandTimeout = 600
                    cmd_simu.CommandType = CommandType.Text
                    cmd_simu.CommandText = m_ssql.Trim

                    Dim rd_sim As SqlClient.SqlDataReader = cmd_simu.ExecuteReader
                    sXML_SQL = ""

                    If rd_sim.Read = True Then
                        'ARMAR CADENA
                        sXML_SQL = IIf(IsDBNull(rd_sim.Item("PLAZO_MIN")), "0", rd_sim.Item("PLAZO_MIN").ToString) & "|\t|"
                        sXML_SQL = sXML_SQL & IIf(IsDBNull(rd_sim.Item("PLAZO_MAX")), "0", rd_sim.Item("PLAZO_MAX").ToString) & "|\t|"
                        sXML_SQL = sXML_SQL & IIf(IsDBNull(rd_sim.Item("ENVIO_EECC")), "0", rd_sim.Item("ENVIO_EECC").ToString) & "|\t|"
                        sXML_SQL = sXML_SQL & IIf(IsDBNull(rd_sim.Item("SEG_DESG")), "0", rd_sim.Item("SEG_DESG").ToString) & "|\t|"
                        sXML_SQL = sXML_SQL & IIf(IsDBNull(rd_sim.Item("SEG_PROT")), "0", rd_sim.Item("SEG_PROT").ToString) & "|\t|"
                        sXML_SQL = sXML_SQL & IIf(IsDBNull(rd_sim.Item("TEM")), "0", rd_sim.Item("TEM").ToString) & "|\t|"
                        sXML_SQL = sXML_SQL & IIf(IsDBNull(rd_sim.Item("TEA")), "0", rd_sim.Item("TEA").ToString) & "|\t|"
                        sXML_SQL = sXML_SQL & IIf(IsDBNull(rd_sim.Item("MONTO_MIN")), "0", rd_sim.Item("MONTO_MIN").ToString)
                    Else
                        sXML_SQL = ""
                    End If

                    sResultado_SQL = sXML_SQL

                    rd_sim.Close()
                    rd_sim = Nothing
                    cmd_simu.Dispose()
                    cmd_simu = Nothing
                    oConexion.Close()
                    oConexion = Nothing

                Else
                    sResultado_SQL = ""
                    oConexion = Nothing
                End If
            End If


        Catch ex As Exception
            sResultado_SQL = ""
            oConexion = Nothing
        End Try

        Return sResultado_SQL

    End Function



    'GUARDAR LOG DE CONSULTAS EN EL KIOSCO RIPLEYMATICO
    <WebMethod(Description:="GUARDAR LOG DE CONSULTAS DEL KIOSCO.")> _
    Public Function SAVE_LOG_CONSULTAS( _
     ByVal ID_SUCURSAL As String, ByVal NOMBRE_KIOSCO As String, ByVal ID_MENU As String, ByVal TIPO_TARJETA As String) As String

        Dim sEstado As String
        Dim sInsert As String

        If OpLog_Consulta = "ON" Then

            Try

                sEstado = "NO"

                Dim oCN As SqlConnection
                Dim oCMD As SqlCommand


                'Conexion a la base de datos
                oCN = SQL_ConnectionOpen(Get_CadenaConexion(), sMensajeError_SQL)

                If TIPO_TARJETA.Trim.Length > 2 Then
                    TIPO_TARJETA = FUN_BUSCAR_TIPO_TARJETA(TIPO_TARJETA.Trim)
                End If

                'Armar cadena Insert
                sInsert = "INSERT INTO KIO_CONSULTAS (ID_SUCURSAL,NOMBRE_KIOSCO,ID_MENU,TIPO_TARJETA)"
                sInsert = sInsert & " VALUES(" & ID_SUCURSAL.Trim & ",'" & NOMBRE_KIOSCO.Trim & "'," & ID_MENU.Trim & "," & TIPO_TARJETA.Trim & ")"


                'Ejecutar el Stored Procedure
                oCMD = New SqlClient.SqlCommand
                oCMD.CommandText = sInsert
                oCMD.CommandType = CommandType.Text
                oCMD.Connection = oCN


                'Ejecutar oCMD
                oCMD.ExecuteNonQuery()
                'operacion con exito
                sEstado = "SI"

                oCMD.Dispose()
                oCN.Close()

                oCMD = Nothing
                oCN = Nothing


            Catch ex As Exception
                sEstado = "NO"
            End Try

        Else
            sEstado = "NO"

        End If

        Return sEstado

    End Function

    'GUARDAR LOG DE CONSULTAS DEL SUPER EFECTIVO SEF
    <WebMethod(Description:="GUARDAR LOG DE CONSULTAS SUPER EFECTIVO SEF.")> _
    Public Function SAVE_LOG_CONSULTAS_SEF( _
     ByVal sCodSucursalBanco As String, ByVal sCodigoKiosco As String, ByVal sTipoDocumento As String, _
     ByVal sNroDocumento As String, ByVal sOpcion As String, ByVal sCodigoAtencion As String) As String

        Dim sEstado As String
        Dim sInsert As String

        Try

            sEstado = "NO"

            Dim oConexion As SqlConnection
            Dim oCMD As SqlCommand


            'Conexion a la base de datos
            sMensajeError_SQL = ""
            oConexion = SQL_ConnectionOpen(Get_CadenaConexion(), sMensajeError_SQL)
            If sMensajeError_SQL <> "" Then
                sEstado = "NO"
            Else

                If oConexion.State = ConnectionState.Open Then

                    'Armar cadena
                    sInsert = "INSERT INTO KIO_CONSULTAS_SEF (COD_SUCURSAL_BAN,CODIGO_KIOSCO,TIPO_DOC,NRO_DOC,OPCION,COD_ATENCION)"
                    sInsert = sInsert & " VALUES(" & sCodSucursalBanco.Trim & ",'" & sCodigoKiosco.Trim & "','" & sTipoDocumento.Trim & "','" & sNroDocumento.Trim & "','" & sOpcion.Trim & "','" & sCodigoAtencion.Trim & "')"


                    'Ejecutar el Stored Procedure
                    oCMD = New SqlClient.SqlCommand
                    oCMD.CommandText = sInsert
                    oCMD.CommandType = CommandType.Text
                    oCMD.Connection = oConexion


                    'Ejecutar oCMD
                    oCMD.ExecuteNonQuery()
                    'operacion con exito
                    sEstado = "SI"

                    oCMD.Dispose()
                    oConexion.Close()

                    oCMD = Nothing
                    oConexion = Nothing
                Else
                    sEstado = "NO"
                    oConexion = Nothing
                End If

            End If


        Catch ex As Exception
            sEstado = "NO"
            oConexion = Nothing
        End Try

        Return sEstado

    End Function


    <WebMethod(Description:="Validar Numero de DNI SICRON | RSAT")> _
    Public Function VALIDAR_DNI_CLASICA(ByVal sNroDNI As String, ByVal sDATA_MONITOR_KIOSCO As String, ByVal sIdTipoTarjeta As String, ByVal sTipoDocumento As String) As String
        'Public Function VALIDAR_DNI_CLASICA(ByVal sNroDNI As String, ByVal sDATA_MONITOR_KIOSCO As String) As String
        ErrorLog("Entro al método VALIDAR_DNI_CLASICA(" + sNroDNI + "," + sDATA_MONITOR_KIOSCO + ")")
        Dim Respuesta, v_prioridad As String
        Dim sDataSuperEfectivoSEF As String
        Respuesta = String.Empty
        sDataSuperEfectivoSEF = String.Empty

        Dim Nro_Contrato As String = String.Empty
        Dim Tipo_Tarjeta_Ofertas As String = String.Empty
        ErrorLog("p_clasica=" & p_clasica)
        p_clasica = p_clasica
        v_prioridad = IIf(p_clasica = "01" Or p_clasica = "02", p_clasica, PRIORIDAD_S)
        ErrorLog("v_prioridad=" & v_prioridad)

        For x As Integer = 0 To 1
            Select Case v_prioridad
                Case "02" 'SICRON
                    Respuesta = VALIDAR_DNI_CLASICA_SICRON(sNroDNI, sDATA_MONITOR_KIOSCO, Nro_Contrato, Tipo_Tarjeta_Ofertas, sDataSuperEfectivoSEF)
                    ErrorLog("Respuesta=" & Respuesta)
                    v_prioridad = "01"

                    If Respuesta.Substring(0, 5) <> "ERROR" Then
                        v_prioridad = "SICRON"
                        Exit For
                    End If

                Case "01" 'RSAT
                    Respuesta = VALIDAR_DNI_CLASICA_RSAT("C", sNroDNI, sDATA_MONITOR_KIOSCO, Nro_Contrato, Tipo_Tarjeta_Ofertas, sDataSuperEfectivoSEF)
                    ErrorLog("Respuesta=" & Respuesta)
                    v_prioridad = "02"

                    If Respuesta.Substring(0, 5) <> "ERROR" Then
                        v_prioridad = "RSAT"
                        Exit For
                    End If

                Case Else
                    Respuesta = "ERROR: DNI no encontrado"
                    v_prioridad = String.Empty
            End Select

        Next

        If Respuesta.Substring(0, 5) <> "ERROR" Then
            Respuesta = Respuesta & "|\t|" & v_prioridad & "|\t|" & sDataSuperEfectivoSEF & "|\t|" & v_prioridad & "|\t|" & Nro_Contrato & "|\t|" & Tipo_Tarjeta_Ofertas

            If sDATA_MONITOR_KIOSCO.Length > 0 Then

                'RECORRER ARREGLO
                Dim sModoEntrada As String
                Dim sIDSucursal As String
                Dim sIDTerminal As String
                Dim sCodigoTransaccion As String
                Dim NroTarjeta As String

                Dim ADATA_MONITOR As Array
                ADATA_MONITOR = Split(sDATA_MONITOR_KIOSCO, "|\t|", , CompareMethod.Text)

                sModoEntrada = ADATA_MONITOR(2)
                sIDSucursal = ADATA_MONITOR(3)
                sIDTerminal = ADATA_MONITOR(4)
                sCodigoTransaccion = ADATA_MONITOR(5)
                NroTarjeta = Respuesta.Substring(0, 16)

                'GUADAR LOG CONSULTAS
                SAVE_LOG_CONSULTAS(sIDSucursal.Trim, sIDTerminal.Trim, "54", FUN_BUSCAR_TIPO_TARJETA(NroTarjeta.Trim))


            End If
        End If

        Return Respuesta

    End Function

    Private Function VALIDAR_DNI_CLASICA_SICRON(ByVal sNroDNI As String, _
                                                ByVal sDATA_MONITOR_KIOSCO As String, _
                                                ByRef Nro_Contrato As String, _
                                                ByRef Tipo_Tarjeta_Ofertas As String, _
                                                Optional ByRef sDataSuperEfectivo As String = "") As String

        ErrorLog("Entro a metodo VALIDAR_DNI_CLASICA_SICRON(" & sNroDNI & "," & sDATA_MONITOR_KIOSCO & "," & Nro_Contrato & "," & Tipo_Tarjeta_Ofertas & "," & sDataSuperEfectivo & ")")

        Dim sRespuesta As String = ""
        Dim sParametros As String = ""
        Dim sMensajeErrorUsuario As String = ""
        Dim sXML As String = ""

        'Recuperar Data
        Dim sNroTarjeta As String = ""
        Dim sNroDocumento As String = ""
        Dim sApellidoParterno As String = ""
        Dim sApellidoMaterno As String = ""
        Dim sNombres As String = ""
        Dim sCliente As String = ""
        Dim sFechaNac As String = ""
        Dim sTipoProducto As String = "" 'Segun el bin de la tarjeta 6 primeros numeros de la tarjeta
        Dim sNroCuenta As String = ""

        'Variables nuevas
        Dim actions As Long = 1
        Dim inetputBuff As String = ""
        Dim outpputBuff As String = ""
        Dim errorMsg As String = "" ':Mensaje de Error


        Try

            If sNroDNI.Trim.Length = 8 Then

                'Instancia al mirapiweb
                Dim obSendMirror As ClsTxMirapi = Nothing
                obSendMirror = New ClsTxMirapi()


                sParametros = "000" & "M" & "00000000" & "0000" & sNroDNI.Trim & "1"
                inetputBuff = "      " + "V303" + sParametros

                sRespuesta = obSendMirror.ExecuteTX(actions.ToString, inetputBuff, outpputBuff, TIME_OUT_SERVER, SERVIDOR_MIRROR_DESTINO, SERVIDOR_MIRROR_NODE, PUERTO, errorMsg)
                ErrorLog("sRespuesta=" & sRespuesta)

                If sRespuesta = "0" Then 'EXITO

                    If outpputBuff.Length > 0 Then sRespuesta = outpputBuff.Substring(8, outpputBuff.Length - 8)

                    'Evaluar Respuesta si es ERROR,  RU, AU (Respuesta correcta)
                    If sRespuesta.Trim.Length > 0 Then
                        If Left(sRespuesta.Trim, 2) = "RU" Then 'Mensaje de Validacion al usuario. RU            000CUENTA BLOQUEADA
                            sMensajeErrorUsuario = Mid(sRespuesta.Trim, 18, Len(sRespuesta.Trim))
                            sRespuesta = "ERROR:NODATA-" & sMensajeErrorUsuario.Trim
                        Else
                            If Left(sRespuesta.Trim, 2) = "AU" Then

                                sNroTarjeta = Mid(sRespuesta, 18, 16)
                                sNroDocumento = Mid(sRespuesta, 40, 8) 'Cortar

                                sApellidoParterno = Mid(sRespuesta, 57, 17)
                                sApellidoMaterno = Mid(sRespuesta, 74, 17)
                                sNombres = Mid(sRespuesta, 91, 17)
                                sCliente = sApellidoParterno.Trim & " " & sApellidoMaterno.Trim & ", " & sNombres.Trim
                                sFechaNac = Mid(sRespuesta, 237, 8)

                                If sFechaNac.Trim.Length > 0 Then 'DDMM
                                    sFechaNac = Right(Trim("00" & Mid(sFechaNac.Trim, 1, 2)), 2) & Right(Trim("00" & Mid(sFechaNac.Trim, 3, 2)), 2)

                                    'VALIDAR CUMPLEAÑOS
                                    Dim vFecHoy_ As Date
                                    Dim factual As String = ""

                                    vFecHoy_ = Date.Now
                                    factual = Right(Trim("00" & Day(vFecHoy_).ToString), 2) & Right(Trim("00" & Month(vFecHoy_).ToString), 2)

                                    If sFechaNac.Trim <> factual.Trim Then
                                        sFechaNac = ""
                                    End If

                                End If

                                sTipoProducto = Mid(sNroTarjeta, 1, 6)
                                sNroCuenta = ("5" & Mid(sNroTarjeta, 9, 7)).Trim 'CORREGIDO POS CORTE

                                Nro_Contrato = sNroCuenta
                                Tipo_Tarjeta_Ofertas = "SICRON"

                                sDataSuperEfectivo = MOSTRAR_SUPEREFECTIVO_SEF("1", "1", sNroDocumento.Trim)
                                sXML = sNroTarjeta.Trim & "|\t|" & sNroDocumento.Trim & "|\t|" & sCliente.Trim & "|\t|" & sFechaNac.Trim & "|\t|" & sTipoProducto.Trim & "|\t|" & MOSTRAR_SUPEREFECTIVO(sNroCuenta, sNroDocumento.Trim, TServidor.SICRON) & "|\t|" & sNroCuenta.Trim


                                sRespuesta = sXML

                            Else 'Cualquier Otro Caso
                                sRespuesta = "ERROR:Servicio no disponible."
                            End If


                        End If

                    Else 'Sino devuelve nada
                        sRespuesta = "ERROR:Servicio no disponible."
                    End If

                ElseIf sRespuesta = "-2" Then 'Ocurrio un error Recuperar el error
                    sRespuesta = "ERROR:" & errorMsg.Trim
                Else
                    sRespuesta = "ERROR:Servicio no disponible."
                End If

            Else
                'Mostrar Mensaje de Error
                sRespuesta = "ERROR:El DNI que ingreso no es valido."
            End If


        Catch ex As Exception
            'save log error

            sRespuesta = "ERROR:" & ex.Message.Trim

        End Try


        '    'armar variable desde flash con: numero de cuenta, numero de tarjeta, nombre del cliente, modo de entrada, canal de atencion, id de sucursal,codigo kiosco y codigo de transaccion (opcion)
        '    If Left(sRespuesta.Trim, 5) <> "ERROR" Then
        '        'VALIDACION CORRECTA

        '        sNroCuenta = Right("          " & sNroCuenta.Trim, 10)
        '        sNroTarjeta = Right("                    " & sNroTarjeta.Trim, 20)

        '        sNombres = Right(Strings.StrDup(26, " ") & sCliente.Trim, 26)

        '        sEstadoCuenta = "01" 'Estado de la cuenta
        '        sRespuestaServidor = "01" 'Atendido

        '        'ENVIAR_MONITOR
        '        sDataMonitor = ""
        '        sDataMonitor = FECHA_TX & HORA_TX & sNroCuenta & sNroTarjeta & sNombres & sModoEntrada & sCanalAtencion & sSaldoDisponible & sTotalDeuda & sIDSucursal & sIDTerminal & sEstadoCuenta & sRespuestaServidor & sCodigoTransaccion

        '        ENVIAR_MONITOR(sDataMonitor)

        '    Else
        '        'ENVIAR_MONITOR
        '        sNroCuenta = Right("          " & sNroCuenta.Trim, 10)
        '        sNroTarjeta = Right("                    " & sNroTarjeta.Trim, 20)
        '        sNombres = Strings.StrDup(26, " ")
        '        sEstadoCuenta = "02" 'Estado de la cuenta
        '        sRespuestaServidor = "02" 'Rechazado

        '        sDataMonitor = ""
        '        sDataMonitor = FECHA_TX & HORA_TX & sNroCuenta & sNroTarjeta & sNombres & sModoEntrada & sCanalAtencion & sSaldoDisponible & sTotalDeuda & sIDSucursal & sIDTerminal & sEstadoCuenta & sRespuestaServidor & sCodigoTransaccion

        '        ENVIAR_MONITOR(sDataMonitor)

        '    End If

        'End If


        Return sRespuesta

    End Function

    <WebMethod(Description:="Identificacion por DNI/Carnet de Extranjeria de Tarjetas Asociadas.")> _
    Public Function VALIDAR_DOC_ASOCIADOS(ByVal sNDocumento As String, ByVal sDATA_MONITOR_KIOSCO As String, ByVal sTProducto As String, ByVal sTDocumento As String) As String
        'Public Function VALIDAR_DOC_ASOCIADOS(ByVal sTProducto As String, ByVal sTDocumento As String, ByVal sNDocumento As String, ByVal sDATA_MONITOR_KIOSCO As String) As String
        ErrorLog("Entro al método VALIDAR_DOC_ASOCIADOS(" + sNDocumento + "," + sDATA_MONITOR_KIOSCO + "," + sTProducto + "," + sTDocumento + ")")
        Dim Respuesta As String = String.Empty
        Dim Nro_Contrato As String = String.Empty
        Dim Tipo_Tarjeta_Ofertas As String = String.Empty

        Respuesta = VALIDAR_DOC_ASOCIADOS_MC_VISA(sTProducto, sTDocumento, sNDocumento, sDATA_MONITOR_KIOSCO)
        ErrorLog("Respuesta de VALIDAR_DOC_ASOCIADOS_MC_VISA=" & Respuesta)
        If (Respuesta.Substring(0, 5) = "ERROR") Then
            Respuesta = VALIDAR_DNI_ABIERTA_RSAT(sTProducto, sTDocumento, sNDocumento, Nro_Contrato, Tipo_Tarjeta_Ofertas)
            ErrorLog("Respuesta de VALIDAR_DNI_ABIERTA_RSAT=" & Respuesta)
        End If

        If (Respuesta.Substring(0, 5) <> "ERROR") Then

            Dim NroTarjeta As String

            If Right(Respuesta, 4) = "RSAT" Then
                NroTarjeta = Respuesta.ToString.Substring(0, 16)
                Tipo_Tarjeta_Ofertas = "RSAT"
                Respuesta = Respuesta & "|\t|" & Nro_Contrato & "|\t|" & Tipo_Tarjeta_Ofertas
            Else
                Dim Arreglo1 As Array
                Arreglo1 = Split(Respuesta, "*$¿*", , CompareMethod.Text)
                NroTarjeta = Arreglo1(1).ToString.Substring(0, 16)
            End If

            Try
                If sDATA_MONITOR_KIOSCO.Length > 0 Then

                    'RECORRER ARREGLO
                    Dim ADATA_MONITOR As Array
                    Dim sModoEntrada As String
                    Dim sIDSucursal As String
                    Dim sIDTerminal As String
                    Dim sCodigoTransaccion As String

                    ADATA_MONITOR = Split(sDATA_MONITOR_KIOSCO, "|\t|", , CompareMethod.Text)

                    sModoEntrada = ADATA_MONITOR(2)
                    sIDSucursal = ADATA_MONITOR(3)
                    sIDTerminal = ADATA_MONITOR(4)
                    sCodigoTransaccion = ADATA_MONITOR(5)


                    'GUADAR LOG CONSULTAS sNroTarjeta
                    If sTDocumento.Trim = "1" Then 'DNI
                        SAVE_LOG_CONSULTAS(sIDSucursal.Trim, sIDTerminal.Trim, "54", FUN_BUSCAR_TIPO_TARJETA(NroTarjeta.Trim))
                    End If

                    If sTDocumento.Trim = "2" Then 'CARNET DE EXTRANJERIA
                        SAVE_LOG_CONSULTAS(sIDSucursal.Trim, sIDTerminal.Trim, "55", FUN_BUSCAR_TIPO_TARJETA(NroTarjeta.Trim))
                    End If
                End If
            Catch ex As Exception

            End Try


        End If

        Return Respuesta


    End Function


    <WebMethod(Description:="Identificacion por DNI/Carnet de Extranjeria de Tarjetas Asociadas.")> _
    Public Function VALIDAR_DOC_ASOCIADOS_MC_VISA(ByVal sTProducto As String, ByVal sTDocumento As String, ByVal sNDocumento As String, ByVal sDATA_MONITOR_KIOSCO As String) As String
        Dim sRespuesta As String = ""
        Dim sParametros As String = ""
        Dim sMensajeErrorUsuario As String = ""
        Dim sXML As String = ""
        Dim sNumeroDoc As String = ""
        Dim sValidado As Boolean = True

        'Recuperar Data
        Dim sNroTarjeta As String = ""
        Dim sNroDocumento As String = ""
        Dim sApellidoParterno As String = ""
        Dim sApellidoMaterno As String = ""
        Dim sNombres As String = ""
        Dim sCliente As String = ""
        Dim sFechaNac As String = ""
        Dim sTipoProducto As String = ""
        Dim sNroCuenta As String = ""

        'Variables nuevas
        Dim actions As Long = 1
        Dim inetputBuff As String = ""
        Dim outpputBuff As String = ""
        Dim errorMsg As String = ""

        Try

            If sTDocumento.Trim = "1" Then 'DNI
                If sNDocumento.Trim.Length <> 8 Then
                    sValidado = False
                Else
                    sNDocumento = "0000" & sNDocumento.Trim
                End If
            End If

            If sTDocumento.Trim = "2" Then 'CE
                sNDocumento = Right(Trim("000000000000" & sNDocumento.Trim).Trim, 12)

            End If


            If sValidado = True Then

                'Instancia al mirapiweb
                Dim obSendMirror As ClsTxMirapi = Nothing
                obSendMirror = New ClsTxMirapi()

                sParametros = "3" & "0000000000000000000" & "0000" & sTDocumento.Trim & sNDocumento.Trim & sTProducto.Trim
                inetputBuff = "      " + "Q001" + sParametros
                ErrorLog("Antes de llamar a obSendMirror")
                sRespuesta = obSendMirror.ExecuteTX(actions.ToString, inetputBuff, outpputBuff, TIME_OUT_SERVER, SERVIDOR_MIRROR_DESTINO, SERVIDOR_MIRROR_NODE, PUERTO, errorMsg)
                ErrorLog("sRespuesta=" & sRespuesta)
                If sRespuesta = "0" Then 'EXITO

                    If outpputBuff.Length > 0 Then sRespuesta = outpputBuff.Substring(8, outpputBuff.Length - 8)

                    'Evaluar Respuesta si es ERROR,  RU, AU (Respuesta correcta)
                    If sRespuesta.Trim.Length > 0 Then
                        If Left(sRespuesta.Trim, 2) = "RU" Then 'Mensaje de Validacion al usuario. RU            000CUENTA BLOQUEADA
                            sMensajeErrorUsuario = Mid(sRespuesta.Trim, 18, Len(sRespuesta.Trim))
                            sRespuesta = "ERROR:NODATA-" & sMensajeErrorUsuario.Trim
                        Else
                            If Left(sRespuesta.Trim, 2) = "AU" Then

                                sNroTarjeta = Mid(sRespuesta, 267, 16)

                                'NO SE ENVIA VARIABLES PARA EL MONITOR YA QUE SE ESTA ENVIADO EN ESTE METODO VALIDAR_DOC_ASOCIADOS
                                Dim oRpta As Respuesta
                                oRpta = SALDO_TARJETA_ASOCIADA(sNroTarjeta.Trim, "")
                                sXML = oRpta.Cadena

                                sRespuesta = sXML

                            Else 'Cualquier Otro Caso
                                sRespuesta = "ERROR:Servicio no disponible."
                            End If


                        End If

                    Else 'Sino devuelve nada
                        sRespuesta = "ERROR:Parametros Incorrectos."
                    End If

                ElseIf sRespuesta = "-2" Then 'Ocurrio un error Recuperar el error
                    sRespuesta = "ERROR:" & errorMsg.Trim
                Else
                    sRespuesta = "ERROR:Servicio no disponible."
                End If


            Else
                'Mostrar Mensaje de Error
                sRespuesta = "ERROR:Parametros Incorrectos."
            End If

        Catch ex As Exception
            'save log error

            sRespuesta = "ERROR:" & ex.Message.Trim

        End Try

        ''Save Monitor
        'Dim sDataMonitor As String = ""
        'Dim FECHA_TX As String = DateTime.Now.Year.ToString("0000").Trim & DateTime.Now.Month.ToString("00").Trim + DateTime.Now.Day.ToString("00").Trim
        'Dim HORA_TX As String = DateTime.Now.Hour.ToString("00").Trim & DateTime.Now.Minute.ToString("00").Trim & DateTime.Now.Second.ToString("00").Trim
        'Dim sModoEntrada As String = ""
        'Dim sCanalAtencion As String = "01" 'RipleyMatico
        'Dim sSaldoDisponible As String = "      "
        'Dim sTotalDeuda As String = "      "
        'Dim sIDSucursal As String = ""
        'Dim sIDTerminal As String = ""
        'Dim sEstadoCuenta As String = ""
        'Dim sRespuestaServidor As String = ""
        'Dim sCodigoTransaccion As String = ""


        'If sDATA_MONITOR_KIOSCO.Length > 0 Then

        '    'RECORRER ARREGLO
        '    Dim ADATA_MONITOR As Array
        '    ADATA_MONITOR = Split(sDATA_MONITOR_KIOSCO, "|\t|", , CompareMethod.Text)

        '    sModoEntrada = ADATA_MONITOR(2)
        '    sIDSucursal = ADATA_MONITOR(3)
        '    sIDTerminal = ADATA_MONITOR(4)
        '    sCodigoTransaccion = ADATA_MONITOR(5)


        '    'GUADAR LOG CONSULTAS sNroTarjeta
        '    If sTDocumento.Trim = "1" Then 'DNI
        '        SAVE_LOG_CONSULTAS(sIDSucursal.Trim, sIDTerminal.Trim, "54", FUN_BUSCAR_TIPO_TARJETA(sNroTarjeta.Trim))
        '    End If

        '    If sTDocumento.Trim = "2" Then 'CARNET DE EXTRANJERIA
        '        SAVE_LOG_CONSULTAS(sIDSucursal.Trim, sIDTerminal.Trim, "55", FUN_BUSCAR_TIPO_TARJETA(sNroTarjeta.Trim))
        '    End If


        'If Left(sRespuesta.Trim, 5) <> "ERROR" Then
        '    'VALIDACION CORRECTA

        '    'RECUPERAR DATOS DE LOS SALDOS DE LA TARJETA DE CREDITO DE MC VISA
        '    Dim ADATA_SALDOS_CLIENTE1 As Array
        '    Dim ADATA_SALDOS_AUX As Array
        '    Dim ADATA_CLIENTE As Array
        '    ADATA_SALDOS_CLIENTE1 = Split(sRespuesta, "*$¿*", , CompareMethod.Text)
        '    ADATA_SALDOS_AUX = Split(ADATA_SALDOS_CLIENTE1(0), "|\t|", , CompareMethod.Text) 'DATOS DE SALDO DE CREDITO DE ASOCIADO
        '    ADATA_CLIENTE = Split(ADATA_SALDOS_CLIENTE1(1), "|\t|", , CompareMethod.Text) 'DATOS DEL CLIENTE ASOCIADO


        '    sNroTarjeta = ADATA_CLIENTE(0)
        '    sNroCuenta = Mid(sNroTarjeta.Trim, 7, 8)
        '    sCliente = ADATA_CLIENTE(2)

        '    sSaldoDisponible = ADATA_SALDOS_AUX(2) 'DISPONIBLE DE COMPRAS SIN DECIMALES Y SIN COMAS
        '    sSaldoDisponible = Replace(sSaldoDisponible.Trim, ",", "") 'QUITAR LA COMA
        '    sSaldoDisponible = Replace(sSaldoDisponible.Trim, ".", "") 'QUITAR EL PUNTO DECIMAL
        '    sSaldoDisponible = Mid(sSaldoDisponible, 1, sSaldoDisponible.Trim.Length - 2) 'QUITAR LOS DOS ULTIMOS DIGITOS QUE SON DECIMALES


        '    sTotalDeuda = ADATA_SALDOS_AUX(10) 'PAGO TOTAL
        '    sTotalDeuda = Replace(sTotalDeuda.Trim, ",", "") 'QUITAR LA COMA
        '    sTotalDeuda = Replace(sTotalDeuda.Trim, ".", "") 'QUITAR EL PUNTO DECIMAL
        '    sTotalDeuda = Mid(sTotalDeuda, 1, sTotalDeuda.Trim.Length - 2) 'QUITAR LOS DOS ULTIMOS DIGITOS QUE SON DECIMALES


        '    sNroCuenta = Right("          " & sNroCuenta.Trim, 10)
        '    sNroTarjeta = Right("                    " & sNroTarjeta.Trim, 20)

        '    sNombres = Right(Strings.StrDup(26, " ") & sCliente.Trim, 26)
        '    sEstadoCuenta = "01" 'Estado de la cuenta
        '    sRespuestaServidor = "01" 'Atendido

        '    'ENVIAR_MONITOR
        '    sDataMonitor = ""
        '    sDataMonitor = FECHA_TX & HORA_TX & sNroCuenta & sNroTarjeta & sNombres & sModoEntrada & sCanalAtencion & sSaldoDisponible & sTotalDeuda & sIDSucursal & sIDTerminal & sEstadoCuenta & sRespuestaServidor & sCodigoTransaccion

        '    ENVIAR_MONITOR(sDataMonitor)

        'Else
        '    'ENVIAR_MONITOR
        '    sNroCuenta = Right("          " & sNroCuenta.Trim, 10)
        '    sNroTarjeta = Right("          " & sNroTarjeta.Trim, 20)


        '    sNombres = Strings.StrDup(26, " ")
        '    sEstadoCuenta = "02" 'Estado de la cuenta
        '    sRespuestaServidor = "02" 'Rechazado

        '    sDataMonitor = ""
        '    sDataMonitor = FECHA_TX & HORA_TX & sNroCuenta & sNroTarjeta & sNombres & sModoEntrada & sCanalAtencion & sSaldoDisponible & sTotalDeuda & sIDSucursal & sIDTerminal & sEstadoCuenta & sRespuestaServidor & sCodigoTransaccion

        '    ENVIAR_MONITOR(sDataMonitor)

        'End If

        ' End If

        Return sRespuesta

    End Function


    <WebMethod(Description:="Identificacion por Carnet de Extranjeria")> _
    Public Function VALIDAR_CARNET_EXTRANJERIA_CLASICA(ByVal sNroCarnet As String, _
                                                       ByVal sDATA_MONITOR_KIOSCO As String) As String
    'Public Function VALIDAR_CARNET_EXTRANJERIA_CLASICA(ByVal sNroCarnet As String, ByVal sDATA_MONITOR_KIOSCO As String) As String

        ErrorLog("Entro al método VALIDAR_CARNET_EXTRANJERIA_CLASICA(" + sNroCarnet + "," + sDATA_MONITOR_KIOSCO + ")")
    Dim sRespuesta As String = ""
    Dim sParametros As String = ""
    Dim sMensajeErrorUsuario As String = ""
    Dim sXML As String = ""

    'Recuperar Data
    Dim sNroTarjeta As String = ""
    Dim sNroDocumento As String = ""
    Dim sApellidoParterno As String = ""
    Dim sApellidoMaterno As String = ""
    Dim sNombres As String = ""
    Dim sCliente As String = ""
    Dim sFechaNac As String = ""
    Dim sTipoProducto As String = "" 'Segun el bin de la tarjeta 6 primeros numeros de la tarjeta
    Dim sNroCuenta As String = ""

    'Variables Nuevas
    Dim actions As Long = 1
    Dim inetputBuff As String = ""
    Dim outpputBuff As String = ""
    Dim errorMsg As String = ""
    Dim Nro_Contrato As String = String.Empty
    Dim Tipo_Tarjeta_Ofertas As String = String.Empty


        Try

            If sNroCarnet.Trim.Length > 0 Then


    'Instancia al mirapiweb
    Dim obSendMirror As ClsTxMirapi = Nothing
                obSendMirror = New ClsTxMirapi()

                sNroCarnet = Right(Trim("000000000000" & sNroCarnet.Trim).Trim, 12)
                sParametros = "000" & "M" & "00000000" & sNroCarnet.Trim & "2"

                inetputBuff = "      " + "V303" + sParametros
                sRespuesta = obSendMirror.ExecuteTX(actions.ToString, inetputBuff, outpputBuff, TIME_OUT_SERVER, SERVIDOR_MIRROR_DESTINO, SERVIDOR_MIRROR_NODE, PUERTO, errorMsg)

                If sRespuesta = "0" Then 'EXITO
                    If outpputBuff.Length > 0 Then sRespuesta = outpputBuff.Substring(8, outpputBuff.Length - 8)

    'Evaluar Respuesta si es ERROR,  RU, AU (Respuesta correcta)
                    If sRespuesta.Trim.Length > 0 Then
                        If Left(sRespuesta.Trim, 2) = "RU" Then 'Mensaje de Validacion al usuario. RU            000CUENTA BLOQUEADA
                            sMensajeErrorUsuario = Mid(sRespuesta.Trim, 18, Len(sRespuesta.Trim))
                            sRespuesta = "ERROR:NODATA-" & sMensajeErrorUsuario.Trim
                        Else
                            If Left(sRespuesta.Trim, 2) = "AU" Then


                                sNroTarjeta = Mid(sRespuesta, 18, 16)
                                sNroDocumento = Mid(sRespuesta, 36, 12) 'QUE DEVUELVE EL DNI O CARNET DE EXTRANJERIA ?
                                sApellidoParterno = Mid(sRespuesta, 57, 17)
                                sApellidoMaterno = Mid(sRespuesta, 74, 17)
                                sNombres = Mid(sRespuesta, 91, 17)
                                sCliente = sApellidoParterno.Trim & " " & sApellidoMaterno.Trim & ", " & sNombres.Trim
                                sFechaNac = Mid(sRespuesta, 237, 8)



                                If sFechaNac.Trim.Length > 0 Then 'DDMM
                                    sFechaNac = Right(Trim("00" & Mid(sFechaNac.Trim, 1, 2)), 2) & Right(Trim("00" & Mid(sFechaNac.Trim, 3, 2)), 2)

    'VALIDAR CUMPLEAÑOS
    Dim vFecHoy_ As Date
    Dim factual As String = ""

                                    vFecHoy_ = Date.Now
                                    factual = Right(Trim("00" & Day(vFecHoy_).ToString), 2) & Right(Trim("00" & Month(vFecHoy_).ToString), 2)

                                    If sFechaNac.Trim <> factual.Trim Then
                                        sFechaNac = ""
                                    End If

                                End If

                                sTipoProducto = Mid(sNroTarjeta, 1, 6)
                                sNroCuenta = ("5" & Mid(sNroTarjeta, 9, 7)).Trim 'CORREGIDO POS CORTE
                                Nro_Contrato = sNroCuenta
                                Tipo_Tarjeta_Ofertas = "SICRON"

                                sXML = sNroTarjeta.Trim & "|\t|" & sNroDocumento.Trim & "|\t|" & sCliente.Trim & "|\t|" & sFechaNac.Trim & "|\t|" & sTipoProducto.Trim & "|\t|" & MOSTRAR_SUPEREFECTIVO(sNroCuenta, sNroDocumento.Trim, TServidor.SICRON) & "|\t|" & sNroCuenta & "|\t|" & "SICRON" & "|\t|" & MOSTRAR_SUPEREFECTIVO_SEF("1", "2", sNroDocumento.Trim) & "|\t|" & "SICRON" & "|\t|" & Nro_Contrato & "|\t|" & Tipo_Tarjeta_Ofertas
                                sRespuesta = sXML

                            Else 'Cualquier Otro Caso
                                sRespuesta = "ERROR:Servicio no disponible."
                            End If


                        End If

                    Else 'Sino devuelve nada
                        sRespuesta = "ERROR:Servicio no disponible."
                    End If

                ElseIf sRespuesta = "-2" Then 'Ocurrio un error Recuperar el error
                    sRespuesta = "ERROR:" & errorMsg.Trim
                Else
                    sRespuesta = "ERROR:Servicio no disponible."
                End If

            Else
    'Mostrar Mensaje de Error
                sRespuesta = "ERROR:El Nro Carnet que ingreso no es valido."
            End If

        Catch ex As Exception
    'save log error

            sRespuesta = "ERROR:" & ex.Message.Trim

        End Try



    'Save Monitor
    Dim sDataMonitor As String = ""
    Dim FECHA_TX As String = DateTime.Now.Year.ToString("0000").Trim & DateTime.Now.Month.ToString("00").Trim + DateTime.Now.Day.ToString("00").Trim
    Dim HORA_TX As String = DateTime.Now.Hour.ToString("00").Trim & DateTime.Now.Minute.ToString("00").Trim & DateTime.Now.Second.ToString("00").Trim
    Dim sModoEntrada As String = ""
    Dim sCanalAtencion As String = "01" 'RipleyMatico
    Dim sSaldoDisponible As String = "      "
    Dim sTotalDeuda As String = "      "
    Dim sIDSucursal As String = ""
    Dim sIDTerminal As String = ""
    Dim sEstadoCuenta As String = ""
    Dim sRespuestaServidor As String = ""
    Dim sCodigoTransaccion As String = ""


        If sDATA_MONITOR_KIOSCO.Length > 0 And sRespuesta.Substring(0, 5) <> "ERROR" Then

    'RECORRER ARREGLO
    Dim ADATA_MONITOR As Array
            ADATA_MONITOR = Split(sDATA_MONITOR_KIOSCO, "|\t|", , CompareMethod.Text)

            sModoEntrada = ADATA_MONITOR(2)
            sIDSucursal = ADATA_MONITOR(3) 'CODIGO DE LA SUCURSAL
            sIDTerminal = ADATA_MONITOR(4) 'CODIGO DEL KIOSCO
            sCodigoTransaccion = ADATA_MONITOR(5)

    'GUADAR LOG CONSULTAS
            SAVE_LOG_CONSULTAS(sIDSucursal.Trim, sIDTerminal.Trim, "55", FUN_BUSCAR_TIPO_TARJETA(sNroTarjeta.Trim))


    ''armar variable desde flash con: numero de cuenta, numero de tarjeta, nombre del cliente, modo de entrada, canal de atencion, id de sucursal,codigo kiosco y codigo de transaccion (opcion)
    'If Left(sRespuesta.Trim, 5) <> "ERROR" Then
    '    'VALIDACION CORRECTA

    '    sNroCuenta = Right("          " & sNroCuenta.Trim, 10)
    '    sNroTarjeta = Right("                    " & sNroTarjeta.Trim, 20)


    '    sNombres = Right(Strings.StrDup(26, " ") & sCliente.Trim, 26)

    '    sEstadoCuenta = "01" 'Estado de la cuenta
    '    sRespuestaServidor = "01" 'Atendido

    '    'ENVIAR_MONITOR
    '    sDataMonitor = ""
    '    sDataMonitor = FECHA_TX & HORA_TX & sNroCuenta & sNroTarjeta & sNombres & sModoEntrada & sCanalAtencion & sSaldoDisponible & sTotalDeuda & sIDSucursal & sIDTerminal & sEstadoCuenta & sRespuestaServidor & sCodigoTransaccion

    '    'ENVIAR_MONITOR(sDataMonitor)

    'Else
    '    'ENVIAR_MONITOR
    '    sNroCuenta = Right("          " & sNroCuenta.Trim, 10)
    '    sNroTarjeta = Right("                    " & sNroTarjeta.Trim, 20)

    '    sNombres = Strings.StrDup(26, " ")
    '    sEstadoCuenta = "02" 'Estado de la cuenta
    '    sRespuestaServidor = "02" 'Rechazado

    '    sDataMonitor = ""
    '    sDataMonitor = FECHA_TX & HORA_TX & sNroCuenta & sNroTarjeta & sNombres & sModoEntrada & sCanalAtencion & sSaldoDisponible & sTotalDeuda & sIDSucursal & sIDTerminal & sEstadoCuenta & sRespuestaServidor & sCodigoTransaccion

    '    'ENVIAR_MONITOR(sDataMonitor)

    'End If

        End If

        Return sRespuesta

    End Function


    <WebMethod(Description:="Validar Numero Tarjeta SICRON | RSAT")> _
    Public Function VALIDAR_TARJETA_CLASICA(ByVal NroTarjeta As String, ByVal sDATA_MONITOR_KIOSCO As String) As Respuesta
        Dim oRespuesta As Respuesta = New Respuesta
        Dim v_prioridad As String
        Dim sTramaSuperEfectivoSEF As String = ""
        Dim Nro_Contrato As String = String.Empty
        Dim Tipo_Tarjeta_Ofertas As String = String.Empty

        'Respuesta = String.Empty

        p_clasica = p_clasica
        v_prioridad = IIf(p_clasica = "01" Or p_clasica = "02", p_clasica, PRIORIDAD_S)

        For x As Integer = 0 To 1
            Select Case v_prioridad
                Case "02" 'SICRON

                    oRespuesta = VALIDAR_TARJETA_CLASICA_SICRON(NroTarjeta, sDATA_MONITOR_KIOSCO, Nro_Contrato, Tipo_Tarjeta_Ofertas, sTramaSuperEfectivoSEF)
                    v_prioridad = "01"

                    If oRespuesta.Estado <> "ERROR" Then
                        v_prioridad = "SICRON"
                        Exit For
                    End If

                Case "01" 'RSAT
                    'Respuesta = VALIDAR_TARJETA_CLASICA_RSAT(NroTarjeta, sDATA_MONITOR_KIOSCO, Nro_Contrato, Tipo_Tarjeta_Ofertas, sTramaSuperEfectivoSEF)
                    oRespuesta = VALIDAR_TARJETA_CLASICA_RSAT(NroTarjeta, sDATA_MONITOR_KIOSCO, Nro_Contrato, Tipo_Tarjeta_Ofertas, sTramaSuperEfectivoSEF)
                    v_prioridad = "02"
                    If oRespuesta.Estado <> "ERROR" Then
                        v_prioridad = "RSAT"
                        Exit For
                    End If

                Case Else
                    oRespuesta.Estado = "ERROR"
                    oRespuesta.Cadena = "ERROR:Servicio no Disponible"
                    'Respuesta = "ERROR:Servicio no disponible"
                    v_prioridad = String.Empty
            End Select

        Next

        If oRespuesta.Estado <> "ERROR" Then
            oRespuesta.Cadena = oRespuesta.Cadena & "|\t|" & v_prioridad & "|\t|" & sTramaSuperEfectivoSEF.Trim & "|\t|" & v_prioridad & "|\t|" & Nro_Contrato & "|\t|" & Tipo_Tarjeta_Ofertas

            If sDATA_MONITOR_KIOSCO.Length > 0 Then
                Dim sIDSucursal As String = ""
                Dim sIDTerminal As String = ""
                'Dim sEstadoCuenta As String = ""
                'Dim sRespuestaServidor As String = ""
                Dim sCodigoTransaccion As String = ""

                ''RECORRER ARREGLO
                Dim ADATA_MONITOR As Array
                ADATA_MONITOR = Split(sDATA_MONITOR_KIOSCO, "|\t|", , CompareMethod.Text)

                sIDSucursal = ADATA_MONITOR(3)
                sIDTerminal = ADATA_MONITOR(4)
                sCodigoTransaccion = ADATA_MONITOR(5)

                ''GUADAR LOG CONSULTAS
                SAVE_LOG_CONSULTAS(sIDSucursal.Trim, sIDTerminal.Trim, "56", FUN_BUSCAR_TIPO_TARJETA(NroTarjeta.Trim))
            End If

        End If

        Return oRespuesta

    End Function


    <WebMethod(Description:="Validacion para Tarjetas Abiertas RSAT")> _
    Public Function VALIDAR_TARJETA_ABIERTA_RSAT(ByVal NroTarjeta As String, ByVal sDATA_MONITOR_KIOSCO As String) As Respuesta
        Dim oRespuesta As Respuesta = New Respuesta
        Dim Respuesta, v_prioridad As String
        Dim sDataSuperEfectivoSEF As String
        Dim Nro_Contrato As String = String.Empty
        Dim Tipo_Tarjeta_Ofertas As String = String.Empty
        sDataSuperEfectivoSEF = String.Empty
        v_prioridad = String.Empty
        Respuesta = String.Empty
        ErrorLog("VALIDAR_TARJETA_ABIERTA_RSAT")
        sDATA_MONITOR_KIOSCO = sDATA_MONITOR_KIOSCO.Replace("|\t|56|\t|", "|\t|57|\t|")

        oRespuesta = VALIDAR_TARJETA_CLASICA_RSAT(NroTarjeta, sDATA_MONITOR_KIOSCO, Nro_Contrato, Tipo_Tarjeta_Ofertas, sDataSuperEfectivoSEF)
        Tipo_Tarjeta_Ofertas = "RSAT"
        If oRespuesta.Estado <> "ERROR" Then
            v_prioridad = "RSAT"
        End If

        If oRespuesta.Estado <> "ERROR" Then
            oRespuesta.Cadena = oRespuesta.Cadena & "|\t|" & v_prioridad & "|\t|" & sDataSuperEfectivoSEF & "|\t|" & v_prioridad & "|\t|" & Nro_Contrato & "|\t|" & Tipo_Tarjeta_Ofertas

            If sDATA_MONITOR_KIOSCO.Length > 0 Then
                Dim sIDSucursal As String = ""
                Dim sIDTerminal As String = ""
                'Dim sEstadoCuenta As String = ""
                'Dim sRespuestaServidor As String = ""
                Dim sCodigoTransaccion As String = ""

                ''RECORRER ARREGLO
                Dim ADATA_MONITOR As Array
                ADATA_MONITOR = Split(sDATA_MONITOR_KIOSCO, "|\t|", , CompareMethod.Text)

                sIDSucursal = ADATA_MONITOR(3)
                sIDTerminal = ADATA_MONITOR(4)
                sCodigoTransaccion = ADATA_MONITOR(5)

                ''GUADAR LOG CONSULTAS
                SAVE_LOG_CONSULTAS(sIDSucursal.Trim, sIDTerminal.Trim, "57", FUN_BUSCAR_TIPO_TARJETA(NroTarjeta.Trim))
            End If

        End If

        Return oRespuesta

    End Function


    Private Function VALIDAR_TARJETA_CLASICA_SICRON(ByVal NroTarjeta As String, _
                                                    ByVal sDATA_MONITOR_KIOSCO As String, _
                                                    ByRef Nro_Contrato As String, _
                                                    ByRef Tipo_Tarjeta_Ofertas As String, _
                                                    Optional ByRef sDataSuperEfectivoSEF As String = "") As Respuesta
        Dim sRespuesta As String = ""
        Dim oRespuesta As Respuesta = New Respuesta
        Dim sParametros As String = ""
        Dim sMensajeErrorUsuario As String = ""
        Dim sXML As String = ""
        Dim sNroCuenta As String = String.Empty
        'Recuperar Data
        Dim sNroTarjeta As String = ""
        Dim sNroDocumento As String = ""
        Dim sApellidoParterno As String = ""
        Dim sApellidoMaterno As String = ""
        Dim sNombres As String = ""
        Dim sCliente As String = ""
        Dim sFechaNac As String = ""
        Dim sTipoProducto As String = "" 'Segun el bin de la tarjeta 6 primeros numeros de la tarjeta
        Dim sTipoDocumento As String = ""

        'Variables nuevas
        Dim actions As Long = 1
        Dim inetputBuff As String = ""
        Dim outpputBuff As String = ""
        Dim errorMsg As String = ""

        sNroCuenta = "5" & NroTarjeta.Substring(8, 7)

        Try

            If sNroCuenta.Trim.Length = 8 Then

                'Instancia al mirapiweb
                Dim obSendMirror As ClsTxMirapi = Nothing
                obSendMirror = New ClsTxMirapi()


                sParametros = "000" & "M" & sNroCuenta.Trim & "0000000000000"

                inetputBuff = "      " + "V303" + sParametros
                sRespuesta = obSendMirror.ExecuteTX(actions.ToString, inetputBuff, outpputBuff, TIME_OUT_SERVER, SERVIDOR_MIRROR_DESTINO, SERVIDOR_MIRROR_NODE, PUERTO, errorMsg)

                If sRespuesta = "0" Then 'EXITO
                    oRespuesta.Estado = "EXITO"
                    If outpputBuff.Length > 0 Then sRespuesta = outpputBuff.Substring(8, outpputBuff.Length - 8)

                    'Evaluar Respuesta si es ERROR,  RU, AU (Respuesta correcta)
                    If sRespuesta.Trim.Length > 0 Then
                        If Left(sRespuesta.Trim, 2) = "RU" Then 'Mensaje de Validacion al usuario. RU            000CUENTA BLOQUEADA
                            sMensajeErrorUsuario = Mid(sRespuesta.Trim, 18, Len(sRespuesta.Trim))
                            sRespuesta = "ERROR:NODATA-" & sMensajeErrorUsuario.Trim
                            oRespuesta.Estado = "ERROR"
                            oRespuesta.Codigo = "MSGMSG"
                            oRespuesta.Cadena = sRespuesta
                            oRespuesta.Mensaje = ValidarMensajeAMostrarEnRipleymatico(True, "", "", "")
                        Else
                            If Left(sRespuesta.Trim, 2) = "AU" Then

                                sNroTarjeta = Mid(sRespuesta, 18, 16)

                                If NroTarjeta <> sNroTarjeta.Trim Then
                                    oRespuesta.Estado = "ERROR"
                                    oRespuesta.Cadena = "ERROR:Tarjeta no valida."
                                    oRespuesta.Codigo = "MSGMSG"
                                    oRespuesta.Mensaje = "Tarjeta no válida."
                                    Return oRespuesta
                                End If

                                sTipoDocumento = Mid(sRespuesta, 37, 1)

                                If sTipoDocumento.Trim = "0" Then 'DNI
                                    sNroDocumento = Mid(sRespuesta, 36, 12)
                                    sNroDocumento = Right(sNroDocumento.Trim, 8)
                                Else
                                    sNroDocumento = Mid(sRespuesta, 36, 12)
                                End If


                                sApellidoParterno = Mid(sRespuesta, 57, 17)
                                sApellidoMaterno = Mid(sRespuesta, 74, 17)
                                sNombres = Mid(sRespuesta, 91, 17)
                                Nro_Contrato = sNroCuenta
                                Tipo_Tarjeta_Ofertas = "SICRON"
                                sCliente = sApellidoParterno.Trim & " " & sApellidoMaterno.Trim & ", " & sNombres.Trim
                                sFechaNac = Mid(sRespuesta, 237, 8) 'DDMMYYYY

                                If sFechaNac.Trim.Length > 0 Then 'DDMM
                                    sFechaNac = Right(Trim("00" & Mid(sFechaNac.Trim, 1, 2)), 2) & Right(Trim("00" & Mid(sFechaNac.Trim, 3, 2)), 2)

                                    'VALIDAR CUMPLEAÑOS
                                    Dim vFecHoy_ As Date
                                    Dim factual As String = ""

                                    vFecHoy_ = Date.Now
                                    factual = Right(Trim("00" & Day(vFecHoy_).ToString), 2) & Right(Trim("00" & Month(vFecHoy_).ToString), 2)

                                    If sFechaNac.Trim <> factual.Trim Then
                                        sFechaNac = ""
                                    End If

                                End If

                                sTipoProducto = Mid(sNroTarjeta, 1, 6)

                                sDataSuperEfectivoSEF = MOSTRAR_SUPEREFECTIVO_SEF("1", IIf(sTipoDocumento.Trim = "0", "1", "2"), sNroDocumento.Trim)

                                sXML = sNroTarjeta.Trim & "|\t|" & sNroDocumento.Trim & "|\t|" & sCliente.Trim & "|\t|" & sFechaNac.Trim & "|\t|" & sTipoProducto.Trim & "|\t|" & MOSTRAR_SUPEREFECTIVO(sNroCuenta.Trim, sNroDocumento.Trim, TServidor.SICRON) & "|\t|" & sNroCuenta.Trim
                                sRespuesta = sXML
                                oRespuesta.Cadena = sRespuesta
                                Dim ADATA_MONITOR As Array
                                Dim sIDSucursal, sIDTerminal As String
                                sIDSucursal = ""
                                sIDTerminal = ""

                                If sDATA_MONITOR_KIOSCO.Length > 0 Then
                                    ADATA_MONITOR = Split(sDATA_MONITOR_KIOSCO, "|\t|", , CompareMethod.Text)
                                    sIDSucursal = ADATA_MONITOR(3)
                                    sIDTerminal = ADATA_MONITOR(4)
                                    getRegistrar_Ingreso_Ripleymatico(sIDSucursal, sIDTerminal, sNroDocumento, IIf(sTipoDocumento.Trim = "0", "1", "2"), sNroTarjeta, sTipoProducto)
                                End If


                            Else 'Cualquier Otro Caso
                                oRespuesta.Estado = "ERROR"
                                oRespuesta.Cadena = "ERROR:Servicio no disponible."
                                oRespuesta.Codigo = "MSGMSG"
                                oRespuesta.Mensaje = ValidarMensajeAMostrarEnRipleymatico(False, "", "", "NoServicio")
                            End If

                        End If

                    Else 'Sino devuelve nada
                        oRespuesta.Estado = "ERROR"
                        oRespuesta.Cadena = "ERROR:Servicio no disponible."
                        oRespuesta.Codigo = "MSGMSG"
                        oRespuesta.Mensaje = ValidarMensajeAMostrarEnRipleymatico(False, "", "", "NoServicio")
                    End If

                ElseIf sRespuesta = "-2" Then 'Ocurrio un error Recuperar el error
                    oRespuesta.Estado = "ERROR"
                    oRespuesta.Cadena = "ERROR:Servicio no disponible."
                Else
                    oRespuesta.Estado = "ERROR"
                    oRespuesta.Cadena = "ERROR:Servicio no disponible."
                End If

            Else
                'Mostrar Mensaje de Error
                oRespuesta.Estado = "ERROR"
                oRespuesta.Cadena = "ERROR:Servicio no disponible."
            End If

        Catch ex As Exception
            'save log error
            oRespuesta.Estado = "ERROR"
            oRespuesta.Cadena = "ERROR:Servicio no disponible."
        End Try


        ' ''Save(Monitor)
        'Dim sDataMonitor As String = ""
        'Dim FECHA_TX As String = DateTime.Now.Year.ToString("0000").Trim & DateTime.Now.Month.ToString("00").Trim + DateTime.Now.Day.ToString("00").Trim
        'Dim HORA_TX As String = DateTime.Now.Hour.ToString("00").Trim & DateTime.Now.Minute.ToString("00").Trim & DateTime.Now.Second.ToString("00").Trim
        'Dim sModoEntrada As String = "01" 'Lectura por Tarjeta
        'Dim sCanalAtencion As String = "01" 'RipleyMatico
        'Dim sSaldoDisponible As String = "      "
        'Dim sTotalDeuda As String = "      "
        'Dim sIDSucursal As String = ""
        'Dim sIDTerminal As String = ""
        ''Dim sEstadoCuenta As String = ""
        ''Dim sRespuestaServidor As String = ""
        'Dim sCodigoTransaccion As String = ""

        'If sDATA_MONITOR_KIOSCO.Length > 0 And sRespuesta.Substring(0, 5) <> "ERROR" Then

        '    ''RECORRER ARREGLO
        '    Dim ADATA_MONITOR As Array
        '    ADATA_MONITOR = Split(sDATA_MONITOR_KIOSCO, "|\t|", , CompareMethod.Text)

        '    sIDSucursal = ADATA_MONITOR(3)
        '    sIDTerminal = ADATA_MONITOR(4)
        '    sCodigoTransaccion = ADATA_MONITOR(5)

        '    ''GUADAR LOG CONSULTAS
        '    SAVE_LOG_CONSULTAS(sIDSucursal.Trim, sIDTerminal.Trim, "56", FUN_BUSCAR_TIPO_TARJETA(sNroTarjeta.Trim))
        'End If

        ''armar variable desde flash con: numero de cuenta, numero de tarjeta, nombre del cliente, modo de entrada, canal de atencion, id de sucursal,codigo kiosco y codigo de transaccion (opcion)
        'If Left(sRespuesta.Trim, 5) <> "ERROR" Then
        '    'VALIDACION CORRECTA

        '    sNroCuenta = Right("          " & sNroCuenta.Trim, 10)
        '    sNroTarjeta = Right("                    " & sNroTarjeta.Trim, 20)

        '    sCliente = Right(Strings.StrDup(26, " ") & sCliente.Trim, 26)

        '    sEstadoCuenta = "01" 'Estado de la cuenta
        '    sRespuestaServidor = "01" 'Atendido

        '    'ENVIAR_MONITOR
        '    sDataMonitor = ""
        '    sDataMonitor = FECHA_TX & HORA_TX & sNroCuenta & sNroTarjeta & sCliente & sModoEntrada & sCanalAtencion & sSaldoDisponible & sTotalDeuda & sIDSucursal & sIDTerminal & sEstadoCuenta & sRespuestaServidor & sCodigoTransaccion

        '    ENVIAR_MONITOR(sDataMonitor)

        'Else
        '    'ENVIAR_MONITOR
        '    sNroCuenta = Right("          " & ADATA_MONITOR(0), 10) 'LEER ARREGLO
        '    sNroTarjeta = Right(Strings.StrDup(20, " ") & ADATA_MONITOR(1), 20)
        '    sCliente = Strings.StrDup(26, " ")
        '    sEstadoCuenta = "02" 'Estado de la cuenta
        '    sRespuestaServidor = "02" 'Rechazado

        '    sDataMonitor = ""
        '    sDataMonitor = FECHA_TX & HORA_TX & sNroCuenta & sNroTarjeta & sCliente & sModoEntrada & sCanalAtencion & sSaldoDisponible & sTotalDeuda & sIDSucursal & sIDTerminal & sEstadoCuenta & sRespuestaServidor & sCodigoTransaccion

        '    ENVIAR_MONITOR(sDataMonitor)

        'End If


        Return oRespuesta

    End Function

    'DEVUELVE EL SALDO DE RIPLEY CLASICA
    '5254740045393942   000002405254    SIRIP-99    000002405254|\t|5254740045393942|\t|PARRA WILLIAMS, RUBEN EMILIO|\t|00|\t|      |\t|      |\t|2|\t|SIRIP-99|\t|11   2    NO
    <WebMethod(Description:="Obtiene la fecha de Facturacion para RipleyMatico")> _
    Public Function OBTENER_FECHA_FACTURACION(ByVal sNroTarjeta As String, _
                                          ByVal sNroCuenta As String, _
                                          ByVal sCodKiosco As String, _
                                          ByVal sDATA_MONITOR_KIOSCO As String, _
                                          ByVal Servidor As TServidor, _
                                          ByVal sMigrado As String, _
                                          ByVal sfechaConsumo As String) As String
        Dim Respuesta As String
        Respuesta = String.Empty

        ErrorLog("Llamar a método OBTENER_FECHA_FACTURACION(" & sNroTarjeta & "," & sNroCuenta & "," & sCodKiosco & "," & sDATA_MONITOR_KIOSCO & "," & Servidor & "," & sMigrado & "," & sfechaConsumo)

        Try
            Select Case Servidor

                Case TServidor.SICRON
                    Respuesta = OBTENER_DIAVENCIMIENTO_TARJETA_CLASICA_SICRON(sNroCuenta, sCodKiosco, sDATA_MONITOR_KIOSCO)
                Case TServidor.RSAT
                    Respuesta = OBTENER_DIAVENCIMIENTO_TARJETA_CLASICA_RSAT(sNroTarjeta, sNroCuenta, sMigrado)
                    ErrorLog("FIN DE OBTENER_DIAVENCIMIENTO_TARJETA_CLASICA_RSAT")
                Case Else
                    Respuesta = "ERROR:Servidor no especificado"
            End Select
            ErrorLog("Respuesta DESPUES DE LLAMAR A LOS SALDOS =" & Respuesta)
            If Respuesta.Length <= 2 Then
                Respuesta = fx_Completar_Campo("0", 2, Respuesta, TYPE_ALINEAR.DERECHA)
                ErrorLog("DIA DE FACT= " & Respuesta)
                Dim anio As String = sfechaConsumo.Substring(0, 4)
                Dim mes As String = sfechaConsumo.Substring(4, 2)
                Dim dia As String = sfechaConsumo.Substring(6, 2)

                If (dia > Respuesta) Then
                    ErrorLog("dia > Respuesta")
                    Dim numeroMes As Integer = CInt(mes)
                    Dim numeroAnio As Integer = CInt(anio)

                    If numeroMes = 12 Then
                        numeroMes = 1
                        numeroAnio += 1
                    Else
                        numeroMes += 1
                    End If

                    mes = fx_Completar_Campo("0", 2, numeroMes.ToString(), TYPE_ALINEAR.DERECHA)
                    anio = fx_Completar_Campo("0", 4, numeroAnio.ToString(), TYPE_ALINEAR.DERECHA)
                    ErrorLog("mes: " & mes)
                    ErrorLog("anio: " & anio)

                End If


                Respuesta = ObtenerFechaFacturacion(Respuesta, mes, anio)
            End If
        Catch ex As Exception
            Respuesta = "ERROR:Hubo un error en el método OBTENER_FECHA_FACTURACION. " & ex.Message
        End Try

        Return Respuesta
        '20141113
    End Function

    Private Function OBTENER_DIAVENCIMIENTO_TARJETA_CLASICA_RSAT(ByVal nrotarjeta As String, _
                                                ByVal nrocuenta As String, _
                                                ByVal sMigrado As String) As String
        ErrorLog("OBTENER_DIAVENCIMIENTO_TARJETA_CLASICA_RSAT")
        Dim objMQ As New MQ
        Dim strServicio As String = String.Empty
        Dim strMensaje As String = String.Empty
        Dim strRespuesta As String = String.Empty

        Try
            nrocuenta = nrocuenta.Trim
            nrotarjeta = nrotarjeta.Trim

            If (nrotarjeta.Substring(0, 6) = "542020" Or nrotarjeta.Substring(0, 6) = "525474" Or _
                nrotarjeta.Substring(0, 6) = "450034" Or nrotarjeta.Substring(0, 6) = "450035") And sMigrado = "SI" Then
                nrocuenta = "000000000000"
            End If
            ErrorLog("ReadAppConfig(SFSCAN_SALDO_FECHACORTE_CLASICA_RSAT): " & ReadAppConfig("SFSCAN_SALDO_FECHACORTE_CLASICA_RSAT"))
            strServicio = ReadAppConfig("SFSCAN_SALDO_FECHACORTE_CLASICA_RSAT")
            strMensaje = "0000000000" & strServicio & "                                       000000069000002USUARIO127  SANI0001PE0001" & nrocuenta & "" & nrotarjeta & "      604"
            ErrorLog("strMensaje= " & strMensaje)
            objMQ.Service = strServicio
            objMQ.Message = strMensaje
            ErrorLog("ANTES DE objMQ.Execute()")
            objMQ.Execute()
            ErrorLog("DESPUES DE objMQ.Execute()")


            If objMQ.ReasonMd <> 0 Then
                strRespuesta = "ERROR"
            End If
            If objMQ.ReasonApp = 0 Then
                strRespuesta = objMQ.Response
            End If

            'strRespuesta = "0000000000SFSCANT0114                                       000005895000002USUARIO127  SANI0001PE00010000000007049604100079592514      6040100000000000300000 000000000000000001120001COMPRAS                       00000000000292620120002EFECTIVO EXPRESS              00000000000292620120003NO EXISTE                     00000000000000000240001COMPRAS                       00000000000292620360001COMPRAS                       0000000000029262000                                  0000000000000000000                                  0000000000000000000                                  0000000000000000000                                  0000000000000000000                                  0000000000000000000                                  0000000000000000000                                  0000000000000000000                                  0000000000000000000                                  0000000000000000000                                  0000000000000000000                                  0000000000000000000                                  0000000000000000000                                  0000000000000000000                                  0000000000000000000                                  00000000000000000 0000000000000000 000000000000000000 000000000000000000000000 000000000000000000000000 000000000000000000000000 000000000000000000000000 000000000000000000000000 000000000000000000000000 000000000000000000000000 000000000000000000000000 000000000000000000000000 000000000000000000000000 0000000000000000 0000000000000000 00000000000000000000000                                                                                             0000000000000000  0000                                                                                             0000000000000000  0000                                                                                             0000000000000000  0000                                                                                             0000000000000000  0000                                                                                             0000000000000000  0000                                                                                             0000000000000000  0000                                                                                             0000000000000000  0000                                                                                             0000000000000000  0000                                                                                             0000000000000000  0000                                                                                             0000000000000000  0000                                                                                             0000000000000000  0000                                                                                             0000000000000000  0000                                                                                             0000000000000000  0000                                                                                             0000000000000000  0000                                                                                             0000000000000000  0000                                                                                             0000000000000000  0000                                                                                             0000000000000000  0000                                                                                             0000000000000000  0000                                                                                             0000000000000000  0000                                                                                             0000000000000000  0000                                                                                             0000000000000000  0000                                                                                             0000000000000000  0000                                                                                             0000000000000000  0000                                                                                             0000000000000000  0000                                                                                             0000000000000000                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                      "
            ErrorLog("strRespuesta=" & strRespuesta)
            If strRespuesta.Substring(0, 5) <> "ERROR" Then

                If Val(Mid(strRespuesta, 1, 10)) = 0 Then

                    strRespuesta = Mid(strRespuesta, 139, 2)
                    ErrorLog("Mid(strRespuesta, 139, 2)=" & strRespuesta)
                End If

                strRespuesta = Val(strRespuesta)
            End If
        Catch ex As Exception
            strRespuesta = "ERROR:" & ex.Message.Trim
        End Try

        Return strRespuesta

    End Function

    Private Function OBTENER_DIAVENCIMIENTO_TARJETA_CLASICA_SICRON(ByVal sNroCuenta As String, ByVal sCodKiosco As String, ByVal sDATA_MONITOR_KIOSCO As String) As String
        ErrorLog("OBTENER_DIAVENCIMIENTO_TARJETA_CLASICA_SICRON")
        Dim sRespuesta As String = ""
        Dim sParametros As String = ""
        Dim sMensajeErrorUsuario As String = ""

        'Variables nuevas
        Dim actions As Long = 1
        Dim inetputBuff As String = ""
        Dim outpputBuff As String = ""
        Dim errorMsg As String = ""

        Try
            If sNroCuenta.Trim.Length = 8 Then

                'Instancia al mirapiweb
                Dim obSendMirror As ClsTxMirapi = Nothing
                obSendMirror = New ClsTxMirapi()

                sParametros = Strings.StrDup(4, " ") & "245                00REU          1" & sNroCuenta.Trim & "000000000000" & "0" & "10000" & "000000000000" & Strings.StrDup(114, " ") & "00"
                inetputBuff = "      " + "V107" + sParametros
                ErrorLog("INICIO ExecuteTX")
                sRespuesta = obSendMirror.ExecuteTX(actions.ToString, inetputBuff, outpputBuff, TIME_OUT_SERVER, SERVIDOR_MIRROR_DESTINO, SERVIDOR_MIRROR_NODE, PUERTO, errorMsg)
                ErrorLog("FIN ExecuteTX")
                If sRespuesta = "0" Then 'EXITO

                    If outpputBuff.Length > 0 Then sRespuesta = outpputBuff.Substring(8, outpputBuff.Length - 8)

                    'Evaluar Respuesta si es ERROR,  RU, AU (Respuesta correcta)
                    If sRespuesta.Trim.Length > 0 Then
                        If Left(sRespuesta.Trim, 2) = "RU" Then 'Mensaje de Validacion al usuario. RU            000CUENTA BLOQUEADA
                            sMensajeErrorUsuario = Mid(sRespuesta.Trim, 18, Len(sRespuesta.Trim))
                            sRespuesta = "ERROR:NODATA-" & sMensajeErrorUsuario.Trim
                        Else
                            If Left(sRespuesta.Trim, 2) = "AU" Then

                                Dim lintDiaVencimiento As Long = 0

                                'dia de vencimiento de pago
                                lintDiaVencimiento = Val(Mid(sRespuesta, 250, 2))

                                sRespuesta = lintDiaVencimiento

                            Else 'Cualquier Otro Caso
                                sRespuesta = "ERROR:Servicio no disponible."
                            End If


                        End If

                    Else 'Sino devuelve nada
                        sRespuesta = "ERROR:Servicio no disponible."
                    End If

                ElseIf sRespuesta = "-2" Then 'Ocurrio un error Recuperar el error
                    sRespuesta = "ERROR:" & errorMsg.Trim
                Else
                    sRespuesta = "ERROR:Servicio no disponible."
                End If

            Else
                sRespuesta = "ERROR:Tarjeta no valida."
            End If

        Catch ex As Exception

            sRespuesta = "ERROR:" & ex.Message.Trim

        End Try

        Return sRespuesta

    End Function

    'DEVUELVE EL SALDO DE RIPLEY CLASICA
    <WebMethod(Description:="Consulta de Saldos Tarjeta Clasica SICRON | RSAT")> _
    Public Function SALDO_TARJETA_CLASICA(ByVal sNroTarjeta As String, _
                                          ByVal sNroCuenta As String, _
                                          ByVal sCodKiosco As String, _
                                          ByVal sDATA_MONITOR_KIOSCO As String, _
                                          ByVal Servidor As TServidor, _
                                          ByVal sMigrado As String) As String
        Dim Respuesta As String
        Respuesta = String.Empty

        Select Case Servidor

            Case TServidor.SICRON
                Respuesta = SALDO_TARJETA_CLASICA_SICRON(sNroCuenta, sCodKiosco, sDATA_MONITOR_KIOSCO)
            Case TServidor.RSAT
                Respuesta = SALDO_TARJETA_CLASICA_RSAT(sNroTarjeta, sNroCuenta, sMigrado)
            Case Else
                Respuesta = "ERROR:Servidor no especificado"
        End Select

        If Respuesta.Substring(0, 5) <> "ERROR" And sDATA_MONITOR_KIOSCO.Length > 0 Then
            'RECORRER ARREGLO
            Dim ADATA_MONITOR As Array
            ADATA_MONITOR = Split(sDATA_MONITOR_KIOSCO, "|\t|", , CompareMethod.Text)

            Dim sNroTarjetax As String = ADATA_MONITOR(1)
            Dim sIDSucursal As String = ADATA_MONITOR(6)
            Dim sIDTerminal As String = ADATA_MONITOR(7)

            'GUADAR LOG CONSULTAS
            SAVE_LOG_CONSULTAS(sIDSucursal.Trim, sIDTerminal.Trim, "30", FUN_BUSCAR_TIPO_TARJETA(sNroTarjetax.Trim))

        End If


        Return Respuesta

    End Function

    Private Function SALDO_TARJETA_CLASICA_SICRON(ByVal sNroCuenta As String, ByVal sCodKiosco As String, ByVal sDATA_MONITOR_KIOSCO As String) As String
        Dim sRespuesta As String = ""
        Dim sParametros As String = ""
        Dim sMensajeErrorUsuario As String = ""
        Dim sXML As String = ""

        'VARIABLES DEL SALDO DE TARJETA CLASICA
        Dim sLineaCredito As String = "" 'LINEA DE CRÉDITO
        Dim dLineaCreditoUtilizada As Double = 0 'LÍNEA DE CRÈDITO UTILIZADA
        Dim sDisponibleCompras As String = "" 'DISPONIBLE COMPRAS
        Dim sDisponibleEfectivo As String = "" 'DISPONIBLE EFECTIVO EXPRESS
        Dim sPagoTotalMes As String = "" 'PAGO TOTAL DEL MES
        Dim sPagoMinimo As String = "" 'PAGO MINIMO DEL MES
        Dim sPeriodo_Facturacion As String = "" 'PERIODO DE FACTURACION
        Dim sFechaPago As String = "" 'FECHA DE PAGO
        Dim sRipleyPuntos As String = "" 'RIPLEY PUNTOS
        Dim sDisponibleSuperEfectivo As String = ""
        Dim sPagoTotal As String = "" 'PAGO TOTAL


        'Variables nuevas
        Dim actions As Long = 1
        Dim inetputBuff As String = ""
        Dim outpputBuff As String = ""
        Dim errorMsg As String = ""

        Try

            If sNroCuenta.Trim.Length = 8 Then

                'Instancia al mirapiweb
                Dim obSendMirror As ClsTxMirapi = Nothing
                obSendMirror = New ClsTxMirapi()

                sParametros = Strings.StrDup(4, " ") & "245                00REU          1" & sNroCuenta.Trim & "000000000000" & "0" & "10000" & "000000000000" & Strings.StrDup(114, " ") & "00"
                inetputBuff = "      " + "V107" + sParametros

                sRespuesta = obSendMirror.ExecuteTX(actions.ToString, inetputBuff, outpputBuff, TIME_OUT_SERVER, SERVIDOR_MIRROR_DESTINO, SERVIDOR_MIRROR_NODE, PUERTO, errorMsg)

                If sRespuesta = "0" Then 'EXITO

                    If outpputBuff.Length > 0 Then sRespuesta = outpputBuff.Substring(8, outpputBuff.Length - 8)

                    'Evaluar Respuesta si es ERROR,  RU, AU (Respuesta correcta)
                    If sRespuesta.Trim.Length > 0 Then
                        If Left(sRespuesta.Trim, 2) = "RU" Then 'Mensaje de Validacion al usuario. RU            000CUENTA BLOQUEADA
                            sMensajeErrorUsuario = Mid(sRespuesta.Trim, 18, Len(sRespuesta.Trim))
                            sRespuesta = "ERROR:NODATA-" & sMensajeErrorUsuario.Trim
                        Else
                            If Left(sRespuesta.Trim, 2) = "AU" Then


                                Dim sFechaKiosco As String = DateTime.Now.Day.ToString("00").Trim + "/" + DateTime.Now.Month.ToString("00").Trim + "/" + DateTime.Now.Year.ToString("0000").Trim
                                Dim sHoraKiosco As String = DateTime.Now.ToString("HH:mm:ss")
                                Dim sNOperacion As String = GET_NUMERO_OPERACION()

                                'VARIABLES PARA CALCULOS
                                Dim sMonto1 As String = "" '
                                Dim sMonto2 As String = "" '
                                Dim lintDiaVencimiento As Long = 0
                                Dim lstrFechaEvaluacion As String = ""
                                Dim lstrFecFactVigente As String = ""
                                Dim lstrFecFactVigenteDDMMYYYY As String = ""
                                Dim lstrAnio As String = ""
                                Dim lstrMes As String = ""
                                Dim dMonto1 As Double = 0
                                Dim dMonto2 As Double = 0


                                Dim sTarjetaRecuperada As String = ""


                                sLineaCredito = Format(Val(Mid(sRespuesta, 181, 6)), "##,##0.00")
                                'Calculo de la Linea Utilizada
                                dMonto1 = CDbl(Trim(Mid(sRespuesta, 123, 13)))
                                dMonto2 = CDbl(Trim(Mid(sRespuesta, 168, 13)))
                                dLineaCreditoUtilizada = dMonto1 + dMonto2

                                sDisponibleCompras = Format(Val(Mid(sRespuesta, 144, 8)), "##,##0.00")

                                sTarjetaRecuperada = Mid(sRespuesta, 28, 16)

                                'dia de vencimiento de pago
                                lintDiaVencimiento = Val(Mid(sRespuesta, 250, 2))


                                Dim vFecCorteIni As Date
                                Dim vFecCorteFin As Date
                                Dim vFecPago As Date
                                Dim vFecHoy As Date
                                Dim vDia As Integer
                                Dim sFechaPagox As String = ""
                                Dim sFecha1 As String = ""
                                Dim sFecha2 As String = ""

                                vFecHoy = Date.Now
                                vDia = lintDiaVencimiento

                                vFecPago = CStr(funDevuelve_FechaPago(vFecHoy, vDia))
                                Call procDevuelve_PeriodoFacturacion(vFecPago, vFecCorteIni, vFecCorteFin)


                                'Formatear DDMMYYYY
                                sFechaPagox = Right(Trim("00" & Day(vFecPago).ToString), 2) & Right(Trim("00" & Month(vFecPago).ToString), 2) & Right(Trim("0000" & Year(vFecPago).ToString), 4)

                                sFecha1 = Right(Trim("00" & Day(vFecCorteIni).ToString), 2) & Right(Trim("00" & Month(vFecCorteIni).ToString), 2) & Right(Trim("0000" & Year(vFecCorteIni).ToString), 4)
                                sFecha2 = Right(Trim("00" & Day(vFecCorteFin).ToString), 2) & Right(Trim("00" & Month(vFecCorteFin).ToString), 2) & Right(Trim("0000" & Year(vFecCorteFin).ToString), 4)

                                'PERIODO DE FACTURACION
                                sFechaPago = Cambia_Fecha(sFechaPagox.Trim)
                                sPeriodo_Facturacion = Cambia_Fecha(sFecha1.Trim) & " - " & Cambia_Fecha(sFecha2.Trim)

                                'Periodo de LLamada Anio y Mes
                                lstrAnio = Right(Trim("0000" & Year(vFecPago).ToString), 4)
                                lstrMes = Right(Trim("00" & Month(vFecPago).ToString), 2)

                                Dim lEstadoCta As Long = 2
                                Dim sDataEECC_SALDO As String = ""

                                'REALIZAR TRES INTENTOS SEGUN LA FECHA ACTUAL
                                Dim sPeriodoFinal As String = ""
                                Dim sPeriodo1 As String = ""
                                Dim sPeriodo2 As String = ""
                                Dim sPeriodo3 As String = ""
                                Dim vFechaHoy As Date
                                Dim fFechaHoyAux1 As Date
                                Dim lEstadoLlamadaEECC As Long = 9 'Variable para iniciar llamadas al EECC 3 veces


                                vFechaHoy = Date.Now
                                fFechaHoyAux1 = vFechaHoy

                                fFechaHoyAux1 = DateAdd("m", 1, fFechaHoyAux1)

                                sPeriodo1 = Right(Trim("0000" & Year(fFechaHoyAux1).ToString), 4) & Right(Trim("00" & Month(fFechaHoyAux1).ToString), 2)

                                sPeriodo2 = Right(Trim("0000" & Year(vFechaHoy).ToString), 4) & Right(Trim("00" & Month(vFechaHoy).ToString), 2)

                                vFechaHoy = DateAdd("m", -1, vFechaHoy)

                                sPeriodo3 = Right(Trim("0000" & Year(vFechaHoy).ToString), 4) & Right(Trim("00" & Month(vFechaHoy).ToString), 2)

                                sPeriodoFinal = sPeriodo1.Trim


                                '' Obtencion de estado de cuenta ...
                                'Primera llamada
                                lEstadoCta = pfnblnObtenerEstadoCuenta(sTarjetaRecuperada, sNroCuenta, sPeriodoFinal.Trim, sDataEECC_SALDO)

                                If sDataEECC_SALDO.Trim = "" Then
                                    'Segunda LLamada segundo periodo
                                    sPeriodoFinal = sPeriodo2.Trim
                                    lEstadoCta = pfnblnObtenerEstadoCuenta(sTarjetaRecuperada, sNroCuenta, sPeriodoFinal.Trim, sDataEECC_SALDO)
                                End If

                                If sDataEECC_SALDO.Trim = "" Then
                                    'Tercera LLamada Tercer periodo
                                    sPeriodoFinal = sPeriodo3.Trim
                                    lEstadoCta = pfnblnObtenerEstadoCuenta(sTarjetaRecuperada, sNroCuenta, sPeriodoFinal.Trim, sDataEECC_SALDO)
                                End If


                                If sDataEECC_SALDO.Trim = "" Then
                                    sPagoTotalMes = "<EECC/Pend>"
                                    sPagoMinimo = "<EECC/Pend>"
                                Else
                                    Dim SDATA_EECC As String = ""
                                    Dim ADATA_EECC As Array

                                    SDATA_EECC = sDataEECC_SALDO
                                    ADATA_EECC = Split(SDATA_EECC, "|\t|", , CompareMethod.Text)

                                    sPagoTotalMes = ADATA_EECC(0)
                                    sPagoMinimo = ADATA_EECC(1)

                                End If


                                Dim sDataDisponiblePagoTotal As String = ""
                                sDataDisponiblePagoTotal = DISPONIBLE_EFEC_EXPRESS_PAGO_TOTAL(sTarjetaRecuperada.Trim, sCodKiosco.Trim)

                                If sDataDisponiblePagoTotal.Trim.Length = 0 Then
                                    sDisponibleEfectivo = "0.00"
                                    sPagoTotal = "0.00"
                                Else

                                    Dim SDATA_EFECTIVO As String = ""
                                    Dim ADATA_EFECTIVO As Array

                                    SDATA_EFECTIVO = sDataDisponiblePagoTotal
                                    ADATA_EFECTIVO = Split(SDATA_EFECTIVO, "|\t|", , CompareMethod.Text)

                                    sDisponibleEfectivo = ADATA_EFECTIVO(0)
                                    sPagoTotal = ADATA_EFECTIVO(1)

                                End If

                                'RIPLEY PUNTOS ACUMULADOS
                                sRipleyPuntos = PUNTOS_RIPLEY_ACUMULADOS(sTarjetaRecuperada.Trim)

                                Dim sDataSuperEfectivo As String = ""
                                sDataSuperEfectivo = DISPONIBLE_SUPER_EFECTIVO(sNroCuenta.Trim)
                                If sDataSuperEfectivo.Trim.Length = 0 Then
                                    sDisponibleSuperEfectivo = "0.00"
                                Else
                                    sDisponibleSuperEfectivo = sDataSuperEfectivo.Trim
                                End If


                                sXML = sLineaCredito.Trim & "|\t|" & Format(dLineaCreditoUtilizada, "##,##0.00") & "|\t|" & sDisponibleCompras.Trim & "|\t|"
                                sXML = sXML & sDisponibleEfectivo.Trim & "|\t|" & sPagoTotalMes.Trim & "|\t|" & sPagoMinimo.Trim & "|\t|"
                                sXML = sXML & sPeriodo_Facturacion.Trim & "|\t|" & sFechaPago.Trim & "|\t|"
                                sXML = sXML & sRipleyPuntos.Trim & "|\t|" & sDisponibleSuperEfectivo.Trim & "|\t|" & sPagoTotal.Trim & "|\t|"
                                sXML = sXML & sFechaKiosco.Trim & "|\t|" & sHoraKiosco.Trim & "|\t|" & sNOperacion.Trim


                                sRespuesta = sXML

                            Else 'Cualquier Otro Caso
                                sRespuesta = "ERROR:Servicio no disponible."
                            End If


                        End If

                    Else 'Sino devuelve nada
                        sRespuesta = "ERROR:Servicio no disponible."
                    End If

                ElseIf sRespuesta = "-2" Then 'Ocurrio un error Recuperar el error
                    sRespuesta = "ERROR:" & errorMsg.Trim
                Else
                    sRespuesta = "ERROR:Servicio no disponible."
                End If

            Else
                'Mostrar Mensaje de Error
                sRespuesta = "ERROR:Tarjeta no valida."
            End If

        Catch ex As Exception
            'save log error

            sRespuesta = "ERROR:" & ex.Message.Trim

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

            sSaldoDisponible = sDisponibleCompras.Trim
            sSaldoDisponible = Replace(sSaldoDisponible.Trim, ",", "") 'QUITAR LA COMA
            sSaldoDisponible = Replace(sSaldoDisponible.Trim, ".", "") 'QUITAR EL PUNTO DECIMAL
            sSaldoDisponible = Mid(sSaldoDisponible, 1, sSaldoDisponible.Trim.Length - 2) 'QUITAR LOS DOS ULTIMOS DIGITOS QUE SON DECIMALES

            sTotalDeuda = sPagoTotal.Trim
            sTotalDeuda = Replace(sTotalDeuda.Trim, ",", "") 'QUITAR LA COMA
            sTotalDeuda = Replace(sTotalDeuda.Trim, ".", "") 'QUITAR EL PUNTO DECIMAL
            sTotalDeuda = Mid(sTotalDeuda, 1, sTotalDeuda.Trim.Length - 2) 'QUITAR LOS DOS ULTIMOS DIGITOS QUE SON DECIMALES



            sIDSucursal = ADATA_MONITOR(6)
            sIDTerminal = ADATA_MONITOR(7)
            sCodigoTransaccion = ADATA_MONITOR(8)


            'GUADAR LOG CONSULTAS
            ''SAVE_LOG_CONSULTAS(sIDSucursal.Trim, sIDTerminal.Trim, "30", FUN_BUSCAR_TIPO_TARJETA(sNroTarjetax.Trim))


            sNroCuentax = Right("          " & sNroCuentax.Trim, 10)
            sNroTarjetax = Right("          " & sNroTarjetax.Trim, 20)
            sNombreCliente = Right(Strings.StrDup(26, " ") & sNombreCliente.Trim, 26)

            sSaldoDisponible = Right("      " & sSaldoDisponible.Trim, 6)
            sTotalDeuda = Right("      " & sTotalDeuda.Trim, 6)


            'armar variable desde flash con: numero de cuenta, numero de tarjeta, nombre del cliente, modo de entrada, canal de atencion, id de sucursal,codigo kiosco y codigo de transaccion (opcion)
            If Left(sRespuesta.Trim, 5) <> "ERROR" Then
                'VALIDACION CORRECTA

                sEstadoCuenta = "01" 'Estado de la cuenta
                sRespuestaServidor = "01" 'Atendido

                'ENVIAR_MONITOR
                sDataMonitor = ""
                sDataMonitor = FECHA_TX & HORA_TX & sNroCuentax & sNroTarjetax & sNombreCliente & sModoEntrada & sCanalAtencion & sSaldoDisponible & sTotalDeuda & sIDSucursal & sIDTerminal & sEstadoCuenta & sRespuestaServidor & sCodigoTransaccion

                ''ENVIAR_MONITOR(sDataMonitor)

            Else
                'ENVIAR_MONITOR

                sEstadoCuenta = "01" 'Estado de la cuenta
                sRespuestaServidor = "02" 'Rechazado

                sDataMonitor = ""
                sDataMonitor = FECHA_TX & HORA_TX & sNroCuentax & sNroTarjetax & sNombreCliente & sModoEntrada & sCanalAtencion & sSaldoDisponible & sTotalDeuda & sIDSucursal & sIDTerminal & sEstadoCuenta & sRespuestaServidor & sCodigoTransaccion

                ''ENVIAR_MONITOR(sDataMonitor)

            End If

        End If

        Return sRespuesta

    End Function

    'Movimientos CLASICA
    <WebMethod(Description:="Consulta de Movimientos Tarjeta Clasica SICRON | RSAT")> _
    Public Function MOVIMIENTOS_CLASICA(ByVal sNroTarjeta As String, ByVal sNroCuenta As String, ByVal sDATA_MONITOR_KIOSCO As String, ByVal Servidor As TServidor, ByVal sMigrado As String) As String
        ErrorLog("MOVIMIENTOS_CLASICA(" + sNroTarjeta + "," + sNroCuenta + "," + sDATA_MONITOR_KIOSCO + "," + Servidor.ToString() + "," + sMigrado + ")")
        Dim Respuesta As String
        Respuesta = String.Empty

        Select Case Servidor
            Case TServidor.SICRON
                Respuesta = MOVIMIENTOS_CLASICA_SICRON(sNroCuenta, sDATA_MONITOR_KIOSCO)
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
            SAVE_LOG_CONSULTAS(sIDSucursal.Trim, sIDTerminal.Trim, "34", FUN_BUSCAR_TIPO_TARJETA(sNroTarjetax.Trim))
        End If

        Return Respuesta
    End Function

    Private Function MOVIMIENTOS_CLASICA_SICRON(ByVal sNroCuenta As String, ByVal sDATA_MONITOR_KIOSCO As String) As String
        Dim sRespuesta As String = ""
        Dim sParametros As String = ""
        Dim sMensajeErrorUsuario As String = ""
        Dim sXML As String = ""

        'Variables nuevas
        Dim actions As Long = 1
        Dim inetputBuff As String = ""
        Dim outpputBuff As String = ""
        Dim errorMsg As String = ""

        Try


            'Instancia al mirapiweb
            Dim obSendMirror As ClsTxMirapi = Nothing
            obSendMirror = New ClsTxMirapi()

            Dim lPagina As Long = 0 'Llamadas 0 1 2 3
            Dim gNumReg As Long = 0
            Dim TotMovimientos As Double = 0
            Dim bContinuaCall As Boolean = True
            Dim sDATA_MOVI As String = "" 'Variable que concatena los movimientos

            Do Until bContinuaCall = False

                sRespuesta = ""
                sParametros = "                                      " & sNroCuenta.Trim & "0000000000000" & lPagina.ToString & "1"
                inetputBuff = "      " + "V500" + sParametros

                sRespuesta = obSendMirror.ExecuteTX(actions.ToString, inetputBuff, outpputBuff, TIME_OUT_SERVER, SERVIDOR_MIRROR_DESTINO, SERVIDOR_MIRROR_NODE, PUERTO, errorMsg)

                If sRespuesta = "0" Then 'EXITO

                    If outpputBuff.Length > 0 Then sRespuesta = outpputBuff.Substring(8, outpputBuff.Length - 8)


                    If Left(sRespuesta.Trim, 2) = "RU" Then 'Mensaje de Validacion al usuario.
                        sMensajeErrorUsuario = Mid(sRespuesta.Trim, 3, Len(sRespuesta.Trim))
                        sRespuesta = "ERROR:NODATA-" & sMensajeErrorUsuario.Trim
                        Exit Do
                    Else
                        If Left(sRespuesta.Trim, 2) = "AU" Then
                            'Recuperar Data

                            Dim sFechaMov As String = ""
                            Dim sDesMov As String = ""
                            Dim sEstablecimiento As String = ""
                            Dim sTicket As String = ""
                            Dim sMonto As String = ""

                            'Data de Movimientos
                            sXML = ""
                            sXML = Cortar_Movimientos_Clasica(sRespuesta, TotMovimientos, bContinuaCall, gNumReg)
                            sDATA_MOVI = sDATA_MOVI & sXML

                        End If
                    End If

                ElseIf sRespuesta = "-2" Then 'Ocurrio un error Recuperar el error
                    sRespuesta = "ERROR:" & errorMsg.Trim
                    Exit Do 'salir del while
                Else
                    sRespuesta = "ERROR:Servicio no disponible."
                    Exit Do 'salir del while
                End If


                If TotMovimientos > 0 Then
                    lPagina = lPagina + 1
                Else
                    sRespuesta = sDATA_MOVI
                    Exit Do 'salir del while
                End If

            Loop

            If sDATA_MOVI.Length > 0 Then

                'Variables de Cabecera de Kiosco
                Dim sFechaKiosco As String = DateTime.Now.Day.ToString("00").Trim + "/" + DateTime.Now.Month.ToString("00").Trim + "/" + DateTime.Now.Year.ToString("0000").Trim
                Dim sHoraKiosco As String = DateTime.Now.ToString("HH:mm:ss")
                Dim sNOperacion As String = GET_NUMERO_OPERACION()

                Dim DAT_NUM_OPERACION As String = sFechaKiosco & "|\t|" & sHoraKiosco & "|\t|" & sNOperacion
                Dim sDataOrdenada As String = Left(sDATA_MOVI, sDATA_MOVI.Trim.Length - 4)

                'CALL FUNCION PARA ODERNAR LA DATA
                sDataOrdenada = fun_OrdenarMovimientos(sDataOrdenada)

                sRespuesta = DAT_NUM_OPERACION & "|¿**?|" & sDataOrdenada

            End If

        Catch ex As Exception

            sRespuesta = "ERROR:" & ex.Message.Trim

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
            sNroTarjetax = Right("          " & sNroTarjetax.Trim, 20)
            sNombreCliente = Right(Strings.StrDup(26, " ") & sNombreCliente.Trim, 26)

            sSaldoDisponible = Right("      " & sSaldoDisponible.Trim, 6)
            sTotalDeuda = Right("      " & sTotalDeuda.Trim, 6)


            'armar variable desde flash con: numero de cuenta, numero de tarjeta, nombre del cliente, modo de entrada, canal de atencion, id de sucursal,codigo kiosco y codigo de transaccion (opcion)
            If Left(sRespuesta.Trim, 5) <> "ERROR" Then
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


        Return sRespuesta

    End Function



    'DEVUELVE EL SALDO DE RIPLEY CLASICA
    <WebMethod(Description:="Consulta de Saldos y Disponibles Tarjeta Ripley Asociada.")> _
    Public Function SALDO_TARJETA_ASOCIADA(ByVal sNroTarjeta As String, ByVal sDATA_MONITOR_KIOSCO As String) As Respuesta

        Dim oRespuesta As Respuesta = New Respuesta
        If sNroTarjeta.Substring(0, 6) = "542020" Or sNroTarjeta.Substring(0, 6) = "525474" Or _
           sNroTarjeta.Substring(0, 6) = "450035" Or sNroTarjeta.Substring(0, 6) = "450034" Then

            oRespuesta = VALIDAR_TARJETA_ABIERTA_RSAT(sNroTarjeta, sDATA_MONITOR_KIOSCO)
            If oRespuesta.Estado <> "ERROR" Then
                Return oRespuesta
            Else
                If oRespuesta.Codigo = "MSGMSG" Then
                    Return oRespuesta
                End If
            End If

        End If

        Dim sGetTramaMC As String = ""
        Dim sParametros As String = ""
        Dim sMensajeErrorUsuario As String = ""
        Dim sXML As String = ""
        Dim sTarjetaRecuperada As String = ""
        Dim sNombres As String = "" 'Nombres del cliente

        'VARIABLES DEL SALDO DE TARJETA ASOCIADA
        Dim sLineaCredito As String = "" 'LINEA DE CRÉDITO
        Dim sLineaCreditoUtilizada As String = "" 'LÍNEA DE CRÈDITO UTILIZADA
        Dim sDisponibleCompras As String = "" 'DISPONIBLE COMPRAS
        Dim sDisponibleEfectivo As String = "" 'DISPONIBLE EFECTIVO EXPRESS
        Dim sPagoTotalMes As String = "" 'PAGO TOTAL DEL MES
        Dim sPagoMinimo As String = "" 'PAGO MINIMO DEL MES
        Dim sPeriodo_Facturacion As String = "" 'PERIODO DE FACTURACION
        Dim sFechaPago As String = "" 'FECHA DE PAGO
        Dim sRipleyPuntos As String = "" 'RIPLEY PUNTOS
        Dim sDisponibleSuperEfectivo As String = ""
        Dim sPagoTotal As String = "" 'PAGO TOTAL

        'Return "2,220.00|\t|1,357.56|\t|862.44|\t|862.00|\t|481.63|\t|0.00|\t|11/09/2013 - 10/10/2013|\t|25/10/2013|\t|2728|\t|0.00|\t|1,357.63|\t|07/11/2013|\t|17:27:17|\t|005138*$¿*5420200004815764|\t|41690197|\t|ALTAMIRANO LEIVA, JUAN ALBERTO|\t||\t|542020|\t||\t||\t|5420200004815764|\t||\t|1|\t|1|\t|41690197|\t|Super Efectivo GOLD MC|\t|JUAN ALBERTO ALTAMIRANO LEIVA|\t|5420200004812308|\t|4000|\t|36|\t|66.59|\t|218.15|\t|18/11/2013|\t|9737965|\t|3.99|\t||\t|5420200004812308|\t|MC"

        Try


            If sNroTarjeta.Trim.Length = 16 Then

                Dim obSendWAS As New WSCONSULTAS_MC_VS.defaultService

                sParametros = "HQ000001073RMDQCS" & sNroTarjeta.Trim & Strings.StrDup(16, " ") & "0000000000" & "0" & "000000" & "000073" & "@"

                sGetTramaMC = obSendWAS.execute(sParametros.Trim)

                'Evaluar Respuesta si es ERROR
                If sGetTramaMC.Trim.Length > 0 Then
                    If Left(sGetTramaMC.Trim, 2) = "HE" Then
                        'sGetTramaMC = "ERROR:" & sGetTramaMC.Trim
                        sGetTramaMC = "ERROR:" & Trim(Mid(sGetTramaMC, 16, 30)) 'CORTAR MENSAJE DE ERROR
                        oRespuesta.Cadena = sGetTramaMC
                        oRespuesta.Codigo = "MSGMSG"
                        oRespuesta.Estado = "ERROR"
                        oRespuesta.Mensaje = Trim(Mid(sGetTramaMC, 16, 30))
                    Else

                        Dim sFechaKiosco As String = DateTime.Now.Day.ToString("00").Trim + "/" + DateTime.Now.Month.ToString("00").Trim + "/" + DateTime.Now.Year.ToString("0000").Trim
                        Dim sHoraKiosco As String = DateTime.Now.ToString("HH:mm:ss")
                        Dim sNOperacion As String = GET_NUMERO_OPERACION()

                        'VARIABLES PARA CALCULOS
                        Dim dPagoTotalMes As Double = 0
                        Dim dITF As Double = 0
                        Dim dPagoMinimoMes As Double = 0
                        Dim sDia As String = ""
                        Dim sMes As String = ""
                        Dim sAnio As String = ""
                        Dim dPagoTotal As Double = 0


                        sTarjetaRecuperada = sNroTarjeta.Trim 'formatear 000-0000-0000-0000
                        sNombres = Trim(Mid(sGetTramaMC, 73, 30))

                        If sNombres.Trim = "" Then
                            sGetTramaMC = "ERROR:NODATA" 'SI NO DEVUELVE NOMBRES LANSAR MENSAJE TERMINA
                            oRespuesta.Cadena = "ERROR:NODATA"
                            oRespuesta.Codigo = "MSGMSG"
                            oRespuesta.Estado = "ERROR"
                            oRespuesta.Mensaje = Constantes.MSG_NO_DATOS_CLIENTE
                            Return oRespuesta
                        End If

                        sNombres = Replace(sNombres, "#", "Ñ")

                        If sNombres Is Nothing Then
                            sNombres = ""
                        End If

                        sLineaCredito = Format(Val(SET_MONTO_DECIMAL(CLng(Mid(sGetTramaMC, 133, 14)).ToString)), "###,##0.00")

                        sLineaCreditoUtilizada = Format(Val(SET_MONTO_DECIMAL(CLng(Mid(sGetTramaMC, 193, 14)).ToString)), "###,##0.00")

                        sDisponibleCompras = Format(Val(SET_MONTO_DECIMAL(CLng(Mid(sGetTramaMC, 147, 14)).ToString)), "###,##0.00")

                        sDisponibleEfectivo = Format(Val(SET_MONTO_DECIMAL(CLng(Mid(sGetTramaMC, 163, 14)).ToString)), "###,##0.00")

                        'Calculo pago total del mes
                        dPagoTotalMes = Val(SET_MONTO_DECIMAL(CLng(Mid(sGetTramaMC, 207, 14)).ToString))
                        'Obtener el ITF
                        dITF = OBTENER_FACTOR_ITF()
                        dPagoTotalMes = dPagoTotalMes * dITF
                        sPagoTotalMes = Format(dPagoTotalMes, "###,##0.00")

                        'Calculo pago minimo del mes
                        dPagoMinimoMes = Val(SET_MONTO_DECIMAL(CLng(Mid(sGetTramaMC, 179, 14)).ToString))
                        'Obtener el ITF
                        dPagoMinimoMes = dPagoMinimoMes * dITF
                        sPagoMinimo = Format(dPagoMinimoMes, "###,##0.00")


                        'Fecha de Pago
                        sDia = Trim(Mid(sGetTramaMC, 227, 2))
                        sMes = Trim(Mid(sGetTramaMC, 225, 2))
                        sAnio = Trim(Mid(sGetTramaMC, 221, 4))

                        sFechaPago = sDia.Trim & "/" & sMes.Trim & "/" & sAnio

                        If Not IsDate(sFechaPago.Trim) Then
                            sFechaPago = "NO DISPONIBLE"
                        End If

                        'PERIODO DE FACTURACION
                        Dim sFecha1 As String = ""
                        Dim sFecha2 As String = ""
                        Dim lMes1 As Integer = 0
                        Dim lAnio As Integer = 0
                        Dim lAnio1 As Integer = 0
                        Dim lAnio2 As Integer = 0


                        If sFechaPago.Trim.ToUpper <> "NO DISPONIBLE" Then

                            If CInt(sMes.Trim) < 2 Then
                                lMes1 = 12 + CInt(sMes.Trim)
                            Else
                                lMes1 = CInt(sMes.Trim)
                            End If

                            lAnio = CInt(sAnio.Trim)

                            'Año1
                            If CInt(sMes.Trim) < 3 Then
                                lAnio1 = lAnio - 1
                            Else
                                lAnio1 = lAnio
                            End If

                            'Año2
                            If CInt(sMes.Trim) < 3 Then
                                Select Case CInt(sMes.Trim)
                                    Case 1
                                        lAnio2 = lAnio - 1
                                    Case 2
                                        lAnio2 = lAnio
                                End Select

                            Else
                                lAnio2 = lAnio
                            End If


                            If CInt(sDia.Trim) = 5 Then
                                If lMes1 = 2 Then
                                    sFecha1 = "21" & "/" & Right("12", 2) & "/" & lAnio1.ToString
                                Else
                                    sFecha1 = "21" & "/" & Right("00" & (lMes1 - 2).ToString, 2) & "/" & lAnio1.ToString
                                End If

                                sFecha2 = "20" & "/" & Right("00" & (lMes1 - 1).ToString, 2) & "/" & lAnio2.ToString
                            End If


                            If CInt(sDia.Trim) = 25 Then
                                sFecha1 = "11" & "/" & Right("00" & (lMes1 - 1).ToString, 2) & "/" & lAnio2.ToString
                                sFecha2 = "10" & "/" & Right("00" & sMes.Trim, 2) & "/" & sAnio.Trim
                            End If


                            sPeriodo_Facturacion = sFecha1.Trim & " - " & sFecha2.Trim
                        Else
                            sPeriodo_Facturacion = "NO DISPONIBLE"
                        End If


                        'Pago Total

                        dPagoTotal = Val(SET_MONTO_DECIMAL(CLng(Mid(sGetTramaMC, 193, 14)).ToString))
                        dPagoTotal = dPagoTotal * dITF
                        sPagoTotal = Format(dPagoTotal, "###,##0.00")


                        'RIPLEY PUNTOS ACUMULADOS
                        sRipleyPuntos = PUNTOS_RIPLEY_ACUMULADOS(sTarjetaRecuperada.Trim)

                        'Buscar Tipo Doc y Nro Documento desde oracle
                        Dim sTipoDocumento As String = "0"
                        Dim sNroDocumento As String = "000000000000"
                        Dim Nro_Contrato As String = String.Empty
                        Dim sDataTipoDocum As String = ""
                        Dim ADATA_TDOC As Array
                        sDataTipoDocum = CONSULTAR_DOC_MC_VISA(sTarjetaRecuperada.Trim, Nro_Contrato)

                        If sDataTipoDocum.Length > 0 Then
                            ADATA_TDOC = Split(sDataTipoDocum, "|\t|", , CompareMethod.Text)
                            sTipoDocumento = ADATA_TDOC(0)
                            sNroDocumento = ADATA_TDOC(1)

                            If sTipoDocumento.Trim = "1" Then 'DNI
                                sNroDocumento = Right(sNroDocumento.Trim, 8)
                            Else
                                'CE
                                sNroDocumento = Right(sNroDocumento.Trim, 12)
                            End If

                        End If

                        Dim sDataSuperEfectivo As String = ""
                        Dim sMontoAprobadoSuperEfectivo As String = ""
                        Dim sFechaVenceSEfectivo As String = ""

                        sDataSuperEfectivo = DISPONIBLE_SUPER_EFECTIVO_MC_VISA(sTarjetaRecuperada.Trim, sTipoDocumento.Trim, sNroDocumento.Trim)
                        If sDataSuperEfectivo.Trim.Length = 0 Then
                            sDisponibleSuperEfectivo = "0.00"
                            sFechaVenceSEfectivo = ""
                            sMontoAprobadoSuperEfectivo = ""
                        Else
                            Dim ADATASUPER As Array
                            ADATASUPER = Split(sDataSuperEfectivo, "|\t|", , CompareMethod.Text)
                            sMontoAprobadoSuperEfectivo = ADATASUPER(0)
                            sFechaVenceSEfectivo = ADATASUPER(1)
                            sDisponibleSuperEfectivo = ADATASUPER(0)

                            If sDisponibleSuperEfectivo.Trim.Length = 0 Then
                                sDisponibleSuperEfectivo = "0.00"
                            End If

                        End If

                        sXML = sLineaCredito.Trim & "|\t|" & sLineaCreditoUtilizada & "|\t|" & sDisponibleCompras.Trim & "|\t|"
                        sXML = sXML & sDisponibleEfectivo.Trim & "|\t|" & sPagoTotalMes.Trim & "|\t|" & sPagoMinimo.Trim & "|\t|"
                        sXML = sXML & sPeriodo_Facturacion.Trim & "|\t|" & sFechaPago.Trim & "|\t|"
                        sXML = sXML & sRipleyPuntos.Trim & "|\t|" & sDisponibleSuperEfectivo.Trim & "|\t|" & sPagoTotal.Trim & "|\t|"
                        sXML = sXML & sFechaKiosco.Trim & "|\t|" & sHoraKiosco.Trim & "|\t|" & sNOperacion.Trim

                        Dim sDatosCliente As String = ""
                        Dim sTipoProducto As String = ""
                        Dim sCuenta As String = ""

                        sTipoProducto = Mid(sNroTarjeta, 1, 6)
                        sDatosCliente = sNroTarjeta.Trim & "|\t|" & sNroDocumento.Trim & "|\t|" & sNombres.Trim & "|\t||\t|" & sTipoProducto.Trim & "|\t|" & sMontoAprobadoSuperEfectivo.Trim & "|\t|" & sFechaVenceSEfectivo.Trim & "|\t|" & sNroTarjeta.Trim & "|\t||\t|" & MOSTRAR_SUPEREFECTIVO_SEF("1", sTipoDocumento.Trim, sNroDocumento.Trim) & "|\t|"
                        sXML = sXML & "*$¿*" & sDatosCliente
                        sGetTramaMC = sXML & "|\t|" & Nro_Contrato & "|\t|" & "MC"

                        oRespuesta.Cadena = sGetTramaMC
                        oRespuesta.Estado = "EXITO"


                        Dim ADATA_MONITOR As Array
                        Dim IDSucursal, IDTerminal As String
                        IDSucursal = ""
                        IDTerminal = ""

                        If sDATA_MONITOR_KIOSCO.Length > 0 Then
                            ADATA_MONITOR = Split(sDATA_MONITOR_KIOSCO, "|\t|", , CompareMethod.Text)
                            IDSucursal = ADATA_MONITOR(3)
                            IDTerminal = ADATA_MONITOR(4)

                            getRegistrar_Ingreso_Ripleymatico(IDSucursal, IDTerminal, sNroDocumento, sTipoDocumento, sNroTarjeta, sTipoProducto)

                        End If



                    End If

                Else 'Sino devuelve nada
                    sGetTramaMC = "ERROR:NODATA"
                    oRespuesta.Cadena = "ERROR:NODATA"
                    oRespuesta.Codigo = "MSGMSG"
                    oRespuesta.Estado = "ERROR"
                    oRespuesta.Mensaje = Constantes.MSG_NO_DATOS_CLIENTE
                End If

            Else
                'Mostrar Mensaje de Error
                sGetTramaMC = "ERROR:NODATA"
                oRespuesta.Cadena = "ERROR:NODATA"
                oRespuesta.Codigo = "MSGMSG"
                oRespuesta.Estado = "ERROR"
                oRespuesta.Mensaje = Constantes.MSG_TARJETA_NO_VALIDA
            End If

        Catch ex As Exception
            'save log error

            sGetTramaMC = "ERROR:NODATA"
            oRespuesta.Cadena = "ERROR:NODATA"
            oRespuesta.Codigo = "MSGMSG"
            oRespuesta.Estado = "ERROR"
            oRespuesta.Mensaje = Constantes.MSG_ERROR_CODIGO
        End Try


        ''Save Monitor
        'Dim sDataMonitor As String = ""
        'Dim FECHA_TX As String = DateTime.Now.Year.ToString("0000").Trim & DateTime.Now.Month.ToString("00").Trim + DateTime.Now.Day.ToString("00").Trim
        'Dim HORA_TX As String = DateTime.Now.Hour.ToString("00").Trim & DateTime.Now.Minute.ToString("00").Trim & DateTime.Now.Second.ToString("00").Trim
        'Dim sModoEntrada As String = ""
        'Dim sCanalAtencion As String = "01" 'RipleyMatico
        'Dim sSaldoDisponible As String = "      " 'Right("      " & sDisponibleCompras.Trim, 6)
        'Dim sTotalDeuda As String = "      " 'Right("      " & sPagoTotal.Trim, 6)
        Dim sIDSucursal As String = ""
        Dim sIDTerminal As String = ""
        'Dim sEstadoCuenta As String = ""
        'Dim sRespuestaServidor As String = ""
        'Dim sCodigoTransaccion As String = ""
        'Dim sNroCuenta As String = ""

        If sDATA_MONITOR_KIOSCO.Length > 0 Then


            'ORDEN DE LAS VARRIABLES
            'DATA_MONITOR=_root.CUENTA_TARJETA_VALIDAR+"|\u005Ct|"+_root.P_NroTarjeta+"|\u005Ct|"+"01"+"|\u005Ct|"+_global.gCODE_SUCURSAL+"|\u005Ct|"+_global.gCODE_KIOSCO+"|\u005Ct|"+"01";
            'RECORRER ARREGLO
            Dim ADATA_MONITOR As Array
            ADATA_MONITOR = Split(sDATA_MONITOR_KIOSCO, "|\t|", , CompareMethod.Text)

            'sNroCuenta = ADATA_MONITOR(0)
            'sNroTarjeta = ADATA_MONITOR(1)
            'sModoEntrada = ADATA_MONITOR(2)
            sIDSucursal = ADATA_MONITOR(3)
            sIDTerminal = ADATA_MONITOR(4)
            'sCodigoTransaccion = ADATA_MONITOR(5)

            '    sSaldoDisponible = sDisponibleCompras.Trim
            '    sSaldoDisponible = Replace(sSaldoDisponible.Trim, ",", "") 'QUITAR LA COMA
            '    sSaldoDisponible = Replace(sSaldoDisponible.Trim, ".", "") 'QUITAR EL PUNTO DECIMAL
            '    sSaldoDisponible = Mid(sSaldoDisponible, 1, sSaldoDisponible.Trim.Length - 2) 'QUITAR LOS DOS ULTIMOS DIGITOS QUE SON DECIMALES


            '    sTotalDeuda = sPagoTotal.Trim
            '    sTotalDeuda = Replace(sTotalDeuda.Trim, ",", "") 'QUITAR LA COMA
            '    sTotalDeuda = Replace(sTotalDeuda.Trim, ".", "") 'QUITAR EL PUNTO DECIMAL
            '    sTotalDeuda = Mid(sTotalDeuda, 1, sTotalDeuda.Trim.Length - 2) 'QUITAR LOS DOS ULTIMOS DIGITOS QUE SON DECIMALES

            '    sSaldoDisponible = Right("      " & sSaldoDisponible, 6)
            '    sTotalDeuda = Right("      " & sTotalDeuda, 6)

            '    'GUADAR LOG CONSULTAS
            SAVE_LOG_CONSULTAS(sIDSucursal.Trim, sIDTerminal.Trim, "57", FUN_BUSCAR_TIPO_TARJETA(sNroTarjeta.Trim)) 'TARJETA ASOCIADA



            '    'armar variable desde flash con: numero de cuenta, numero de tarjeta, nombre del cliente, modo de entrada, canal de atencion, id de sucursal,codigo kiosco y codigo de transaccion (opcion)
            '    If Left(sGetTramaMC.Trim, 5) <> "ERROR" Then
            '        'VALIDACION CORRECTA

            '        sNroCuenta = Right("          " & sNroCuenta.Trim, 10)
            '        sNroTarjeta = Right("                    " & sNroTarjeta.Trim, 20)

            '        sNombres = Right(Strings.StrDup(26, " ") & sNombres.Trim, 26)

            '        sEstadoCuenta = "01" 'Estado de la cuenta
            '        sRespuestaServidor = "01" 'Atendido

            '        'ENVIAR_MONITOR
            '        sDataMonitor = ""
            '        sDataMonitor = FECHA_TX & HORA_TX & sNroCuenta & sNroTarjeta & sNombres & sModoEntrada & sCanalAtencion & sSaldoDisponible & sTotalDeuda & sIDSucursal & sIDTerminal & sEstadoCuenta & sRespuestaServidor & sCodigoTransaccion

            '        'ENVIAR_MONITOR(sDataMonitor)
            '    Else
            '        'ENVIAR_MONITOR
            '        sNroCuenta = Right("          " & sNroCuenta.Trim, 10)
            '        sNroTarjeta = Right("                    " & sNroTarjeta.Trim, 20)

            '        sNombres = Strings.StrDup(26, " ")
            '        sEstadoCuenta = "02" 'Estado de la cuenta
            '        sRespuestaServidor = "02" 'Rechazado

            '        sDataMonitor = ""
            '        sDataMonitor = FECHA_TX & HORA_TX & sNroCuenta & sNroTarjeta & sNombres & sModoEntrada & sCanalAtencion & sSaldoDisponible & sTotalDeuda & sIDSucursal & sIDTerminal & sEstadoCuenta & sRespuestaServidor & sCodigoTransaccion

            '        'ENVIAR_MONITOR(sDataMonitor)

            '    End If

        End If

        Return oRespuesta

        'Return "1,800.00|\t|334.62|\t|1,465.38|\t|0.00|\t|334.64|\t|307.17|\t|11/06/2013 - 10/07/2013|\t|25/07/2013|\t|1400|\t|7,500.00|\t|334.64|\t|22/07/2013|\t|17:28:52|\t|001500*$¿*5254740005648129|\t|09924784|\t|RUBEN PARRA|\t||\t|525474|\t|7,500.00|\t|15/08/2013|\t|5254740005648129|\t||\t|1|\t|1|\t|09924784|\t|Super Efectivo SILVER MC|\t|RUBEN PARRA|\t|5254740005648129|\t|7500|\t|36|\t|62.13|\t|396.1|\t|15/08/2013|\t|2243185|\t|3.99|\t|5254740005648129|\t|MC"

    End Function



    Private Function SALDO_TARJETA_ASOCIADA_OLD(ByVal sNroTarjeta As String, ByVal sDATA_MONITOR_KIOSCO As String) As String


        Dim sGetTramaMC As String = ""
        Dim sParametros As String = ""
        Dim sMensajeErrorUsuario As String = ""
        Dim sXML As String = ""
        Dim sTarjetaRecuperada As String = ""
        Dim sNombres As String = "" 'Nombres del cliente

        'VARIABLES DEL SALDO DE TARJETA ASOCIADA
        Dim sLineaCredito As String = "" 'LINEA DE CRÉDITO
        Dim sLineaCreditoUtilizada As String = "" 'LÍNEA DE CRÈDITO UTILIZADA
        Dim sDisponibleCompras As String = "" 'DISPONIBLE COMPRAS
        Dim sDisponibleEfectivo As String = "" 'DISPONIBLE EFECTIVO EXPRESS
        Dim sPagoTotalMes As String = "" 'PAGO TOTAL DEL MES
        Dim sPagoMinimo As String = "" 'PAGO MINIMO DEL MES
        Dim sPeriodo_Facturacion As String = "" 'PERIODO DE FACTURACION
        Dim sFechaPago As String = "" 'FECHA DE PAGO
        Dim sRipleyPuntos As String = "" 'RIPLEY PUNTOS
        Dim sDisponibleSuperEfectivo As String = ""
        Dim sPagoTotal As String = "" 'PAGO TOTAL

        'Return "2,220.00|\t|1,357.56|\t|862.44|\t|862.00|\t|481.63|\t|0.00|\t|11/09/2013 - 10/10/2013|\t|25/10/2013|\t|2728|\t|0.00|\t|1,357.63|\t|07/11/2013|\t|17:27:17|\t|005138*$¿*5420200004815764|\t|41690197|\t|ALTAMIRANO LEIVA, JUAN ALBERTO|\t||\t|542020|\t||\t||\t|5420200004815764|\t||\t|1|\t|1|\t|41690197|\t|Super Efectivo GOLD MC|\t|JUAN ALBERTO ALTAMIRANO LEIVA|\t|5420200004812308|\t|4000|\t|36|\t|66.59|\t|218.15|\t|18/11/2013|\t|9737965|\t|3.99|\t||\t|5420200004812308|\t|MC"

        Try


            If sNroTarjeta.Trim.Length = 16 Then

                Dim obSendWAS As New WSCONSULTAS_MC_VS.defaultService

                sParametros = "HQ000001073RMDQCS" & sNroTarjeta.Trim & Strings.StrDup(16, " ") & "0000000000" & "0" & "000000" & "000073" & "@"

                sGetTramaMC = obSendWAS.execute(sParametros.Trim)

                'Evaluar Respuesta si es ERROR
                If sGetTramaMC.Trim.Length > 0 Then
                    If Left(sGetTramaMC.Trim, 2) = "HE" Then
                        'sGetTramaMC = "ERROR:" & sGetTramaMC.Trim
                        sGetTramaMC = "ERROR:" & Trim(Mid(sGetTramaMC, 16, 30)) 'CORTAR MENSAJE DE ERROR

                    Else

                        Dim sFechaKiosco As String = DateTime.Now.Day.ToString("00").Trim + "/" + DateTime.Now.Month.ToString("00").Trim + "/" + DateTime.Now.Year.ToString("0000").Trim
                        Dim sHoraKiosco As String = DateTime.Now.ToString("HH:mm:ss")
                        Dim sNOperacion As String = GET_NUMERO_OPERACION()

                        'VARIABLES PARA CALCULOS
                        Dim dPagoTotalMes As Double = 0
                        Dim dITF As Double = 0
                        Dim dPagoMinimoMes As Double = 0
                        Dim sDia As String = ""
                        Dim sMes As String = ""
                        Dim sAnio As String = ""
                        Dim dPagoTotal As Double = 0


                        sTarjetaRecuperada = sNroTarjeta.Trim 'formatear 000-0000-0000-0000
                        sNombres = Trim(Mid(sGetTramaMC, 73, 30))

                        If sNombres.Trim = "" Then
                            sGetTramaMC = "ERROR:NODATA" 'SI NO DEVUELVE NOMBRES LANSAR MENSAJE TERMINA
                            Return sGetTramaMC
                        End If

                        sNombres = Replace(sNombres, "#", "Ñ")

                        If sNombres Is Nothing Then
                            sNombres = ""
                        End If

                        sLineaCredito = Format(Val(SET_MONTO_DECIMAL(CLng(Mid(sGetTramaMC, 133, 14)).ToString)), "###,##0.00")

                        sLineaCreditoUtilizada = Format(Val(SET_MONTO_DECIMAL(CLng(Mid(sGetTramaMC, 193, 14)).ToString)), "###,##0.00")

                        sDisponibleCompras = Format(Val(SET_MONTO_DECIMAL(CLng(Mid(sGetTramaMC, 147, 14)).ToString)), "###,##0.00")

                        sDisponibleEfectivo = Format(Val(SET_MONTO_DECIMAL(CLng(Mid(sGetTramaMC, 163, 14)).ToString)), "###,##0.00")

                        'Calculo pago total del mes
                        dPagoTotalMes = Val(SET_MONTO_DECIMAL(CLng(Mid(sGetTramaMC, 207, 14)).ToString))
                        'Obtener el ITF
                        dITF = OBTENER_FACTOR_ITF()
                        dPagoTotalMes = dPagoTotalMes * dITF
                        sPagoTotalMes = Format(dPagoTotalMes, "###,##0.00")

                        'Calculo pago minimo del mes
                        dPagoMinimoMes = Val(SET_MONTO_DECIMAL(CLng(Mid(sGetTramaMC, 179, 14)).ToString))
                        'Obtener el ITF
                        dPagoMinimoMes = dPagoMinimoMes * dITF
                        sPagoMinimo = Format(dPagoMinimoMes, "###,##0.00")


                        'Fecha de Pago
                        sDia = Trim(Mid(sGetTramaMC, 227, 2))
                        sMes = Trim(Mid(sGetTramaMC, 225, 2))
                        sAnio = Trim(Mid(sGetTramaMC, 221, 4))

                        sFechaPago = sDia.Trim & "/" & sMes.Trim & "/" & sAnio

                        If Not IsDate(sFechaPago.Trim) Then
                            sFechaPago = "NO DISPONIBLE"
                        End If

                        'PERIODO DE FACTURACION
                        Dim sFecha1 As String = ""
                        Dim sFecha2 As String = ""
                        Dim lMes1 As Integer = 0
                        Dim lAnio As Integer = 0
                        Dim lAnio1 As Integer = 0
                        Dim lAnio2 As Integer = 0


                        If sFechaPago.Trim.ToUpper <> "NO DISPONIBLE" Then

                            If CInt(sMes.Trim) < 2 Then
                                lMes1 = 12 + CInt(sMes.Trim)
                            Else
                                lMes1 = CInt(sMes.Trim)
                            End If

                            lAnio = CInt(sAnio.Trim)

                            'Año1
                            If CInt(sMes.Trim) < 3 Then
                                lAnio1 = lAnio - 1
                            Else
                                lAnio1 = lAnio
                            End If

                            'Año2
                            If CInt(sMes.Trim) < 3 Then
                                Select Case CInt(sMes.Trim)
                                    Case 1
                                        lAnio2 = lAnio - 1
                                    Case 2
                                        lAnio2 = lAnio
                                End Select

                            Else
                                lAnio2 = lAnio
                            End If


                            If CInt(sDia.Trim) = 5 Then
                                If lMes1 = 2 Then
                                    sFecha1 = "21" & "/" & Right("12", 2) & "/" & lAnio1.ToString
                                Else
                                    sFecha1 = "21" & "/" & Right("00" & (lMes1 - 2).ToString, 2) & "/" & lAnio1.ToString
                                End If

                                sFecha2 = "20" & "/" & Right("00" & (lMes1 - 1).ToString, 2) & "/" & lAnio2.ToString
                            End If


                            If CInt(sDia.Trim) = 25 Then
                                sFecha1 = "11" & "/" & Right("00" & (lMes1 - 1).ToString, 2) & "/" & lAnio2.ToString
                                sFecha2 = "10" & "/" & Right("00" & sMes.Trim, 2) & "/" & sAnio.Trim
                            End If


                            sPeriodo_Facturacion = sFecha1.Trim & " - " & sFecha2.Trim
                        Else
                            sPeriodo_Facturacion = "NO DISPONIBLE"
                        End If


                        'Pago Total

                        dPagoTotal = Val(SET_MONTO_DECIMAL(CLng(Mid(sGetTramaMC, 193, 14)).ToString))
                        dPagoTotal = dPagoTotal * dITF
                        sPagoTotal = Format(dPagoTotal, "###,##0.00")


                        'RIPLEY PUNTOS ACUMULADOS
                        sRipleyPuntos = PUNTOS_RIPLEY_ACUMULADOS(sTarjetaRecuperada.Trim)

                        'Buscar Tipo Doc y Nro Documento desde oracle
                        Dim sTipoDocumento As String = "0"
                        Dim sNroDocumento As String = "000000000000"
                        Dim Nro_Contrato As String = String.Empty
                        Dim sDataTipoDocum As String = ""
                        Dim ADATA_TDOC As Array
                        sDataTipoDocum = CONSULTAR_DOC_MC_VISA(sTarjetaRecuperada.Trim, Nro_Contrato)

                        If sDataTipoDocum.Length > 0 Then
                            ADATA_TDOC = Split(sDataTipoDocum, "|\t|", , CompareMethod.Text)
                            sTipoDocumento = ADATA_TDOC(0)
                            sNroDocumento = ADATA_TDOC(1)

                            If sTipoDocumento.Trim = "1" Then 'DNI
                                sNroDocumento = Right(sNroDocumento.Trim, 8)
                            Else
                                'CE
                                sNroDocumento = Right(sNroDocumento.Trim, 12)
                            End If

                        End If

                        Dim sDataSuperEfectivo As String = ""
                        Dim sMontoAprobadoSuperEfectivo As String = ""
                        Dim sFechaVenceSEfectivo As String = ""

                        sDataSuperEfectivo = DISPONIBLE_SUPER_EFECTIVO_MC_VISA(sTarjetaRecuperada.Trim, sTipoDocumento.Trim, sNroDocumento.Trim)
                        If sDataSuperEfectivo.Trim.Length = 0 Then
                            sDisponibleSuperEfectivo = "0.00"
                            sFechaVenceSEfectivo = ""
                            sMontoAprobadoSuperEfectivo = ""
                        Else
                            Dim ADATASUPER As Array
                            ADATASUPER = Split(sDataSuperEfectivo, "|\t|", , CompareMethod.Text)
                            sMontoAprobadoSuperEfectivo = ADATASUPER(0)
                            sFechaVenceSEfectivo = ADATASUPER(1)
                            sDisponibleSuperEfectivo = ADATASUPER(0)

                            If sDisponibleSuperEfectivo.Trim.Length = 0 Then
                                sDisponibleSuperEfectivo = "0.00"
                            End If

                        End If

                        sXML = sLineaCredito.Trim & "|\t|" & sLineaCreditoUtilizada & "|\t|" & sDisponibleCompras.Trim & "|\t|"
                        sXML = sXML & sDisponibleEfectivo.Trim & "|\t|" & sPagoTotalMes.Trim & "|\t|" & sPagoMinimo.Trim & "|\t|"
                        sXML = sXML & sPeriodo_Facturacion.Trim & "|\t|" & sFechaPago.Trim & "|\t|"
                        sXML = sXML & sRipleyPuntos.Trim & "|\t|" & sDisponibleSuperEfectivo.Trim & "|\t|" & sPagoTotal.Trim & "|\t|"
                        sXML = sXML & sFechaKiosco.Trim & "|\t|" & sHoraKiosco.Trim & "|\t|" & sNOperacion.Trim

                        Dim sDatosCliente As String = ""
                        Dim sTipoProducto As String = ""
                        Dim sCuenta As String = ""

                        sTipoProducto = Mid(sNroTarjeta, 1, 6)
                        sDatosCliente = sNroTarjeta.Trim & "|\t|" & sNroDocumento.Trim & "|\t|" & sNombres.Trim & "|\t||\t|" & sTipoProducto.Trim & "|\t|" & sMontoAprobadoSuperEfectivo.Trim & "|\t|" & sFechaVenceSEfectivo.Trim & "|\t|" & sNroTarjeta.Trim & "|\t||\t|" & MOSTRAR_SUPEREFECTIVO_SEF("1", sTipoDocumento.Trim, sNroDocumento.Trim) & "|\t|"
                        sXML = sXML & "*$¿*" & sDatosCliente
                        sGetTramaMC = sXML & "|\t|" & Nro_Contrato & "|\t|" & "MC"

                        Dim ADATA_MONITOR As Array
                        Dim IDSucursal, IDTerminal As String
                        IDSucursal = ""
                        IDTerminal = ""

                        If sDATA_MONITOR_KIOSCO.Length > 0 Then
                            ADATA_MONITOR = Split(sDATA_MONITOR_KIOSCO, "|\t|", , CompareMethod.Text)
                            IDSucursal = ADATA_MONITOR(3)
                            IDTerminal = ADATA_MONITOR(4)

                            getRegistrar_Ingreso_Ripleymatico(IDSucursal, IDTerminal, sNroDocumento, sTipoDocumento, sNroTarjeta, sTipoProducto)

                        End If



                    End If

                Else 'Sino devuelve nada
                    sGetTramaMC = "ERROR:NODATA"
                End If

            Else
                'Mostrar Mensaje de Error
                sGetTramaMC = "ERROR:NODATA"
            End If

        Catch ex As Exception
            'save log error

            sGetTramaMC = "ERROR:NODATA"

        End Try


        ''Save Monitor
        'Dim sDataMonitor As String = ""
        'Dim FECHA_TX As String = DateTime.Now.Year.ToString("0000").Trim & DateTime.Now.Month.ToString("00").Trim + DateTime.Now.Day.ToString("00").Trim
        'Dim HORA_TX As String = DateTime.Now.Hour.ToString("00").Trim & DateTime.Now.Minute.ToString("00").Trim & DateTime.Now.Second.ToString("00").Trim
        'Dim sModoEntrada As String = ""
        'Dim sCanalAtencion As String = "01" 'RipleyMatico
        'Dim sSaldoDisponible As String = "      " 'Right("      " & sDisponibleCompras.Trim, 6)
        'Dim sTotalDeuda As String = "      " 'Right("      " & sPagoTotal.Trim, 6)
        Dim sIDSucursal As String = ""
        Dim sIDTerminal As String = ""
        'Dim sEstadoCuenta As String = ""
        'Dim sRespuestaServidor As String = ""
        'Dim sCodigoTransaccion As String = ""
        'Dim sNroCuenta As String = ""

        If sDATA_MONITOR_KIOSCO.Length > 0 Then


            'ORDEN DE LAS VARRIABLES
            'DATA_MONITOR=_root.CUENTA_TARJETA_VALIDAR+"|\u005Ct|"+_root.P_NroTarjeta+"|\u005Ct|"+"01"+"|\u005Ct|"+_global.gCODE_SUCURSAL+"|\u005Ct|"+_global.gCODE_KIOSCO+"|\u005Ct|"+"01";
            'RECORRER ARREGLO
            Dim ADATA_MONITOR As Array
            ADATA_MONITOR = Split(sDATA_MONITOR_KIOSCO, "|\t|", , CompareMethod.Text)

            'sNroCuenta = ADATA_MONITOR(0)
            'sNroTarjeta = ADATA_MONITOR(1)
            'sModoEntrada = ADATA_MONITOR(2)
            sIDSucursal = ADATA_MONITOR(3)
            sIDTerminal = ADATA_MONITOR(4)
            'sCodigoTransaccion = ADATA_MONITOR(5)

            '    sSaldoDisponible = sDisponibleCompras.Trim
            '    sSaldoDisponible = Replace(sSaldoDisponible.Trim, ",", "") 'QUITAR LA COMA
            '    sSaldoDisponible = Replace(sSaldoDisponible.Trim, ".", "") 'QUITAR EL PUNTO DECIMAL
            '    sSaldoDisponible = Mid(sSaldoDisponible, 1, sSaldoDisponible.Trim.Length - 2) 'QUITAR LOS DOS ULTIMOS DIGITOS QUE SON DECIMALES


            '    sTotalDeuda = sPagoTotal.Trim
            '    sTotalDeuda = Replace(sTotalDeuda.Trim, ",", "") 'QUITAR LA COMA
            '    sTotalDeuda = Replace(sTotalDeuda.Trim, ".", "") 'QUITAR EL PUNTO DECIMAL
            '    sTotalDeuda = Mid(sTotalDeuda, 1, sTotalDeuda.Trim.Length - 2) 'QUITAR LOS DOS ULTIMOS DIGITOS QUE SON DECIMALES

            '    sSaldoDisponible = Right("      " & sSaldoDisponible, 6)
            '    sTotalDeuda = Right("      " & sTotalDeuda, 6)

            '    'GUADAR LOG CONSULTAS
            SAVE_LOG_CONSULTAS(sIDSucursal.Trim, sIDTerminal.Trim, "57", FUN_BUSCAR_TIPO_TARJETA(sNroTarjeta.Trim)) 'TARJETA ASOCIADA



            '    'armar variable desde flash con: numero de cuenta, numero de tarjeta, nombre del cliente, modo de entrada, canal de atencion, id de sucursal,codigo kiosco y codigo de transaccion (opcion)
            '    If Left(sGetTramaMC.Trim, 5) <> "ERROR" Then
            '        'VALIDACION CORRECTA

            '        sNroCuenta = Right("          " & sNroCuenta.Trim, 10)
            '        sNroTarjeta = Right("                    " & sNroTarjeta.Trim, 20)

            '        sNombres = Right(Strings.StrDup(26, " ") & sNombres.Trim, 26)

            '        sEstadoCuenta = "01" 'Estado de la cuenta
            '        sRespuestaServidor = "01" 'Atendido

            '        'ENVIAR_MONITOR
            '        sDataMonitor = ""
            '        sDataMonitor = FECHA_TX & HORA_TX & sNroCuenta & sNroTarjeta & sNombres & sModoEntrada & sCanalAtencion & sSaldoDisponible & sTotalDeuda & sIDSucursal & sIDTerminal & sEstadoCuenta & sRespuestaServidor & sCodigoTransaccion

            '        'ENVIAR_MONITOR(sDataMonitor)
            '    Else
            '        'ENVIAR_MONITOR
            '        sNroCuenta = Right("          " & sNroCuenta.Trim, 10)
            '        sNroTarjeta = Right("                    " & sNroTarjeta.Trim, 20)

            '        sNombres = Strings.StrDup(26, " ")
            '        sEstadoCuenta = "02" 'Estado de la cuenta
            '        sRespuestaServidor = "02" 'Rechazado

            '        sDataMonitor = ""
            '        sDataMonitor = FECHA_TX & HORA_TX & sNroCuenta & sNroTarjeta & sNombres & sModoEntrada & sCanalAtencion & sSaldoDisponible & sTotalDeuda & sIDSucursal & sIDTerminal & sEstadoCuenta & sRespuestaServidor & sCodigoTransaccion

            '        'ENVIAR_MONITOR(sDataMonitor)

            '    End If

        End If

        Return sGetTramaMC

        'Return "1,800.00|\t|334.62|\t|1,465.38|\t|0.00|\t|334.64|\t|307.17|\t|11/06/2013 - 10/07/2013|\t|25/07/2013|\t|1400|\t|7,500.00|\t|334.64|\t|22/07/2013|\t|17:28:52|\t|001500*$¿*5254740005648129|\t|09924784|\t|RUBEN PARRA|\t||\t|525474|\t|7,500.00|\t|15/08/2013|\t|5254740005648129|\t||\t|1|\t|1|\t|09924784|\t|Super Efectivo SILVER MC|\t|RUBEN PARRA|\t|5254740005648129|\t|7500|\t|36|\t|62.13|\t|396.1|\t|15/08/2013|\t|2243185|\t|3.99|\t|5254740005648129|\t|MC"

    End Function


    'MOVIMIENTOS DE TARJETA ASOCIADA
    'DEVUELVE EL SALDO DE RIPLEY CLASICA
    <WebMethod(Description:="Consulta de Movimientos de Tarjeta Ripley Asociada.")> _
    Public Function MOVIMIENTOS_ASOCIADA(ByVal sNroTarjeta As String, ByVal sDATA_MONITOR_KIOSCO As String) As String
        Dim sGetTramaMC_MOV As String = ""
        Dim sParametros_MC As String = ""
        Dim sMensajeErrorUsuario As String = ""
        Dim sXML As String = ""

        'Return "09/01/2014|\t|20:31:10|\t|001398|¿**?|09/01/2014|\t|CARGOS POR PROCESAR|\t||\t||\t|36.80|\n|06/01/2014|\t|WONG SAN MIGUEL T009|\t||\t||\t|33.26|\n|02/01/2014|\t|PAGOS EN TIENDAS RIP|\t||\t||\t|103.00|\n|27/12/2013|\t|BANCO CREDITO/PAGO A|\t||\t||\t|2445.59|\n|26/12/2013|\t|OPENENGLISH.COM|\t||\t||\t|202.50|\n|26/12/2013|\t|FRUTIX|\t||\t||\t|23.70|\n|24/12/2013|\t|RIPLEY MIRAFLORES|\t||\t||\t|252.20|\n|24/12/2013|\t|RIPLEY MIRAFLORES|\t||\t||\t|140.00|\n|19/12/2013|\t|RIPLEY SAN MIGUEL|\t||\t||\t|234.32|\n|16/12/2013|\t|REST JHONNY Y JENNIF|\t||\t||\t|28.00|\n|15/12/2013|\t|RIPLEY LIMA NORTE|\t||\t||\t|254.20|\n|14/12/2013|\t|BOTICAS INKAFARMA|\t||\t||\t|69.63|\n|13/12/2013|\t|RIPLEY SAN MIGUEL|\t||\t||\t|136.88|\n|13/12/2013|\t|RIPLEY|\t||\t||\t|427.83|\n|12/12/2013|\t|PARDOS CHICKEN|\t||\t||\t|53.40|\n|09/12/2013|\t|CLINICA STELLA MARIS|\t||\t||\t|40.00|\n|09/12/2013|\t|WONG SAN MIGUEL T009|\t||\t||\t|61.49|\n|07/12/2013|\t|MONTALVO HAIR PERU|\t||\t||\t|45.00|\n|05/12/2013|\t|REST LA CARAVANA|\t||\t||\t|30.50|\n|01/12/2013|\t|BANCO CREDITO/PAGO A|\t||\t||\t|2794.89"


        Try

            If sNroTarjeta.Trim.Length = 16 Then

                Dim obSendWAS_ As New WSCONSULTAS_MC_VS.defaultService

                sParametros_MC = "HQ000001073RMDQCM" & sNroTarjeta.Trim & Strings.StrDup(16, " ") & "0000000000" & "0" & "000000" & "000073" & "@"

                sGetTramaMC_MOV = obSendWAS_.execute(sParametros_MC.Trim)

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


                        lLongitudBuffer = Val(Trim(Mid(sGetTramaMC_MOV, 16, 6))).ToString

                        If lLongitudBuffer > 0 Then
                            'EXTRAER EL ENCABEZADO Y DETALLE DE MOVIMIENTOS
                            sBufferHeaderMovimientos = Mid(sGetTramaMC_MOV, 22, lLongitudBuffer)
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
            SAVE_LOG_CONSULTAS(sIDSucursal.Trim, sIDTerminal.Trim, "34", FUN_BUSCAR_TIPO_TARJETA(sNroTarjetax.Trim))

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


    <WebMethod(Description:="Estado de Cuenta Clasica")> _
    Public Function ESTADO_CUENTA_CLASICA(ByVal sNroTarjeta As String, ByVal sNroCuenta As String, ByVal sDATA_MONITOR_KIOSCO As String, ByVal SERVIDOR As TServidor) As String
        Dim Respuesta As String
        Respuesta = String.Empty

        ErrorLog("ESTADO_CUENTA_CLASICA_RSAT(" & sNroTarjeta & ", " & sNroCuenta & ", " & sDATA_MONITOR_KIOSCO & ", " & SERVIDOR & ")")
        Respuesta = ESTADO_CUENTA_CLASICA_RSAT(sNroTarjeta, sNroCuenta, sDATA_MONITOR_KIOSCO, SERVIDOR)
        Return Respuesta
    End Function

    'Buscar el estado de cuenta..
    <WebMethod(Description:="Estado de Cuenta Clasica")> _
    Private Function ESTADO_CUENTA_CLASICA_SICRON(ByVal sNroTarjeta As String, ByVal sNroCuenta As String, ByVal sDATA_MONITOR_KIOSCO As String) As String
        Dim sRespuesta As String = ""
        Dim sParametros As String = ""
        Dim sMensajeErrorUsuario As String = ""
        Dim sXML As String = ""
        Dim lContador As Long = 0
        Dim sPagina As String = ""

        'Variables nuevas
        Dim actions As Long = 1
        Dim inetputBuff As String = ""
        Dim outpputBuff As String = ""
        Dim errorMsg As String = ""

        Try
            If sNroCuenta.Trim.Length = 8 And sNroTarjeta.Length = 16 Then
                'Instancia al mirapiweb
                Dim obSendMirror As ClsTxMirapi = Nothing
                obSendMirror = New ClsTxMirapi()

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


                vFechaHoy = Date.Now
                fFechaHoyAux1 = vFechaHoy
                fFechaHoyAux1 = DateAdd("m", 1, fFechaHoyAux1)
                sPeriodo1 = Right(Trim("0000" & Year(fFechaHoyAux1).ToString), 4) & Right(Trim("00" & Month(fFechaHoyAux1).ToString), 2)
                sPeriodo2 = Right(Trim("0000" & Year(vFechaHoy).ToString), 4) & Right(Trim("00" & Month(vFechaHoy).ToString), 2)
                vFechaHoy = DateAdd("m", -1, vFechaHoy)
                sPeriodo3 = Right(Trim("0000" & Year(vFechaHoy).ToString), 4) & Right(Trim("00" & Month(vFechaHoy).ToString), 2)
                sPeriodoFinal = sPeriodo1.Trim

                Do
                    lContador = lContador + 1
                    If lContador > 9 Then
                        sPagina = lContador.ToString
                    Else
                        sPagina = "0" & lContador.ToString
                    End If

                    sParametros = "00000000000" & sNroTarjeta.Trim & sNroCuenta.Trim & sPeriodoFinal.Trim & sPagina
                    inetputBuff = "      " + "R192" + sParametros

                    sTrama = ""
                    sTrama = obSendMirror.ExecuteTX(actions.ToString, inetputBuff, outpputBuff, TIME_OUT_SERVER, SERVIDOR_MIRROR_DESTINO, SERVIDOR_MIRROR_NODE, PUERTO, errorMsg)
                    outpputBuff = RTrim(outpputBuff)

                    'sTrama = "0"
                    'outpputBuff = "12345678AU            006LLANCO ALBORNOZ, ANGEL                            MZ.L LT.5 URB.EL PACIFICO                                                                           SAN MARTIN DE PORRES                              96041-000128560-34                   8,500.00      0.00      0.00  8,500.0016/ABR/2012-17/MAY/201201/JUN/2012    203.98      0.00    203.98    324.33     66.14     66.14    803.50      0.00      0.00      0.00      0.00    324.33      0.00     58.66      7.48      0.00    324.33     66.14                0.00                7.48      0.00     66.14                              1D           *                            JUL/2012     66.14AGO/2012     45.33SET/2012     35.83              0.00              0.00              0.00              0.00              0.00              0.00              0.00              0.00              0.00                                  SALDO ANTERIOR                                       324.3313/MAR/201217/MAY/2012  8819COMPRA EN CUOTAS    T    75.9003/04   19.31    1.50   20.8113/MAR/201217/MAY/2012  4673COMPRA EN CUOTAS    T   189.0503/06   30.86    4.97   35.8327/MAR/201217/MAY/2012 11442COMPRA EN CUOTAS    T    34.6502/04    8.49    1.01    9.5001/MAY/201201/MAY/2012      RET.ITF PAGO              0.00                         0.0201/MAY/201201/MAY/2012      PAGO TIENDA  RIPLEY       0.00                      -324.35                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                  U"
                    If sTrama = "0" Then 'EXITO
                        If outpputBuff.Length > 0 Then sTrama = outpputBuff.Substring(8, outpputBuff.Length - 8)

                        'Intentos de call EECC
                        If lEstadoLlamadaEECC = 9 Then
                            lEstadoLlamadaEECC = 0 'Para que no vuelva a ingresar a esta logica

                            If Left(sTrama, 2) = "RU" Then
                                'Segunda LLamada segundo periodo
                                sPeriodoFinal = sPeriodo2.Trim
                                sParametros = "00000000000" & sNroTarjeta.Trim & sNroCuenta.Trim & sPeriodoFinal.Trim & sPagina
                                inetputBuff = "      " + "R192" + sParametros

                                sTrama = ""
                                sTrama = obSendMirror.ExecuteTX(actions.ToString, inetputBuff, outpputBuff, TIME_OUT_SERVER, SERVIDOR_MIRROR_DESTINO, SERVIDOR_MIRROR_NODE, PUERTO, errorMsg)
                                outpputBuff = RTrim(outpputBuff)
                                If sTrama = "0" Then 'EXITO
                                    If outpputBuff.Length > 0 Then sTrama = outpputBuff.Substring(8, outpputBuff.Length - 8)

                                ElseIf sTrama = "-2" Then 'Ocurrio un error Recuperar el error
                                    'sTrama = "ERROR:" & errorMsg.Trim
                                    sTrama = ""
                                    Exit Do
                                Else
                                    'sTrama = "ERROR:Servicio no disponible."
                                    sTrama = ""
                                    Exit Do
                                End If


                            End If

                            If Left(sTrama, 2) = "RU" Then
                                'Tercera LLamada Tercer periodo
                                sPeriodoFinal = sPeriodo3.Trim
                                sParametros = "00000000000" & sNroTarjeta.Trim & sNroCuenta.Trim & sPeriodoFinal.Trim & sPagina
                                inetputBuff = "      " + "R192" + sParametros

                                sTrama = ""
                                sTrama = obSendMirror.ExecuteTX(actions.ToString, inetputBuff, outpputBuff, TIME_OUT_SERVER, SERVIDOR_MIRROR_DESTINO, SERVIDOR_MIRROR_NODE, PUERTO, errorMsg)
                                outpputBuff = RTrim(outpputBuff)
                                If sTrama = "0" Then 'EXITO
                                    If outpputBuff.Length > 0 Then sTrama = outpputBuff.Substring(8, outpputBuff.Length - 8)

                                ElseIf sTrama = "-2" Then 'Ocurrio un error Recuperar el error
                                    'sTrama = "ERROR:" & errorMsg.Trim
                                    sTrama = ""
                                    Exit Do
                                Else
                                    'sTrama = "ERROR:Servicio no disponible."
                                    sTrama = ""
                                    Exit Do
                                End If

                            End If

                        End If 'fin de los intentos en periodos diferentes EECC

                    ElseIf sTrama = "-2" Then 'Ocurrio un error en la consulta de la transaccion
                        'sTrama = "ERROR:" & errorMsg.Trim
                        sTrama = ""
                        Exit Do
                    Else
                        'sTrama = "ERROR:Servicio no disponible."
                        sTrama = ""
                        Exit Do
                    End If


                    If Left(sTrama, 2) <> "RU" Then
                        'CONSTRUIR CADENA DE LA CABECERA Y PIE DE PAGINA
                        If lContador = 1 Then

                            'VARIABLES RESUMEN DE ESTADO DE CUENTA
                            sLineaCredito = Mid(sTrama, 253, 10)
                            sPeriodoFacturacion = Mid(sTrama, 293, 23)
                            sUltimoDiaPago = Mid(sTrama, 316, 11)

                            sCreditoUtilizado = Mid(sTrama, 327, 10)
                            sComisionCargos = Mid(sTrama, 337, 10)
                            sDeudaTotal = Mid(sTrama, 347, 10)
                            sDeudaVencida = Mid(sTrama, 357, 10)
                            sPagoMinimoMes = Mid(sTrama, 367, 10)
                            sPagoTotalMes = Mid(sTrama, 377, 10)

                            sDisponibleCompras = Mid(sTrama, 387, 10)
                            sDispEfectivoExpress = Mid(sTrama, 427, 10)


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
                            sDataMov = Mid(sTrama, 861, sTrama.Length)

                            Incrementa = 1
                            If sDataMov.Length > 0 Then
                                If sPeriodoFinal >= Constantes.PeriodoInclusionTEA Then
                                    For lFila = 1 To 13
                                        If Trim(Mid(sDataMov, Incrementa, tamanioFilaDetalle)) <> "" Then
                                            sDataMovAux = sDataMovAux & Mid(sDataMov, Incrementa, tamanioFilaDetalle) & "|\n|"
                                            ErrorLog("sDataMovAux = " & sDataMovAux)
                                        End If

                                        Incrementa = Incrementa + tamanioFilaDetalle
                                    Next
                                Else
                                    For lFila = 1 To 13
                                        If Trim(Mid(sDataMov, Incrementa, tamanioFilaDetalleAnt)) <> "" Then
                                            sDataMovAux = sDataMovAux & Mid(sDataMov, Incrementa, tamanioFilaDetalleAnt) & "|\n|"
                                            ErrorLog("sDataMovAux = " & sDataMovAux)
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
                            sXMLPIE = sXMLPIE & "|\t|0|\t|0|\t|0" 'DEVOLVER CAMPOS DE PAGO MINIMO QUE ESPERA EL FLASH RIPLEYMATICO
                        Else
                            'EVALUAR LA SEGUNDA CALL SOLO MOVIMIENTOS CONTADOR DE LLAMADAS

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
                                            ErrorLog("sDataMovAux = " & sDataMovAux)
                                        End If

                                        Incrementa = Incrementa + tamanioFilaDetalle
                                    Next
                                Else
                                    For lFila = 1 To 13
                                        If Trim(Mid(sDataMov, Incrementa, tamanioFilaDetalleAnt)) <> "" Then
                                            sDataMovAux = sDataMovAux & Mid(sDataMov, Incrementa, tamanioFilaDetalleAnt) & "|\n|"
                                            ErrorLog("sDataMovAux = " & sDataMovAux)
                                        End If

                                        Incrementa = Incrementa + tamanioFilaDetalleAnt
                                    Next
                                End If
                            End If
                        End If 'FIN DE VALIDACION DEL CONTADOR

                        Dim caracter As String = Right(sTrama, 1)
                        If Right(sTrama, 1) = "C" Then
                            'If Left(sTrama, 2) = "AC" Then 'HAY MOVIMIENTOS PENDIENTES POR LLAMAR
                            sXMLMOV = sXMLMOV & sDataMovAux
                        End If

                        If Right(sTrama, 1) = "U" Then
                            'If Left(sTrama, 2) = "AU" Then 'FIN DE LA SOLICITUD
                            If sDataMovAux.Length > 0 Then
                                sDataMovAux = Left(sDataMovAux, sDataMovAux.Length - 4)
                                sXMLMOV = sXMLMOV & sDataMovAux
                            Else
                                sDataMovAux = "" 'NO HAY MOVIMIENTOS
                                sXMLMOV = sXMLMOV & sDataMovAux
                            End If

                            If sXMLMOV.Length > 0 Then
                                'Armar las columnas de los movimientos de EECC
                                sXMLMOV_FINAL = FUN_DETALLE_EECC_TEA(sXMLMOV, sPeriodoFinal, TServidor.SICRON)
                            End If
                        End If

                        'CADENA FINAL CON LOS DATOS FINALES
                        sRespuesta = sXMLCAB.Trim & "*$¿*" & sXMLPIE.Trim & "*$¿*" & sXMLMOV_FINAL.Trim
                    End If

                    If lContador > 20 Then
                        Exit Do
                    End If
                Loop Until Right(sTrama, 1) <> "C"

                'Evaluar Respuesta si es ERROR,  RU, AU (Respuesta correcta)
                If sTrama.Trim.Length > 0 Then
                    If Left(sTrama.Trim, 2) = "RU" Then 'Mensaje de Validacion al usuario. RU            000CUENTA BLOQUEADA
                        sMensajeErrorUsuario = Mid(sTrama.Trim, 18, Len(sTrama.Trim))
                        sRespuesta = "ERROR:NODATA-" & sMensajeErrorUsuario.Trim
                    End If
                Else 'Sino devuelve nada
                    sRespuesta = "ERROR:Servicio no disponible."
                End If
            Else
                'Mostrar Mensaje de Error
                sRespuesta = "ERROR:Parametros Incompletos"
            End If
        Catch ex As Exception
            'save log error
            sRespuesta = "ERROR:" & ex.Message.Trim
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
            'SAVE_LOG_CONSULTAS(sIDSucursal.Trim, sIDTerminal.Trim, "32", FUN_BUSCAR_TIPO_TARJETA(sNroTarjetax.Trim))

            sNroCuentax = Right("          " & sNroCuentax.Trim, 10)
            sNroTarjetax = Right(Strings.StrDup(20, " ") & sNroTarjetax.Trim, 20)
            sNombreCliente = Right(Strings.StrDup(26, " ") & sNombreCliente.Trim, 26)

            sSaldoDisponible = Right("      " & sSaldoDisponible.Trim, 6)
            sTotalDeuda = Right("      " & sTotalDeuda.Trim, 6)


            'armar variable desde flash con: numero de cuenta, numero de tarjeta, nombre del cliente, modo de entrada, canal de atencion, id de sucursal,codigo kiosco y codigo de transaccion (opcion)
            If Left(sRespuesta.Trim, 5) <> "ERROR" Then
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

        Return sRespuesta

    End Function

    'Buscar el estado de cuenta..
    <WebMethod(Description:="Estado de Cuenta Asociados")> _
    Public Function ESTADO_CUENTA_ASOCIADOS(ByVal sNroTarjeta As String, ByVal sNroCuenta As String, ByVal sDATA_MONITOR_KIOSCO As String) As String
        Dim sRespuesta_AS As String = ""
        Dim sParametros As String = ""
        Dim sMensajeErrorUsuario As String = ""
        Dim sXML As String = ""
        Dim lContador As Long = 0
        Dim sPagina As String = ""

        'Variables nuevas
        Dim actions As Long = 1
        Dim inetputBuff As String = ""
        Dim outpputBuff As String = ""
        Dim errorMsg As String = ""

        Try
            If sNroCuenta.Trim.Length = 8 And sNroTarjeta.Length = 16 Then
                'Instancia al mirapiweb
                Dim obSendMirror As ClsTxMirapi = Nothing
                obSendMirror = New ClsTxMirapi()

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

                'VARIABLES RESUMEN DE ESTADO DE CUENTA
                Dim sLineaCredito As String = ""
                Dim sDispEfectivoExpress As String = ""
                Dim sDisponibleCompras As String = ""
                Dim sPeriodoFacturacion As String = ""
                Dim sUltimoDiaPago As String = ""
                Dim sCreditoUtilizado As String = ""
                Dim sComisionCargos As String = ""
                Dim sDeudaTotal As String = ""
                Dim sMoraSobreGiro As String = ""
                Dim sPagoMinimoMes As String = ""
                Dim sPagoTotalMes As String = ""

                'CALCULO DE PAGO TOTAL DEL MES
                Dim sSaldoAnterior As String = ""
                Dim sConsumoMes As String = ""
                Dim sCuotaMes As String = ""
                Dim sInteres As String = ""
                Dim sComisionCargos_Mes As String = ""
                Dim sPagosAbonos As String = ""
                Dim sPagoTotal_Mes As String = ""

                'CALCULO DE PAGO MINIMO DEL MES
                Dim sSobreGiro As String = ""
                Dim sMontoMinimo As String = ""
                Dim sCuotaMes_Min As String = ""
                Dim sInteres_Min As String = ""
                Dim sComisionCargos_Min As String = ""
                Dim sPagoMinimoMes_Min As String = ""

                'CALCULO DEL PAGO MINIMO OPCIONAL
                Dim sSobreGiro_op As String = ""
                Dim sDeudaVencida_op As String = ""
                Dim sDeudaRevolvente_op As String = ""
                Dim s50CapCuotas_op As String = ""
                Dim sInteres_op As String = ""
                Dim sComisionesCargos_op As String = ""
                Dim sPagoMinimo_op As String = ""

                'PLAN DE CUOTAS DE LOS PROXIMOS 3 MESES
                Dim sMes1 As String = ""
                Dim sMonto1 As String = ""
                Dim sMes2 As String = ""
                Dim sMonto2 As String = ""
                Dim sMes3 As String = ""
                Dim sMonto3 As String = ""

                'DATOS DE RENOVACION POR MEMBRESIA
                Dim sFechaCargo As String = ""

                Dim sPeriodoFinal As String = ""
                Dim sPeriodo1 As String = ""
                Dim sPeriodo2 As String = ""
                Dim sPeriodo3 As String = ""
                Dim vFechaHoy As Date
                Dim fFechaHoyAux1 As Date
                Dim lEstadoLlamadaEECC As Long = 9 'Variable para iniciar llamadas al EECC 3 veces

                vFechaHoy = Date.Now
                fFechaHoyAux1 = vFechaHoy 'Fecha actual
                fFechaHoyAux1 = DateAdd("m", 1, fFechaHoyAux1) 'Aumentar un mes a la fecha actual
                sPeriodo1 = Right(Trim("0000" & Year(fFechaHoyAux1).ToString), 4) & Right(Trim("00" & Month(fFechaHoyAux1).ToString), 2) 'Año mes primera llamada
                sPeriodo2 = Right(Trim("0000" & Year(vFechaHoy).ToString), 4) & Right(Trim("00" & Month(vFechaHoy).ToString), 2) 'Año mes segunda llamada
                vFechaHoy = DateAdd("m", -1, vFechaHoy) 'Restar un mes a la fecha actual...
                sPeriodo3 = Right(Trim("0000" & Year(vFechaHoy).ToString), 4) & Right(Trim("00" & Month(vFechaHoy).ToString), 2) 'Año mes tercera llamada
                sPeriodoFinal = sPeriodo1.Trim 'Set Primera llamada
                'Agregar 201503 para que funcione EECC
                'sPeriodoFinal = "201503"

                Do
                    lContador = lContador + 1
                    If lContador > 9 Then
                        sPagina = lContador.ToString
                    Else
                        sPagina = "0" & lContador.ToString
                    End If

                    sParametros = "00000000000" & sNroTarjeta.Trim & sNroCuenta.Trim & sPeriodoFinal.Trim & sPagina
                    inetputBuff = "      " + "R192" + sParametros
                    sTrama = ""
                    sTrama = obSendMirror.ExecuteTX(actions.ToString, inetputBuff, outpputBuff, TIME_OUT_SERVER, SERVIDOR_MIRROR_DESTINO, SERVIDOR_MIRROR_NODE, PUERTO, errorMsg)
                    outpputBuff = RTrim(outpputBuff)

                    If sTrama = "0" Then 'EXITO
                        If outpputBuff.Length > 0 Then sTrama = outpputBuff.Substring(8, outpputBuff.Length - 8)

                        'Intentos de call EECC
                        If lEstadoLlamadaEECC = 9 Then
                            lEstadoLlamadaEECC = 0 'Para que no vuelva a ingresar a esta logica

                            If Left(sTrama, 2) = "RU" Then
                                'Segunda LLamada segundo periodo
                                sPeriodoFinal = sPeriodo2.Trim
                                sParametros = "00000000000" & sNroTarjeta.Trim & sNroCuenta.Trim & sPeriodoFinal.Trim & sPagina
                                inetputBuff = "      " + "R192" + sParametros
                                sTrama = ""
                                sTrama = obSendMirror.ExecuteTX(actions.ToString, inetputBuff, outpputBuff, TIME_OUT_SERVER, SERVIDOR_MIRROR_DESTINO, SERVIDOR_MIRROR_NODE, PUERTO, errorMsg)
                                outpputBuff = RTrim(outpputBuff)
                                If sTrama = "0" Then 'EXITO
                                    If outpputBuff.Length > 0 Then sTrama = outpputBuff.Substring(8, outpputBuff.Length - 8)

                                ElseIf sTrama = "-2" Then 'Ocurrio un error Recuperar el error
                                    'sTrama = "ERROR:" & errorMsg.Trim
                                    sTrama = ""
                                    Exit Do
                                Else
                                    'sTrama = "ERROR:Servicio no disponible."
                                    sTrama = ""
                                    Exit Do
                                End If
                            End If

                            If Left(sTrama, 2) = "RU" Then
                                'Tercera LLamada Tercer periodo
                                sPeriodoFinal = sPeriodo3.Trim
                                sParametros = "00000000000" & sNroTarjeta.Trim & sNroCuenta.Trim & sPeriodoFinal.Trim & sPagina
                                inetputBuff = "      " + "R192" + sParametros
                                sTrama = ""
                                sTrama = obSendMirror.ExecuteTX(actions.ToString, inetputBuff, outpputBuff, TIME_OUT_SERVER, SERVIDOR_MIRROR_DESTINO, SERVIDOR_MIRROR_NODE, PUERTO, errorMsg)
                                outpputBuff = RTrim(outpputBuff)
                                If sTrama = "0" Then 'EXITO
                                    If outpputBuff.Length > 0 Then sTrama = outpputBuff.Substring(8, outpputBuff.Length - 8)
                                ElseIf sTrama = "-2" Then 'Ocurrio un error Recuperar el error
                                    'sTrama = "ERROR:" & errorMsg.Trim
                                    sTrama = ""
                                    Exit Do
                                Else
                                    'sTrama = "ERROR:Servicio no disponible."
                                    sTrama = ""
                                    Exit Do
                                End If

                            End If
                        End If 'FIN DE INTENTOS DE LLAMADAS A EECC

                    ElseIf sTrama = "-2" Then 'Ocurrio un error en la consulta de la transaccion
                        'sTrama = "ERROR:" & errorMsg.Trim
                        sTrama = ""
                        Exit Do
                    Else
                        'sTrama = "ERROR:Servicio no disponible."
                        sTrama = ""
                        Exit Do
                    End If


                    If Left(sTrama, 2) <> "RU" Then
                        'CONSTRUIR CADENA DE LA CABECERA Y PIE DE PAGINA
                        If lContador = 1 Then
                            'VARIABLES RESUMEN DE ESTADO DE CUENTA
                            sLineaCredito = Mid(sTrama, 253, 10)
                            sPeriodoFacturacion = Mid(sTrama, 293, 23)
                            sUltimoDiaPago = Mid(sTrama, 316, 11)
                            sCreditoUtilizado = Mid(sTrama, 327, 10)
                            sComisionCargos = Mid(sTrama, 337, 10)
                            sDeudaTotal = Mid(sTrama, 347, 10)
                            sMoraSobreGiro = ""
                            sPagoMinimoMes = Mid(sTrama, 367, 10)
                            sPagoTotalMes = Mid(sTrama, 377, 10)
                            sDisponibleCompras = Mid(sTrama, 387, 10)
                            sDispEfectivoExpress = Mid(sTrama, 427, 10)

                            'CALCULO DE PAGO TOTAL DEL MES
                            sSaldoAnterior = Mid(sTrama, 437, 10)
                            sConsumoMes = Mid(sTrama, 447, 10)
                            sCuotaMes = Mid(sTrama, 457, 10)
                            sInteres = Mid(sTrama, 467, 10)
                            sComisionCargos_Mes = Mid(sTrama, 477, 10)
                            sPagosAbonos = Mid(sTrama, 487, 10)
                            sPagoTotal_Mes = Mid(sTrama, 497, 10)

                            'CALCULO DE PAGO MINIMO DEL MES
                            sSobreGiro = ""
                            sMontoMinimo = Mid(sTrama, 517, 10)
                            sCuotaMes_Min = Mid(sTrama, 527, 10)
                            sInteres_Min = Mid(sTrama, 537, 10)
                            sComisionCargos_Min = Mid(sTrama, 547, 10)
                            sPagoMinimoMes_Min = Mid(sTrama, 557, 10)

                            'CALCULO DEL PAGO MINIMO OPCIONAL
                            sSobreGiro_op = ""
                            sDeudaVencida_op = ""
                            sDeudaRevolvente_op = ""
                            s50CapCuotas_op = ""
                            sInteres_op = ""
                            sComisionesCargos_op = ""
                            sPagoMinimo_op = ""

                            'PLAN DE CUOTAS DE LOS PROXIMOS 3 MESES
                            sMes1 = Mid(sTrama, 615, 8)
                            sMonto1 = Mid(sTrama, 623, 10)
                            sMes2 = Mid(sTrama, 633, 8)
                            sMonto2 = Mid(sTrama, 641, 10)
                            sMes3 = Mid(sTrama, 651, 8)
                            sMonto3 = Mid(sTrama, 659, 10)

                            'DATOS DE RENOVACION POR MEMBRESIA
                            sFechaCargo = ""

                            'MOVIMIENTOS
                            sDataMovAux = ""
                            sDataMov = Mid(sTrama, 861, sTrama.Length)

                            Incrementa = 1
                            If sDataMov.Length > 0 Then
                                If sPeriodoFinal >= Constantes.PeriodoInclusionTEA Then
                                    For lFila = 1 To 13
                                        If Trim(Mid(sDataMov, Incrementa, tamanioFilaDetalle)) <> "" Then
                                            sDataMovAux = sDataMovAux & Mid(sDataMov, Incrementa, tamanioFilaDetalle) & "|\n|"
                                            ErrorLog("sDataMovAux = " & sDataMovAux)
                                        End If

                                        Incrementa = Incrementa + tamanioFilaDetalle
                                    Next
                                Else
                                    For lFila = 1 To 13
                                        If Trim(Mid(sDataMov, Incrementa, tamanioFilaDetalleAnt)) <> "" Then
                                            sDataMovAux = sDataMovAux & Mid(sDataMov, Incrementa, tamanioFilaDetalleAnt) & "|\n|"
                                            ErrorLog("sDataMovAux = " & sDataMovAux)
                                        End If

                                        Incrementa = Incrementa + tamanioFilaDetalleAnt
                                    Next
                                End If
                            End If

                            sXMLCAB = ""
                            sXMLCAB = sLineaCredito.Trim & "|\t|" & sDispEfectivoExpress.Trim & "|\t|" & sDisponibleCompras.Trim
                            sXMLCAB = sXMLCAB & "|\t|" & sPeriodoFacturacion.Trim & "|\t|" & sUltimoDiaPago.Trim
                            sXMLCAB = sXMLCAB & "|\t|" & sCreditoUtilizado.Trim & "|\t|" & sComisionCargos.Trim & "|\t|" & sDeudaTotal.Trim
                            sXMLCAB = sXMLCAB & "|\t|" & sMoraSobreGiro.Trim & "|\t|" & sPagoMinimoMes.Trim & "|\t|" & sPagoTotalMes.Trim

                            sXMLPIE = ""
                            sXMLPIE = sSaldoAnterior.Trim & "|\t|" & sConsumoMes.Trim & "|\t|" & sCuotaMes.Trim
                            sXMLPIE = sXMLPIE & "|\t|" & sInteres.Trim & "|\t|" & sComisionCargos_Mes.Trim & "|\t|" & sPagosAbonos.Trim
                            sXMLPIE = sXMLPIE & "|\t|" & sPagoTotal_Mes.Trim & "|\t|" & sSobreGiro.Trim & "|\t|" & sMontoMinimo.Trim
                            sXMLPIE = sXMLPIE & "|\t|" & sCuotaMes_Min.Trim & "|\t|" & sInteres_Min.Trim & "|\t|" & sComisionCargos_Min.Trim
                            sXMLPIE = sXMLPIE & "|\t|" & sPagoMinimoMes_Min.Trim
                            sXMLPIE = sXMLPIE & "|\t|" & sSobreGiro_op.Trim & "|\t|" & sDeudaVencida_op.Trim & "|\t|" & sDeudaRevolvente_op.Trim
                            sXMLPIE = sXMLPIE & "|\t|" & s50CapCuotas_op.Trim & "|\t|" & sInteres_op.Trim & "|\t|" & sComisionesCargos_op.Trim
                            sXMLPIE = sXMLPIE & "|\t|" & sPagoMinimo_op.Trim & "|\t|" & sMes1.Trim & "|\t|" & sMonto1.Trim
                            sXMLPIE = sXMLPIE & "|\t|" & sMes2.Trim & "|\t|" & sMonto2.Trim & "|\t|" & sMes3.Trim & "|\t|" & sMonto3.Trim & "|\t|" & sFechaCargo.Trim
                            sXMLPIE = sXMLPIE & "|\t|0|\t|0|\t|0" 'DEVOLVER CAMPOS DE PAGO MINIMO QUE ESPERA EL FLASH RIPLEYMATICO
                        Else
                            'EVALUAR LA SEGUNDA CALL SOLO MOVIMIENTOS CONTADOR DE LLAMADAS

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
                                            ErrorLog("sDataMovAux = " & sDataMovAux)
                                        End If

                                        Incrementa = Incrementa + tamanioFilaDetalle
                                    Next
                                Else
                                    For lFila = 1 To 13
                                        If Trim(Mid(sDataMov, Incrementa, tamanioFilaDetalleAnt)) <> "" Then
                                            sDataMovAux = sDataMovAux & Mid(sDataMov, Incrementa, tamanioFilaDetalleAnt) & "|\n|"
                                            ErrorLog("sDataMovAux = " & sDataMovAux)
                                        End If

                                        Incrementa = Incrementa + tamanioFilaDetalleAnt
                                    Next
                                End If
                            End If
                        End If 'FIN DE VALIDACION DEL CONTADOR

                        'If Left(sTrama, 2) = "AC" Then 'HAY MOVIMIENTOS PENDIENTES POR LLAMAR
                        If Right(sTrama, 1) = "C" Then
                            sXMLMOV = sXMLMOV & sDataMovAux
                        End If

                        'If Left(sTrama, 2) = "AU" Then 'FIN DE LA SOLICITUD
                        If Right(sTrama, 1) = "U" Then
                            If sDataMovAux.Length > 0 Then
                                sDataMovAux = Left(sDataMovAux, sDataMovAux.Length - 4)
                                sXMLMOV = sXMLMOV & sDataMovAux
                            Else
                                sDataMovAux = "" 'NO HAY MOVIMIENTOS
                                sXMLMOV = sXMLMOV & sDataMovAux
                            End If

                            If sXMLMOV.Length > 0 Then
                                'Armar las columnas de los movimientos de EECC
                                sXMLMOV_FINAL = FUN_DETALLE_EECC_TEA(sXMLMOV, sPeriodoFinal, TServidor.SICRON)
                            End If
                        End If

                        'CADENA FINAL CON LOS DATOS FINALES
                        sRespuesta_AS = sXMLCAB.Trim & "*$¿*" & sXMLPIE.Trim & "*$¿*" & sXMLMOV_FINAL.Trim
                    End If

                    'Loop Until Left(sTrama, 2) <> "AC"
                Loop Until Right(sTrama, 1) <> "C"

                'Evaluar Respuesta si es ERROR,  RU, AU (Respuesta correcta)
                If sTrama.Trim.Length > 0 Then
                    If Left(sTrama.Trim, 2) = "RU" Then 'Mensaje de Validacion al usuario. RU            000CUENTA BLOQUEADA
                        sMensajeErrorUsuario = Mid(sRespuesta_AS.Trim, 3, Len(sRespuesta_AS.Trim))
                        sRespuesta_AS = "ERROR:NODATA-" & sMensajeErrorUsuario.Trim
                    End If
                Else 'Sino devuelve nada
                    sRespuesta_AS = "ERROR:Servicio no disponible."
                End If
            Else
                'Mostrar Mensaje de Error
                sRespuesta_AS = "ERROR:Parametros Incompletos"
            End If
        Catch ex As Exception
            'save log error
            sRespuesta_AS = "ERROR:" & ex.Message.Trim
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
            SAVE_LOG_CONSULTAS(sIDSucursal.Trim, sIDTerminal.Trim, "32", FUN_BUSCAR_TIPO_TARJETA(sNroTarjetax.Trim))

            sNroCuentax = Right("          " & sNroCuentax.Trim, 10)
            sNroTarjetax = Right(Strings.StrDup(20, " ") & sNroTarjetax.Trim, 20)
            sNombreCliente = Right(Strings.StrDup(26, " ") & sNombreCliente.Trim, 26)
            sSaldoDisponible = Right("      " & sSaldoDisponible.Trim, 6)
            sTotalDeuda = Right("      " & sTotalDeuda.Trim, 6)

            'armar variable desde flash con: numero de cuenta, numero de tarjeta, nombre del cliente, modo de entrada, canal de atencion, id de sucursal,codigo kiosco y codigo de transaccion (opcion)
            If Left(sRespuesta_AS.Trim, 5) <> "ERROR" Then
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

        Return sRespuesta_AS
    End Function

    <WebMethod(Description:="Estado de Cuenta Tarjeta Clasica SICRON | RSAT")> _
    Public Function ESTADO_CUENTA_CLASICA_RSAT(ByVal sNroTarjeta As String, ByVal sNroCuenta As String, ByVal sDATA_MONITOR_KIOSCO As String, ByVal Servidor As TServidor) As String
        Dim sRespuesta As String = ""
        Dim sParametros As String = ""
        Dim sMensajeErrorUsuario As String = ""
        Dim sXML As String = ""
        Dim lContador As Long = 0
        Dim sPagina As String = ""

        'Variables nuevas
        Dim actions As Long = 1
        Dim inetputBuff As String = ""
        Dim outpputBuff As String = ""
        Dim errorMsg As String = ""

        Try
            sNroCuenta = Right(sNroCuenta.Trim, 8)

            If sNroCuenta.Trim.Length = 8 And sNroTarjeta.Length = 16 Then
                'Instancia al mirapiweb
                '**Dim obSendMirror As ClsTxMirapi = Nothing
                '**obSendMirror = New ClsTxMirapi()

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

                'Cambios en Trama 04/06/2015
                Dim sCntMeses As String = ""
                Dim sIntPagMin As String = ""
                Dim sComPagMin As String = ""

                Dim sPeriodoFinal As String = ""
                Dim sPeriodo1 As String = ""
                Dim sPeriodo2 As String = ""
                Dim sPeriodo3 As String = ""
                Dim vFechaHoy As Date
                Dim fFechaHoyAux1 As Date
                Dim lEstadoLlamadaEECC As Long = 9 'Variable para iniciar llamadas al EECC 3 veces

                vFechaHoy = Date.Now
                fFechaHoyAux1 = vFechaHoy
                fFechaHoyAux1 = DateAdd("m", 1, fFechaHoyAux1)
                sPeriodo1 = Right(Trim("0000" & Year(fFechaHoyAux1).ToString), 4) & Right(Trim("00" & Month(fFechaHoyAux1).ToString), 2)
                sPeriodo2 = Right(Trim("0000" & Year(vFechaHoy).ToString), 4) & Right(Trim("00" & Month(vFechaHoy).ToString), 2)
                vFechaHoy = DateAdd("m", -1, vFechaHoy)
                sPeriodo3 = Right(Trim("0000" & Year(vFechaHoy).ToString), 4) & Right(Trim("00" & Month(vFechaHoy).ToString), 2)
                sPeriodoFinal = sPeriodo1.Trim
                'Agregar 201503 para que funcione EECC
                'sPeriodoFinal = ConfigurationSettings.AppSettings.Get("PeriodoInclusionTEA")

                Do
                    lContador = lContador + 1
                    If lContador > 9 Then
                        sPagina = lContador.ToString
                    Else
                        sPagina = "0" & lContador.ToString
                    End If

                    Dim CServidor As String = "0"
                    Select Case Servidor
                        Case TServidor.SICRON
                            CServidor = "0"
                        Case TServidor.RSAT
                            CServidor = "S"
                    End Select

                    sTrama = GetTramaR192(CServidor, sNroTarjeta, sNroCuenta, sPeriodoFinal, sPagina, outpputBuff)
                    If sTrama = "0" Then 'EXITO
                        If outpputBuff.Length > 0 Then sTrama = outpputBuff.Substring(8, outpputBuff.Length - 8)

                        'Intentos de call EECC
                        If lEstadoLlamadaEECC = 9 Then
                            lEstadoLlamadaEECC = 0 'Para que no vuelva a ingresar a esta logica

                            If Left(sTrama, 2) = "RU" Then
                                'Segunda LLamada segundo periodo
                                sPeriodoFinal = sPeriodo2.Trim
                                sTrama = GetTramaR192(CServidor, sNroTarjeta, sNroCuenta, sPeriodoFinal, sPagina, outpputBuff)

                                If sTrama = "0" Then 'EXITO
                                    If outpputBuff.Length > 0 Then sTrama = outpputBuff.Substring(8, outpputBuff.Length - 8)

                                ElseIf sTrama = "-2" Then 'Ocurrio un error Recuperar el error
                                    'sTrama = "ERROR:" & errorMsg.Trim
                                    sTrama = ""
                                    Exit Do
                                Else
                                    'sTrama = "ERROR:Servicio no disponible."
                                    sTrama = ""
                                    Exit Do
                                End If
                            End If

                            If Left(sTrama, 2) = "RU" Then
                                'Tercera LLamada Tercer periodo
                                sPeriodoFinal = sPeriodo3.Trim
                                sTrama = GetTramaR192(CServidor, sNroTarjeta, sNroCuenta, sPeriodoFinal, sPagina, outpputBuff)

                                If sTrama = "0" Then 'EXITO
                                    If outpputBuff.Length > 0 Then sTrama = outpputBuff.Substring(8, outpputBuff.Length - 8)
                                ElseIf sTrama = "-2" Then 'Ocurrio un error Recuperar el error
                                    'sTrama = "ERROR:" & errorMsg.Trim
                                    sTrama = ""
                                    Exit Do
                                Else
                                    'sTrama = "ERROR:Servicio no disponible."
                                    sTrama = ""
                                    Exit Do
                                End If
                            End If
                        End If 'fin de los intentos en periodos diferentes EECC
                    ElseIf sTrama = "-2" Then 'Ocurrio un error en la consulta de la transaccion
                        'sTrama = "ERROR:" & errorMsg.Trim
                        sTrama = ""
                        Exit Do
                    Else
                        'sTrama = "ERROR:Servicio no disponible."
                        sTrama = ""
                        Exit Do
                    End If

                    If Left(sTrama, 2) <> "RU" Then
                        'CONSTRUIR CADENA DE LA CABECERA Y PIE DE PAGINA
                        If lContador = 1 Then
                            'VARIABLES RESUMEN DE ESTADO DE CUENTA
                            sLineaCredito = Mid(sTrama, 253, 10)
                            sPeriodoFacturacion = Mid(sTrama, 293, 23)
                            sUltimoDiaPago = Mid(sTrama, 316, 11)
                            sCreditoUtilizado = Mid(sTrama, 327, 10)
                            sComisionCargos = Mid(sTrama, 337, 10)
                            sDeudaTotal = Mid(sTrama, 347, 10)
                            sDeudaVencida = Mid(sTrama, 357, 10)
                            sPagoMinimoMes = Mid(sTrama, 367, 10)
                            sPagoTotalMes = Mid(sTrama, 377, 10)
                            sDisponibleCompras = Mid(sTrama, 387, 10)
                            sDispEfectivoExpress = Mid(sTrama, 427, 10)

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

                            'Cambios en Trama 04/06/2015
                            Dim iCntMeses As Integer = 0
                            Dim dIntPagMin As Double = 0
                            Dim dComPagMin As Double = 0
                            If Servidor = TServidor.RSAT And sPeriodoFinal >= ReadAppConfig("PeriodoAgrandarGlosa") Then
                                Integer.TryParse(Mid(sTrama, 861, 5), iCntMeses)
                                Double.TryParse(Mid(sTrama, 866, 8), dIntPagMin)
                                Double.TryParse(Mid(sTrama, 874, 8), dComPagMin)
                            End If
                            sCntMeses = iCntMeses.ToString()
                            sIntPagMin = dIntPagMin.ToString()
                            sComPagMin = dComPagMin.ToString()

                            'MOVIMIENTOS
                            sDataMovAux = ""
                            If Servidor = TServidor.RSAT And sPeriodoFinal >= ReadAppConfig("PeriodoAgrandarGlosa") Then
                                sDataMov = Mid(sTrama, 882, sTrama.Length)
                            Else
                                sDataMov = Mid(sTrama, 861, sTrama.Length)
                            End If
                            ErrorLog("Movimiento sDataMov = " & sDataMov)

                            Incrementa = 1
                            If sDataMov.Length > 0 Then
                                If Servidor = TServidor.SICRON Or sPeriodoFinal < Constantes.PeriodoInclusionTEA Then
                                    For lFila = 1 To 13
                                        If Trim(Mid(sDataMov, Incrementa, tamanioFilaDetalleAnt)) <> "" Then
                                            sDataMovAux = sDataMovAux & Mid(sDataMov, Incrementa, tamanioFilaDetalleAnt) & "|\n|"
                                            ErrorLog("sDataMovAux = " & sDataMovAux)
                                        End If
                                        Incrementa = Incrementa + tamanioFilaDetalleAnt
                                    Next
                                ElseIf sPeriodoFinal >= ReadAppConfig("PeriodoAgrandarGlosa") Then
                                    '10-03-2015 Ocultar para pase por partes
                                    tamanioFilaDetalle = 159
                                    For lFila = 1 To 7
                                        If Trim(Mid(sDataMov, Incrementa, tamanioFilaDetalle)) <> "" Then
                                            sDataMovAux = sDataMovAux & Mid(sDataMov, Incrementa, tamanioFilaDetalle) & "|\n|"
                                            ErrorLog("sDataMovAux = " & sDataMovAux)
                                        End If
                                        Incrementa = Incrementa + tamanioFilaDetalle
                                    Next

                                    If Right(sTrama, 1) = "C" Then
                                        ErrorLog("Llamar a TramaSiguiente")
                                        sTrama = GetTramaR192(Constantes.TramaTrece, sNroTarjeta, sNroCuenta, sPeriodoFinal, sPagina, outpputBuff)
                                        If sTrama = "0" Then 'EXITO
                                            If outpputBuff.Length > 0 Then sTrama = outpputBuff.Substring(8, outpputBuff.Length - 8)
                                        ElseIf sTrama = "-2" Then
                                            sTrama = ""
                                            Exit Do
                                        Else
                                            sTrama = ""
                                            Exit Do
                                        End If

                                        If Left(sTrama, 2) <> "RU" Then
                                            sDataMov = ""
                                            sDataMov = Mid(sTrama, 882, sTrama.Length)
                                            ErrorLog("Movimiento sDataMov = " & sDataMov)
                                            Incrementa = 1
                                            If sDataMov.Length > 0 Then
                                                For lFila = 1 To 6
                                                    If Trim(Mid(sDataMov, Incrementa, tamanioFilaDetalle)) <> "" Then
                                                        sDataMovAux = sDataMovAux & Mid(sDataMov, Incrementa, tamanioFilaDetalle) & "|\n|"
                                                        ErrorLog("sDataMovAux = " & sDataMovAux)
                                                    End If
                                                    Incrementa = Incrementa + tamanioFilaDetalle
                                                Next
                                            End If
                                        End If
                                    End If
                                Else
                                    For lFila = 1 To 12
                                        If Trim(Mid(sDataMov, Incrementa, tamanioFilaDetalle)) <> "" Then
                                            sDataMovAux = sDataMovAux & Mid(sDataMov, Incrementa, tamanioFilaDetalle) & "|\n|"
                                            ErrorLog("sDataMovAux = " & sDataMovAux)
                                        End If
                                        Incrementa = Incrementa + tamanioFilaDetalle
                                    Next

                                    If Right(sTrama, 1) = "C" Then
                                        ErrorLog("Llamar a TramaSiguiente")
                                        sTrama = GetTramaR192(Constantes.TramaTrece, sNroTarjeta, sNroCuenta, sPeriodoFinal, sPagina, outpputBuff)
                                        If sTrama = "0" Then 'EXITO
                                            If outpputBuff.Length > 0 Then sTrama = outpputBuff.Substring(8, outpputBuff.Length - 8)
                                        ElseIf sTrama = "-2" Then
                                            sTrama = ""
                                            Exit Do
                                        Else
                                            sTrama = ""
                                            Exit Do
                                        End If

                                        If Left(sTrama, 2) <> "RU" Then
                                            sDataMov = ""
                                            sDataMov = Mid(sTrama, 861, sTrama.Length)
                                            ErrorLog("Movimiento sDataMov = " & sDataMov)
                                            If sDataMov.Length > 0 Then
                                                If Trim(Mid(sDataMov, 1, tamanioFilaDetalle)) <> "" Then
                                                    sDataMovAux = sDataMovAux & Mid(sDataMov, 1, tamanioFilaDetalle) & "|\n|"
                                                    ErrorLog("sDataMovAux = " & sDataMovAux)
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If

                            sXMLCAB = ""
                            sXMLCAB = sLineaCredito.Trim & "|\t|" & sDispEfectivoExpress.Trim & "|\t|" & sDisponibleCompras.Trim
                            sXMLCAB = sXMLCAB & "|\t|" & sPeriodoFacturacion.Trim & "|\t|" & sUltimoDiaPago.Trim
                            sXMLCAB = sXMLCAB & "|\t|" & sCreditoUtilizado.Trim & "|\t|" & sComisionCargos.Trim & "|\t|" & sDeudaTotal.Trim
                            sXMLCAB = sXMLCAB & "|\t|" & sDeudaVencida.Trim & "|\t|" & sPagoMinimoMes.Trim & "|\t|" & sPagoTotalMes.Trim
                            ErrorLog("sXMLCAB=" & sXMLCAB)

                            sXMLPIE = ""
                            sXMLPIE = sSaldoAnterior.Trim & "|\t|" & sConsumoMes.Trim & "|\t|" & sCuotaMes.Trim
                            sXMLPIE = sXMLPIE & "|\t|" & sInteres.Trim & "|\t|" & sComisionCargos_Mes.Trim & "|\t|" & sPagosAbonos.Trim
                            sXMLPIE = sXMLPIE & "|\t|" & sPagoTotal_Mes.Trim & "|\t|" & sSaldoFavor.Trim & "|\t|" & sMontoMinimo.Trim
                            sXMLPIE = sXMLPIE & "|\t|" & sCuotaMes_Min.Trim & "|\t|" & sInteres_Min.Trim & "|\t|" & sComisionCargos_Min.Trim
                            sXMLPIE = sXMLPIE & "|\t|" & sPagoMinimoMes_Min.Trim & "|\t|" & sMes1.Trim & "|\t|" & sMonto1.Trim
                            sXMLPIE = sXMLPIE & "|\t|" & sMes2.Trim & "|\t|" & sMonto2.Trim & "|\t|" & sMes3.Trim & "|\t|" & sMonto3.Trim
                            sXMLPIE = sXMLPIE & "|\t|" & sCntMeses.Trim & "|\t|" & sIntPagMin.Trim & "|\t|" & sComPagMin.Trim & "|\t|"
                            sXMLPIE = sXMLPIE & IIf(Servidor = TServidor.RSAT And sPeriodoFinal >= ReadAppConfig("PeriodoAgrandarGlosa"), "1", "0")

                            ErrorLog("sXMLPIE=" & sXMLPIE)
                        Else
                            'EVALUAR LA SEGUNDA CALL SOLO MOVIMIENTOS CONTADOR DE LLAMADAS
                            sDataMovAux = ""
                            If Servidor = TServidor.RSAT And sPeriodoFinal >= ReadAppConfig("PeriodoAgrandarGlosa") Then
                                sDataMov = Mid(sTrama, 882, sTrama.Length)
                            Else
                                sDataMov = Mid(sTrama, 861, sTrama.Length)
                            End If
                            ErrorLog("Movimiento sDataMov = " & sDataMov)

                            Incrementa = 1
                            If sDataMov.Length > 0 Then
                                If Servidor = TServidor.SICRON Or sPeriodoFinal < Constantes.PeriodoInclusionTEA Then
                                    For lFila = 1 To 13
                                        If Trim(Mid(sDataMov, Incrementa, tamanioFilaDetalleAnt)) <> "" Then
                                            sDataMovAux = sDataMovAux & Mid(sDataMov, Incrementa, tamanioFilaDetalleAnt) & "|\n|"
                                            ErrorLog("sDataMovAux = " & sDataMovAux)
                                        End If
                                        Incrementa = Incrementa + tamanioFilaDetalleAnt
                                    Next
                                ElseIf sPeriodoFinal >= ReadAppConfig("PeriodoAgrandarGlosa") Then
                                    '10-03-2015 Ocultar para pase por partes
                                    tamanioFilaDetalle = 159
                                    For lFila = 1 To 7
                                        If Trim(Mid(sDataMov, Incrementa, tamanioFilaDetalle)) <> "" Then
                                            sDataMovAux = sDataMovAux & Mid(sDataMov, Incrementa, tamanioFilaDetalle) & "|\n|"
                                            ErrorLog("sDataMovAux = " & sDataMovAux)
                                        End If
                                        Incrementa = Incrementa + tamanioFilaDetalle
                                    Next

                                    If Right(sTrama, 1) = "C" Then
                                        ErrorLog("Llamar a TramaSiguiente")
                                        sTrama = GetTramaR192(Constantes.TramaTrece, sNroTarjeta, sNroCuenta, sPeriodoFinal, sPagina, outpputBuff)
                                        If sTrama = "0" Then 'EXITO
                                            If outpputBuff.Length > 0 Then sTrama = outpputBuff.Substring(8, outpputBuff.Length - 8)
                                        ElseIf sTrama = "-2" Then
                                            sTrama = ""
                                            Exit Do
                                        Else
                                            sTrama = ""
                                            Exit Do
                                        End If

                                        If Left(sTrama, 2) <> "RU" Then
                                            sDataMov = ""
                                            sDataMov = Mid(sTrama, 882, sTrama.Length)
                                            ErrorLog("Movimiento sDataMov = " & sDataMov)
                                            Incrementa = 1
                                            If sDataMov.Length > 0 Then
                                                For lFila = 1 To 6
                                                    If Trim(Mid(sDataMov, Incrementa, tamanioFilaDetalle)) <> "" Then
                                                        sDataMovAux = sDataMovAux & Mid(sDataMov, Incrementa, tamanioFilaDetalle) & "|\n|"
                                                        ErrorLog("sDataMovAux = " & sDataMovAux)
                                                    End If
                                                    Incrementa = Incrementa + tamanioFilaDetalle
                                                Next
                                            End If
                                        End If
                                    End If
                                Else
                                    For lFila = 1 To 12
                                        If Trim(Mid(sDataMov, Incrementa, tamanioFilaDetalle)) <> "" Then
                                            sDataMovAux = sDataMovAux & Mid(sDataMov, Incrementa, tamanioFilaDetalle) & "|\n|"
                                            ErrorLog("sDataMovAux = " & sDataMovAux)
                                        End If
                                        Incrementa = Incrementa + tamanioFilaDetalle
                                    Next

                                    If Right(sTrama, 1) = "C" Then
                                        ErrorLog("Llamar a TramaSiguiente")
                                        sTrama = GetTramaR192(Constantes.TramaTrece, sNroTarjeta, sNroCuenta, sPeriodoFinal, sPagina, outpputBuff)
                                        If sTrama = "0" Then 'EXITO
                                            If outpputBuff.Length > 0 Then sTrama = outpputBuff.Substring(8, outpputBuff.Length - 8)
                                        ElseIf sTrama = "-2" Then
                                            sTrama = ""
                                            Exit Do
                                        Else
                                            sTrama = ""
                                            Exit Do
                                        End If

                                        If Left(sTrama, 2) <> "RU" Then
                                            sDataMov = ""
                                            sDataMov = Mid(sTrama, 861, sTrama.Length)
                                            ErrorLog("Movimiento sDataMov = " & sDataMov)
                                            If sDataMov.Length > 0 Then
                                                If Trim(Mid(sDataMov, 1, tamanioFilaDetalle)) <> "" Then
                                                    sDataMovAux = sDataMovAux & Mid(sDataMov, 1, tamanioFilaDetalle) & "|\n|"
                                                    ErrorLog("sDataMovAux = " & sDataMovAux)
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If 'FIN DE VALIDACION DEL CONTADOR
                        End If

                        If Right(sTrama, 1) = "C" Then
                            'If Left(sTrama, 2) = "AC" Then 'HAY MOVIMIENTOS PENDIENTES POR LLAMAR
                            sXMLMOV = sXMLMOV & sDataMovAux
                            ErrorLog("sXMLMOV = " & sXMLMOV)
                        End If
                        If Right(sTrama, 1) = "U" Then
                            'If Left(sTrama, 2) = "AU" Then 'FIN DE LA SOLICITUD
                            If sDataMovAux.Length > 0 Then
                                sDataMovAux = Left(sDataMovAux, sDataMovAux.Length - 4)
                                sXMLMOV = sXMLMOV & sDataMovAux
                                ErrorLog("sXMLMOV = " & sXMLMOV)
                            Else
                                sDataMovAux = "" 'NO HAY MOVIMIENTOS
                                sXMLMOV = sXMLMOV & sDataMovAux
                                ErrorLog("sXMLMOV = " & sXMLMOV)
                            End If

                            If sXMLMOV.Length > 0 Then
                                'Armar las columnas de los movimientos de EECC
                                sXMLMOV_FINAL = FUN_DETALLE_EECC_TEA(sXMLMOV, sPeriodoFinal, Servidor)
                                ErrorLog("sXMLMOV_FINAL = " & sXMLMOV_FINAL)
                            End If
                        End If

                        'CADENA FINAL CON LOS DATOS FINALES
                        sRespuesta = sXMLCAB.Trim & "*$¿*" & sXMLPIE.Trim & "*$¿*" & sXMLMOV_FINAL.Trim
                    End If

                    If lContador > 20 Then
                        Exit Do
                    End If
                Loop Until Right(sTrama, 1) <> "C"

                'Evaluar Respuesta si es ERROR,  RU, AU (Respuesta correcta)
                If sTrama.Trim.Length > 0 Then
                    If Left(sTrama.Trim, 2) = "RU" Then 'Mensaje de Validacion al usuario. RU            000CUENTA BLOQUEADA
                        sMensajeErrorUsuario = Mid(sTrama.Trim, 18, Len(sTrama.Trim))
                        sRespuesta = "ERROR:NODATA-" & sMensajeErrorUsuario.Trim
                    End If
                Else 'Sino devuelve nada
                    sRespuesta = "ERROR:Servicio no disponible."
                End If
            Else
                'Mostrar Mensaje de Error
                sRespuesta = "ERROR:Parametros Incompletos"
            End If
        Catch ex As Exception
            'save log error
            sRespuesta = "ERROR:" & ex.Message.Trim
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
            SAVE_LOG_CONSULTAS(sIDSucursal.Trim, sIDTerminal.Trim, "32", FUN_BUSCAR_TIPO_TARJETA(sNroTarjetax.Trim))

            sNroCuentax = Right("          " & sNroCuentax.Trim, 10)
            sNroTarjetax = Right(Strings.StrDup(20, " ") & sNroTarjetax.Trim, 20)
            sNombreCliente = Right(Strings.StrDup(26, " ") & sNombreCliente.Trim, 26)
            sSaldoDisponible = Right("      " & sSaldoDisponible.Trim, 6)
            sTotalDeuda = Right("      " & sTotalDeuda.Trim, 6)

            'armar variable desde flash con: numero de cuenta, numero de tarjeta, nombre del cliente, modo de entrada, canal de atencion, id de sucursal,codigo kiosco y codigo de transaccion (opcion)
            If Left(sRespuesta.Trim, 5) <> "ERROR" Then
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

        Return sRespuesta
    End Function

    Private Function GetTramaR192(ByVal CServidor As String, ByVal sNroTarjeta As String, ByVal sNroCuenta As String, ByVal sPeriodoFinal As String, ByVal sPagina As String, ByRef outpputBuff As String) As String
        Dim sParametros As String = ""
        'Variables nuevas
        Dim actions As Long = 1
        Dim inetputBuff As String = ""
        Dim errorMsg As String = ""

        'Instancia al mirapiweb
        Dim obSendMirror As ClsTxMirapi = Nothing
        obSendMirror = New ClsTxMirapi()

        Dim sTrama As String = ""

        sParametros = "0000000000" & CServidor & sNroTarjeta & sNroCuenta.Trim & sPeriodoFinal.Trim & sPagina
        inetputBuff = "      " + "R192" + sParametros
        ErrorLog("inetputBuff=" & inetputBuff)

        sTrama = ""
        sTrama = obSendMirror.ExecuteTX(actions.ToString, inetputBuff, outpputBuff, TIME_OUT_SERVER, SERVIDOR_MIRROR_DESTINO, SERVIDOR_MIRROR_NODE, PUERTO, errorMsg)
        ErrorLog("sTrama=" & sTrama)

        outpputBuff = RTrim(outpputBuff)
        ErrorLog("outpputBuff=" & outpputBuff)

        Return sTrama
    End Function

    'FUNCION QUE REALIZA EL CALCULO DEL SIMULADOR DEPOSITO A PLAZO

    <WebMethod(Description:="Simulador deposito a plazo")> _
    Public Function SIMULADOR_DEPOSITO_PLAZO(ByVal sMoneda As String, ByVal sMonto As String, ByVal sPlazoDias As String, ByVal sDATA_MONITOR_KIOSCO As String) As String
        Dim sRespuesta As String = ""
        Dim sParametros As String = ""
        Dim sMensajeErrorUsuario As String = ""
        Dim sXML As String = ""

        Try

            If Not IsNumeric(sMonto) Then
                sRespuesta = "ERROR:NODATA El monto ingresado no es válido !!!"
            Else
                'Validar Plazo en dias y obtener la tasa para realizar el calculo
                'sPlazoDias  sMoneda
                Dim STREA As String = ""

                STREA = GET_TASA_SIMULADOR_DPF(sMoneda, CDbl(sMonto.Trim), CLng(sPlazoDias.Trim))

                If STREA.Trim.Length > 0 Then
                    Dim dTasa As Double = 0
                    Dim dMonto As Double = 0
                    Dim iDiasPlazo As Integer = 0
                    Dim dInteresGanado As Double = 0

                    Dim sFechaKiosco As String = DateTime.Now.Day.ToString("00").Trim + "/" + DateTime.Now.Month.ToString("00").Trim + "/" + DateTime.Now.Year.ToString("0000").Trim
                    Dim sNOperacion As String = GET_NUMERO_OPERACION()
                    Dim sHoraKiosco As String = DateTime.Now.ToString("HH:mm:ss")

                    dTasa = CDbl(STREA)
                    dMonto = CDbl(sMonto.Trim)
                    iDiasPlazo = CInt(sPlazoDias.Trim)

                    'Calculo del interes ganado
                    dInteresGanado = dMonto * ((1 + dTasa / 100) ^ (iDiasPlazo / 360) - 1)

                    sRespuesta = STREA.Trim & "|\t|" & Format(dInteresGanado, "##,##0.00") & "|\t|" & STREA.Trim & "|\t|" & sFechaKiosco.Trim & "|\t|" & sNOperacion & "|\t|" & sHoraKiosco


                Else
                    sRespuesta = "ERROR:NODATA El plazo en dias ingresado no es válido !!!"

                End If


            End If


        Catch ex As Exception
            'save log error

            sRespuesta = "ERROR:" & ex.Message.Trim

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
            'SAVE_LOG_CONSULTAS(sIDSucursal.Trim, sIDTerminal.Trim, "52", FUN_BUSCAR_TIPO_TARJETA(sNroTarjetax.Trim))


            sNroCuentax = Right("          " & sNroCuentax.Trim, 10)
            sNroTarjetax = Right(Strings.StrDup(20, " ") & sNroTarjetax.Trim, 20)

            sNombreCliente = Right(Strings.StrDup(26, " ") & sNombreCliente.Trim, 26)

            sSaldoDisponible = Right("      " & sSaldoDisponible.Trim, 6)
            sTotalDeuda = Right("      " & sTotalDeuda.Trim, 6)


            'armar variable desde flash con: numero de cuenta, numero de tarjeta, nombre del cliente, modo de entrada, canal de atencion, id de sucursal,codigo kiosco y codigo de transaccion (opcion)
            If Left(sRespuesta.Trim, 5) <> "ERROR" Then
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


        Return sRespuesta

    End Function


    'FUNCION QUE REALIZA EL CALCULO DEL SIMULADOR DE OFERTA SEF
    '5,000.00|\t|36|\t|APLICA|\t|434.64|\t|260.00|\t|927.12|\t|12.01|\t|18.09|\t|03/05/2013|\t|006006|\t|12:31:32
    <WebMethod(Description:="Simulador para Super Efectivo")> _
    Public Function SIMULADOR_OFERTA_SUPEREFECTIVO(ByVal sTarjeta As String, ByVal sTasaTEM As String, ByVal sMonto As String, ByVal sPlazoMeses As String, ByVal sMontoOfertado As String) As String
        Dim sRespuesta As String = ""
        Dim sParametros As String = ""
        Dim sMensajeErrorUsuario As String = ""
        Dim sXML As String = ""
        Dim dMontoOfertado As Double = 0
        Dim dMontoMinimoPermitido As Double = 1
        Dim dProcIncrementoOferta As Single = 1


        Try

            If Not IsNumeric(sMonto) Then
                sRespuesta = "ERROR:NODATA-El monto ingresado no es válido."
            Else

                dMontoOfertado = IIf(sMontoOfertado.Trim.Length = 0, 0, CDbl(sMontoOfertado.Trim))

                'Consultar datos para el calculo
                Dim STRAMA_TEA_SIMULADOR As String = ""
                STRAMA_TEA_SIMULADOR = BUSCAR_DATOS_SIMULADOR_OFERTA_SEF(sTasaTEM, sTarjeta)

                Dim ADATA_TASAS As Array

                If STRAMA_TEA_SIMULADOR.Trim.Length > 0 Then
                    ADATA_TASAS = Split(STRAMA_TEA_SIMULADOR, "|\t|", , CompareMethod.Text)
                    'Leer el monto minimo permitido
                    dMontoMinimoPermitido = CDbl(ADATA_TASAS(7).ToString.Trim)
                Else
                    ADATA_TASAS = Nothing
                End If

                'Validar que el monto a simular no sea mayor al monto ofertado (El monto * 1 es para ampliar en algun momento el monto permitido)
                If CDbl(sMonto.Trim) <= (dMontoOfertado * dProcIncrementoOferta) Then

                    'Validar el monto minimo permitido
                    If CDbl(sMonto.Trim) >= dMontoMinimoPermitido Then

                        'Validar Plazo en Meses

                        If STRAMA_TEA_SIMULADOR.Trim.Length > 0 Then

                            Dim iPlazoMinimo As Integer = 0
                            Dim iPlazoMaximo As Integer = 0
                            Dim dCostoEnvioEECC As Double = 0
                            Dim dTasaSegDesgravamen As Single = 0
                            Dim dTasaSegPagos As Single = 0
                            Dim dTasaTEM As Single = 0
                            Dim dTasaTEA As Single = 0

                            Dim dMonto As Double = 0
                            Dim iMesesPlazo As Integer = 0
                            Dim dValorSegDesgravamen As Double = 0
                            Dim sSeguroProteccion As String = ""
                            Dim dValorSegProteccion As Double = 0
                            Dim dCostoSeguros As Double = 0
                            Dim dIntereses As Double = 0
                            Dim dTasaTCEA As Double = 0
                            Dim dPago As Double = 0
                            Dim dCuotaMensual As Double = 0
                            Dim dCuota2Mensual As Double = 0
                            Dim dProcTIR As Double = 0


                            Dim sFechaKiosco As String = DateTime.Now.Day.ToString("00").Trim + "/" + DateTime.Now.Month.ToString("00").Trim + "/" + DateTime.Now.Year.ToString("0000").Trim
                            Dim sNOperacion As String = GET_NUMERO_OPERACION()
                            Dim sHoraKiosco As String = DateTime.Now.ToString("HH:mm:ss")



                            'DATOS PARA REALIZAR EL CALCULO
                            iPlazoMinimo = CInt(ADATA_TASAS(0).ToString.Trim)
                            iPlazoMaximo = CInt(ADATA_TASAS(1).ToString.Trim)
                            dCostoEnvioEECC = CDbl(ADATA_TASAS(2).ToString.Trim)
                            dTasaSegDesgravamen = CSng(ADATA_TASAS(3).ToString.Trim)
                            dTasaSegPagos = CSng(ADATA_TASAS(4).ToString.Trim)
                            dTasaTEM = CSng(ADATA_TASAS(5).ToString.Trim)
                            dTasaTEA = CSng(ADATA_TASAS(6).ToString.Trim)

                            'Validar Plazo ingresado
                            If (CInt(sPlazoMeses.Trim) >= iPlazoMinimo And CInt(sPlazoMeses.Trim) <= iPlazoMaximo) Then

                                'Hallar el monto del pago incluido el interes
                                dPago = Financial.Pmt((dTasaTEM / 100), CDbl(sPlazoMeses.Trim), -CDbl(sMonto.Trim))

                                'Calculo del seguro de desgravamen
                                dValorSegDesgravamen = (dTasaSegDesgravamen / 100) * CDbl(sMonto.Trim)

                                'Calculo del seguro de proteccion
                                If CDbl(sMonto.Trim) <= 5000 Then
                                    sSeguroProteccion = "APLICA"
                                    dValorSegProteccion = (dTasaSegPagos / 100) * CDbl(sMonto.Trim)
                                Else
                                    sSeguroProteccion = "NO APLICA"
                                    dValorSegProteccion = 0
                                End If

                                'Total costo seguros
                                dCostoSeguros = dValorSegDesgravamen + dValorSegProteccion

                                'Cuota Mensual a pagar  Pago + Total Costo Seguros + Costo de Envio de EECC
                                dCuotaMensual = dPago + dCostoSeguros + dCostoEnvioEECC

                                dCuota2Mensual = dPago + dCostoEnvioEECC

                                'Hallar el total de intereses a pagar
                                dIntereses = (dPago * CDbl(sPlazoMeses.Trim)) - CDbl(sMonto.Trim)


                                '***************** INCIO Calculo del TCEA

                                '* Calcular el TIR
                                If CDbl(sPlazoMeses.Trim) > 0 Then

                                    'Variables para hallar el TIR
                                    Dim iFila As Integer = 0
                                    Dim DataHallarTIR(CDbl(sPlazoMeses.Trim)) As Double
                                    Dim dValor As Double = 0
                                    Dim Guess As Double = 0.1

                                    For iFila = 0 To CLng(sPlazoMeses.Trim)
                                        If iFila = 0 Then
                                            'Agregar el monto del prestamo en negativo
                                            dValor = -CDbl(sMonto.Trim)
                                        ElseIf iFila = 1 Then
                                            'Primera Cuota Mensual sin seguro de proteccion
                                            dValor = dCuotaMensual - dValorSegProteccion
                                        Else
                                            'Demas cuota mensual - total costo de seguros
                                            dValor = dCuotaMensual - dCostoSeguros
                                        End If
                                        DataHallarTIR(iFila) = dValor
                                    Next

                                    'Hallar TIR
                                    dProcTIR = Financial.IRR(DataHallarTIR, Guess) * 100

                                End If

                                dTasaTCEA = (((dProcTIR / 100) + 1) ^ 12 - 1) * 100

                                '***************** FIN Calculo del TCEA

                                sRespuesta = Format(CDbl(sMonto.Trim), "##,##0.00") & "|\t|" & sPlazoMeses.Trim & "|\t|" & sSeguroProteccion.Trim & "|\t|" & Format(dCuota2Mensual, "##,##0.00")
                                sRespuesta = sRespuesta & "|\t|" & Format(dCostoSeguros, "##,##0.00") & "|\t|" & Format(dIntereses, "##,##0.00") & "|\t|" & Format(dTasaTEA, "##,##0.00")
                                sRespuesta = sRespuesta & "|\t|" & Format(dTasaTCEA, "##,##0.00") & "|\t|" & sFechaKiosco.Trim & "|\t|" & sNOperacion & "|\t|" & sHoraKiosco

                            Else
                                'El plazo ingresado no esta en el rango de minimo y maximo configurado para esta tasa TEM
                                sRespuesta = "ERROR:NODATA-El plazo ingresado no está permitido, puede ingresar de (" & iPlazoMinimo.ToString & "-" & iPlazoMaximo.ToString & ") meses."
                            End If

                        Else
                            sRespuesta = "ERROR:NODATA-No se encontrarón los parámetros para los cálculos de la simulación."
                        End If


                    Else
                        'Mostrar mensaje de validacion del monto minimo permitido
                        sRespuesta = "ERROR:NODATA-El monto ingresado no está permitido, puede ingresar hasta un mínimo de ( S/." & Format(dMontoMinimoPermitido, "##,##0.00") & ")"
                    End If

                Else
                    'Mostrar mensaje el monto ingresado no está permitido, es mayor a la oferta presentada. 
                    sRespuesta = "ERROR:NODATA-El monto ingresado no está permitido, puede ingresar hasta un máximo de ( S/." & Format((dMontoOfertado * dProcIncrementoOferta), "##,##0.00") & ")"
                End If

            End If 'Fin validar monto numerico

        Catch ex As Exception
            'save log error
            sRespuesta = "ERROR:" & ex.Message.Trim
        End Try

        Return sRespuesta
    End Function

    'FUNCION QUE REALIZA EL CALCULO DEL SIMULADOR DE OFERTA SEF
    '5,000.00|\t|36|\t|APLICA|\t|434.64|\t|260.00|\t|927.12|\t|12.01|\t|18.09|\t|03/05/2013|\t|006006|\t|12:31:32
    <WebMethod(Description:="Simulador para Super Efectivo")> _
    Public Function SIMULADOR_OFERTA_SEF_CUOTA(ByVal sTarjeta As String, _
                                               ByVal sTasaTEM As String, _
                                               ByVal sMonto As String, _
                                               ByVal sPlazoMeses As String, _
                                               ByVal sMontoOfertado As String, _
                                               ByVal sDiferido As String, _
                                               ByVal sFechaConsumo As String, _
                                               ByVal sNroCuenta As String, _
                                               ByVal sCodKiosco As String, _
                                               ByVal sDATA_MONITOR_KIOSCO As String, _
                                               ByVal Servidor As TServidor, _
                                               ByVal sMigrado As String) As String
        Dim sRespuesta As String = ""
        Dim sParametros As String = ""
        Dim sMensajeErrorUsuario As String = ""
        Dim sXML As String = ""
        Dim dMontoOfertado As Double = 0
        Dim dMontoMinimoPermitido As Double = 1
        Dim dProcIncrementoOferta As Single = 1
        Dim dMontoMaximoPorFactor As Double = 1
        Dim bExcedeMaximoFactor As Boolean = False

        ErrorLog("sFechaConsumo=" & sFechaConsumo)
        Dim fechaDesem As DateTime = Helpers.GetDatetimeFromString(Constantes.FORMATO_FECHASDESEMBOLSO, sFechaConsumo)
        ErrorLog("fechaDesem=" & fechaDesem.ToShortDateString())
        Dim validFechaDesem As Boolean = ValidarFechaDesembolso(fechaDesem)
        ErrorLog("validFechaDesem=" & validFechaDesem.ToString())


        Try
            If validFechaDesem Then
                If Not IsNumeric(sMonto) Then
                    sRespuesta = "ERROR:NODATA-El monto ingresado no es válido."
                Else

                    dMontoOfertado = IIf(sMontoOfertado.Trim.Length = 0, 0, CDbl(sMontoOfertado.Trim))

                    'Consultar datos para el calculo
                    Dim STRAMA_TEA_SIMULADOR As String = ""
                    STRAMA_TEA_SIMULADOR = BUSCAR_DATOS_SIMULADOR_OFERTA_SEF(sTasaTEM, sTarjeta)

                    Dim ADATA_TASAS As Array

                    If STRAMA_TEA_SIMULADOR.Trim.Length > 0 Then
                        ADATA_TASAS = Split(STRAMA_TEA_SIMULADOR, "|\t|", , CompareMethod.Text)
                        'Leer el monto minimo permitido
                        dMontoMinimoPermitido = CDbl(ADATA_TASAS(7).ToString.Trim)
                    Else
                        ADATA_TASAS = Nothing
                    End If

                    'Validar que el monto a simular no sea mayor al monto ofertado (El monto * 1 es para ampliar en algun momento el monto permitido)
                    If CDbl(sMonto.Trim) <= (dMontoOfertado * dProcIncrementoOferta) Then

                        'Validar el monto minimo permitido
                        If CDbl(sMonto.Trim) >= dMontoMinimoPermitido Then

                            'Validar Plazo en Meses

                            If STRAMA_TEA_SIMULADOR.Trim.Length > 0 Then

                                Dim iPlazoMinimo As Integer = 0
                                Dim iPlazoMaximo As Integer = 0
                                Dim dCostoEnvioEECC As Double = 0
                                Dim dTasaSegDesgravamen As Single = 0
                                Dim dTasaSegPagos As Single = 0
                                Dim dTasaTEM As Single = 0
                                Dim dTasaTEA As Single = 0

                                Dim dMonto As Double = 0
                                Dim iMesesPlazo As Integer = 0
                                Dim dValorSegDesgravamen As Double = 0
                                Dim sSeguroProteccion As String = ""
                                Dim dValorSegProteccion As Double = 0
                                Dim dCostoSeguros As Double = 0
                                Dim dIntereses As Double = 0
                                Dim dTasaTCEA As Double = 0
                                Dim dPago As Double = 0
                                Dim dCuotaMensual As Double = 0
                                Dim dCuota2Mensual As Double = 0
                                Dim dProcTIR As Double = 0
                                Dim dCuotaMesCobol As Double = 0


                                Dim sFechaKiosco As String = DateTime.Now.Day.ToString("00").Trim + "/" + DateTime.Now.Month.ToString("00").Trim + "/" + DateTime.Now.Year.ToString("0000").Trim
                                Dim sNOperacion As String = GET_NUMERO_OPERACION()
                                Dim sHoraKiosco As String = DateTime.Now.ToString("HH:mm:ss")



                                'DATOS PARA REALIZAR EL CALCULO
                                iPlazoMinimo = CInt(ADATA_TASAS(0).ToString.Trim)
                                iPlazoMaximo = CInt(ADATA_TASAS(1).ToString.Trim)
                                dCostoEnvioEECC = CDbl(ADATA_TASAS(2).ToString.Trim)
                                dTasaSegDesgravamen = CSng(ADATA_TASAS(3).ToString.Trim)
                                dTasaSegPagos = CSng(ADATA_TASAS(4).ToString.Trim)
                                dTasaTEM = CSng(ADATA_TASAS(5).ToString.Trim)
                                dTasaTEA = CSng(ADATA_TASAS(6).ToString.Trim)

                                'Validar Plazo ingresado
                                If (CInt(sPlazoMeses.Trim) >= iPlazoMinimo And CInt(sPlazoMeses.Trim) <= iPlazoMaximo) Then

                                    'Validación de Plazos mediante Factores
                                    If CDbl(sMonto.Trim) > 5000 Then
                                        Dim factorPlazo As String = ObtenerFactorPorPlazo(sPlazoMeses)
                                        dMontoMaximoPorFactor = (dMontoOfertado / 2) * CDbl(factorPlazo.Trim())
                                        If dMontoMaximoPorFactor < CDbl(sMonto.Trim) Then
                                            bExcedeMaximoFactor = True
                                        End If
                                    End If

                                    If bExcedeMaximoFactor Then
                                        sRespuesta = "ERROR:NODATA-A un plazo de " & sPlazoMeses.Trim & " meses puede solicitar un monto máximo de S/. " & Format(dMontoMaximoPorFactor, "##,##0.00")
                                    Else

                                        ''Inicio Invocación a COBOL SFSCANT0013
                                        Dim respuestaFecha As String = String.Empty
                                        sFechaConsumo = fx_Completar_Campo("0", 4, fechaDesem.Year, TYPE_ALINEAR.DERECHA)
                                        sFechaConsumo += fx_Completar_Campo("0", 2, fechaDesem.Month, TYPE_ALINEAR.DERECHA)
                                        sFechaConsumo += fx_Completar_Campo("0", 2, fechaDesem.Day, TYPE_ALINEAR.DERECHA)
                                        ErrorLog("sFechaConsumo=" & sFechaConsumo)
                                        respuestaFecha = OBTENER_FECHA_FACTURACION(sTarjeta, sNroCuenta, sCodKiosco, sDATA_MONITOR_KIOSCO, Servidor, sMigrado, sFechaConsumo)
                                        ErrorLog("respuestaFecha=" & respuestaFecha)
                                        ''Fin Invocación a COBOL SFSCANT0013
                                        If respuestaFecha.Substring(0, 5) = "ERROR" Then
                                            sRespuesta = respuestaFecha
                                        Else

                                            ''Inicio Invocación a COBOL SFSCANT0013
                                            Dim respuestaCuota As String = String.Empty

                                            respuestaCuota = ObtenerCuotaMes(sMonto.Trim, sPlazoMeses.Trim, sDiferido, CStr(dTasaTEA), sFechaConsumo, respuestaFecha)
                                            ErrorLog("respuestaCuota=" & respuestaCuota)
                                            ''Fin Invocación a COBOL SFSCANT0013
                                            If respuestaCuota.Substring(0, 5) = "ERROR" Then
                                                sRespuesta = respuestaCuota
                                            Else
                                                dCuotaMesCobol = CDbl(respuestaCuota) / 100
                                                'Hallar el monto del pago incluido el interes
                                                dPago = Financial.Pmt((dTasaTEM / 100), CDbl(sPlazoMeses.Trim), -CDbl(sMonto.Trim))

                                                'Calculo del seguro de desgravamen
                                                dValorSegDesgravamen = (dTasaSegDesgravamen / 100) * CDbl(sMonto.Trim)

                                                'Calculo del seguro de proteccion
                                                If CDbl(sMonto.Trim) <= 5000 Then
                                                    sSeguroProteccion = "APLICA"
                                                    dValorSegProteccion = (dTasaSegPagos / 100) * CDbl(sMonto.Trim)
                                                Else
                                                    sSeguroProteccion = "NO APLICA"
                                                    dValorSegProteccion = 0
                                                End If

                                                'Total costo seguros
                                                dCostoSeguros = dValorSegDesgravamen + dValorSegProteccion

                                                'Cuota Mensual a pagar  Pago + Total Costo Seguros + Costo de Envio de EECC
                                                dCuotaMensual = dPago + dCostoSeguros + dCostoEnvioEECC

                                                dCuota2Mensual = dPago + dCostoEnvioEECC

                                                'Hallar el total de intereses a pagar
                                                dIntereses = (dPago * CDbl(sPlazoMeses.Trim)) - CDbl(sMonto.Trim)


                                                '***************** INCIO Calculo del TCEA

                                                '* Calcular el TIR
                                                If CDbl(sPlazoMeses.Trim) > 0 Then

                                                    'Variables para hallar el TIR
                                                    Dim iFila As Integer = 0
                                                    Dim DataHallarTIR(CDbl(sPlazoMeses.Trim)) As Double
                                                    Dim dValor As Double = 0
                                                    Dim Guess As Double = 0.1

                                                    For iFila = 0 To CLng(sPlazoMeses.Trim)
                                                        If iFila = 0 Then
                                                            'Agregar el monto del prestamo en negativo
                                                            dValor = -CDbl(sMonto.Trim)
                                                        ElseIf iFila = 1 Then
                                                            'Primera Cuota Mensual sin seguro de proteccion
                                                            dValor = dCuotaMensual - dValorSegProteccion
                                                        Else
                                                            'Demas cuota mensual - total costo de seguros
                                                            dValor = dCuotaMensual - dCostoSeguros
                                                        End If
                                                        DataHallarTIR(iFila) = dValor
                                                    Next

                                                    'Hallar TIR
                                                    dProcTIR = Financial.IRR(DataHallarTIR, Guess) * 100

                                                End If

                                                dTasaTCEA = (((dProcTIR / 100) + 1) ^ 12 - 1) * 100

                                                '***************** FIN Calculo del TCEA

                                                sRespuesta = Format(CDbl(sMonto.Trim), "##,##0.00") & "|\t|" & sPlazoMeses.Trim & "|\t|" & sSeguroProteccion.Trim & "|\t|" & Format(dCuotaMesCobol, "##,##0.00")
                                                sRespuesta = sRespuesta & "|\t|" & Format(dCostoSeguros, "##,##0.00") & "|\t|" & Format(dIntereses, "##,##0.00") & "|\t|" & Format(dTasaTEA, "##,##0.00")
                                                sRespuesta = sRespuesta & "|\t|" & Format(dTasaTCEA, "##,##0.00") & "|\t|" & sFechaKiosco.Trim & "|\t|" & sNOperacion & "|\t|" & sHoraKiosco
                                            End If

                                        End If

                                    End If
                                Else
                                    'El plazo ingresado no esta en el rango de minimo y maximo configurado para esta tasa TEM
                                    sRespuesta = "ERROR:NODATA-El plazo ingresado no está permitido, puede ingresar de (" & iPlazoMinimo.ToString & "-" & iPlazoMaximo.ToString & ") meses."
                                End If

                            Else
                                sRespuesta = "ERROR:NODATA-No se encontrarón los parámetros para los cálculos de la simulación."
                            End If


                        Else
                            'Mostrar mensaje de validacion del monto minimo permitido
                            sRespuesta = "ERROR:NODATA-El monto ingresado no está permitido, puede ingresar hasta un mínimo de ( S/." & Format(dMontoMinimoPermitido, "##,##0.00") & ")"
                        End If

                    Else
                        'Mostrar mensaje el monto ingresado no está permitido, es mayor a la oferta presentada. 
                        sRespuesta = "ERROR:NODATA-El monto ingresado no está permitido, puede ingresar hasta un máximo de ( S/." & Format((dMontoOfertado * dProcIncrementoOferta), "##,##0.00") & ")"
                    End If

                End If 'Fin validar monto numerico
            Else
                sRespuesta = "ERROR:NODATA-La fecha de desembolso es incorrecta."
            End If

        Catch ex As Exception
            'save log error
            sRespuesta = "ERROR:" & ex.Message.Trim
        End Try

        Return sRespuesta
    End Function

    'FUNCION QUE REALIZA EL CALCULO DEL SIMULADOR DE OFERTA SEF
    '5,000.00|\t|36|\t|APLICA|\t|434.64|\t|260.00|\t|927.12|\t|12.01|\t|18.09|\t|03/05/2013|\t|006006|\t|12:31:32
    <WebMethod(Description:="Simulador para Super Efectivo considerando el plazo")> _
    Public Function SIMULADOR_OFERTA_SEF_CUOTA_PLAZO(ByVal sTarjeta As String, _
                                               ByVal sTasaTEM As String, _
                                               ByVal sMonto As String, _
                                               ByVal sPlazoMeses As String, _
                                               ByVal sMontoOfertado As String, _
                                               ByVal sPlazoOfertado As String, _
                                               ByVal sDiferido As String, _
                                               ByVal sFechaConsumo As String, _
                                               ByVal sNroCuenta As String, _
                                               ByVal sCodKiosco As String, _
                                               ByVal sDATA_MONITOR_KIOSCO As String, _
                                               ByVal Servidor As TServidor, _
                                               ByVal sMigrado As String) As String
        Dim sRespuesta As String = ""
        Dim sParametros As String = ""
        Dim sMensajeErrorUsuario As String = ""
        Dim sXML As String = ""
        Dim dMontoOfertado As Double = 0
        Dim dPlazoOfertado As Double = 0
        Dim dMontoMinimoPermitido As Double = 1
        Dim dProcIncrementoOferta As Single = 1
        Dim dMontoMaximoPorFactor As Double = 1
        Dim bExcedeMaximoFactor As Boolean = False

        ErrorLog("sFechaConsumo=" & sFechaConsumo)
        Dim fechaDesem As DateTime = Helpers.GetDatetimeFromString(Constantes.FORMATO_FECHASDESEMBOLSO, sFechaConsumo)
        ErrorLog("fechaDesem=" & fechaDesem.ToShortDateString())
        Dim validFechaDesem As Boolean = ValidarFechaDesembolso(fechaDesem)
        ErrorLog("validFechaDesem=" & validFechaDesem.ToString())


        Try
            If validFechaDesem Then
                If Not IsNumeric(sMonto) Then
                    sRespuesta = "ERROR:NODATA-El monto ingresado no es válido."
                Else

                    dMontoOfertado = IIf(sMontoOfertado.Trim.Length = 0, 0, CDbl(sMontoOfertado.Trim))
                    dPlazoOfertado = IIf(sPlazoOfertado.Trim.Length = 0, 0, CDbl(sPlazoOfertado.Trim))

                    'Consultar datos para el calculo
                    Dim STRAMA_TEA_SIMULADOR As String = ""
                    STRAMA_TEA_SIMULADOR = BUSCAR_DATOS_SIMULADOR_OFERTA_SEF(sTasaTEM, sTarjeta)

                    Dim ADATA_TASAS As Array

                    If STRAMA_TEA_SIMULADOR.Trim.Length > 0 Then
                        ADATA_TASAS = Split(STRAMA_TEA_SIMULADOR, "|\t|", , CompareMethod.Text)
                        'Leer el monto minimo permitido
                        dMontoMinimoPermitido = CDbl(ADATA_TASAS(7).ToString.Trim)
                    Else
                        ADATA_TASAS = Nothing
                    End If

                    'Validar que el monto a simular no sea mayor al monto ofertado (El monto * 1 es para ampliar en algun momento el monto permitido)
                    If CDbl(sMonto.Trim) <= (dMontoOfertado * dProcIncrementoOferta) Then

                        'Validar el monto minimo permitido
                        If CDbl(sMonto.Trim) >= dMontoMinimoPermitido Then

                            'Validar Plazo en Meses

                            If STRAMA_TEA_SIMULADOR.Trim.Length > 0 Then

                                Dim iPlazoMinimo As Integer = 0
                                Dim iPlazoMaximo As Integer = 0
                                Dim dCostoEnvioEECC As Double = 0
                                Dim dTasaSegDesgravamen As Single = 0
                                Dim dTasaSegPagos As Single = 0
                                Dim dTasaTEM As Single = 0
                                Dim dTasaTEA As Single = 0

                                Dim dMonto As Double = 0
                                Dim iMesesPlazo As Integer = 0
                                Dim dValorSegDesgravamen As Double = 0
                                Dim sSeguroProteccion As String = ""
                                Dim dValorSegProteccion As Double = 0
                                Dim dCostoSeguros As Double = 0
                                Dim dIntereses As Double = 0
                                Dim dTasaTCEA As Double = 0
                                Dim dPago As Double = 0
                                Dim dCuotaMensual As Double = 0
                                Dim dCuota2Mensual As Double = 0
                                Dim dProcTIR As Double = 0
                                Dim dCuotaMesCobol As Double = 0


                                Dim sFechaKiosco As String = DateTime.Now.Day.ToString("00").Trim + "/" + DateTime.Now.Month.ToString("00").Trim + "/" + DateTime.Now.Year.ToString("0000").Trim
                                Dim sNOperacion As String = GET_NUMERO_OPERACION()
                                Dim sHoraKiosco As String = DateTime.Now.ToString("HH:mm:ss")



                                'DATOS PARA REALIZAR EL CALCULO
                                iPlazoMinimo = CInt(ADATA_TASAS(0).ToString.Trim)
                                iPlazoMaximo = CInt(ADATA_TASAS(1).ToString.Trim)
                                dCostoEnvioEECC = CDbl(ADATA_TASAS(2).ToString.Trim)
                                dTasaSegDesgravamen = CSng(ADATA_TASAS(3).ToString.Trim)
                                dTasaSegPagos = CSng(ADATA_TASAS(4).ToString.Trim)
                                dTasaTEM = CSng(ADATA_TASAS(5).ToString.Trim)
                                dTasaTEA = CSng(ADATA_TASAS(6).ToString.Trim)

                                'Validar Plazo ingresado
                                If (CInt(sPlazoMeses.Trim) >= iPlazoMinimo And CInt(sPlazoMeses.Trim) <= dPlazoOfertado) Then

                                    'Validación de Plazos mediante Factores
                                    If CDbl(sMonto.Trim) > 5000 Then
                                        Dim factorPlazo As String = ObtenerFactorPorPlazo(sPlazoMeses)
                                        dMontoMaximoPorFactor = (dMontoOfertado / 2) * CDbl(factorPlazo.Trim())
                                        If dMontoMaximoPorFactor < CDbl(sMonto.Trim) Then
                                            bExcedeMaximoFactor = True
                                        End If
                                    End If

                                    If bExcedeMaximoFactor Then
                                        sRespuesta = "ERROR:NODATA-A un plazo de " & sPlazoMeses.Trim & " meses puede solicitar un monto máximo de S/. " & Format(dMontoMaximoPorFactor, "##,##0.00")
                                    Else

                                        ''Inicio Invocación a COBOL SFSCANT0013
                                        Dim respuestaFecha As String = String.Empty
                                        sFechaConsumo = fx_Completar_Campo("0", 4, fechaDesem.Year, TYPE_ALINEAR.DERECHA)
                                        sFechaConsumo += fx_Completar_Campo("0", 2, fechaDesem.Month, TYPE_ALINEAR.DERECHA)
                                        sFechaConsumo += fx_Completar_Campo("0", 2, fechaDesem.Day, TYPE_ALINEAR.DERECHA)
                                        ErrorLog("sFechaConsumo=" & sFechaConsumo)
                                        respuestaFecha = OBTENER_FECHA_FACTURACION(sTarjeta, sNroCuenta, sCodKiosco, sDATA_MONITOR_KIOSCO, Servidor, sMigrado, sFechaConsumo)
                                        ErrorLog("respuestaFecha=" & respuestaFecha)
                                        ''Fin Invocación a COBOL SFSCANT0013
                                        If respuestaFecha.Substring(0, 5) = "ERROR" Then
                                            sRespuesta = respuestaFecha
                                        Else

                                            ''Inicio Invocación a COBOL SFSCANT0013
                                            Dim respuestaCuota As String = String.Empty

                                            respuestaCuota = ObtenerCuotaMes(sMonto.Trim, sPlazoMeses.Trim, sDiferido, CStr(dTasaTEA), sFechaConsumo, respuestaFecha)
                                            ErrorLog("respuestaCuota=" & respuestaCuota)
                                            ''Fin Invocación a COBOL SFSCANT0013
                                            If respuestaCuota.Substring(0, 5) = "ERROR" Then
                                                sRespuesta = respuestaCuota
                                            Else
                                                dCuotaMesCobol = CDbl(respuestaCuota) / 100
                                                'Hallar el monto del pago incluido el interes
                                                dPago = Financial.Pmt((dTasaTEM / 100), CDbl(sPlazoMeses.Trim), -CDbl(sMonto.Trim))

                                                'Calculo del seguro de desgravamen
                                                dValorSegDesgravamen = (dTasaSegDesgravamen / 100) * CDbl(sMonto.Trim)

                                                'Calculo del seguro de proteccion
                                                If CDbl(sMonto.Trim) <= 5000 Then
                                                    sSeguroProteccion = "APLICA"
                                                    dValorSegProteccion = (dTasaSegPagos / 100) * CDbl(sMonto.Trim)
                                                Else
                                                    sSeguroProteccion = "NO APLICA"
                                                    dValorSegProteccion = 0
                                                End If

                                                'Total costo seguros
                                                dCostoSeguros = dValorSegDesgravamen + dValorSegProteccion

                                                'Cuota Mensual a pagar  Pago + Total Costo Seguros + Costo de Envio de EECC
                                                dCuotaMensual = dPago + dCostoSeguros + dCostoEnvioEECC

                                                dCuota2Mensual = dPago + dCostoEnvioEECC

                                                'Hallar el total de intereses a pagar
                                                dIntereses = (dPago * CDbl(sPlazoMeses.Trim)) - CDbl(sMonto.Trim)


                                                '***************** INCIO Calculo del TCEA

                                                '* Calcular el TIR
                                                If CDbl(sPlazoMeses.Trim) > 0 Then

                                                    'Variables para hallar el TIR
                                                    Dim iFila As Integer = 0
                                                    Dim DataHallarTIR(CDbl(sPlazoMeses.Trim)) As Double
                                                    Dim dValor As Double = 0
                                                    Dim Guess As Double = 0.1

                                                    For iFila = 0 To CLng(sPlazoMeses.Trim)
                                                        If iFila = 0 Then
                                                            'Agregar el monto del prestamo en negativo
                                                            dValor = -CDbl(sMonto.Trim)
                                                        ElseIf iFila = 1 Then
                                                            'Primera Cuota Mensual sin seguro de proteccion
                                                            dValor = dCuotaMensual - dValorSegProteccion
                                                        Else
                                                            'Demas cuota mensual - total costo de seguros
                                                            dValor = dCuotaMensual - dCostoSeguros
                                                        End If
                                                        DataHallarTIR(iFila) = dValor
                                                    Next

                                                    'Hallar TIR
                                                    dProcTIR = Financial.IRR(DataHallarTIR, Guess) * 100

                                                End If

                                                dTasaTCEA = (((dProcTIR / 100) + 1) ^ 12 - 1) * 100

                                                '***************** FIN Calculo del TCEA

                                                sRespuesta = Format(CDbl(sMonto.Trim), "##,##0.00") & "|\t|" & sPlazoMeses.Trim & "|\t|" & sSeguroProteccion.Trim & "|\t|" & Format(dCuotaMesCobol, "##,##0.00")
                                                sRespuesta = sRespuesta & "|\t|" & Format(dCostoSeguros, "##,##0.00") & "|\t|" & Format(dIntereses, "##,##0.00") & "|\t|" & Format(dTasaTEA, "##,##0.00")
                                                sRespuesta = sRespuesta & "|\t|" & Format(dTasaTCEA, "##,##0.00") & "|\t|" & sFechaKiosco.Trim & "|\t|" & sNOperacion & "|\t|" & sHoraKiosco
                                            End If

                                        End If

                                    End If
                                Else
                                    'El plazo ingresado no esta en el rango de minimo y maximo configurado para esta tasa TEM
                                    sRespuesta = "ERROR:NODATA-El plazo ingresado no está permitido, puede ingresar de (" & iPlazoMinimo.ToString & "-" & dPlazoOfertado.ToString & ") meses."
                                End If

                            Else
                                sRespuesta = "ERROR:NODATA-No se encontrarón los parámetros para los cálculos de la simulación."
                            End If


                        Else
                            'Mostrar mensaje de validacion del monto minimo permitido
                            sRespuesta = "ERROR:NODATA-El monto ingresado no está permitido, puede ingresar hasta un mínimo de ( S/." & Format(dMontoMinimoPermitido, "##,##0.00") & ")"
                        End If

                    Else
                        'Mostrar mensaje el monto ingresado no está permitido, es mayor a la oferta presentada. 
                        sRespuesta = "ERROR:NODATA-El monto ingresado no está permitido, puede ingresar hasta un máximo de ( S/." & Format((dMontoOfertado * dProcIncrementoOferta), "##,##0.00") & ")"
                    End If

                End If 'Fin validar monto numerico
            Else
                sRespuesta = "ERROR:NODATA-La fecha de desembolso es incorrecta."
            End If

        Catch ex As Exception
            'save log error
            sRespuesta = "ERROR:" & ex.Message.Trim
        End Try

        Return sRespuesta
    End Function


    'FUNCION QUE ACTUALIZA OFERTA CON LA ÚLTIMA SIMULACIÓN ACEPTADA
    '5,000.00|\t|36|\t|APLICA|\t|434.64|\t|260.00|\t|927.12|\t|12.01|\t|18.09|\t|03/05/2013|\t|006006|\t|12:31:32
    <WebMethod(Description:="Actualización Oferta con la última simulación aceptada")> _
    Public Function SIMULADOR_OFERTA_SEF_ACEPTADO(ByVal sTipoProducto As String, ByVal sTipoDocumento As String, ByVal sNroDocumento As String, ByVal sMonto As String, ByVal sPlazoMeses As String, ByVal sTasaTCEA As String, ByVal sCuota As String, ByVal sUsuario As String, ByVal codigoSucursal As String) As String
        ErrorLog(" sTipoProducto: " & sTipoProducto & " sTipoDocumento: " & sTipoDocumento & " sNroDocumento: " & sNroDocumento & " sMonto: " & sMonto & " sPlazoMeses: " & sPlazoMeses & " sTasa: " & sTasaTCEA & " sCuota: " & sCuota & " sUsuario: " & sUsuario & " codigoSucursal: " & codigoSucursal)
        Dim sRespuesta As String = ""
        Dim oferta As New OfertaCliente

        Try
            'Actualizar la tabla de VEX_OFERTAS_PROD con los nuevos valores del Simulador
            ErrorLog("Llamar al método: ActualizarOfertaPorSilumacion(" & sTipoProducto & "," & sTipoDocumento & "," & sNroDocumento & "," & sMonto & "," & sPlazoMeses & "," & sTasaTCEA & "," & sCuota & "," & sUsuario & "," & codigoSucursal & ")")
            oferta = BNOfertaCliente.Instancia.ActualizarOfertaPorSilumacion(sTipoProducto, sTipoDocumento, sNroDocumento, sMonto, sPlazoMeses, sTasaTCEA, sCuota, sUsuario, codigoSucursal)
            If oferta.Estado <> "" Then
                sRespuesta = "ERROR:NODATA-" & oferta.Estado
            End If
            ErrorLog("Valor de srespuesta: " & sRespuesta)

        Catch ex As Exception
            'save log error
            sRespuesta = "ERROR:" & ex.Message.Trim
        End Try

        Return sRespuesta
    End Function

    <WebMethod(Description:="Consolidacion del Cliente.")> _
    Public Function CONSOLIDADO_CLIENTE(ByVal sTipoProducto As String, ByVal sNroTarjeta As String, ByVal sNroCuenta As String, ByVal sDNI As String, ByVal sDATA_MONITOR_KIOSCO As String, ByVal Servidor As TServidor) As String

        Dim sRespuesta As String = ""
        Dim sDataSaldosTarjeta As String = ""
        Dim sXML_DATA As String = ""
        Dim sDataPEfectivo As String = ""

        Try

            'SALDO DE TARJETA DE CREDITO
            Dim sSaldosTarjeta As String = ""

            If sTipoProducto.Trim.ToUpper = "C" Then 'CLASICA
                Select Case Servidor
                    Case TServidor.SICRON
                        sSaldosTarjeta = FUN_BUSCAR_LINEA_DISPONIBLE_COMPRAS_CLASICA_SICRON(sNroCuenta.Trim)
                    Case TServidor.RSAT
                        sSaldosTarjeta = FUN_BUSCAR_LINEA_DISPONIBLE_COMPRAS_CLASICA_RSAT(sNroTarjeta, sNroCuenta.Trim)
                End Select
            Else 'ASOCIADAS
                sSaldosTarjeta = FUN_BUSCAR_LINEA_DISPONIBLE_COMPRAS_ASOCIADA(sNroTarjeta.Trim)
            End If

            If sSaldosTarjeta.Trim.Length > 0 Then
                sDataSaldosTarjeta = ""
                sDataSaldosTarjeta = "PRODUCTO|\t|NRO. DE TARJETA|\t|LÍNEA DISPONIBLE|\n|"
                sDataSaldosTarjeta = sDataSaldosTarjeta & sSaldosTarjeta.Trim
            End If

            sXML_DATA = sDataSaldosTarjeta


            'PRESTAMO EN EFECTIVO
            sDataPEfectivo = CONSULTA_PRESTAMO_EFECTIVO(sDNI.Trim, "")

            If Left(sDataPEfectivo.Trim.ToUpper, 5) <> "ERROR" Then
                'Sacar los datos del Prestamo
                Dim ADATA_PEF As Array
                Dim ADATA_PEF_ As Array
                Dim sDATA_PEF As String = ""
                Dim lIndice As Long = 0
                Dim SDATA_REGISTROS As String = ""
                Dim ADATA_REGISTROS As Array
                Dim sDataRecuperada As String = ""

                ADATA_PEF = Split(sDataPEfectivo, "*$¿*", , CompareMethod.Text)
                sDATA_PEF = ADATA_PEF(1)

                ADATA_PEF_ = Split(sDATA_PEF, "|\n|", , CompareMethod.Text)

                lIndice = 0
                For lIndice = 0 To ADATA_PEF_.Length - 1
                    SDATA_REGISTROS = ADATA_PEF_(lIndice)
                    ADATA_REGISTROS = Split(SDATA_REGISTROS, "|\t|", , CompareMethod.Text)  'Separar en columnas de array

                    sDataRecuperada = sDataRecuperada & "PRÉSTAMO EN EFECTIVO" & "|\t|" & ADATA_REGISTROS(0) & "|\t|" & ADATA_REGISTROS(2) & "|\n|"  'NRO CUENTA Y PRESTAMO SOLICITADO

                Next

                If sDataRecuperada.Trim.Length > 0 Then
                    sDataRecuperada = "PRODUCTO|\t|NRO. DE CUENTA|\t|PRÉSTAMO SOLICITADO|\n|" & Left(sDataRecuperada, sDataRecuperada.Trim.Length - 4)
                Else
                    sDataRecuperada = ""
                End If

                If sXML_DATA.Trim.Length > 0 Then
                    If sDataRecuperada.Trim.Length > 0 Then
                        sXML_DATA = sXML_DATA & "|\n|" & sDataRecuperada
                    End If
                End If

            Else
                sDataPEfectivo = ""
            End If


            'CUENTA DE AHORROS
            'CONSULTA_CUENTA_AHORROS
            Dim ADATA_AHO As Array
            Dim ADATA_AHO_ As Array
            Dim sDATA_AHO As String = ""
            Dim lIndice_aho As Long = 0
            Dim SDATA_REGISTROS_AHO As String = ""
            Dim ADATA_REGISTROS_AHO As Array
            Dim sDataRecuperada_AHO As String = ""
            Dim sDataAhorros As String = ""

            sDataAhorros = CONSULTA_CUENTA_AHORROS(sDNI.Trim, "")

            If Left(sDataAhorros.Trim.ToUpper, 5) <> "ERROR" Then
                'Sacar los datos del ahorro

                ADATA_AHO = Split(sDataPEfectivo, "*$¿*", , CompareMethod.Text)
                sDATA_AHO = ADATA_AHO(1)

                ADATA_AHO_ = Split(sDATA_AHO, "|\n|", , CompareMethod.Text)

                lIndice_aho = 0
                For lIndice_aho = 0 To ADATA_AHO_.Length - 1
                    SDATA_REGISTROS_AHO = ADATA_AHO_(lIndice_aho)
                    ADATA_REGISTROS_AHO = Split(SDATA_REGISTROS_AHO, "|\t|", , CompareMethod.Text)  'Separar en columnas de array

                    sDataRecuperada_AHO = sDataRecuperada_AHO & "CUENTA DE AHORRO" & "|\t|" & ADATA_REGISTROS_AHO(0) & "|\t|" & ADATA_REGISTROS_AHO(1) & "|\n|"  'NRO CUENTA Y SALDO

                Next

                If sDataRecuperada_AHO.Trim.Length > 0 Then
                    sDataRecuperada_AHO = "PRODUCTO|\t|NRO. DE CUENTA|\t|SALDO ACTUAL|\n|" & Left(sDataRecuperada_AHO, sDataRecuperada_AHO.Trim.Length - 4)
                Else
                    sDataRecuperada_AHO = ""
                End If

                If sXML_DATA.Trim.Length > 0 Then
                    If sDataRecuperada_AHO.Trim.Length > 0 Then
                        sXML_DATA = sXML_DATA & "|\n|" & sDataRecuperada_AHO
                    End If
                End If


            Else
                sDataRecuperada_AHO = ""
            End If

            'RIPLEY PUNTOS
            Dim sDataRipleyPuntos As String = "PRODUCTO|\t|NRO. DE TARJETA|\t|PUNTOS ACTIVOS|\n|"
            sDataRipleyPuntos = sDataRipleyPuntos & "RIPLEY PUNTOS|\t|" & sNroTarjeta.Trim & "|\t|" & FUN_PUNTOS_ACTIVOS(sNroTarjeta.Trim) 'PUNTOS ACTIVOS

            sXML_DATA = sXML_DATA & "|\n|" & sDataRipleyPuntos


            Dim sFechaKiosco As String = DateTime.Now.Day.ToString("00").Trim + "/" + DateTime.Now.Month.ToString("00").Trim + "/" + DateTime.Now.Year.ToString("0000").Trim
            Dim sHoraKiosco As String = DateTime.Now.ToString("HH:mm:ss")

            sXML_DATA = sFechaKiosco.Trim & "|\t|" & sHoraKiosco.Trim & "*$¿*" & sXML_DATA

            sRespuesta = ""
            sRespuesta = sXML_DATA.Trim

        Catch ex As Exception
            'save log error

            sRespuesta = "ERROR:" & ex.Message.Trim

        End Try


        'Save Monitor
        'Dim sDataMonitor As String = ""
        'Dim FECHA_TX As String = DateTime.Now.Year.ToString("0000").Trim & DateTime.Now.Month.ToString("00").Trim + DateTime.Now.Day.ToString("00").Trim
        'Dim HORA_TX As String = DateTime.Now.Hour.ToString("00").Trim & DateTime.Now.Minute.ToString("00").Trim & DateTime.Now.Second.ToString("00").Trim
        'Dim sNroCuentax As String = ""
        'Dim sNroTarjetax As String = ""
        'Dim sNombreCliente As String = ""
        'Dim sModoEntrada As String = ""
        'Dim sSaldoDisponible As String = ""
        'Dim sTotalDeuda As String = ""
        'Dim sIDSucursal As String = ""
        'Dim sIDTerminal As String = ""
        'Dim sCodigoTransaccion As String = ""

        'Dim sRespuestaServidor As String = ""
        'Dim sCanalAtencion As String = "01" 'RipleyMatico
        'Dim sEstadoCuenta As String = ""



        'If sDATA_MONITOR_KIOSCO.Length > 0 Then

        '    'RECORRER ARREGLO
        '    Dim ADATA_MONITOR As Array
        '    ADATA_MONITOR = Split(sDATA_MONITOR_KIOSCO, "|\t|", , CompareMethod.Text)


        '    sNroCuentax = ADATA_MONITOR(0)
        '    sNroTarjetax = ADATA_MONITOR(1)
        '    sNombreCliente = ADATA_MONITOR(2)
        '    sModoEntrada = ADATA_MONITOR(3)
        '    sSaldoDisponible = ADATA_MONITOR(4)
        '    sTotalDeuda = ADATA_MONITOR(5)
        '    sIDSucursal = ADATA_MONITOR(6)
        '    sIDTerminal = ADATA_MONITOR(7)
        '    sCodigoTransaccion = ADATA_MONITOR(8)


        '    'GUADAR LOG CONSULTAS
        '    'SAVE_LOG_CONSULTAS(sIDSucursal.Trim, sIDTerminal.Trim, "28", FUN_BUSCAR_TIPO_TARJETA(sNroTarjetax.Trim))

        '    sNroCuentax = Right("          " & sNroCuentax.Trim, 10)
        '    sNroTarjetax = Right(Strings.StrDup(20, " ") & sNroTarjetax.Trim, 20)

        '    sNombreCliente = Right(Strings.StrDup(26, " ") & sNombreCliente.Trim, 26)

        '    sSaldoDisponible = Right("      " & sSaldoDisponible.Trim, 6)
        '    sTotalDeuda = Right("      " & sTotalDeuda.Trim, 6)


        '    'armar variable desde flash con: numero de cuenta, numero de tarjeta, nombre del cliente, modo de entrada, canal de atencion, id de sucursal,codigo kiosco y codigo de transaccion (opcion)
        '    If Left(sRespuesta.Trim, 5) <> "ERROR" Then
        '        'VALIDACION CORRECTA

        '        sEstadoCuenta = "01" 'Estado de la cuenta
        '        sRespuestaServidor = "01" 'Atendido

        '        'ENVIAR_MONITOR
        '        sDataMonitor = ""
        '        sDataMonitor = FECHA_TX & HORA_TX & sNroCuentax & sNroTarjetax & sNombreCliente & sModoEntrada & sCanalAtencion & sSaldoDisponible & sTotalDeuda & sIDSucursal & sIDTerminal & sEstadoCuenta & sRespuestaServidor & sCodigoTransaccion

        '        'ENVIAR_MONITOR(sDataMonitor)

        '    Else
        '        'ENVIAR_MONITOR

        '        sEstadoCuenta = "01" 'Estado de la cuenta
        '        sRespuestaServidor = "02" 'Rechazado

        '        sDataMonitor = ""
        '        sDataMonitor = FECHA_TX & HORA_TX & sNroCuentax & sNroTarjetax & sNombreCliente & sModoEntrada & sCanalAtencion & sSaldoDisponible & sTotalDeuda & sIDSucursal & sIDTerminal & sEstadoCuenta & sRespuestaServidor & sCodigoTransaccion

        '        'ENVIAR_MONITOR(sDataMonitor)

        '    End If

        'End If


        Return sRespuesta

    End Function


    <WebMethod(Description:="Consulta de Deposito a Plazo.")> _
    Public Function CONSULTA_DEPOSITO_PLAZO(ByVal sDNI As String, ByVal sDATA_MONITOR_KIOSCO As String) As String


        Dim conConexion_Ora_DPF As OracleConnection

        Try
            conConexion_Ora_DPF = New OracleConnection(mCadenaConexion_ORA)

            If Not conConexion_Ora_DPF Is Nothing Then
                conConexion_Ora_DPF.Open()

                If conConexion_Ora_DPF.State = ConnectionState.Open Then

                    'LLAMAR AL PROCEDIMIENTO DEPOSITO A PLAZO
                    Dim cmd_ora_DPF As New OracleCommand
                    Dim param1 As New OracleParameter
                    Dim param2 As New OracleParameter
                    Dim param_curx As New OracleParameter


                    cmd_ora_DPF.Connection = conConexion_Ora_DPF
                    cmd_ora_DPF.CommandText = "RIPLEY_DA.RPM_PKG_CONSULTA.RPM_PRC_DEPOSITOS"
                    cmd_ora_DPF.CommandType = CommandType.StoredProcedure

                    param1.ParameterName = "DNI"
                    param1.OracleType = OracleType.VarChar
                    param1.Direction = ParameterDirection.Input
                    param1.Value = sDNI.Trim
                    cmd_ora_DPF.Parameters.Add(param1)


                    param2.ParameterName = "p_FECHA"
                    param2.OracleType = OracleType.DateTime
                    param2.Direction = ParameterDirection.Input

                    Dim vFecHoy As Date
                    Dim sFechaParam As String = ""

                    vFecHoy = Date.Now
                    sFechaParam = Right(Trim("00" & Day(vFecHoy).ToString), 2) & "/" & Right(Trim("00" & Month(vFecHoy).ToString), 2) & "/" & Right(Trim("0000" & Year(vFecHoy).ToString), 4)

                    param2.Value = sFechaParam.Trim
                    cmd_ora_DPF.Parameters.Add(param2)

                    param_curx.ParameterName = "V_CURSOR"
                    param_curx.OracleType = OracleType.Cursor
                    param_curx.Direction = ParameterDirection.Output
                    param_curx.Value = vbNull

                    cmd_ora_DPF.Parameters.Add(param_curx)

                    Dim rd_cur_DPF As OracleDataReader = cmd_ora_DPF.ExecuteReader

                    Dim sFechaKiosco As String = DateTime.Now.Day.ToString("00").Trim + "/" + DateTime.Now.Month.ToString("00").Trim + "/" + DateTime.Now.Year.ToString("0000").Trim
                    Dim sHoraKiosco As String = DateTime.Now.ToString("HH:mm:ss")
                    Dim sNOperacion As String = GET_NUMERO_OPERACION()

                    sResultado_ORA = ""
                    Do While rd_cur_DPF.Read

                        'sResultado_ORA = sResultado_ORA & sFechaKiosco.Trim & "|\t|" & sHoraKiosco.Trim & "|\t|" & sNOperacion.Trim & "|\t|"
                        sResultado_ORA = sResultado_ORA & IIf(IsDBNull(rd_cur_DPF.Item("PDA_CUENTA")), "", rd_cur_DPF.Item("PDA_CUENTA").ToString) & "|\t|" 'NUMERO DE CUENTA
                        sResultado_ORA = sResultado_ORA & IIf(IsDBNull(rd_cur_DPF.Item("FECHAPER")), "", rd_cur_DPF.Item("FECHAPER").ToString) & "|\t|" 'FECHA
                        sResultado_ORA = sResultado_ORA & IIf(IsDBNull(rd_cur_DPF.Item("MONTODEPINICIAL")), "0.00", Format(CDbl(rd_cur_DPF.Item("MONTODEPINICIAL").ToString.Trim), "##,##0.00")) & "|\t|" 'MONTO INICIAL DEPOSITADO
                        sResultado_ORA = sResultado_ORA & IIf(IsDBNull(rd_cur_DPF.Item("SALDOANTERIOR")), "0.00", Format(CDbl(rd_cur_DPF.Item("SALDOANTERIOR").ToString.Trim), "##,##0.00")) & "|\t|" 'SALDO ANTERIOR
                        sResultado_ORA = sResultado_ORA & IIf(IsDBNull(rd_cur_DPF.Item("INTERESALVENC")), "0.00", Format(CDbl(rd_cur_DPF.Item("INTERESALVENC").ToString.Trim), "##,##0.00")) & "|\t|" 'INTERES AL VENCIMIENTO
                        sResultado_ORA = sResultado_ORA & IIf(IsDBNull(rd_cur_DPF.Item("TOTALDEPACT")), "0.00", Format(CDbl(rd_cur_DPF.Item("TOTALDEPACT").ToString.Trim), "##,##0.00")) & "|\n|" 'MONTO A COBRAR

                    Loop


                    If sResultado_ORA.Trim.Length > 0 Then
                        sResultado_ORA = sFechaKiosco.Trim & "|\t|" & sHoraKiosco.Trim & "|\t|" & sNOperacion.Trim & "*$¿*" & Left(sResultado_ORA, sResultado_ORA.Trim.Length - 4)
                    Else
                        sResultado_ORA = "ERROR:NODATA"
                    End If


                    rd_cur_DPF.Close()
                    rd_cur_DPF = Nothing
                    cmd_ora_DPF.Dispose()
                    cmd_ora_DPF = Nothing
                    conConexion_Ora_DPF.Close()
                    conConexion_Ora_DPF = Nothing

                End If

            Else

                sResultado_ORA = "ERROR:No se pudo conectar al servidor de base de datos Oracle."

            End If

        Catch ex As Exception
            sResultado_ORA = "ERROR:" & ex.Message.Trim
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
            'SAVE_LOG_CONSULTAS(sIDSucursal.Trim, sIDTerminal.Trim, "40", FUN_BUSCAR_TIPO_TARJETA(sNroTarjetax.Trim))

            sNroCuentax = Right("          " & sNroCuentax.Trim, 10)
            sNroTarjetax = Right(Strings.StrDup(20, " ") & sNroTarjetax.Trim, 20)

            sNombreCliente = Right(Strings.StrDup(26, " ") & sNombreCliente.Trim, 26)

            sSaldoDisponible = Right("      " & sSaldoDisponible.Trim, 6)
            sTotalDeuda = Right("      " & sTotalDeuda.Trim, 6)


            'armar variable desde flash con: numero de cuenta, numero de tarjeta, nombre del cliente, modo de entrada, canal de atencion, id de sucursal,codigo kiosco y codigo de transaccion (opcion)
            If Left(sResultado_ORA.Trim, 5) <> "ERROR" Then
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


        Return sResultado_ORA


    End Function


    <WebMethod(Description:="Consulta de Cuentas de Ahorros.")> _
    Public Function CONSULTA_CUENTA_AHORROS(ByVal sDNI As String, ByVal sDATA_MONITOR_KIOSCO As String) As String
        Dim conConexion_Ora As OracleConnection


        Try
            conConexion_Ora = New OracleConnection(mCadenaConexion_ORA)

            If Not conConexion_Ora Is Nothing Then
                conConexion_Ora.Open()

                If conConexion_Ora.State = ConnectionState.Open Then

                    'LLAMAR AL PROCEDIMIENTO CUENTAS DE AHORRO
                    Dim cmd_ora As New OracleCommand
                    Dim param1 As New OracleParameter
                    Dim param2 As New OracleParameter
                    Dim param_cur As New OracleParameter


                    cmd_ora.Connection = conConexion_Ora
                    cmd_ora.CommandText = "RIPLEY_DA.RPM_PKG_CONSULTA.RPM_PRC_AHORRO"
                    cmd_ora.CommandType = CommandType.StoredProcedure


                    param1.ParameterName = "DNI"
                    param1.OracleType = OracleType.VarChar
                    param1.Direction = ParameterDirection.Input
                    param1.Value = sDNI.Trim
                    cmd_ora.Parameters.Add(param1)


                    param2.ParameterName = "p_FECHA"
                    param2.OracleType = OracleType.DateTime
                    param2.Direction = ParameterDirection.Input

                    Dim vFecHoy As Date
                    Dim sFechaParam As String = ""

                    vFecHoy = Date.Now
                    sFechaParam = Right(Trim("00" & Day(vFecHoy).ToString), 2) & "/" & Right(Trim("00" & Month(vFecHoy).ToString), 2) & "/" & Right(Trim("0000" & Year(vFecHoy).ToString), 4)

                    param2.Value = sFechaParam.Trim
                    cmd_ora.Parameters.Add(param2)


                    param_cur.ParameterName = "V_CURSOR"
                    param_cur.OracleType = OracleType.Cursor
                    param_cur.Direction = ParameterDirection.Output
                    param_cur.Value = vbNull

                    cmd_ora.Parameters.Add(param_cur)

                    Dim rd_cur As OracleDataReader = cmd_ora.ExecuteReader

                    Dim sFechaKiosco As String = DateTime.Now.Day.ToString("00").Trim + "/" + DateTime.Now.Month.ToString("00").Trim + "/" + DateTime.Now.Year.ToString("0000").Trim
                    Dim sHoraKiosco As String = DateTime.Now.ToString("HH:mm:ss")
                    Dim sNOperacion As String = GET_NUMERO_OPERACION()

                    sResultado_ORA = ""
                    Do While rd_cur.Read

                        sResultado_ORA = sResultado_ORA & IIf(IsDBNull(rd_cur.Item("NUMCTA")), "", rd_cur.Item("NUMCTA").ToString) & "|\t|" 'NUMERO DE CUENTA
                        sResultado_ORA = sResultado_ORA & IIf(IsDBNull(rd_cur.Item("SALDOTOT")), "0.00", Format(CDbl(rd_cur.Item("SALDOTOT").ToString.Trim), "##,##0.00")) & "|\n|" 'SALDO ACTUAL

                    Loop

                    If sResultado_ORA.Trim.Length > 0 Then
                        sResultado_ORA = sFechaKiosco.Trim & "|\t|" & sHoraKiosco.Trim & "|\t|" & sNOperacion.Trim & "*$¿*" & Left(sResultado_ORA, sResultado_ORA.Trim.Length - 4)
                    Else
                        sResultado_ORA = "ERROR:NODATA"
                    End If


                    rd_cur.Close()
                    rd_cur = Nothing
                    cmd_ora.Dispose()
                    cmd_ora = Nothing
                    conConexion_Ora.Close()
                    conConexion_Ora = Nothing

                End If

            Else

                sResultado_ORA = "ERROR:No se pudo conectar al servidor de base de datos Oracle."

            End If


        Catch ex As Exception
            sResultado_ORA = "ERROR:" & ex.Message.Trim
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
            'SAVE_LOG_CONSULTAS(sIDSucursal.Trim, sIDTerminal.Trim, "42", FUN_BUSCAR_TIPO_TARJETA(sNroTarjetax.Trim))

            sNroCuentax = Right("          " & sNroCuentax.Trim, 10)
            sNroTarjetax = Right(Strings.StrDup(20, " ") & sNroTarjetax.Trim, 20)

            sNombreCliente = Right(Strings.StrDup(26, " ") & sNombreCliente.Trim, 26)

            sSaldoDisponible = Right("      " & sSaldoDisponible.Trim, 6)
            sTotalDeuda = Right("      " & sTotalDeuda.Trim, 6)


            'armar variable desde flash con: numero de cuenta, numero de tarjeta, nombre del cliente, modo de entrada, canal de atencion, id de sucursal,codigo kiosco y codigo de transaccion (opcion)
            If Left(sResultado_ORA.Trim, 5) <> "ERROR" Then
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


        Return sResultado_ORA


    End Function



    <WebMethod(Description:="Consulta de Prestamo en Efectivo.")> _
    Public Function CONSULTA_PRESTAMO_EFECTIVO(ByVal sDNI As String, ByVal sDATA_MONITOR_KIOSCO As String) As String

        Dim conConexion_Ora As OracleConnection

        Try
            conConexion_Ora = New OracleConnection(mCadenaConexion_ORA)

            If Not conConexion_Ora Is Nothing Then
                conConexion_Ora.Open()

                If conConexion_Ora.State = ConnectionState.Open Then

                    'LLAMAR AL PROCEDIMIENTO PRESTAMO EN EFECTIVO
                    Dim cmd_ora As New OracleCommand
                    Dim param1 As New OracleParameter
                    Dim param2 As New OracleParameter
                    Dim param_cur As New OracleParameter

                    cmd_ora.Connection = conConexion_Ora
                    cmd_ora.CommandText = "RIPLEY_DA.RPM_PKG_CONSULTA.RPM_PRC_PRESTAMOS"
                    cmd_ora.CommandType = CommandType.StoredProcedure

                    param1.ParameterName = "DNI"
                    param1.OracleType = OracleType.VarChar
                    param1.Direction = ParameterDirection.Input
                    param1.Value = sDNI.Trim
                    cmd_ora.Parameters.Add(param1)

                    param2.ParameterName = "p_FECHA"
                    param2.OracleType = OracleType.DateTime
                    param2.Direction = ParameterDirection.Input

                    Dim vFecHoy As Date
                    Dim sFechaParam As String = ""

                    vFecHoy = Date.Now
                    sFechaParam = Right(Trim("00" & Day(vFecHoy).ToString), 2) & "/" & Right(Trim("00" & Month(vFecHoy).ToString), 2) & "/" & Right(Trim("0000" & Year(vFecHoy).ToString), 4)

                    param2.Value = sFechaParam.Trim
                    cmd_ora.Parameters.Add(param2)

                    param_cur.ParameterName = "V_CURSOR"
                    param_cur.OracleType = OracleType.Cursor
                    param_cur.Direction = ParameterDirection.Output
                    param_cur.Value = vbNull

                    cmd_ora.Parameters.Add(param_cur)

                    Dim rd_cur As OracleDataReader = cmd_ora.ExecuteReader

                    Dim sFechaKiosco As String = DateTime.Now.Day.ToString("00").Trim + "/" + DateTime.Now.Month.ToString("00").Trim + "/" + DateTime.Now.Year.ToString("0000").Trim
                    Dim sHoraKiosco As String = DateTime.Now.ToString("HH:mm:ss")
                    Dim sNOperacion As String = GET_NUMERO_OPERACION()


                    sResultado_ORA = ""
                    Do While rd_cur.Read

                        sResultado_ORA = sResultado_ORA & IIf(IsDBNull(rd_cur.Item("PRE_CREDITO")), "", rd_cur.Item("PRE_CREDITO").ToString) & "|\t|" 'NUMERO DE CUENTA
                        sResultado_ORA = sResultado_ORA & IIf(IsDBNull(rd_cur.Item("FECHAPER")), "", rd_cur.Item("FECHAPER").ToString) & "|\t|" 'FECHA  DAR FORMATO A DD/MM/YYYY
                        sResultado_ORA = sResultado_ORA & IIf(IsDBNull(rd_cur.Item("MONTOPRESTAMO")), "0.00", Format(CDbl(rd_cur.Item("MONTOPRESTAMO").ToString.Trim), "##,##0.00")) & "|\t|" 'PRESTAMO SOLICITADO
                        sResultado_ORA = sResultado_ORA & IIf(IsDBNull(rd_cur.Item("FECHAPROXCUOTA")), "", rd_cur.Item("FECHAPROXCUOTA").ToString.Trim) & "|\t|" 'FECHA A PAGAR PROXIMA CUOTA
                        sResultado_ORA = sResultado_ORA & IIf(IsDBNull(rd_cur.Item("CUOTAAPAGAR")), "0.00", Format(CDbl(rd_cur.Item("CUOTAAPAGAR").ToString.Trim), "##,##0.00")) & "|\n|" 'MONTO A PAGAR

                    Loop


                    If sResultado_ORA.Trim.Length > 0 Then
                        sResultado_ORA = sFechaKiosco.Trim & "|\t|" & sHoraKiosco.Trim & "|\t|" & sNOperacion.Trim & "*$¿*" & Left(sResultado_ORA, sResultado_ORA.Trim.Length - 4)
                    Else
                        sResultado_ORA = "ERROR:NODATA"
                    End If


                    rd_cur.Close()
                    rd_cur = Nothing
                    cmd_ora.Dispose()
                    cmd_ora = Nothing
                    conConexion_Ora.Close()
                    conConexion_Ora = Nothing

                End If

            Else

                sResultado_ORA = "ERROR:No se pudo conectar al servidor de base de datos Oracle."

            End If

        Catch ex As Exception
            sResultado_ORA = "ERROR:" & ex.Message.Trim
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
            'SAVE_LOG_CONSULTAS(sIDSucursal.Trim, sIDTerminal.Trim, "36", FUN_BUSCAR_TIPO_TARJETA(sNroTarjetax.Trim))

            sNroCuentax = Right("          " & sNroCuentax.Trim, 10)
            sNroTarjetax = Right(Strings.StrDup(20, " ") & sNroTarjetax.Trim, 20)

            sNombreCliente = Right(Strings.StrDup(26, " ") & sNombreCliente.Trim, 26)


            sSaldoDisponible = Right("      " & sSaldoDisponible.Trim, 6)
            sTotalDeuda = Right("      " & sTotalDeuda.Trim, 6)


            'armar variable desde flash con: numero de cuenta, numero de tarjeta, nombre del cliente, modo de entrada, canal de atencion, id de sucursal,codigo kiosco y codigo de transaccion (opcion)
            If Left(sResultado_ORA.Trim, 5) <> "ERROR" Then
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


        Return sResultado_ORA

    End Function


    'CONSULTA DE PUNTOS
    <WebMethod(Description:="Consulta de Ripley Puntos.")> _
    Public Function CONSULTA_RIPLEY_PUNTOS(ByVal sTipoTarjeta As String, ByVal sNroTarjeta As String, ByVal sNRO_CUENTA As String, ByVal sDATA_MONITOR_KIOSCO As String, ByVal Servidor As TServidor) As String
        'C = CLASICA
        'A = ASOCIADA
        Dim conConexion_ripley As OracleConnection

        Try
            conConexion_ripley = New OracleConnection(mCadenaConexion_ORA_PUNTOS)
            If Not conConexion_ripley Is Nothing Then
                conConexion_ripley.Open()
                If conConexion_ripley.State = ConnectionState.Open Then
                    'LLAMAR AL PROCEDIMIENTO RIPEY PUNTOS
                    Dim cmd_ora_ As New OracleCommand
                    Dim param1_ As New OracleParameter
                    Dim param2_ As New OracleParameter
                    Dim param3_ As New OracleParameter
                    Dim param4_ As New OracleParameter
                    Dim param5_ As New OracleParameter
                    Dim param6_ As New OracleParameter
                    Dim param7_ As New OracleParameter
                    Dim param8_ As New OracleParameter
                    Dim sFecha_Pago As String = ""

                    'VARIABLES DE SALIDA
                    Dim sPuntosGanados As String = ""
                    Dim sPuntosCanjeados As String = ""
                    Dim sPuntosVencidos As String = ""
                    Dim sPuntosXvencer As String = ""

                    cmd_ora_.Connection = conConexion_ripley
                    cmd_ora_.CommandText = "PKG_RIPLEY_PUNTOS_EPUS.PRC_PUNTOS_RIPLEYMATICO"
                    cmd_ora_.CommandType = CommandType.StoredProcedure

                    param1_.ParameterName = "pTarjeta"
                    param1_.OracleType = OracleType.VarChar
                    param1_.Size = 16
                    param1_.Direction = ParameterDirection.Input
                    param1_.Value = sNroTarjeta.Trim
                    cmd_ora_.Parameters.Add(param1_)

                    'OBTENER FECHA FINAL DE FACTURACION
                    If sTipoTarjeta.Trim.ToUpper = "C" Then
                        Select Case Servidor
                            Case TServidor.SICRON
                                sFecha_Pago = FUN_BUSCAR_FECHA_CORTE_CLASICA_SICRON(sNRO_CUENTA.Trim)
                            Case TServidor.RSAT
                                sFecha_Pago = FUN_BUSCAR_FECHA_CORTE_CLASICA_RSAT(sNroTarjeta.Trim, sNRO_CUENTA.Trim)
                        End Select
                    Else
                        sFecha_Pago = FUN_BUSCAR_FECHA_CORTE_ASOCIADA(sNroTarjeta.Trim)
                    End If
                    Log.ErrorLog("Function CONSULTA_RIPLEY_PUNTOS: sFecha_Pago = " & sFecha_Pago)

                    If sFecha_Pago.Trim.Length > 0 Then
                        param2_.ParameterName = "vFechaCorte"
                        param2_.OracleType = OracleType.DateTime
                        param2_.Direction = ParameterDirection.Input
                        param2_.Value = sFecha_Pago.Trim
                        cmd_ora_.Parameters.Add(param2_)

                        param3_.ParameterName = "nPtosCons"
                        param3_.OracleType = OracleType.Double
                        param3_.Direction = ParameterDirection.Output
                        cmd_ora_.Parameters.Add(param3_)

                        param4_.ParameterName = "nPtosAcum"
                        param4_.OracleType = OracleType.Double
                        param4_.Direction = ParameterDirection.Output
                        cmd_ora_.Parameters.Add(param4_)

                        param5_.ParameterName = "nPtoVcdo"
                        param5_.OracleType = OracleType.Double
                        param5_.Direction = ParameterDirection.Output
                        cmd_ora_.Parameters.Add(param5_)

                        param6_.ParameterName = "nPtosSaldoAnt"
                        param6_.OracleType = OracleType.Double
                        param6_.Direction = ParameterDirection.Output
                        cmd_ora_.Parameters.Add(param6_)

                        param7_.ParameterName = "nPtosxVencer"
                        param7_.OracleType = OracleType.Double
                        param7_.Direction = ParameterDirection.Output
                        cmd_ora_.Parameters.Add(param7_)

                        param8_.ParameterName = "vFechxVencer"
                        param8_.OracleType = OracleType.VarChar
                        param8_.Size = 16
                        param8_.Direction = ParameterDirection.Output
                        cmd_ora_.Parameters.Add(param8_)

                        cmd_ora_.ExecuteNonQuery()

                        sPuntosGanados = cmd_ora_.Parameters("nPtosAcum").Value.ToString
                        sPuntosCanjeados = cmd_ora_.Parameters("nPtosCons").Value.ToString
                        sPuntosVencidos = cmd_ora_.Parameters("nPtoVcdo").Value.ToString
                        sPuntosXvencer = cmd_ora_.Parameters("nPtosxVencer").Value.ToString
                        Log.ErrorLog("Function CONSULTA_RIPLEY_PUNTOS: sPuntosGanados = " & sPuntosGanados)
                        Log.ErrorLog("Function CONSULTA_RIPLEY_PUNTOS: sPuntosCanjeados = " & sPuntosCanjeados)
                        Log.ErrorLog("Function CONSULTA_RIPLEY_PUNTOS: sPuntosVencidos = " & sPuntosVencidos)
                        Log.ErrorLog("Function CONSULTA_RIPLEY_PUNTOS: sPuntosXvencer = " & sPuntosXvencer)
                        Log.ErrorLog("Function CONSULTA_RIPLEY_PUNTOS: nPtosSaldoAnt = " & cmd_ora_.Parameters("nPtosSaldoAnt").Value.ToString)
                        Log.ErrorLog("Function CONSULTA_RIPLEY_PUNTOS: vFechxVencer = " & cmd_ora_.Parameters("vFechxVencer").Value.ToString)

                        sResultado_ORA = ""
                        sResultado_ORA = sPuntosGanados.Trim & "|\t|" & sPuntosCanjeados.Trim & "|\t|" & sPuntosVencidos.Trim & "|\t|" & sPuntosXvencer.Trim & "|\t|" & FUN_PUNTOS_ACTIVOS(sNroTarjeta.Trim) 'PUNTOS ACTIVOS

                        Dim sFechaKiosco As String = DateTime.Now.Day.ToString("00").Trim + "/" + DateTime.Now.Month.ToString("00").Trim + "/" + DateTime.Now.Year.ToString("0000").Trim
                        Dim sHoraKiosco As String = DateTime.Now.ToString("HH:mm:ss")
                        Dim sNOperacion As String = GET_NUMERO_OPERACION()
                        sResultado_ORA = sFechaKiosco.Trim & "|\t|" & sHoraKiosco.Trim & "|\t|" & sNOperacion.Trim & "|\t|" & sResultado_ORA.Trim
                    End If

                    cmd_ora_.Dispose()
                    cmd_ora_ = Nothing
                    conConexion_ripley.Close()
                    conConexion_ripley = Nothing
                End If
            Else
                sResultado_ORA = "ERROR:NODATA"
            End If
        Catch ex As Exception
            Log.ErrorLog("Function CONSULTA_RIPLEY_PUNTOS: Exception = " & ex.Message)
            sResultado_ORA = "ERROR:NODATA"
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
            SAVE_LOG_CONSULTAS(sIDSucursal.Trim, sIDTerminal.Trim, "46", FUN_BUSCAR_TIPO_TARJETA(sNroTarjetax.Trim))

            sNroCuentax = Right("          " & sNroCuentax.Trim, 10)
            sNroTarjetax = Right(Strings.StrDup(20, " ") & sNroTarjetax.Trim, 20)
            sNombreCliente = Right(Strings.StrDup(26, " ") & sNombreCliente.Trim, 26)

            sSaldoDisponible = Right("      " & sSaldoDisponible.Trim, 6)
            sTotalDeuda = Right("      " & sTotalDeuda.Trim, 6)


            'armar variable desde flash con: numero de cuenta, numero de tarjeta, nombre del cliente, modo de entrada, canal de atencion, id de sucursal,codigo kiosco y codigo de transaccion (opcion)
            If Left(sResultado_ORA.Trim, 5) <> "ERROR" Then
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

        Log.ErrorLog(sResultado_ORA)
        Return sResultado_ORA
    End Function


    'FUNCION QUE DEVUELVE EL TOTAL DE PUNTOS ACTIVOS ESTA FUNCION ES PARA LA CONSULTA DE PUNTOS
    <WebMethod(Description:="Consulta de Puntos Activos")> _
    Public Function FUN_PUNTOS_ACTIVOS(ByVal sNroTarjeta As String) As String

        Dim conConexion_puntos As OracleConnection


        Try
            conConexion_puntos = New OracleConnection(mCadenaConexion_ORA_PUNTOS)

            If Not conConexion_puntos Is Nothing Then
                conConexion_puntos.Open()

                If conConexion_puntos.State = ConnectionState.Open Then

                    'LLAMAR AL PROCEDIMIENTO PRESTAMO EN EFECTIVO
                    Dim cmd_ora_ptos As New OracleCommand
                    Dim param1__ As New OracleParameter
                    Dim param2__ As New OracleParameter
                    Dim sTotalPuntosActivos As String = ""

                    cmd_ora_ptos.Connection = conConexion_puntos
                    cmd_ora_ptos.CommandText = "FEFID_SP_CONSULTA_PTOS_RIPLEY"
                    cmd_ora_ptos.CommandType = CommandType.StoredProcedure

                    param1__.ParameterName = "echr_num_tar"

                    param1__.OracleType = OracleType.Char
                    param1__.Size = 16
                    param1__.Direction = ParameterDirection.Input
                    param1__.Value = sNroTarjeta.Trim
                    cmd_ora_ptos.Parameters.Add(param1__)


                    param2__.ParameterName = "onum_ptos_disp"
                    param2__.OracleType = OracleType.Double
                    param2__.Direction = ParameterDirection.Output
                    cmd_ora_ptos.Parameters.Add(param2__)


                    cmd_ora_ptos.ExecuteNonQuery()

                    sTotalPuntosActivos = cmd_ora_ptos.Parameters("onum_ptos_disp").Value.ToString

                    sResultado_ORA = ""
                    sResultado_ORA = sTotalPuntosActivos.Trim  'PUNTOS GANADOS


                    cmd_ora_ptos.Dispose()
                    cmd_ora_ptos = Nothing

                End If

                conConexion_puntos.Close()
                conConexion_puntos = Nothing

            Else

                sResultado_ORA = "0"

            End If

        Catch ex As Exception
            sResultado_ORA = "0" 'ex.Message.Trim
        End Try


        Return sResultado_ORA


    End Function

    'CONSULTA SUPER EFECTIVO
    <WebMethod(Description:="Consulta de SuperEfectivo.")> _
    Public Function CONSULTA_SUPER_EFECTIVO(ByVal sDATA_MONITOR_KIOSCO As String) As String

        'DATOS DE PRUEBA
        Dim sFechaKiosco As String = DateTime.Now.Day.ToString("00").Trim + "/" + DateTime.Now.Month.ToString("00").Trim + "/" + DateTime.Now.Year.ToString("0000").Trim
        Dim sHoraKiosco As String = DateTime.Now.ToString("HH:mm:ss")
        Dim sNOperacion As String = GET_NUMERO_OPERACION()

        sResultado_ORA = ""

        sResultado_ORA = sFechaKiosco.Trim & "|\t|" & sHoraKiosco.Trim & "|\t|" & sNOperacion.Trim


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
            'SAVE_LOG_CONSULTAS(sIDSucursal.Trim, sIDTerminal.Trim, "38", FUN_BUSCAR_TIPO_TARJETA(sNroTarjetax.Trim))

            sNroCuentax = Right("          " & sNroCuentax.Trim, 10)
            sNroTarjetax = Right(Strings.StrDup(20, " ") & sNroTarjetax.Trim, 20)

            sNombreCliente = Right(Strings.StrDup(26, " ") & sNombreCliente.Trim, 26)

            sSaldoDisponible = Right("      " & sSaldoDisponible.Trim, 6)
            sTotalDeuda = Right("      " & sTotalDeuda.Trim, 6)


            'armar variable desde flash con: numero de cuenta, numero de tarjeta, nombre del cliente, modo de entrada, canal de atencion, id de sucursal,codigo kiosco y codigo de transaccion (opcion)
            If Left(sResultado_ORA.Trim, 5) <> "ERROR" Then
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

        Return sResultado_ORA

    End Function

    '04-06-2015 Nuevo formato de tramas
    'DEVUELVE DETALLE DE EECC
    '<WebMethod(Description:="Devuelve el detalle del EECC Clasica")> _    
    Private Function FUN_DETALLE_EECC_TEA(ByVal sDataDetalle As String, ByVal sPeriodoFinal As String, ByVal Servidor As TServidor) As String
        Dim sResult As String = ""
        Dim SDATA_MOVIMIENTOS As String = ""
        Dim ADATA_MOVIMIENTO As Array
        Dim lIndice As Long = 0
        Dim sRegistroAux As String = ""
        Dim sRegistro As String = ""
        Dim sMOV As String = ""

        'CAMPOS MOVIMIENTOS
        Dim sFechaConsumo As String = ""
        Dim sFechaProceso As String = ""
        Dim sNroTicket As String = ""
        Dim sDescripcion As String = ""
        Dim sTA As String = ""
        Dim sMonto As String = ""
        Dim sTEA As String = ""
        Dim sNroCuotas As String = ""
        Dim sCapital As String = ""
        Dim sInteres As String = ""
        Dim sTotal As String = ""

        SDATA_MOVIMIENTOS = sDataDetalle
        If SDATA_MOVIMIENTOS.Trim.Length > 0 Then 'IF HAY REGISTROS
            ADATA_MOVIMIENTO = Split(SDATA_MOVIMIENTOS, "|\n|", , CompareMethod.Text)

            For lIndice = 0 To ADATA_MOVIMIENTO.Length - 1
                sRegistro = ADATA_MOVIMIENTO(lIndice)

                sFechaConsumo = Mid(sRegistro, 1, 11)
                sFechaProceso = Mid(sRegistro, 12, 11)
                sNroTicket = Mid(sRegistro, 23, 6)

                If Servidor = TServidor.RSAT And sPeriodoFinal >= ReadAppConfig("PeriodoAgrandarGlosa") Then
                    sDescripcion = Mid(sRegistro, 29, 85)
                    sTA = Mid(sRegistro, 114, 1)
                    sMonto = Mid(sRegistro, 115, 9)
                    sTEA = Mid(sRegistro, 124, 7)
                    sNroCuotas = Mid(sRegistro, 131, 5)
                    sCapital = Mid(sRegistro, 136, 8)
                    sInteres = Mid(sRegistro, 144, 7)
                    sTotal = Mid(sRegistro, 151, 9)
                Else
                    sDescripcion = Mid(sRegistro, 29, 20)
                    sTA = Mid(sRegistro, 49, 1)
                    sMonto = Mid(sRegistro, 50, 9)

                    If Servidor = TServidor.RSAT And sPeriodoFinal >= Constantes.PeriodoInclusionTEA Then
                        sTEA = Mid(sRegistro, 59, 7)
                        sNroCuotas = Mid(sRegistro, 66, 5)
                        sCapital = Mid(sRegistro, 71, 8)
                        sInteres = Mid(sRegistro, 79, 8)
                        sTotal = Mid(sRegistro, 87, 8)
                    Else
                        sTEA = "       "
                        sNroCuotas = Mid(sRegistro, 59, 5)
                        sCapital = Mid(sRegistro, 64, 8)
                        sInteres = Mid(sRegistro, 72, 8)
                        sTotal = Mid(sRegistro, 80, 8)
                    End If
                End If

                sRegistroAux = ""
                sRegistroAux = sFechaConsumo.Trim & "|\t|" & sFechaProceso.Trim & "|\t|" & sNroTicket.Trim & "|\t|" & sDescripcion.Trim & "|\t|" & sTA.Trim & "|\t|"
                sRegistroAux = sRegistroAux & sMonto.Trim & "|\t|" & sTEA.Trim & "|\t|" & sNroCuotas.Trim & "|\t|" & sCapital.Trim & "|\t|" & sInteres.Trim & "|\t|"
                sRegistroAux = sRegistroAux & sTotal.Trim & "|\n|"
                sMOV = sMOV & sRegistroAux.Trim
            Next

            If sMOV.Trim.Length > 0 Then
                sMOV = Microsoft.VisualBasic.Strings.Left(sMOV, sMOV.Length - 4)
            Else
                sMOV = "" 'NO HAY MOVIMIENTOS
            End If
            sResult = sMOV
        End If
        Return sResult.Trim
    End Function

    'DEVUELVE DETALLE DE EECC
    '<WebMethod(Description:="Devuelve el detalle del EECC Clasica")> _
    Private Function FUN_DETALLE_EECC(ByVal sDataDetalle As String) As String
        Dim sResult As String = ""
        Dim SDATA_MOVIMIENTOS As String = ""
        Dim ADATA_MOVIMIENTO As Array
        Dim lIndice As Long = 0
        Dim sRegistroAux As String = ""
        Dim sRegistro As String = ""
        Dim sMOV As String = ""

        'CAMPOS MOVIMIENTOS
        Dim sFechaConsumo As String = ""
        Dim sFechaProceso As String = ""
        Dim sNroTicket As String = ""
        Dim sDescripcion As String = ""
        Dim sTA As String = ""
        Dim sMonto As String = ""
        Dim sTEA As String = ""
        Dim sNroCuotas As String = ""
        Dim sCapital As String = ""
        Dim sInteres As String = ""
        Dim sTotal As String = ""

        SDATA_MOVIMIENTOS = sDataDetalle
        If SDATA_MOVIMIENTOS.Trim.Length > 0 Then 'IF HAY REGISTROS
            ADATA_MOVIMIENTO = Split(SDATA_MOVIMIENTOS, "|\n|", , CompareMethod.Text)

            For lIndice = 0 To ADATA_MOVIMIENTO.Length - 1
                sRegistro = ""
                sRegistro = ADATA_MOVIMIENTO(lIndice)

                sFechaConsumo = Mid(sRegistro, 1, 11)
                sFechaProceso = Mid(sRegistro, 12, 11)
                sNroTicket = Mid(sRegistro, 23, 6)
                sDescripcion = Mid(sRegistro, 29, 20)
                sTA = Mid(sRegistro, 49, 1)
                sMonto = Mid(sRegistro, 50, 9)
                sTEA = Mid(sRegistro, 59, 7)
                sNroCuotas = Mid(sRegistro, 66, 5)
                sCapital = Mid(sRegistro, 71, 8)
                sInteres = Mid(sRegistro, 79, 8)
                sTotal = Mid(sRegistro, 87, 8)

                sRegistroAux = ""
                sRegistroAux = sFechaConsumo.Trim & "|\t|" & sFechaProceso.Trim & "|\t|" & sNroTicket.Trim & "|\t|" & sDescripcion.Trim & "|\t|" & sTA.Trim & "|\t|"
                sRegistroAux = sRegistroAux & sMonto.Trim & "|\t|" & sTEA.Trim & "|\t|" & sNroCuotas.Trim & "|\t|" & sCapital.Trim & "|\t|" & sInteres.Trim & "|\t|"
                sRegistroAux = sRegistroAux & sTotal.Trim & "|\n|"

                sMOV = sMOV & sRegistroAux.Trim
            Next

            If sMOV.Trim.Length > 0 Then
                sMOV = Left(sMOV, sMOV.Length - 4)
            Else
                sMOV = "" 'NO HAY MOVIMIENTOS
            End If
            sResult = sMOV
        End If
        Return sResult.Trim
    End Function


    Public Function pfnFecFactDDMMYYYY(ByVal estrFecProcYYYYMMDD As String, _
                                        ByVal eintDiaVenc As Integer) As String
        Dim ldatFecProc As Date
        Dim ldatFecVenc As Date
        Dim ldif As Long = 0
        Dim sResultFecha As String = ""

        Try


            ldatFecProc = DateSerial(Mid(estrFecProcYYYYMMDD, 1, 4), _
                                     Mid(estrFecProcYYYYMMDD, 5, 2), _
                                     Mid(estrFecProcYYYYMMDD, 7, 2))
            ldatFecVenc = DateSerial(Mid(estrFecProcYYYYMMDD, 1, 4), _
                                     Mid(estrFecProcYYYYMMDD, 5, 2), _
                                     eintDiaVenc)

            If ldatFecVenc >= ldatFecProc Then
                ldif = DateDiff(DateInterval.Day, ldatFecVenc, ldatFecProc)
                If ldif >= 15 Then
                    ldatFecVenc = DateAdd("m", -1, ldatFecVenc)
                End If

            Else

                ldatFecVenc = DateAdd("m", 1, ldatFecVenc)
                ldif = DateDiff(DateInterval.Day, ldatFecProc, ldatFecVenc)
                If ldif >= 15 Then
                    ldatFecVenc = DateAdd("m", -1, ldatFecVenc)
                End If

            End If
            sResultFecha = Right("00" & Day(ldatFecVenc), 2) & "/" _
                                 & Right("00" & Month(ldatFecVenc), 2) & "/" _
                                 & Right("0000" & Year(ldatFecVenc), 4)
        Catch ex As Exception

        End Try

        Return sResultFecha

    End Function


    Private Function FormatearFecha(ByVal sFecha As String) As String
        Dim sResul As String = ""
        Dim sDia As String = ""
        Dim sMes As String = ""
        Dim sAnio As String = ""

        If sFecha.Trim.Length > 0 Then
            sDia = Left(sFecha.Trim, 2)
            sMes = Mid(sFecha.Trim, 3, 2)
            sAnio = Right(sFecha.Trim, 4)

            sResul = sDia.Trim & "/" & sMes.Trim & "/" & sAnio.Trim

        End If


        Return sResul.Trim

    End Function


    Private Function FormatearFecha_DDMMYYYY(ByVal sFecha As String) As String
        Dim sResul As String = ""
        Dim sDia As String = ""
        Dim sMes As String = ""
        Dim sAnio As String = ""
        '20111005
        If sFecha.Trim.Length > 0 Then
            sDia = Right(sFecha.Trim, 2)
            sMes = Mid(sFecha.Trim, 5, 2)
            sAnio = Left(sFecha.Trim, 4)

            sResul = sDia.Trim & "/" & sMes.Trim & "/" & sAnio.Trim

        End If

        Return sResul.Trim

    End Function


    Private Function Cambia_Fecha(ByVal p_Fecha As String) As String
        Dim r_Result As String
        '01122011
        r_Result = Mid(p_Fecha, 1, 2) & "/"

        Select Case Mid(p_Fecha, 3, 2)
            Case "01"
                r_Result = r_Result & "ENE/"
            Case "02"
                r_Result = r_Result & "FEB/"
            Case "03"
                r_Result = r_Result & "MAR/"
            Case "04"
                r_Result = r_Result & "ABR/"
            Case "05"
                r_Result = r_Result & "MAY/"
            Case "06"
                r_Result = r_Result & "JUN/"
            Case "07"
                r_Result = r_Result & "JUL/"
            Case "08"
                r_Result = r_Result & "AGO/"
            Case "09"
                r_Result = r_Result & "SET/"
            Case "10"
                r_Result = r_Result & "OCT/"
            Case "11"
                r_Result = r_Result & "NOV/"
            Case "12"
                r_Result = r_Result & "DIC/"
        End Select
        r_Result = r_Result & Mid(p_Fecha, 5, 4)

        Cambia_Fecha = r_Result

        Return Cambia_Fecha.Trim

    End Function

    Private Function Prox_Vcto(ByVal p_Fecha As String, ByVal p_Flag As String) As String
        Dim r_Anio As String
        Dim r_Mes As String
        Dim r_Dia As String

        r_Anio = Mid(p_Fecha, 5, 4)
        r_Mes = Mid(p_Fecha, 3, 2)
        r_Dia = Mid(p_Fecha, 1, 2)
        If p_Flag = "+" Then
            If r_Mes = "12" Then
                r_Anio = Right("0000" & Trim(Str(Val(r_Anio) + 1)), 4)
                r_Mes = "01"
            Else
                r_Mes = Right("00" & Trim(Str(Val(r_Mes) + 1)), 2)
            End If
        Else
            If r_Mes = "01" Then
                r_Anio = Right("0000" & Trim(Str(Val(r_Anio) - 1)), 4)
                r_Mes = "12"
            Else
                r_Mes = Right("00" & Trim(Str(Val(r_Mes) - 1)), 2)
            End If
        End If
        Prox_Vcto = r_Dia & r_Mes & r_Anio
    End Function


    'OBTIENE EL PAGO TOTAL DEL MES Y EL PAGO MINIMO DEL MES
    Private Function pfnblnObtenerEstadoCuenta(ByVal sNroTarjeta As String, ByVal sNroCuenta As String, _
                                              ByVal vstrPeriodo As String, ByRef sDataEECC As String) As Integer
        Dim sRep As String = ""
        Dim sParam As String = ""
        Dim sMensajeUser As String = ""
        Dim sXML As String = ""
        Dim lEstado As Long = 0

        'Variables nuevas
        Dim actions_ As Long = 1
        Dim inetputBuff_ As String = ""
        Dim outpputBuff_ As String = ""
        Dim errorMsg_ As String = ""

        Try
            'Instancia al mirapiweb
            Dim obSendMirror_ As ClsTxMirapi = Nothing
            obSendMirror_ = New ClsTxMirapi()


            sParam = "00000000000" & sNroTarjeta.Trim & sNroCuenta.Trim & vstrPeriodo.Trim & "01"
            inetputBuff_ = "      " + "R192" + sParam

            sRep = obSendMirror_.ExecuteTX(actions_.ToString, inetputBuff_, outpputBuff_, TIME_OUT_SERVER, SERVIDOR_MIRROR_DESTINO, SERVIDOR_MIRROR_NODE, PUERTO, errorMsg_)
            outpputBuff_ = RTrim(outpputBuff_)
            If sRep = "0" Then 'EXITO

                If outpputBuff_.Length > 0 Then sRep = outpputBuff_.Substring(8, outpputBuff_.Length - 8)

                'EVALUAR LA RESPUESTA SI ES RU, AU (Respuesta correcta)
                If sRep.Trim.Length > 0 Then
                    If Left(sRep.Trim, 2) = "RU" Then 'Mensaje de Validacion al usuario. RU            000CUENTA BLOQUEADA
                        sMensajeUser = Mid(sRep.Trim, 18, Len(sRep.Trim))
                        sRep = ""
                        lEstado = 2
                    Else

                        'VALIDAR SI LA LLAMADA ES LARGA O CORTA
                        Dim sPagoTotalMes As String = ""
                        Dim sPagoMinimoMes As String = ""

                        sPagoMinimoMes = Trim(Mid(sRep, 367, 10))
                        If Left(sPagoMinimoMes.Trim, 1) = "-" Then
                            sPagoMinimoMes = "0"
                        End If

                        sPagoTotalMes = Trim(Mid(sRep, 377, 10))

                        If Left(sPagoTotalMes.Trim, 1) = "-" Then
                            sPagoTotalMes = "0"
                        End If

                        sDataEECC = sPagoTotalMes & "|\t|" & sPagoMinimoMes

                    End If


                Else
                    lEstado = 2

                End If

            ElseIf sRep = "-2" Then 'Ocurrio un error Recuperar el error
                'sRep = "ERROR:" & errorMsg_.Trim
                lEstado = 2
            Else
                'sRep = "ERROR:Servicio no disponible."
                lEstado = 2
            End If


        Catch ex As Exception
            lEstado = 2
        End Try


        Return lEstado

    End Function



    ' Descripción    : Procedimiento público que permite la obtención de las fechas de evaluación
    '                  de corte inicial y final en formato numérico y también la fecha de corte
    '                  inicial en formato dd/mm/yyyy
    Private Function psubFecEvaDDMMYYYY(ByVal vstrFecProcYYYYMMDD As String, ByVal vintDiaVenc As Integer) As String

        '' Declaracion de variables de trabajo locales
        Dim ldatFecVenc As Date
        Dim ldatFecProc As Date
        Dim ldatFecIniEva As Date
        Dim ldatFecIniEvaDDMMYYYY As Date
        Dim ldatFecFinEva As Date
        Dim ldblFecIniEvaDDMMYYYY As Double
        Dim sPeriodoFacturacion1 As String = ""
        Dim sPeriodoFacturacion2 As String = ""
        Dim sPeriodoFacturacion As String = ""
        Dim FechaAUX As Date
        'Resultados Finales
        Dim pdblFecIniEva As Double = 0
        Dim pdblFecFinEva As Double = 0
        Dim idif As Integer = 0
        Dim sFecha1 As String = ""
        Dim sFecha2 As String = ""



        Try


            '' Conversión de parametros de entrada
            ldatFecVenc = DateSerial(Mid(vstrFecProcYYYYMMDD, 1, 4), _
                                     Mid(vstrFecProcYYYYMMDD, 5, 2), _
                                     vintDiaVenc)

            ldatFecProc = DateSerial(Mid(vstrFecProcYYYYMMDD, 1, 4), _
                                     Mid(vstrFecProcYYYYMMDD, 5, 2), _
                                     Mid(vstrFecProcYYYYMMDD, 7, 2))

            '' Evaluacion de parametros de corte (Inicial y Final)...
            If ldatFecVenc >= ldatFecProc Then
                idif = DateDiff(DateInterval.Day, ldatFecVenc, ldatFecProc)
                If idif >= 15 Then
                    ldatFecVenc = DateAdd("m", -1, ldatFecVenc)
                End If
            Else
                ldatFecVenc = DateAdd("m", 1, ldatFecVenc)
                idif = DateDiff(DateInterval.Day, ldatFecProc, ldatFecVenc)
                If idif >= 15 Then
                    ldatFecVenc = DateAdd("m", -1, ldatFecVenc)
                End If
            End If

            '' Calculo de fechas de evaluacion...}
            FechaAUX = DateAdd("d", -15, ldatFecVenc)

            ldatFecIniEva = FechaAUX 'fecha final de facturacion
            ldatFecIniEvaDDMMYYYY = DateAdd("d", -15, ldatFecVenc) 'fecha final de facturacion
            'Formato DDMMYYYY
            sFecha2 = Right("00" & Str(Day(ldatFecIniEvaDDMMYYYY)).Trim, 2) & Right("00" & Str(Month(ldatFecIniEvaDDMMYYYY)).Trim, 2) & Right("0000" & Str(Year(ldatFecIniEvaDDMMYYYY)).Trim, 4)


            Dim fecha_eva As Date
            fecha_eva = DateAdd("m", -1, FechaAUX)
            fecha_eva = DateAdd("d", 1, fecha_eva)
            ldatFecFinEva = fecha_eva 'fecha final

            'Formato DDMMYYYY
            sFecha1 = Right("00" & Str(Day(ldatFecFinEva)).Trim, 2) & Right("00" & Str(Month(ldatFecFinEva)).Trim, 2) & Right("0000" & Str(Year(ldatFecFinEva)).Trim, 4)


            'Obtención de parametros de corte en Formato Numérico
            pdblFecIniEva = CDbl(Mid(ldatFecIniEva, 7, 4) & Mid(ldatFecIniEva, 4, 2) & Mid(ldatFecIniEva, 1, 2))
            pdblFecFinEva = CDbl(Mid(ldatFecFinEva, 7, 4) & Mid(ldatFecFinEva, 4, 2) & Mid(ldatFecFinEva, 1, 2))
            ldblFecIniEvaDDMMYYYY = CDbl(Mid(ldatFecIniEvaDDMMYYYY, 7, 4) & Mid(ldatFecIniEvaDDMMYYYY, 4, 2) & Mid(ldatFecIniEvaDDMMYYYY, 1, 2))


            sPeriodoFacturacion1 = sFecha1
            sPeriodoFacturacion2 = sFecha2

            sPeriodoFacturacion = Cambia_Fecha(sPeriodoFacturacion1) & " - " & Cambia_Fecha(sPeriodoFacturacion2)

        Catch ex As Exception

        End Try


        Return sPeriodoFacturacion

    End Function

    'BUSCAR PUNTOS RIPLEY

    Private Function PUNTOS_RIPLEY_ACUMULADOS(ByVal sNroTarjeta As String) As String

        Dim sRespuesta_ As String = ""
        Dim sParametros_ As String = ""
        Dim sMensajeErrorUsuario As String = ""
        Dim sXML As String = ""


        'Variables nuevas
        Dim actions__ As Long = 1
        Dim inetputBuff__ As String = ""
        Dim outpputBuff__ As String = ""
        Dim errorMsg__ As String = ""

        Try

            If sNroTarjeta.Trim.Length = 16 Then

                'Instancia al mirapiweb
                Dim obSendMirror__ As ClsTxMirapi = Nothing
                obSendMirror__ = New ClsTxMirapi()

                sParametros_ = "1" & sNroTarjeta.Trim
                inetputBuff__ = "      " + "FI01" + sParametros_

                sRespuesta_ = obSendMirror__.ExecuteTX(actions__.ToString, inetputBuff__, outpputBuff__, TIME_OUT_SERVER, SERVIDOR_MIRROR_DESTINO, SERVIDOR_MIRROR_NODE, PUERTO, errorMsg__)

                If sRespuesta_ = "0" Then 'EXITO

                    If outpputBuff__.Length > 0 Then sRespuesta_ = outpputBuff__.Substring(8, outpputBuff__.Length - 8)

                    'Evaluar Respuesta si es ERROR,  RU, AU (Respuesta correcta)
                    If sRespuesta_.Trim.Length > 0 Then
                        If Left(sRespuesta_.Trim, 2) = "RU" Then 'Mensaje de Validacion al usuario. 
                            sMensajeErrorUsuario = Mid(sRespuesta_.Trim, 18, Len(sRespuesta_.Trim))
                            sRespuesta_ = "0"
                        Else
                            If Left(sRespuesta_.Trim, 2) = "AU" Then
                                'Recuperar Data
                                Dim sPUNTOS As String = ""

                                sPUNTOS = Val(Mid(sRespuesta_, 15, 15)).ToString


                                sXML = sPUNTOS

                                sRespuesta_ = sXML

                            Else 'Cualquier Otro Caso
                                sRespuesta_ = "0"
                            End If

                        End If

                    Else 'Sino devuelve nada
                        sRespuesta_ = "0"
                    End If

                ElseIf sRespuesta_ = "-2" Then 'Ocurrio un error Recuperar el error
                    'sRespuesta_ = "ERROR:" & errorMsg__.Trim
                    sRespuesta_ = "0"
                Else
                    'sRespuesta_ = "ERROR:Servicio no disponible."
                    sRespuesta_ = "0"
                End If

            Else
                'Mostrar Mensaje de Error
                sRespuesta_ = "0"
            End If

        Catch ex As Exception
            'save log error

            sRespuesta_ = "0"

        End Try


        Return sRespuesta_

    End Function


    'OBTENER DISPONIBLE EFECTIVO EXPRESS Y PAGO TOTAL   PRUEBA DE EJEMPLO
    '<WebMethod(Description:="DISPONIBLE_EFEC_EXPRESS_PAGO_TOTAL")> _
    Private Function DISPONIBLE_EFEC_EXPRESS_PAGO_TOTAL(ByVal sNroTarjeta As String, ByVal IDkiosco As String) As String

        Dim sRespuesta_ As String = ""
        Dim sParametros_ As String = ""
        Dim sMensajeErrorUsuario As String = ""
        Dim sXML As String = ""
        Dim CodigoKiosco As String = IDkiosco.Trim & "               "


        'Variables nuevas
        Dim actions_ As Long = 1
        Dim inetputBuff_ As String = ""
        Dim outpputBuff_ As String = ""
        Dim errorMsg_ As String = ""

        Try

            If sNroTarjeta.Trim.Length = 16 Then

                'Instancia al mirapiweb
                Dim obSendMirror_ As ClsTxMirapi = Nothing
                obSendMirror_ = New ClsTxMirapi()

                sParametros_ = "   " & sNroTarjeta.Trim & "   " & "0           " & "0" & "0" & Left(CodigoKiosco, 15)
                inetputBuff_ = "      " + "SRM3" + sParametros_

                sRespuesta_ = obSendMirror_.ExecuteTX(actions_.ToString, inetputBuff_, outpputBuff_, TIME_OUT_SERVER, SERVIDOR_MIRROR_DESTINO, SERVIDOR_MIRROR_NODE, PUERTO, errorMsg_)

                If sRespuesta_ = "0" Then 'EXITO

                    If outpputBuff_.Length > 0 Then sRespuesta_ = outpputBuff_.Substring(8, outpputBuff_.Length - 8)

                    'Evaluar Respuesta si es RU, AU (Respuesta correcta)
                    If sRespuesta_.Trim.Length > 0 Then
                        If Left(sRespuesta_.Trim, 2) = "RU" Then 'Mensaje de Validacion al usuario. 
                            sMensajeErrorUsuario = Mid(sRespuesta_, 18, Len(sRespuesta_.Trim))
                            sRespuesta_ = ""
                        Else

                            'Recuperar Data
                            Dim sDisponibleEfectivo As String = ""
                            Dim sPagoTotal As String = ""
                            Dim sDato As String = ""
                            Dim dSaldoTotal As Double = 0
                            Dim dMontoFavor As Double = 0
                            Dim dPagoTotal As Double = 0
                            Dim dITF As Double = 0

                            sDato = CLng(Trim(Mid(sRespuesta_, 344, 9))).ToString
                            sDisponibleEfectivo = Format(Val(SET_MONTO_DECIMAL(sDato)), "##,##0.00")

                            sDato = ""
                            sDato = CLng(Trim(Mid(sRespuesta_, 46, 9))).ToString
                            dSaldoTotal = 0

                            dSaldoTotal = CDbl(SET_MONTO_DECIMAL(sDato))

                            sDato = ""
                            sDato = CLng(Trim(Mid(sRespuesta_, 353, 10))).ToString
                            dMontoFavor = 0

                            dMontoFavor = CDbl(SET_MONTO_DECIMAL(sDato))

                            'Obtener el ITF
                            dITF = OBTENER_FACTOR_ITF()
                            dPagoTotal = (dSaldoTotal - dMontoFavor) * dITF


                            sPagoTotal = Format(dPagoTotal, "##,##0.00")

                            sXML = sDisponibleEfectivo & "|\t|" & sPagoTotal


                            sRespuesta_ = sXML

                        End If
                    End If

                ElseIf sRespuesta_ = "-2" Then 'Ocurrio un error Recuperar el error
                    'sRespuesta_ = "ERROR:" & errorMsg_.Trim
                    sRespuesta_ = ""
                Else
                    'sRespuesta_ = "ERROR:Servicio no disponible."
                    sRespuesta_ = ""
                End If

            Else
                'Mostrar Mensaje de Error
                sRespuesta_ = ""
            End If

        Catch ex As Exception
            'save log error

            sRespuesta_ = ""

        End Try


        Return sRespuesta_

    End Function



    'DISPONIBLE SUPER EFECTIVO
    '<WebMethod(Description:="PRUEBAS DISPO SUPER EFEC")> _
    Private Function DISPONIBLE_SUPER_EFECTIVO(ByVal sNroCuenta As String) As String
        Dim sRespuesta_ As String = ""
        Dim sParametros_ As String = ""
        Dim sMensajeErrorUsuario As String = ""
        Dim sXML As String = ""



        'Variables nuevas
        Dim actions_ As Long = 1
        Dim inetputBuff_ As String = ""
        Dim outpputBuff_ As String = ""
        Dim errorMsg_ As String = ""

        Try

            If sNroCuenta.Trim.Length = 8 Then
                'Instancia al mirapiweb
                Dim obSendMirror_ As ClsTxMirapi = Nothing
                obSendMirror_ = New ClsTxMirapi()

                sParametros_ = "00002" & sNroCuenta.Trim
                inetputBuff_ = "      " + "V162" + sParametros_
                sRespuesta_ = obSendMirror_.ExecuteTX(actions_.ToString, inetputBuff_, outpputBuff_, TIME_OUT_SERVER, SERVIDOR_MIRROR_DESTINO, SERVIDOR_MIRROR_NODE, PUERTO, errorMsg_)

                'sRespuesta_ = "0"
                'outpputBuff_ = "12345678AU            0091294511B000000000750000000000000000020120615000000000000000003800002012051500000000000000000000000101                                                                                                                                                                                                                                                                28100000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000"

                If sRespuesta_ = "0" Then 'EXITO
                    If outpputBuff_.Length > 0 Then sRespuesta_ = outpputBuff_.Substring(8, outpputBuff_.Length - 8)

                    'Evaluar Respuesta si es RU, AU (Respuesta correcta)
                    If sRespuesta_.Trim.Length > 0 Then
                        If Left(sRespuesta_.Trim, 2) = "RU" Then 'Mensaje de Validacion al usuario. 
                            sMensajeErrorUsuario = Mid(sRespuesta_.Trim, 18, Len(sRespuesta_.Trim))
                            sRespuesta_ = ""
                        Else
                            If Left(sRespuesta_.Trim, 2) = "AU" Then
                                'Recuperar Data
                                Dim sDisponibleSuperEfectivo As String = ""
                                Dim sDato As String = ""
                                Dim sFechaVence As String = ""

                                sDato = CLng(Trim(Mid(sRespuesta_, 26, 14))).ToString
                                sDisponibleSuperEfectivo = Format(Val(SET_MONTO_DECIMAL(sDato)), "##,##0.00")

                                sFechaVence = Mid(sRespuesta_, 54, 8)


                                Dim vFecHoy As Date
                                Dim vFecVence As Date

                                vFecHoy = Date.Now
                                sFechaVence = FormatearFecha_DDMMYYYY(sFechaVence.Trim)

                                If sFechaVence.Trim <> "" Then
                                    'VALIDAR FECHAS
                                    If sFechaVence.Trim.Length = 10 Then
                                        vFecVence = CDate(sFechaVence.Trim)

                                        If vFecVence > vFecHoy Then 'SI LA FECHA DE VENCIMIENTO ES MAYOR A LA FECHA ACTUAL ENTONCES MOSTRAR LOS DATOS DEL SUPEREFECTIVO
                                            sXML = sDisponibleSuperEfectivo.Trim
                                        Else
                                            sXML = ""
                                        End If

                                    End If
                                Else
                                    sXML = ""
                                End If

                                sRespuesta_ = sXML

                            Else 'Cualquier Otro Caso
                                sRespuesta_ = ""
                            End If

                        End If

                    Else 'Sino devuelve nada
                        sRespuesta_ = ""
                    End If

                ElseIf sRespuesta_ = "-2" Then 'Ocurrio un error Recuperar el error
                    'sRespuesta_ = "ERROR:" & errorMsg_.Trim
                    sRespuesta_ = ""
                Else
                    'sRespuesta_ = "ERROR:Servicio no disponible."
                    sRespuesta_ = ""
                End If

            Else
                'Mostrar Mensaje de Error
                sRespuesta_ = ""
            End If

        Catch ex As Exception
            sRespuesta_ = ""
        End Try

        Return sRespuesta_

    End Function

    'DISPONIBLE SUPER EFECTIVO MC VISA
    '<WebMethod(Description:="PRUEBAS DISPO SUPER EFEC MC")> _
    Private Function DISPONIBLE_SUPER_EFECTIVO_MC_VISA(ByVal sNroTarjeta As String, ByVal sTipoDoc As String, ByVal sNroDoc As String) As String
        Dim sRespuesta_ As String = ""
        Dim sParametros_ As String = ""
        Dim sMensajeErrorUsuario As String = ""
        Dim sXML As String = ""

        'Variables nuevas
        Dim actions_ As Long = 1
        Dim inetputBuff_ As String = ""
        Dim outpputBuff_ As String = ""
        Dim errorMsg_ As String = ""

        Try

            If sNroTarjeta.Trim.Length = 16 Then

                'Instancia al mirapiweb
                Dim obSendMirror_ As ClsTxMirapi = Nothing
                obSendMirror_ = New ClsTxMirapi()

                sParametros_ = "0000" & "2" & Right(Trim("0" & sTipoDoc.Trim), 1) & Right(Trim("000000000000" & sNroDoc.Trim), 12) & "000" & sNroTarjeta.Trim
                inetputBuff_ = "      " + "V172" + sParametros_
                sRespuesta_ = obSendMirror_.ExecuteTX(actions_.ToString, inetputBuff_, outpputBuff_, TIME_OUT_SERVER, SERVIDOR_MIRROR_DESTINO, SERVIDOR_MIRROR_NODE, PUERTO, errorMsg_)

                If sRespuesta_ = "0" Then 'EXITO

                    If outpputBuff_.Length > 0 Then sRespuesta_ = outpputBuff_.Substring(8, outpputBuff_.Length - 8)

                    'Evaluar Respuesta si es ERROR,  RU, AU (Respuesta correcta)
                    If sRespuesta_.Trim.Length > 0 Then
                        If Left(sRespuesta_.Trim, 2) = "RU" Then 'Mensaje de Validacion al usuario. 
                            sMensajeErrorUsuario = Mid(sRespuesta_.Trim, 18, Len(sRespuesta_.Trim))
                            sRespuesta_ = "|\t|"
                        Else
                            If Left(sRespuesta_.Trim, 2) = "AU" Then
                                'Recuperar Data
                                Dim sDisponibleSuperEfectivo As String = ""
                                Dim sDato As String = ""
                                Dim sFechaVence As String = ""
                                Dim sCondicion As String = ""

                                sDato = CLng(Trim(Mid(sRespuesta_, 26, 14))).ToString
                                sDisponibleSuperEfectivo = Format(Val(SET_MONTO_DECIMAL(sDato)), "##,##0.00")

                                sFechaVence = Mid(sRespuesta_, 54, 8)

                                Dim vFecHoy As Date
                                Dim vFecVence As Date

                                vFecHoy = Date.Now
                                sFechaVence = FormatearFecha_DDMMYYYY(sFechaVence.Trim)

                                If sFechaVence.Trim <> "" Then
                                    'VALIDAR FECHAS
                                    If sFechaVence.Trim.Length = 10 Then
                                        vFecVence = CDate(sFechaVence.Trim)
                                        If vFecVence > vFecHoy Then 'SI LA FECHA DE VENCIMIENTO ES MAYOR A LA FECHA ACTUAL ENTONCES MOSTRAR LOS DATOS DEL SUPEREFECTIVO
                                            sXML = sDisponibleSuperEfectivo.Trim & "|\t|" & sFechaVence.Trim
                                        Else
                                            sXML = "|\t|"
                                        End If

                                    End If
                                Else
                                    sXML = "|\t|"
                                End If

                                sRespuesta_ = sXML

                            Else 'Cualquier Otro Caso
                                sRespuesta_ = "|\t|"
                            End If


                        End If

                    Else 'Sino devuelve nada
                        sRespuesta_ = "|\t|"
                    End If


                ElseIf sRespuesta_ = "-2" Then 'Ocurrio un error Recuperar el error
                    'sRespuesta_ = "ERROR:" & errorMsg_.Trim
                    sRespuesta_ = "|\t|"
                Else
                    'sRespuesta_ = "ERROR:Servicio no disponible."
                    sRespuesta_ = "|\t|"
                End If

            Else
                'Mostrar Mensaje de Error
                sRespuesta_ = "|\t|"
            End If

        Catch ex As Exception
            'save log error

            sRespuesta_ = "|\t|"

        End Try

        Return sRespuesta_

    End Function

    ''' <summary>
    ''' Obtiene las promociones que tiene un cliente
    ''' </summary>
    ''' <param name="tipoProducto">Tipo de Producto</param>
    ''' <param name="tipoDocumento">Tipo de Documento</param>
    ''' <param name="nroDocumento">Numero de Documento</param>
    ''' <param name="nroContrato">Numero de Contrato</param>
    ''' <returns>Array de 2 objetos con las promociones de super efectivo e incremento de linea</returns>
    ''' <remarks>En caso el cliente no tenga alguna promocion, la respuesta para SEF será una cadena sin datos, y para INC será un objeto con un codigo y sin datos</remarks>
    <WebMethod(Description:="Obtiene las promociones disponibles del cliente")> _
    Public Function OBTENER_PROMOCIONES(ByVal tipoProducto As String, ByVal tipoDocumento As String, _
                                        ByVal nroDocumento As String, ByVal nroContrato As String, _
                                        ByVal tipoTarjeta As String, _
                                        ByVal linea_actual As String, _
                                        ByVal nro_tarjeta As String) As Object()
        Dim promociones As Object() = New Object() {}
        '02/03/2015 CP
        'ReDim promociones(2)
        ReDim promociones(1)

        '17/12/2014 CAMBIO CONSULTA DE SUPEREFECTIVO SEF PARA QUE FUNCIONE TEMPORALMENTE SIN OFERTA INICIAL
        'promociones(0) = ObtenerOfertaSEF(tipoProducto, tipoDocumento, nroDocumento)
        promociones(0) = MOSTRAR_SUPEREFECTIVO_SEF_OFERTAINICIALFICTICIA(tipoProducto, tipoDocumento, nroDocumento)
        'respuesta = "1|\t|1|\t|09924784|\t|1|\t|RUBEN EMILIO PARRA WILLIAMS|\t|5254740045393940|\t|4500.00|\t|36.00|\t|55.93|\t|21.19|\t|31/12/2014|\t|001|\t|1.99|\t|4500|\t|12|\t|33.5|\t|170|\t|3000"
        promociones(1) = MOSTRAR_INCREMENTO_LINEA_INC(nroContrato, tipoTarjeta, linea_actual, nro_tarjeta)
        '02/03/2015 CP
        'promociones(2) = ObtenerOfertaCambioProductoComercial(nroContrato)
        Return promociones
    End Function

    '17/12/2014 CAMBIO CONSULTA DE SUPEREFECTIVO SEF PARA QUE FUNCIONE TEMPORALMENTE SIN OFERTA INICIAL
    <WebMethod(Description:="Consulta de SuperEfectivo SEF.")> _
    Public Function MOSTRAR_SUPEREFECTIVO_SEF_OFERTAINICIALFICTICIA(ByVal sTipoProducto As String, ByVal sTipoDocumento As String, ByVal sNroDocumento As String) As String

        Dim conConexion_sef As OracleConnection
        Dim fechaActual As Date

        Try
            conConexion_sef = New OracleConnection(mCadenaConexion_ORA_SEF)

            If Not conConexion_sef Is Nothing Then
                conConexion_sef.Open()

                If conConexion_sef.State = ConnectionState.Open Then

                    'LLAMAR AL PROCEDIMIENTO RIPEY PUNTOS
                    Dim cmd_sef As New OracleCommand
                    Dim param1 As New OracleParameter
                    Dim param2 As New OracleParameter
                    Dim param3 As New OracleParameter
                    Dim param4 As New OracleParameter
                    Dim param5 As New OracleParameter
                    Dim param6 As New OracleParameter
                    Dim param7 As New OracleParameter
                    Dim param8 As New OracleParameter
                    Dim param9 As New OracleParameter
                    Dim param10 As New OracleParameter
                    Dim param11 As New OracleParameter
                    Dim param12 As New OracleParameter
                    Dim param13 As New OracleParameter
                    Dim param14 As New OracleParameter
                    Dim param15 As New OracleParameter
                    Dim param16 As New OracleParameter
                    Dim param17 As New OracleParameter


                    'VARIABLES DE SALIDA
                    Dim sExiste As String = ""
                    Dim sTipoProd As String = ""
                    Dim sNomCliente As String = ""
                    Dim sNumeroCuenta As String = ""
                    Dim sLineaOferta As String = ""
                    Dim sPlazo As String = ""
                    Dim sTasa As String = ""
                    Dim sCuota As String = ""
                    Dim sTEM As String = ""
                    Dim sTEA As String = ""
                    Dim sOK As String = ""
                    Dim sFechaIniVigencia As String = ""
                    Dim sFechaFinVigencia As String = ""
                    'Dim sLineaOfertaInicial As String = ""



                    cmd_sef.Connection = conConexion_sef
                    cmd_sef.CommandText = "PKG_OFERTAS_PRODUCTOS.PRC_CONS_OFERTA_SEF"
                    cmd_sef.CommandType = CommandType.StoredProcedure

                    param1.ParameterName = "I_TIP_PROD"
                    param1.OracleType = OracleType.Double
                    param1.Direction = ParameterDirection.Input
                    param1.Value = CDbl(sTipoProducto.Trim) 'Tipo de producto
                    cmd_sef.Parameters.Add(param1)

                    param2.ParameterName = "I_TIP_DOC"
                    param2.OracleType = OracleType.Double
                    param2.Direction = ParameterDirection.Input
                    param2.Value = CDbl(sTipoDocumento.Trim)
                    cmd_sef.Parameters.Add(param2)

                    param3.ParameterName = "I_NUM_DOC"
                    param3.OracleType = OracleType.VarChar
                    param3.Size = 15
                    param3.Direction = ParameterDirection.Input
                    param3.Value = sNroDocumento.Trim
                    cmd_sef.Parameters.Add(param3)

                    param4.ParameterName = "O_EXISTE"
                    param4.OracleType = OracleType.Double
                    param4.Direction = ParameterDirection.Output
                    cmd_sef.Parameters.Add(param4)

                    param5.ParameterName = "O_STIP_PROD"
                    param5.OracleType = OracleType.VarChar
                    param5.Size = 50
                    param5.Direction = ParameterDirection.Output
                    cmd_sef.Parameters.Add(param5)

                    param6.ParameterName = "O_NOM_CLIE"
                    param6.OracleType = OracleType.VarChar
                    param6.Size = 80
                    param6.Direction = ParameterDirection.Output
                    cmd_sef.Parameters.Add(param6)

                    param7.ParameterName = "O_NUM_CUENTA"
                    param7.OracleType = OracleType.VarChar
                    param7.Size = 50
                    param7.Direction = ParameterDirection.Output
                    cmd_sef.Parameters.Add(param7)

                    param8.ParameterName = "O_LINEA_OFE"
                    param8.OracleType = OracleType.Double
                    param8.Direction = ParameterDirection.Output
                    cmd_sef.Parameters.Add(param8)

                    param9.ParameterName = "O_PLAZO"
                    param9.OracleType = OracleType.Double
                    param9.Direction = ParameterDirection.Output
                    cmd_sef.Parameters.Add(param9)

                    param10.ParameterName = "O_TASA"
                    param10.OracleType = OracleType.Double
                    param10.Direction = ParameterDirection.Output
                    cmd_sef.Parameters.Add(param10)

                    param11.ParameterName = "O_CUOTA"
                    param11.OracleType = OracleType.Double
                    param11.Direction = ParameterDirection.Output
                    cmd_sef.Parameters.Add(param11)


                    param12.ParameterName = "O_TEM"
                    param12.OracleType = OracleType.Double
                    param12.Direction = ParameterDirection.Output
                    cmd_sef.Parameters.Add(param12)

                    param13.ParameterName = "O_TEA"
                    param13.OracleType = OracleType.Double
                    param13.Direction = ParameterDirection.Output
                    cmd_sef.Parameters.Add(param13)

                    param14.ParameterName = "O_FEC_INI_VIG"
                    param14.OracleType = OracleType.VarChar
                    param14.Size = 10
                    param14.Direction = ParameterDirection.Output
                    cmd_sef.Parameters.Add(param14)

                    param15.ParameterName = "O_FEC_FIN_VIG"
                    param15.OracleType = OracleType.VarChar
                    param15.Size = 10
                    param15.Direction = ParameterDirection.Output
                    cmd_sef.Parameters.Add(param15)

                    'param16.ParameterName = "O_LINEA_OFE_INI"
                    'param16.OracleType = OracleType.Double
                    'param16.Direction = ParameterDirection.Output
                    'cmd_sef.Parameters.Add(param16)

                    param16.ParameterName = "O_OK"
                    param16.OracleType = OracleType.Double
                    param16.Direction = ParameterDirection.Output
                    cmd_sef.Parameters.Add(param16)

                    cmd_sef.ExecuteNonQuery()

                    sExiste = cmd_sef.Parameters("O_EXISTE").Value.ToString
                    sTipoProd = cmd_sef.Parameters("O_STIP_PROD").Value.ToString
                    sNomCliente = cmd_sef.Parameters("O_NOM_CLIE").Value.ToString
                    sNumeroCuenta = cmd_sef.Parameters("O_NUM_CUENTA").Value.ToString
                    sLineaOferta = cmd_sef.Parameters("O_LINEA_OFE").Value.ToString
                    sPlazo = cmd_sef.Parameters("O_PLAZO").Value.ToString
                    sTasa = cmd_sef.Parameters("O_TASA").Value.ToString
                    sCuota = cmd_sef.Parameters("O_CUOTA").Value.ToString
                    sTEM = cmd_sef.Parameters("O_TEM").Value.ToString
                    sTEA = cmd_sef.Parameters("O_TEA").Value.ToString
                    sFechaIniVigencia = cmd_sef.Parameters("O_FEC_INI_VIG").Value.ToString.Trim
                    sFechaFinVigencia = cmd_sef.Parameters("O_FEC_FIN_VIG").Value.ToString.Trim
                    'sLineaOfertaInicial = cmd_sef.Parameters("O_LINEA_OFE_INI").Value.ToString


                    sOK = cmd_sef.Parameters("O_OK").Value.ToString
                    sResultado_ORA = ""
                    Dim codigoBarra As String = "3000"

                    sResultado_ORA = sTipoProducto.Trim & "|\t|" & sTipoDocumento.Trim & "|\t|" & sNroDocumento.Trim & "|\t|" & sTipoProd.Trim & "|\t|" & sNomCliente.Trim
                    sResultado_ORA = sResultado_ORA & "|\t|" & sNumeroCuenta.Trim & "|\t|" & sLineaOferta.Trim & "|\t|" & sPlazo.Trim & "|\t|" & sTasa.Trim & "|\t|" & sCuota.Trim & "|\t|" & sFechaFinVigencia.Trim & "|\t|" & GET_CODIGO_ATENCION_SEF() & "|\t|" & sTEM.Trim
                    sResultado_ORA = sResultado_ORA & "|\t|" & sLineaOferta.Trim & "|\t|" & sPlazo.Trim & "|\t|" & sTasa.Trim & "|\t|" & sCuota.Trim & "|\t|" & codigoBarra.Trim
                    sResultado_ORA = sResultado_ORA & "|\t|" & fechaActual.Now.ToShortDateString()

                    'Demo'
                    'sResultado_ORA = "1|\t|1|\t|09924784|\t|1|\t|RUBEN EMILIO PARRA WILLIAMS|\t|5254740045393940|\t|4500.00|\t|36.00|\t|55.93|\t|21.19|\t|31/12/2014|\t|001|\t|1.99|\t|4500|\t|12|\t|33.5|\t|170|\t|3000"
                    Dim vFecVence As Date
                    Dim vFechaActual As Date

                    vFechaActual = New Date(Now.Year, Now.Month, Now.Day)

                    If sFechaFinVigencia.Trim <> "" Then
                        'VALIDAR FECHAS
                        If sFechaFinVigencia.Trim.Length = 10 Then

                            Dim aFecha() As String = sFechaFinVigencia.Split("/")
                            vFecVence = New Date(aFecha(2), aFecha(1), aFecha(0))

                            If vFecVence < vFechaActual Then 'SI LA FECHA DE VENCIMIENTO YA VENCIO NO MOSTRAR NINGUN DATO
                                sResultado_ORA = "|\t||\t||\t||\t||\t||\t||\t||\t||\t||\t||\t||\t||\t||\t||\t||\t||\t||\t|"
                            End If
                        End If
                    Else
                        sResultado_ORA = "|\t||\t||\t||\t||\t||\t||\t||\t||\t||\t||\t||\t||\t||\t||\t||\t||\t||\t|"
                    End If

                    cmd_sef.Dispose()
                    cmd_sef = Nothing
                    conConexion_sef.Close()
                    conConexion_sef = Nothing

                End If

            Else

                sResultado_ORA = "|\t||\t||\t||\t||\t||\t||\t||\t||\t||\t||\t||\t||\t||\t||\t||\t||\t||\t|"

            End If

        Catch ex As Exception
            sResultado_ORA = "|\t||\t||\t||\t||\t||\t||\t||\t||\t||\t||\t||\t||\t||\t||\t||\t||\t||\t|"
            'sResultado_ORA = ex.Message
            'sResultado_ORA = ex.Message
        End Try

        Return sResultado_ORA

    End Function


    'CONSULTA DE SUPEREFECTIVO SEF
    <WebMethod(Description:="Consulta de SuperEfectivo SEF.")> _
    Public Function MOSTRAR_SUPEREFECTIVO_SEF(ByVal sTipoProducto As String, ByVal sTipoDocumento As String, ByVal sNroDocumento As String) As String

        Dim conConexion_sef As OracleConnection

        Try
            conConexion_sef = New OracleConnection(mCadenaConexion_ORA_SEF)

            If Not conConexion_sef Is Nothing Then
                conConexion_sef.Open()

                If conConexion_sef.State = ConnectionState.Open Then

                    'LLAMAR AL PROCEDIMIENTO RIPEY PUNTOS
                    Dim cmd_sef As New OracleCommand
                    Dim param1 As New OracleParameter
                    Dim param2 As New OracleParameter
                    Dim param3 As New OracleParameter
                    Dim param4 As New OracleParameter
                    Dim param5 As New OracleParameter
                    Dim param6 As New OracleParameter
                    Dim param7 As New OracleParameter
                    Dim param8 As New OracleParameter
                    Dim param9 As New OracleParameter
                    Dim param10 As New OracleParameter
                    Dim param11 As New OracleParameter
                    Dim param12 As New OracleParameter
                    Dim param13 As New OracleParameter
                    Dim param14 As New OracleParameter
                    Dim param15 As New OracleParameter
                    Dim param16 As New OracleParameter
                    Dim param17 As New OracleParameter


                    'VARIABLES DE SALIDA
                    Dim sExiste As String = ""
                    Dim sTipoProd As String = ""
                    Dim sNomCliente As String = ""
                    Dim sNumeroCuenta As String = ""
                    Dim sLineaOferta As String = ""
                    Dim sPlazo As String = ""
                    Dim sTasa As String = ""
                    Dim sCuota As String = ""
                    Dim sTEM As String = ""
                    Dim sTEA As String = ""
                    Dim sOK As String = ""
                    Dim sFechaIniVigencia As String = ""
                    Dim sFechaFinVigencia As String = ""
                    'Dim sLineaOfertaInicial As String = ""



                    cmd_sef.Connection = conConexion_sef
                    cmd_sef.CommandText = "PKG_OFERTAS_PRODUCTOS.PRC_CONS_OFERTA_SEF"
                    cmd_sef.CommandType = CommandType.StoredProcedure

                    param1.ParameterName = "I_TIP_PROD"
                    param1.OracleType = OracleType.Double
                    param1.Direction = ParameterDirection.Input
                    param1.Value = CDbl(sTipoProducto.Trim) 'Tipo de producto
                    cmd_sef.Parameters.Add(param1)

                    param2.ParameterName = "I_TIP_DOC"
                    param2.OracleType = OracleType.Double
                    param2.Direction = ParameterDirection.Input
                    param2.Value = CDbl(sTipoDocumento.Trim)
                    cmd_sef.Parameters.Add(param2)

                    param3.ParameterName = "I_NUM_DOC"
                    param3.OracleType = OracleType.VarChar
                    param3.Size = 15
                    param3.Direction = ParameterDirection.Input
                    param3.Value = sNroDocumento.Trim
                    cmd_sef.Parameters.Add(param3)

                    param4.ParameterName = "O_EXISTE"
                    param4.OracleType = OracleType.Double
                    param4.Direction = ParameterDirection.Output
                    cmd_sef.Parameters.Add(param4)

                    param5.ParameterName = "O_STIP_PROD"
                    param5.OracleType = OracleType.VarChar
                    param5.Size = 50
                    param5.Direction = ParameterDirection.Output
                    cmd_sef.Parameters.Add(param5)

                    param6.ParameterName = "O_NOM_CLIE"
                    param6.OracleType = OracleType.VarChar
                    param6.Size = 80
                    param6.Direction = ParameterDirection.Output
                    cmd_sef.Parameters.Add(param6)

                    param7.ParameterName = "O_NUM_CUENTA"
                    param7.OracleType = OracleType.VarChar
                    param7.Size = 50
                    param7.Direction = ParameterDirection.Output
                    cmd_sef.Parameters.Add(param7)

                    param8.ParameterName = "O_LINEA_OFE"
                    param8.OracleType = OracleType.Double
                    param8.Direction = ParameterDirection.Output
                    cmd_sef.Parameters.Add(param8)

                    param9.ParameterName = "O_PLAZO"
                    param9.OracleType = OracleType.Double
                    param9.Direction = ParameterDirection.Output
                    cmd_sef.Parameters.Add(param9)

                    param10.ParameterName = "O_TASA"
                    param10.OracleType = OracleType.Double
                    param10.Direction = ParameterDirection.Output
                    cmd_sef.Parameters.Add(param10)

                    param11.ParameterName = "O_CUOTA"
                    param11.OracleType = OracleType.Double
                    param11.Direction = ParameterDirection.Output
                    cmd_sef.Parameters.Add(param11)


                    param12.ParameterName = "O_TEM"
                    param12.OracleType = OracleType.Double
                    param12.Direction = ParameterDirection.Output
                    cmd_sef.Parameters.Add(param12)

                    param13.ParameterName = "O_TEA"
                    param13.OracleType = OracleType.Double
                    param13.Direction = ParameterDirection.Output
                    cmd_sef.Parameters.Add(param13)

                    param14.ParameterName = "O_FEC_INI_VIG"
                    param14.OracleType = OracleType.VarChar
                    param14.Size = 10
                    param14.Direction = ParameterDirection.Output
                    cmd_sef.Parameters.Add(param14)

                    param15.ParameterName = "O_FEC_FIN_VIG"
                    param15.OracleType = OracleType.VarChar
                    param15.Size = 10
                    param15.Direction = ParameterDirection.Output
                    cmd_sef.Parameters.Add(param15)

                    'param16.ParameterName = "O_LINEA_OFE_INI"
                    'param16.OracleType = OracleType.Double
                    'param16.Direction = ParameterDirection.Output
                    'cmd_sef.Parameters.Add(param16)

                    param16.ParameterName = "O_OK"
                    param16.OracleType = OracleType.Double
                    param16.Direction = ParameterDirection.Output
                    cmd_sef.Parameters.Add(param16)

                    cmd_sef.ExecuteNonQuery()

                    sExiste = cmd_sef.Parameters("O_EXISTE").Value.ToString
                    sTipoProd = cmd_sef.Parameters("O_STIP_PROD").Value.ToString
                    sNomCliente = cmd_sef.Parameters("O_NOM_CLIE").Value.ToString
                    sNumeroCuenta = cmd_sef.Parameters("O_NUM_CUENTA").Value.ToString
                    sLineaOferta = cmd_sef.Parameters("O_LINEA_OFE").Value.ToString
                    sPlazo = cmd_sef.Parameters("O_PLAZO").Value.ToString
                    sTasa = cmd_sef.Parameters("O_TASA").Value.ToString
                    sCuota = cmd_sef.Parameters("O_CUOTA").Value.ToString
                    sTEM = cmd_sef.Parameters("O_TEM").Value.ToString
                    sTEA = cmd_sef.Parameters("O_TEA").Value.ToString
                    sFechaIniVigencia = cmd_sef.Parameters("O_FEC_INI_VIG").Value.ToString.Trim
                    sFechaFinVigencia = cmd_sef.Parameters("O_FEC_FIN_VIG").Value.ToString.Trim
                    'sLineaOfertaInicial = cmd_sef.Parameters("O_LINEA_OFE_INI").Value.ToString


                    sOK = cmd_sef.Parameters("O_OK").Value.ToString
                    sResultado_ORA = ""

                    sResultado_ORA = sTipoProducto.Trim & "|\t|" & sTipoDocumento.Trim & "|\t|" & sNroDocumento.Trim & "|\t|" & sTipoProd.Trim & "|\t|" & sNomCliente.Trim
                    sResultado_ORA = sResultado_ORA & "|\t|" & sNumeroCuenta.Trim & "|\t|" & sLineaOferta.Trim & "|\t|" & sPlazo.Trim & "|\t|" & sTasa.Trim & "|\t|" & sCuota.Trim & "|\t|" & sFechaFinVigencia.Trim & "|\t|" & GET_CODIGO_ATENCION_SEF() & "|\t|" & sTEM.Trim '& "|\t|" & sLineaOfertaInicial.Trim
                    'Demo'
                    'sResultado_ORA = "1|\t|1|\t|09924784|\t|1|\t|RUBEN EMILIO PARRA WILLIAMS|\t|5254740045393940|\t|4500.00|\t|36.00|\t|55.93|\t|21.19|\t|31/12/2014|\t|001|\t|1.99"
                    Dim vFecVence As Date
                    Dim vFechaActual As Date

                    vFechaActual = New Date(Now.Year, Now.Month, Now.Day)

                    If sFechaFinVigencia.Trim <> "" Then
                        'VALIDAR FECHAS
                        If sFechaFinVigencia.Trim.Length = 10 Then

                            Dim aFecha() As String = sFechaFinVigencia.Split("/")
                            vFecVence = New Date(aFecha(2), aFecha(1), aFecha(0))

                            If vFecVence < vFechaActual Then 'SI LA FECHA DE VENCIMIENTO YA VENCIO NO MOSTRAR NINGUN DATO
                                sResultado_ORA = "|\t||\t||\t||\t||\t||\t||\t||\t||\t||\t||\t||\t|"
                            End If
                        End If
                    Else
                        sResultado_ORA = "|\t||\t||\t||\t||\t||\t||\t||\t||\t||\t||\t||\t|"
                    End If

                    cmd_sef.Dispose()
                    cmd_sef = Nothing
                    conConexion_sef.Close()
                    conConexion_sef = Nothing

                End If

            Else

                sResultado_ORA = "|\t||\t||\t||\t||\t||\t||\t||\t||\t||\t||\t||\t|"

            End If

        Catch ex As Exception
            sResultado_ORA = "|\t||\t||\t||\t||\t||\t||\t||\t||\t||\t||\t||\t|"
            'sResultado_ORA = ex.Message
            'sResultado_ORA = ex.Message
        End Try

        Return sResultado_ORA

    End Function

    ''' <summary>
    ''' Consulta y obtiene los datos para la promoción de incremento de línea del cliente
    ''' </summary>
    ''' <param name="nroContrato">Numero de Documento del cliente</param>
    ''' <param name="tipoTarjeta">Tipo de Tarjeta "RSAT", "SICRON" o "MC"</param>
    ''' <returns>Un onjeto de tipo salida con los resultados del a consulta</returns>
    ''' <remarks></remarks>
    <WebMethod(Description:="Consulta de Incremento de Linea")> _
    Public Function MOSTRAR_INCREMENTO_LINEA_INC(ByVal nroContrato As String, _
                                                 ByVal tipoTarjeta As String, _
                                                 ByVal linea_actual As String, _
                                                 ByVal nro_tarjeta As String) As Salida
        Dim conexionOraInc As OracleConnection
        Dim salida As New Salida

        'salida.Code = Constantes.CODE_EXITO
        'If tipoTarjeta = "RSAT" Then
        '    salida.Mensaje = "Sr(a). Juan Perez, le ofrecemos incrementar su línea de crédito" _
        '                                            & " hasta S/. 1800. Autorice su incremento sólo hasta el 30/10/2013."
        'Else
        '    salida.Mensaje = "Sr(a). Juan Perez, le ofrecemos incrementar su línea de crédito" _
        '                    & " hasta S/. 1800. Autorice su incremento sólo hasta el 30/10/2013. En caso acepte el incremento de su línea de crédito, este se aplicará" _
        '                    & " 48 horas después o 02 días hábiles después de su aceptación."
        'End If
        'salida.Ok = 1
        'salida.CodigoAtencion = 1125
        'salida.Existe = 1
        'salida.NroTarjeta = "1384935466"
        'salida.NroDocumento = "95436413"
        'salida.NombreCliente = "Juan Perez"
        'salida.LineaInicial = 1000
        'salida.LineaEgm = 1800
        'salida.LineaCons = 1500
        'salida.LineaEfex = 300
        'salida.FechaInicioVigencia = "01/01/2013"
        'salida.FechaFinVigencia = "20/12/2013"
        'Return salida

        Dim mensajeErrorInc As String = "Ocurrio un error inesperado"

        Try
            Dim mensajeSinPromocion As String = "El cliente no tiene la promocion de incremento de linea"
            If tipoTarjeta <> "RSAT" And tipoTarjeta <> "MC" Then
                salida.Code = Constantes.CODE_NULL
                salida.Mensaje = mensajeSinPromocion
            Else
                conexionOraInc = New OracleConnection(cadenaConexionOraInc)

                If Not conexionOraInc Is Nothing Then
                    conexionOraInc.Open()
                    If conexionOraInc.State = ConnectionState.Open Then
                        Dim comando As New OracleCommand

                        Dim param1 As New OracleParameter
                        Dim param2 As New OracleParameter
                        Dim param3 As New OracleParameter
                        Dim param4 As New OracleParameter
                        Dim param5 As New OracleParameter
                        Dim param6 As New OracleParameter
                        Dim param7 As New OracleParameter
                        Dim param8 As New OracleParameter
                        Dim param9 As New OracleParameter
                        Dim param10 As New OracleParameter
                        Dim param11 As New OracleParameter
                        Dim param12 As New OracleParameter
                        Dim param13 As New OracleParameter
                        Dim param14 As New OracleParameter
                        Dim param15 As New OracleParameter

                        comando.Connection = conexionOraInc
                        comando.CommandText = "PKG_OFERTAS_INC_LINEA.PRC_CONS_OFERTA_INC_LINEA"
                        comando.CommandType = CommandType.StoredProcedure

                        param1.ParameterName = "I_NUM_CONTRATO"
                        param1.OracleType = OracleType.VarChar
                        param1.Direction = ParameterDirection.Input
                        param1.Value = nroContrato.Trim
                        comando.Parameters.Add(param1)

                        param2.ParameterName = "O_EXISTE"
                        param2.OracleType = OracleType.Int32
                        param2.Direction = ParameterDirection.Output
                        comando.Parameters.Add(param2)

                        param3.ParameterName = "O_NUM_TARJETA"
                        param3.OracleType = OracleType.VarChar
                        param3.Size = 19
                        param3.Direction = ParameterDirection.Output
                        comando.Parameters.Add(param3)

                        param4.ParameterName = "O_NUM_DOCU"
                        param4.OracleType = OracleType.VarChar
                        param4.Size = 12
                        param4.Direction = ParameterDirection.Output
                        comando.Parameters.Add(param4)

                        param15.ParameterName = "O_NOM_CLIE"
                        param15.OracleType = OracleType.VarChar
                        param15.Size = 60
                        param15.Direction = ParameterDirection.Output
                        comando.Parameters.Add(param15)

                        param5.ParameterName = "O_LINEA_INI"
                        param5.OracleType = OracleType.Number
                        param5.Direction = ParameterDirection.Output
                        comando.Parameters.Add(param5)

                        param6.ParameterName = "O_LINEA_EGM"
                        param6.OracleType = OracleType.Number
                        param6.Direction = ParameterDirection.Output
                        comando.Parameters.Add(param6)

                        param7.ParameterName = "O_LINEA_CONS"
                        param7.OracleType = OracleType.Number
                        param7.Direction = ParameterDirection.Output
                        comando.Parameters.Add(param7)

                        param8.ParameterName = "O_LINEA_EFEX"
                        param8.OracleType = OracleType.Number
                        param8.Direction = ParameterDirection.Output
                        comando.Parameters.Add(param8)

                        param12.ParameterName = "O_FEC_INI_VIG"
                        param12.OracleType = OracleType.DateTime
                        param12.Direction = ParameterDirection.Output
                        comando.Parameters.Add(param12)

                        param13.ParameterName = "O_FEC_FIN_VIG"
                        param13.OracleType = OracleType.DateTime
                        param13.Direction = ParameterDirection.Output
                        comando.Parameters.Add(param13)

                        param14.ParameterName = "O_OK"
                        param14.OracleType = OracleType.Double
                        param14.Direction = ParameterDirection.Output
                        comando.Parameters.Add(param14)

                        comando.ExecuteNonQuery()

                        salida.Ok = CDbl(comando.Parameters("O_OK").Value)
                        salida.CodigoAtencion = CInt(GET_CODIGO_ATENCION_SEF()) 'se usara el mismo de SEF
                        If salida.Ok = 1 Then
                            salida.Existe = CInt(comando.Parameters("O_EXISTE").Value)
                            salida.NroTarjeta = comando.Parameters("O_NUM_TARJETA").Value.ToString
                            salida.NroDocumento = comando.Parameters("O_NUM_DOCU").Value.ToString
                            salida.NombreCliente = comando.Parameters("O_NOM_CLIE").Value.ToString
                            salida.LineaInicial = CDbl(comando.Parameters("O_LINEA_INI").Value)
                            salida.LineaEgm = CDbl(comando.Parameters("O_LINEA_EGM").Value)
                            salida.LineaCons = CDbl(comando.Parameters("O_LINEA_CONS").Value)
                            salida.LineaEfex = CDbl(comando.Parameters("O_LINEA_EFEX").Value)
                            salida.NroTarjeta = nro_tarjeta
                            Dim fechaInicioVigencia As Date = CDate(comando.Parameters("O_FEC_INI_VIG").Value)
                            Dim fechaFinVigencia As Date = CDate(comando.Parameters("O_FEC_FIN_VIG").Value)

                            If fechaFinVigencia < DateTime.Today Then 'SI LA FECHA DE VENCIMIENTO YA VENCIO NO MOSTRAR NINGUN DATO
                                salida = New Salida
                                salida.Code = Constantes.CODE_NULL
                                salida.Mensaje = mensajeSinPromocion
                            Else

                                Dim dLineaConsumo As Double = CDbl(salida.LineaCons)
                                Dim dLineaActual As Double
                                Dim TramaSaldosRSAT() As String
                                Dim sSaldoRSAT As String
                                Dim nro_cuenta As String

                                If tipoTarjeta = "RSAT" Then
                                    nro_cuenta = Right(nroContrato, 12)
                                    TramaSaldosRSAT = SALDO_TARJETA_CLASICA_RSAT(nro_tarjeta, nro_cuenta, "").Split("|\t|")
                                    'TramaSaldosRSAT = ("500.00|\t|0.00|\t|500.02|\t|0.00|\t|0.00|\t|0.00|\t|17/SET/2013 - 17/OCT/2013|\t|01/NOV/2013|\t|0|\t|0.00|\t|-0.02|\t|06/11/2013|\t|17:11:20|\t|005772").Split("|\t|")
                                    sSaldoRSAT = TramaSaldosRSAT(0)
                                    sSaldoRSAT = sSaldoRSAT.Replace(",", "")
                                    dLineaActual = IIf(sSaldoRSAT.Substring(0, 5) <> "ERROR", CDbl(sSaldoRSAT), -1)
                                Else
                                    linea_actual = linea_actual.Replace(",", "")
                                    dLineaActual = CDbl(linea_actual)

                                End If

                                salida.Code = Constantes.CODE_EXITO

                                If dLineaActual < dLineaConsumo And dLineaActual > -1 Then

                                    If tipoTarjeta = "RSAT" Then
                                        salida.Mensaje = "Sr(a). " & salida.NombreCliente & "," & vbCr & "Le ofrecemos incrementar su línea de crédito hasta" _
                                                        & " S/. " & FormatNumber(salida.LineaCons, Constantes.DECIMALES_SALIDA) _
                                                        & ". Autorice su incremento sólo hasta el " _
                                                        & fechaFinVigencia.ToString(Constantes.FORMATO_FECHAS) & "."
                                    Else
                                        salida.Mensaje = "Sr(a). " & salida.NombreCliente & "," & vbCr & "Le ofrecemos incrementar su línea de crédito hasta" _
                                                        & " S/. " & FormatNumber(salida.LineaCons, Constantes.DECIMALES_SALIDA) _
                                                        & ". Autorice su incremento sólo hasta el " _
                                                        & fechaFinVigencia.ToString(Constantes.FORMATO_FECHAS) _
                                                        & ". En caso acepte el incremento de su línea de crédito, este se aplicará" _
                                                        & " 48 horas o 2 días hábiles después de su aceptación."
                                    End If

                                    salida.FechaInicioVigencia = fechaInicioVigencia.ToString(Constantes.FORMATO_FECHAS)
                                    salida.FechaFinVigencia = fechaFinVigencia.ToString(Constantes.FORMATO_FECHAS)

                                Else
                                    salida = New Salida
                                    salida.Code = Constantes.CODE_NULL
                                    salida.Mensaje = mensajeSinPromocion
                                End If


                            End If
                        Else
                            salida = New Salida
                            salida.Code = Constantes.CODE_NULL
                            salida.Mensaje = mensajeSinPromocion
                        End If

                        comando.Dispose()
                        comando = Nothing
                        conexionOraInc.Close()
                        conexionOraInc = Nothing
                    End If
                Else
                    salida.Code = Constantes.CODE_ERROR
                    salida.Mensaje = mensajeErrorInc
                End If
            End If
        Catch ex As Exception
            salida.Code = Constantes.CODE_ERROR
            salida.Mensaje = mensajeErrorInc
        End Try

        salida.Promocion = Constantes.PROMOCION_INC
        Return salida
    End Function

    'FUNCION PARA OBTENER EL CODIGO DE ATENCION DE LA OFERTA DE SUPER EFECTIVO SEF
    Private Function GET_CODIGO_ATENCION_SEF() As String

        Try
            'Realizar Conexion a la base de datos
            sMensajeError_SQL = ""

            oConexion = SQL_ConnectionOpen(Get_CadenaConexion(), sMensajeError_SQL)

            If sMensajeError_SQL <> "" Then
                sResultado_SQL = ""
            Else

                If oConexion.State = ConnectionState.Open Then

                    m_ssql = "SP_GET_CODIGO_ATENCION_SEF"

                    Dim cmd As SqlClient.SqlCommand = oConexion.CreateCommand
                    Dim lTotal As Double = 0

                    cmd.CommandTimeout = 600
                    cmd.CommandType = CommandType.StoredProcedure
                    cmd.CommandText = m_ssql



                    Dim rd_cod As SqlClient.SqlDataReader = cmd.ExecuteReader
                    sXML_SQL = ""

                    If rd_cod.Read = True Then
                        sXML_SQL = rd_cod.GetValue(0)
                    End If

                    sResultado_SQL = sXML_SQL

                    rd_cod.Close()
                    rd_cod = Nothing
                    cmd.Dispose()
                    cmd = Nothing
                    oConexion.Close()
                    oConexion = Nothing
                Else
                    sResultado_SQL = ""
                    oConexion = Nothing
                End If
            End If


        Catch ex As Exception
            sResultado_SQL = ""
            oConexion = Nothing
        End Try

        Return sResultado_SQL

    End Function

    '1|\t|1|\t|09526898|\t|SIRIP-01|\t|2|\t|999|\t|120
    <WebMethod(Description:="Actualizar oferta super efectivo.")> _
    Public Function ACTUALIZAR_OFERTA_SUPEREFECTIVO_SEF(ByVal sDataParametros As String) As String

        Dim conConexion_sef As OracleConnection
        Dim aDataParametros As Array

        'Parametros INPUT
        Dim sTipoProducto As String = "" 'Tipo de producto
        Dim sTipoDocumento As String = ""
        Dim sNroDocumento As String = ""
        Dim sNombreUsuario As String = "" 'Codigo del kiosco
        Dim sCodigoSucursal As String = "" 'sucursal de banco de acuerdo a la tabla donde están registradas las sucursales base de datos del kiosco
        Dim sNumeroCaja As String = "" 'id kiosco
        Dim sNumeroTransaccion As String = "" 'opcional si existe un numero de transacción/ticket mandar.


        Dim sTipoUsuario As String = "R"
        Dim sCodigoVendedor As String = "0"
        Dim sFechaTransaccion As String = DateTime.Now.Day.ToString("00").Trim + DateTime.Now.Month.ToString("00").Trim + DateTime.Now.Year.ToString("0000").Trim
        Dim sHoraTransaccion As String = DateTime.Now.Hour.ToString("00").Trim + DateTime.Now.Minute.ToString("00").Trim + DateTime.Now.Second.ToString("00").Trim

        aDataParametros = Split(sDataParametros.Trim, "|\t|", , CompareMethod.Text)

        sTipoProducto = aDataParametros(0)
        sTipoDocumento = aDataParametros(1)
        sNroDocumento = aDataParametros(2)
        sNombreUsuario = aDataParametros(3)
        sCodigoSucursal = aDataParametros(4)
        sNumeroCaja = aDataParametros(5) 'id kiosco
        sNumeroTransaccion = aDataParametros(6)


        Try
            conConexion_sef = New OracleConnection(mCadenaConexion_ORA_SEF)

            If Not conConexion_sef Is Nothing Then
                conConexion_sef.Open()

                If conConexion_sef.State = ConnectionState.Open Then

                    'LLAMAR AL PROCEDIMIENTO
                    Dim cmd_sef As New OracleCommand
                    Dim param1 As New OracleParameter
                    Dim param2 As New OracleParameter
                    Dim param3 As New OracleParameter
                    Dim param4 As New OracleParameter
                    Dim param5 As New OracleParameter
                    Dim param6 As New OracleParameter
                    Dim param7 As New OracleParameter
                    Dim param8 As New OracleParameter
                    Dim param9 As New OracleParameter
                    Dim param10 As New OracleParameter
                    Dim param11 As New OracleParameter
                    Dim param12 As New OracleParameter

                    'VARIABLES DE SALIDA
                    Dim sEstado As String = ""

                    cmd_sef.Connection = conConexion_sef
                    cmd_sef.CommandText = "PKG_OFERTAS_PRODUCTOS.PRC_UPDATE_OFERTA_TICKET"
                    cmd_sef.CommandType = CommandType.StoredProcedure

                    param1.ParameterName = "I_TIP_PROD"
                    param1.OracleType = OracleType.Double
                    param1.Direction = ParameterDirection.Input
                    param1.Value = CDbl(sTipoProducto.Trim) 'Tipo de producto
                    cmd_sef.Parameters.Add(param1)

                    param2.ParameterName = "I_TIP_DOC"
                    param2.OracleType = OracleType.Double
                    param2.Direction = ParameterDirection.Input
                    param2.Value = CDbl(sTipoDocumento.Trim)
                    cmd_sef.Parameters.Add(param2)

                    param3.ParameterName = "I_NUM_DOC"
                    param3.OracleType = OracleType.VarChar
                    param3.Size = 15
                    param3.Direction = ParameterDirection.Input
                    param3.Value = sNroDocumento.Trim
                    cmd_sef.Parameters.Add(param3)

                    param4.ParameterName = "I_TIP_USR"
                    param4.OracleType = OracleType.VarChar
                    param4.Size = 5
                    param4.Direction = ParameterDirection.Input
                    param4.Value = sTipoUsuario.Trim
                    cmd_sef.Parameters.Add(param4)

                    param5.ParameterName = "I_NOM_USR"
                    param5.OracleType = OracleType.VarChar
                    param5.Size = 30
                    param5.Direction = ParameterDirection.Input
                    param5.Value = sNombreUsuario.Trim
                    cmd_sef.Parameters.Add(param5)

                    param6.ParameterName = "I_COD_VEN"
                    param6.OracleType = OracleType.VarChar
                    param6.Size = 1
                    param6.Direction = ParameterDirection.Input
                    param6.Value = sCodigoVendedor.Trim
                    cmd_sef.Parameters.Add(param6)

                    param7.ParameterName = "I_COD_SUC"
                    param7.OracleType = OracleType.Double
                    param7.Direction = ParameterDirection.Input
                    param7.Value = CDbl(sCodigoSucursal.Trim)
                    cmd_sef.Parameters.Add(param7)

                    param8.ParameterName = "I_NUM_CAJ"
                    param8.OracleType = OracleType.Double
                    param8.Direction = ParameterDirection.Input
                    param8.Value = CDbl(sNumeroCaja.Trim)
                    cmd_sef.Parameters.Add(param8)

                    param9.ParameterName = "I_NUM_TRX"
                    param9.OracleType = OracleType.Double
                    param9.Direction = ParameterDirection.Input
                    param9.Value = CDbl(sNumeroTransaccion.Trim)
                    cmd_sef.Parameters.Add(param9)

                    param10.ParameterName = "I_FEC_TRX"
                    param10.OracleType = OracleType.Double
                    param10.Direction = ParameterDirection.Input
                    param10.Value = sFechaTransaccion.Trim
                    cmd_sef.Parameters.Add(param10)

                    param11.ParameterName = "I_HOR_TRX"
                    param11.OracleType = OracleType.Double
                    param11.Direction = ParameterDirection.Input
                    param11.Value = sHoraTransaccion.Trim
                    cmd_sef.Parameters.Add(param11)

                    param12.ParameterName = "O_OK"
                    param12.OracleType = OracleType.Double
                    param12.Direction = ParameterDirection.Output
                    cmd_sef.Parameters.Add(param12)

                    cmd_sef.ExecuteNonQuery()

                    sEstado = cmd_sef.Parameters("O_OK").Value.ToString


                    If sEstado.Trim = "1" Then 'OK
                        sResultado_ORA = ""
                    Else
                        sResultado_ORA = "ERROR:No se actualizó la oferta de super efectivo."
                    End If


                    cmd_sef.Dispose()
                    cmd_sef = Nothing

                    conConexion_sef.Close()
                    conConexion_sef = Nothing

                End If

            Else

                sResultado_ORA = "ERROR:No se pudo conectar al servidor de base de datos Oracle."

            End If

        Catch ex As Exception
            sResultado_ORA = "ERROR:" & ex.Message.Trim
        End Try


        Return sResultado_ORA

    End Function

    '1|\t|1|\t|09526898|\t|SIRIP-01|\t|2|\t|999|\t|120
    <WebMethod(Description:="Actualizar oferta super efectivo.")> _
    Public Function ACTUALIZAR_OFERTA_SUPEREFECTIVO_SEF_TICKET(ByVal sDataParametros As String) As String

        Dim conConexion_sef As OracleConnection
        Dim aDataParametros As Array

        'Parametros INPUT
        Dim sTipoProducto As String = "" 'Tipo de producto
        Dim sTipoDocumento As String = ""
        Dim sNroDocumento As String = ""
        Dim sNombreUsuario As String = "" 'Codigo del kiosco
        Dim sCodigoSucursal As String = "" 'sucursal de banco de acuerdo a la tabla donde están registradas las sucursales base de datos del kiosco
        Dim sNumeroCaja As String = "" 'id kiosco
        Dim sNumeroTransaccion As String = "" 'opcional si existe un numero de transacción/ticket mandar.


        Dim sTipoUsuario As String = "R"
        Dim sCodigoVendedor As String = "0"
        Dim sFechaTransaccion As String = DateTime.Now.Day.ToString("00").Trim + DateTime.Now.Month.ToString("00").Trim + DateTime.Now.Year.ToString("0000").Trim
        Dim sHoraTransaccion As String = DateTime.Now.Hour.ToString("00").Trim + DateTime.Now.Minute.ToString("00").Trim + DateTime.Now.Second.ToString("00").Trim

        aDataParametros = Split(sDataParametros.Trim, "|\t|", , CompareMethod.Text)

        sTipoProducto = aDataParametros(0)
        sTipoDocumento = aDataParametros(1)
        sNroDocumento = aDataParametros(2)
        sNombreUsuario = aDataParametros(3)
        sCodigoSucursal = aDataParametros(4)
        sNumeroCaja = aDataParametros(5) 'id kiosco
        sNumeroTransaccion = aDataParametros(6)


        Try
            conConexion_sef = New OracleConnection(mCadenaConexion_ORA_SEF)

            If Not conConexion_sef Is Nothing Then
                conConexion_sef.Open()

                If conConexion_sef.State = ConnectionState.Open Then

                    'LLAMAR AL PROCEDIMIENTO
                    Dim cmd_sef As New OracleCommand
                    Dim param1 As New OracleParameter
                    Dim param2 As New OracleParameter
                    Dim param3 As New OracleParameter
                    Dim param4 As New OracleParameter
                    Dim param5 As New OracleParameter
                    Dim param6 As New OracleParameter
                    Dim param7 As New OracleParameter
                    Dim param8 As New OracleParameter
                    Dim param9 As New OracleParameter
                    Dim param10 As New OracleParameter
                    Dim param11 As New OracleParameter
                    Dim param12 As New OracleParameter

                    'VARIABLES DE SALIDA
                    Dim sEstado As String = ""

                    cmd_sef.Connection = conConexion_sef
                    '17/12/2014 CAMBIO PARA ACTUALIZAR OFERTA SEF ANTERIOR
                    'cmd_sef.CommandText = "PKG_OFERTAS_PRODUCTOS.PRC_UPDATE_OFERTA_TICKET_SEF"
                    cmd_sef.CommandText = "PKG_OFERTAS_PRODUCTOS.PRC_UPDATE_OFERTA_TICKET"
                    cmd_sef.CommandType = CommandType.StoredProcedure

                    param1.ParameterName = "I_TIP_PROD"
                    param1.OracleType = OracleType.Double
                    param1.Direction = ParameterDirection.Input
                    param1.Value = CDbl(sTipoProducto.Trim) 'Tipo de producto
                    cmd_sef.Parameters.Add(param1)

                    param2.ParameterName = "I_TIP_DOC"
                    param2.OracleType = OracleType.Double
                    param2.Direction = ParameterDirection.Input
                    param2.Value = CDbl(sTipoDocumento.Trim)
                    cmd_sef.Parameters.Add(param2)

                    param3.ParameterName = "I_NUM_DOC"
                    param3.OracleType = OracleType.VarChar
                    param3.Size = 15
                    param3.Direction = ParameterDirection.Input
                    param3.Value = sNroDocumento.Trim
                    cmd_sef.Parameters.Add(param3)

                    param4.ParameterName = "I_TIP_USR"
                    param4.OracleType = OracleType.VarChar
                    param4.Size = 5
                    param4.Direction = ParameterDirection.Input
                    param4.Value = sTipoUsuario.Trim
                    cmd_sef.Parameters.Add(param4)

                    param5.ParameterName = "I_NOM_USR"
                    param5.OracleType = OracleType.VarChar
                    param5.Size = 30
                    param5.Direction = ParameterDirection.Input
                    param5.Value = sNombreUsuario.Trim
                    cmd_sef.Parameters.Add(param5)

                    param6.ParameterName = "I_COD_VEN"
                    param6.OracleType = OracleType.VarChar
                    param6.Size = 1
                    param6.Direction = ParameterDirection.Input
                    param6.Value = sCodigoVendedor.Trim
                    cmd_sef.Parameters.Add(param6)

                    param7.ParameterName = "I_COD_SUC"
                    param7.OracleType = OracleType.Double
                    param7.Direction = ParameterDirection.Input
                    param7.Value = CDbl(sCodigoSucursal.Trim)
                    cmd_sef.Parameters.Add(param7)

                    param8.ParameterName = "I_NUM_CAJ"
                    param8.OracleType = OracleType.Double
                    param8.Direction = ParameterDirection.Input
                    param8.Value = CDbl(sNumeroCaja.Trim)
                    cmd_sef.Parameters.Add(param8)

                    param9.ParameterName = "I_NUM_TRX"
                    param9.OracleType = OracleType.Double
                    param9.Direction = ParameterDirection.Input
                    param9.Value = CDbl(sNumeroTransaccion.Trim)
                    cmd_sef.Parameters.Add(param9)

                    param10.ParameterName = "I_FEC_TRX"
                    param10.OracleType = OracleType.Double
                    param10.Direction = ParameterDirection.Input
                    param10.Value = sFechaTransaccion.Trim
                    cmd_sef.Parameters.Add(param10)

                    param11.ParameterName = "I_HOR_TRX"
                    param11.OracleType = OracleType.Double
                    param11.Direction = ParameterDirection.Input
                    param11.Value = sHoraTransaccion.Trim
                    cmd_sef.Parameters.Add(param11)

                    param12.ParameterName = "O_OK"
                    param12.OracleType = OracleType.Double
                    param12.Direction = ParameterDirection.Output
                    cmd_sef.Parameters.Add(param12)

                    cmd_sef.ExecuteNonQuery()

                    sEstado = cmd_sef.Parameters("O_OK").Value.ToString


                    If sEstado.Trim = "1" Then 'OK
                        sResultado_ORA = ""
                    Else
                        sResultado_ORA = "ERROR:No se actualizó la oferta de super efectivo."
                    End If


                    cmd_sef.Dispose()
                    cmd_sef = Nothing

                    conConexion_sef.Close()
                    conConexion_sef = Nothing

                End If

            Else

                sResultado_ORA = "ERROR:No se pudo conectar al servidor de base de datos Oracle."

            End If

        Catch ex As Exception
            sResultado_ORA = "ERROR:" & ex.Message.Trim
        End Try


        Return sResultado_ORA

    End Function

    ''' <summary>
    ''' Actualiza la oferta de incremento de linea en la base de datos Oracle mediante un package
    ''' </summary>
    ''' <param name="nroContrato">Numero de Contrato</param>
    ''' <param name="codigoSucursal">Codigo de Sucursal</param>
    ''' <param name="nroCaja">Numero de Caja</param>
    ''' <param name="nroTransaccion">Numero de Transaccion</param>
    ''' <returns>Cadena de texto con error o exito de la actualización</returns>
    ''' <remarks></remarks>
    <WebMethod(Description:="Actualizar oferta de incremento de linea.")> _
    Public Function ACTUALIZAR_OFERTA_INCREMENTO_LINEA(ByVal nroContrato As String, ByVal codigoSucursal As Double, _
                                                       ByVal nroCaja As String, ByVal nroTransaccion As Double, _
                                                       ByVal nroTarjeta As String, ByVal nombreCliente As String, _
                                                       ByVal lineaInicial As Double, ByVal lineaEgm As Double, ByVal lineaCons As Double, _
                                                       ByVal lineaEfex As Double, ByVal fechaInicioVigencia As String, _
                                                       ByVal fechaFinVigencia As String, ByVal tipoTarjeta As String) As Salida
        Dim resultado As New Salida
        'Dim eexitoMC As String = "Sr(a). " & nombreCliente & ", Felicitaciones ¡!" & vbCrLf & vbCrLf _
        '                            & "Su incremento ha sido procesado. En 48 horas su nueva línea de" _
        '                            & " crédito será S/. " & FormatNumber(lineaCons, Constantes.DECIMALES_SALIDA) _
        '                            & vbCrLf & vbCrLf & "BANCO RIPLEY"
        'Dim eexitoImpresionMC As String = "Sr(a). " & nombreCliente & ", Felicitaciones ¡!" & vbCrLf & vbCrLf _
        '                    & "Su incremento ha sido procesado. En 48 horas su nueva linea de" _
        '                    & " credito sera S/. " & FormatNumber(lineaCons, Constantes.DECIMALES_SALIDA)
        'Dim eexitoRSAT As String = "Sr(a). " & nombreCliente & ", Felicitaciones ¡!" & vbCrLf & vbCrLf _
        '                            & "Su solicitud de incremento de línea ha sido realizada satisfactoriamente." & vbCrLf & vbCrLf _
        '                            & "Su nueva línea de crédito es S/. " & FormatNumber(lineaCons, Constantes.DECIMALES_SALIDA) _
        '                            & vbCrLf & vbCrLf & "BANCO RIPLEY"
        'Dim eexitoImpresionRSAT As String = "Sr(a). " & nombreCliente & ", Felicitaciones ¡!" & vbCrLf & vbCrLf _
        '                        & "Su solicitud de incremento de linea ha sido realizada satisfactoriamente." & vbCrLf & vbCrLf _
        '                        & "Su nueva linea de credito es S/. " & FormatNumber(lineaCons, Constantes.DECIMALES_SALIDA)
        'resultado.Code = Constantes.CODE_EXITO
        'If tipoTarjeta = "RSAT" Then
        '    resultado.Mensaje = eexitoRSAT
        '    resultado.MensajeImpresion = eexitoImpresionRSAT
        'ElseIf tipoTarjeta = "MC" Then
        '    resultado.Mensaje = eexitoMC
        '    resultado.MensajeImpresion = eexitoImpresionMC
        'End If
        'Return resultado

        Try


            Dim conexionOracle As OracleConnection

            Dim tipoUsuario As String = Constantes.TIPO_USUARIO
            Dim codigoVendedor As String = Constantes.CODIGO_VENDEDOR

            Dim fechaTransaccion As Double = CDbl(DateTime.Now.ToString("ddMMyyyy"))
            Dim horaTransaccion As Double = CDbl(DateTime.Now.ToString("HHmmss"))

            If tipoTarjeta = "RSAT" Then 'Para el caso de RSAT
                Dim errorRSAT As String = "Estimado cliente, su línea de crédito no ha podido ser incrementada, " & _
                                            "por favor comuníquese con nosotros llamándonos al ripleyfono en Lima 611-5757, " & _
                                            "Chimbote 60-4407 y otras provincias al 60-5757, o acérquese a las agencias del " & _
                                            "Banco ubicadas en Tiendas Ripley y Max."

                Dim exitoRSAT As String = "Sr(a). " & nombreCliente & ", ¡Felicitaciones!" & vbCrLf & vbCrLf _
                                        & "Su solicitud de incremento de línea ha sido realizada satisfactoriamente." & vbCrLf & vbCrLf _
                                        & "Su nueva línea de crédito es S/. " & FormatNumber(lineaCons, Constantes.DECIMALES_SALIDA) _
                                        & vbCrLf & vbCrLf & "BANCO RIPLEY"
                Dim exitoImpresionRSAT As String = "Sr(a). " & nombreCliente & "," & vbCrLf & vbCrLf & "¡Felicitaciones!" & vbCrLf & vbCrLf _
                                        & "Su solicitud de incremento de linea ha sido realizada" & vbCrLf & "satisfactoriamente." & vbCrLf & vbCrLf _
                                        & "Su nueva linea de crédito es S/. " & FormatNumber(lineaCons, Constantes.DECIMALES_SALIDA)

                '<INI PRY-0794-01 - Danny Herrera 01/08/2014
                Dim oBN_Incremento_Linea As New BN_Incremento_Linea
                Dim oParametros As New Parametros
                Dim oResultado As New Aprobacion

                oParametros.Nro_Contrato = nroContrato
                oParametros.Cod_Sucursal = codigoSucursal
                oParametros.Cod_Oficina = codigoSucursal
                oParametros.Cod_Usuario_Plataforma = "RPMATICO"
                oParametros.Fecha_Conta = Date.Now
                oParametros.Linea_EGM = lineaEgm
                oParametros.Linea_1 = lineaCons
                oParametros.Linea_2 = lineaEfex
                oParametros.Nro_Caja = nroCaja
                oParametros.Nro_Transaccion = nroTransaccion

                oResultado = oBN_Incremento_Linea.Aprobar_Oferta(oParametros)

                If oResultado.Cod_Respuesta = "00" Then

                    resultado.Code = Constantes.CODE_EXITO
                    resultado.Mensaje = exitoRSAT
                    resultado.MensajeImpresion = exitoImpresionRSAT

                Else
                    resultado.Code = 2
                    resultado.Mensaje = errorRSAT

                End If
                '<FIN PRY-0794-01 - Danny Herrera 01/08/2014

                '<INI PRY-0794-01 - Danny Herrera 01/08/2014
                ''Dim servicio As New ServiceRSat
                'Dim objIn As New BEInIncrementoLinea
                'Dim objOUT As New BEOutIncrementoLinea
                'objIn.sCodUsuario = codigoVendedor
                'objIn.sContrato = nroContrato
                'objIn.sSucursal = codigoSucursal
                'objIn.dLineaCreditoNueva_0 = lineaEgm
                'objIn.dLineaCreditoNueva_1 = lineaCons
                'objIn.dLineaCreditoNueva_2 = lineaEfex

                '    objOUT = ActualizarLinea(objIn)

                'If objOUT.sResultadoActualizacionLinea_1 = "1" Then 'Exito
                '    Try
                '        conexionOracle = New OracleConnection(cadenaConexionOraInc)

                '        If Not conexionOracle Is Nothing Then
                '            conexionOracle.Open()
                '            If conexionOracle.State = ConnectionState.Open Then
                '                'LLAMAR AL PROCEDIMIENTO
                '                Dim comando As New OracleCommand

                '                Dim param1 As New OracleParameter
                '                Dim param2 As New OracleParameter
                '                Dim param3 As New OracleParameter
                '                Dim param4 As New OracleParameter
                '                Dim param5 As New OracleParameter
                '                Dim param6 As New OracleParameter
                '                Dim param7 As New OracleParameter
                '                Dim param8 As New OracleParameter
                '                Dim param9 As New OracleParameter
                '                Dim param10 As New OracleParameter
                '                Dim param11 As New OracleParameter
                '                Dim param12 As New OracleParameter
                '                Dim param13 As New OracleParameter

                '                'VARIABLES DE SALIDA
                '                Dim estado As Double

                '                comando.Connection = conexionOracle
                '                comando.CommandText = "PKG_OFERTAS_INC_LINEA.PRC_UPDATE_OFERTA_INC_LINEA"
                '                comando.CommandType = CommandType.StoredProcedure

                '                param1.ParameterName = "I_NUM_CONTRATO"
                '                param1.OracleType = OracleType.VarChar
                '                param1.Direction = ParameterDirection.Input
                '                param1.Size = 20
                '                param1.Value = nroContrato
                '                comando.Parameters.Add(param1)

                '                param2.ParameterName = "I_COD_SUC"
                '                param2.OracleType = OracleType.Number
                '                param2.Direction = ParameterDirection.Input
                '                param2.Value = codigoSucursal
                '                comando.Parameters.Add(param2)

                '                param3.ParameterName = "I_NUM_CAJ"
                '                param3.OracleType = OracleType.VarChar
                '                param3.Direction = ParameterDirection.Input
                '                param3.Size = 20
                '                param3.Value = nroCaja
                '                comando.Parameters.Add(param3)

                '                param4.ParameterName = "I_NUM_TRX"
                '                param4.OracleType = OracleType.Number
                '                param4.Direction = ParameterDirection.Input
                '                param4.Value = nroTransaccion
                '                comando.Parameters.Add(param4)

                '                param8.ParameterName = "I_LINEA_EGM_OLD"
                '                param8.OracleType = OracleType.Number
                '                param8.Direction = ParameterDirection.Input
                '                param8.Value = objOUT.dLineaCreditoAnterior_0
                '                comando.Parameters.Add(param8)

                '                param9.ParameterName = "I_LINEA1_OLD"
                '                param9.OracleType = OracleType.Number
                '                param9.Direction = ParameterDirection.Input
                '                param9.Value = objOUT.dLineaCreditoAnterior_1
                '                comando.Parameters.Add(param9)

                '                param10.ParameterName = "I_LINEA2_OLD"
                '                param10.OracleType = OracleType.Number
                '                param10.Direction = ParameterDirection.Input
                '                param10.Value = objOUT.dLineaCreditoAnterior_2
                '                comando.Parameters.Add(param10)

                '                param11.ParameterName = "I_UPD_LIM_OK"
                '                param11.OracleType = OracleType.Number
                '                param11.Direction = ParameterDirection.Input
                '                param11.Value = CDbl(objOUT.sResultadoActualizacionLinea_0)
                '                comando.Parameters.Add(param11)

                '                param12.ParameterName = "I_UPD_LIN_OK_1"
                '                param12.OracleType = OracleType.Number
                '                param12.Direction = ParameterDirection.Input
                '                param12.Value = CDbl(objOUT.sResultadoActualizacionLinea_1)
                '                comando.Parameters.Add(param12)

                '                param13.ParameterName = "I_UPD_LIN_OK_2"
                '                param13.OracleType = OracleType.Number
                '                param13.Direction = ParameterDirection.Input
                '                param13.Value = CDbl(objOUT.sResultadoActualizacionLinea_2)
                '                comando.Parameters.Add(param13)

                '                param5.ParameterName = "I_FEC_TRX"
                '                param5.OracleType = OracleType.Number
                '                param5.Direction = ParameterDirection.Input
                '                param5.Value = fechaTransaccion
                '                comando.Parameters.Add(param5)

                '                param6.ParameterName = "I_HOR_TRX"
                '                param6.OracleType = OracleType.Number
                '                param6.Direction = ParameterDirection.Input
                '                param6.Value = horaTransaccion
                '                comando.Parameters.Add(param6)

                '                param7.ParameterName = "O_OK"
                '                param7.OracleType = OracleType.Double
                '                param7.Direction = ParameterDirection.Output
                '                comando.Parameters.Add(param7)

                '                comando.ExecuteNonQuery()

                '                estado = CDbl(comando.Parameters("O_OK").Value)

                '                If estado = 1 Then 'OK
                '                    resultado.Code = Constantes.CODE_EXITO
                '                    resultado.Mensaje = exitoRSAT
                '                    resultado.MensajeImpresion = exitoImpresionRSAT
                '                Else
                '                        resultado.Code = 1
                '                    resultado.Mensaje = errorRSAT
                '                End If

                '                comando.Dispose()
                '                comando = Nothing

                '                conexionOracle.Close()
                '                conexionOracle = Nothing
                '            End If
                '        Else
                '                resultado.Code = 2
                '            resultado.Mensaje = errorRSAT
                '        End If
                '        Catch EX As Exception
                '            resultado.Code = 3
                '            resultado.Mensaje = EX.Message
                '    End Try
                'Else 'Error
                '        resultado.Code = 4
                '        resultado.Mensaje = objOUT.sMsgRespuesta
                'End If
                '<FIN PRY-0794-01 - Danny Herrera 01/08/2014

            Else
                If tipoTarjeta = "MC" And fechaInicioVigencia.Length = 10 And fechaFinVigencia.Length = 10 Then 'Para el caso de Marca Abierta
                    Dim errorMC As String = "Estimado cliente, su línea de crédito no ha podido ser incrementada, " & _
                                            "por favor comuníquese con nosotros llamándonos al ripleyfono en Lima 611-5757, " & _
                                            "Chimbote 60-4407 y otras provincias al 60-5757, o acérquese a las agencias del " & _
                                            "Banco ubicadas en Tiendas Ripley y Max."
                    Dim exitoMC As String = "Sr(a). " & nombreCliente & "," & vbCrLf & "¡Felicitaciones!" & vbCrLf _
                                        & "Su incremento ha sido procesado. En 48 horas o 2 días hábiles su nueva línea de" _
                                        & " crédito será S/. " & FormatNumber(lineaCons, Constantes.DECIMALES_SALIDA) _
                                        & vbCrLf & vbCrLf & "BANCO RIPLEY"
                    Dim exitoImpresionMC As String = "Sr(a). " & nombreCliente & "," & vbCrLf & vbCrLf & "¡FELICITACIONES!" & vbCrLf & vbCrLf _
                                        & "SU INCREMENTO HA SIDO PROCESADO. EN 48 HORAS O 2 DIAS" & vbCrLf & _
                                        "HABILES SU NUEVA LINEA DE CREDITO SERA S/.  " & FormatNumber(lineaCons, Constantes.DECIMALES_SALIDA)

                    Try
                        conexionOracle = New OracleConnection(cadenaConexionOraInc)

                        If Not conexionOracle Is Nothing Then
                            conexionOracle.Open()
                            If conexionOracle.State = ConnectionState.Open Then
                                'LLAMAR AL PROCEDIMIENTO
                                Dim comando As New OracleCommand

                                Dim param1 As New OracleParameter
                                Dim param2 As New OracleParameter
                                Dim param3 As New OracleParameter
                                Dim param4 As New OracleParameter
                                Dim param5 As New OracleParameter
                                Dim param6 As New OracleParameter
                                Dim param7 As New OracleParameter
                                Dim param8 As New OracleParameter
                                Dim param9 As New OracleParameter
                                Dim param10 As New OracleParameter
                                Dim param11 As New OracleParameter
                                Dim param12 As New OracleParameter
                                Dim param13 As New OracleParameter

                                'VARIABLES DE SALIDA
                                Dim estado As Double

                                comando.Connection = conexionOracle
                                comando.CommandText = "PKG_OFERTAS_INC_LINEA.PRC_OFERTA_INC_LINEA_BATCH"
                                comando.CommandType = CommandType.StoredProcedure

                                param1.ParameterName = "I_NUM_CUENTA"
                                param1.OracleType = OracleType.VarChar
                                param1.Direction = ParameterDirection.Input
                                param1.Size = 19
                                param1.Value = nroContrato
                                comando.Parameters.Add(param1)

                                param2.ParameterName = "I_NUM_TARJETA"
                                param2.OracleType = OracleType.VarChar
                                param2.Direction = ParameterDirection.Input
                                param2.Size = 19
                                param2.Value = nroTarjeta
                                comando.Parameters.Add(param2)

                                param3.ParameterName = "I_NOM_CLIE"
                                param3.OracleType = OracleType.VarChar
                                param3.Direction = ParameterDirection.Input
                                param3.Size = 60
                                param3.Value = nombreCliente
                                comando.Parameters.Add(param3)

                                param4.ParameterName = "I_COD_SUC"
                                param4.OracleType = OracleType.Number
                                param4.Direction = ParameterDirection.Input
                                param4.Value = codigoSucursal
                                comando.Parameters.Add(param4)

                                param8.ParameterName = "I_NUM_CAJ"
                                param8.OracleType = OracleType.VarChar
                                param8.Direction = ParameterDirection.Input
                                param8.Size = 20
                                param8.Value = nroCaja
                                comando.Parameters.Add(param8)

                                param9.ParameterName = "I_LINEA_INICIAL"
                                param9.OracleType = OracleType.Number
                                param9.Direction = ParameterDirection.Input
                                param9.Value = lineaInicial
                                comando.Parameters.Add(param9)

                                param10.ParameterName = "I_LINEA_FINAL"
                                param10.OracleType = OracleType.Number
                                param10.Direction = ParameterDirection.Input
                                param10.Value = lineaCons
                                comando.Parameters.Add(param10)

                                Dim partsFechaInicio() As String = fechaInicioVigencia.Split("/")
                                Dim fechaInicio As Date
                                fechaInicio = New Date(partsFechaInicio(2), partsFechaInicio(1), partsFechaInicio(0))

                                param11.ParameterName = "I_FEC_INI_VIG"
                                param11.OracleType = OracleType.DateTime
                                param11.Direction = ParameterDirection.Input
                                param11.Value = fechaInicio
                                comando.Parameters.Add(param11)

                                Dim partsFechaFin() As String = fechaFinVigencia.Split("/")
                                Dim fechaFin As Date
                                fechaFin = New Date(partsFechaFin(2), partsFechaFin(1), partsFechaFin(0))

                                param12.ParameterName = "I_FEC_FIN_VIG"
                                param12.OracleType = OracleType.DateTime
                                param12.Direction = ParameterDirection.Input
                                param12.Value = fechaFin
                                comando.Parameters.Add(param12)

                                param5.ParameterName = "I_FEC_TRX"
                                param5.OracleType = OracleType.Double
                                param5.Direction = ParameterDirection.Input
                                param5.Value = fechaTransaccion
                                comando.Parameters.Add(param5)

                                param6.ParameterName = "I_HOR_TRX"
                                param6.OracleType = OracleType.Double
                                param6.Direction = ParameterDirection.Input
                                param6.Value = horaTransaccion
                                comando.Parameters.Add(param6)

                                param13.ParameterName = "I_NUM_TCK"
                                param13.OracleType = OracleType.Double
                                param13.Direction = ParameterDirection.Input
                                param13.Value = nroTransaccion
                                comando.Parameters.Add(param13)


                                param7.ParameterName = "O_OK"
                                param7.OracleType = OracleType.Double
                                param7.Direction = ParameterDirection.Output
                                comando.Parameters.Add(param7)

                                comando.ExecuteNonQuery()

                                estado = CDbl(comando.Parameters("O_OK").Value)

                                If estado = 1 Then 'OK
                                    resultado.Code = Constantes.CODE_EXITO
                                    resultado.Mensaje = exitoMC
                                    resultado.MensajeImpresion = exitoImpresionMC
                                Else
                                    resultado.Code = Constantes.CODE_ERROR
                                    resultado.Mensaje = errorMC
                                End If

                                comando.Dispose()
                                comando = Nothing

                                conexionOracle.Close()
                                conexionOracle = Nothing
                            End If
                        Else
                            resultado.Code = Constantes.CODE_ERROR
                            resultado.Mensaje = errorMC
                        End If
                    Catch ex As Exception
                        resultado.Code = Constantes.CODE_ERROR
                        resultado.Mensaje = errorMC
                    End Try

                Else 'Para el caso de SICRON. Aca nunca debe entrar, ya que incremento de linea es solo para RSAT o MC
                    resultado.Code = Constantes.CODE_ERROR
                    resultado.Mensaje = "Ocurrio un error inesperado"
                End If
            End If
        Catch ex As Exception
            resultado.Code = 5
            resultado.Mensaje = ex.Message
        End Try

        Return resultado
    End Function


    '<WebMethod(Description:="GUARDAR LOG DE CONSULTAS EN INCREMENTO DE LINEA.")> _
    'Public Function SAVE_LOG_CONSULTAS_INC(ByVal codigoSucursalBanco As Double, ByVal codigoKiosko As String, ByVal codigoTipoDocumento As Double, _
    '                                        ByVal nroDocumento As String, ByVal opcion As String, ByVal codigoAtencion As Integer) As Boolean
    '    Dim estado As Boolean

    '    Try
    '        estado = False

    '        Dim conexionSql As SqlConnection
    '        Dim comando As SqlCommand

    '        'Conexion a la base de datos
    '        Dim mensajeErrorSql As String = String.Empty
    '        conexionSql = SQL_ConnectionOpen(Get_CadenaConexion(), mensajeErrorSql)
    '        If Not String.IsNullOrEmpty(mensajeErrorSql) Then
    '            estado = False
    '        Else
    '            If conexionSql.State = ConnectionState.Open Then
    '                Dim sp_nombre As String = "SP_LOG_CONSULTAS_INC"

    '                comando = conexionSql.CreateCommand
    '                comando.CommandText = sp_nombre
    '                comando.CommandType = CommandType.StoredProcedure

    '                comando.Parameters.AddWithValue("@COD_SUCURSAL_BAN", codigoSucursalBanco.ToString)
    '                comando.Parameters.AddWithValue("@CODIGO_KIOSCO", codigoKiosko)
    '                comando.Parameters.AddWithValue("@TIPO_DOC", codigoTipoDocumento.ToString)
    '                comando.Parameters.AddWithValue("@NRO_DOC", nroDocumento)
    '                comando.Parameters.AddWithValue("@OPCION", opcion)
    '                comando.Parameters.AddWithValue("@COD_ATENCION", codigoAtencion)

    '                comando.ExecuteNonQuery()
    '                'operacion con exito
    '                estado = True

    '                comando.Dispose()
    '                conexionSql.Close()

    '                comando = Nothing
    '                conexionSql = Nothing
    '            Else
    '                estado = False
    '                conexionSql = Nothing
    '            End If
    '        End If
    '    Catch ex As Exception
    '        estado = False
    '        oConexion = Nothing
    '    End Try

    '    Return estado
    'End Function


    <WebMethod(Description:="GUARDAR LOG DE CONSULTAS EN INCREMENTO DE LINEA.")> _
    Public Function SAVE_LOG_CONSULTAS_INC(ByVal codigoSucursalBanco As Double, _
                                           ByVal codigoKiosko As String, _
                                           ByVal codigoTipoDocumento As Double, _
                                            ByVal nroDocumento As String, _
                                            ByVal opcion As String, _
                                            ByVal codigoAtencion As Integer, _
                                            ByVal numeroTarjeta As String, _
                                            ByVal tipoSistema As String, _
                                            ByVal lineaInicial As String, _
                                            ByVal lineaFinal As String, _
                                            ByVal inicioVigencia As String, _
                                            ByVal finVigencia As String) As Boolean
        Dim estado As Boolean

        Try
            estado = False

            Dim conexionSql As SqlConnection
            Dim comando As SqlCommand
            Dim TarjetaBin As New TarjetaBin

            Dim dFechaIni As Date = New Date(inicioVigencia.Substring(6, 4), inicioVigencia.Substring(3, 2), inicioVigencia.Substring(0, 2))
            Dim dFechaFin As Date = New Date(finVigencia.Substring(6, 4), finVigencia.Substring(3, 2), finVigencia.Substring(0, 2))
            Dim binnTarjeta As String = numeroTarjeta.Substring(0, 6)

            TarjetaBin = BNTarjetaBin.Instancia.ObtenerTarjetaBin(binnTarjeta)

            'Conexion a la base de datos
            Dim mensajeErrorSql As String = String.Empty
            conexionSql = SQL_ConnectionOpen(Get_CadenaConexion(), mensajeErrorSql)
            If Not String.IsNullOrEmpty(mensajeErrorSql) Then
                estado = False
            Else
                If conexionSql.State = ConnectionState.Open Then
                    Dim sp_nombre As String = "SP_LOG_CONSULTAS_INC_EST"

                    comando = conexionSql.CreateCommand
                    comando.CommandText = sp_nombre
                    comando.CommandType = CommandType.StoredProcedure

                    comando.Parameters.AddWithValue("@COD_SUCURSAL_BAN", codigoSucursalBanco.ToString)
                    comando.Parameters.AddWithValue("@CODIGO_KIOSCO", codigoKiosko)
                    comando.Parameters.AddWithValue("@TIPO_DOC", codigoTipoDocumento.ToString)
                    comando.Parameters.AddWithValue("@NRO_DOC", nroDocumento)
                    comando.Parameters.AddWithValue("@OPCION", opcion)
                    comando.Parameters.AddWithValue("@COD_ATENCION", codigoAtencion)
                    comando.Parameters.AddWithValue("@NUMERO_TARJETA", numeroTarjeta)
                    comando.Parameters.AddWithValue("@BIN_TARJETA", binnTarjeta)
                    comando.Parameters.AddWithValue("@TIPO_TARJETA", TarjetaBin.Nombre)
                    comando.Parameters.AddWithValue("@TIPO_SISTEMA", tipoSistema)
                    comando.Parameters.AddWithValue("@LINEA_INICIAL", lineaInicial)
                    comando.Parameters.AddWithValue("@LINEA_FINAL", lineaFinal)
                    comando.Parameters.AddWithValue("@INICIO_VIGENCIA", dFechaIni)
                    comando.Parameters.AddWithValue("@FIN_VIGENCIA", dFechaFin)
                    comando.ExecuteNonQuery()
                    'operacion con exito
                    estado = True

                    comando.Dispose()
                    conexionSql.Close()

                    comando = Nothing
                    conexionSql = Nothing
                Else
                    estado = False
                    conexionSql = Nothing
                End If
            End If
        Catch ex As Exception
            ErrorLog(ex.Message)
            estado = False
            oConexion = Nothing
        End Try

        Return estado
    End Function

    Private Function MOSTRAR_SUPEREFECTIVO(ByVal sNroCuenta As String, ByVal NroDNI As String, ByVal SERVIDOR As TServidor) As String
        Dim Respuesta As String
        Respuesta = String.Empty

        Select Case SERVIDOR

            Case TServidor.SICRON
                Respuesta = MOSTRAR_SUPEREFECTIVO_SICRON(sNroCuenta)
            Case TServidor.RSAT
                Respuesta = MOSTRAR_SUPEREFECTIVO_RSAT(NroDNI)
            Case Else
                Respuesta = "ERROR:Servidor no especificado"
        End Select

        Return Respuesta


    End Function


    'ESTA FUNCION SIRVE PARA BUSCAR SI TIENE SUPEREFECTIVO A MOSTRAR DEVUELVE EL MONTO DE SUPEREFECTIVO Y LA FECHA DE VENCIMIENTO.
    Private Function MOSTRAR_SUPEREFECTIVO_SICRON(ByVal sNroCuenta As String) As String
        Dim sRespuesta_ As String = ""
        Dim sParametros_ As String = ""
        Dim sMensajeErrorUsuario As String = ""
        Dim sXML As String = ""

        Dim actions_ As Long = 1
        Dim inetputBuff_ As String = ""
        Dim outpputBuff_ As String = ""
        Dim errorMsg_ As String = ""


        Try

            If sNroCuenta.Trim.Length = 8 Then

                'Instancia al mirapiweb
                Dim obSendMirror_ As ClsTxMirapi = Nothing
                obSendMirror_ = New ClsTxMirapi()

                sParametros_ = "00002" & sNroCuenta.Trim
                inetputBuff_ = "      " + "V162" + sParametros_

                sRespuesta_ = obSendMirror_.ExecuteTX(actions_.ToString, inetputBuff_, outpputBuff_, TIME_OUT_SERVER, SERVIDOR_MIRROR_DESTINO, SERVIDOR_MIRROR_NODE, PUERTO, errorMsg_)

                'sRespuesta_ = "1"
                'outpputBuff_ = "12345678AU            0091294511B000000000750000000000000000020120615000000000000000003800002012051500000000000000000000000101                                                                                                                                                                                                                                                                28100000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000"


                If sRespuesta_ = "0" Then 'EXITO

                    If outpputBuff_.Length > 0 Then sRespuesta_ = outpputBuff_.Substring(8, outpputBuff_.Length - 8)

                    'Evaluar Respuesta si es ERROR,  RU, AU (Respuesta correcta)
                    If sRespuesta_.Trim.Length > 0 Then
                        If Left(sRespuesta_.Trim, 2) = "RU" Then 'Mensaje de Validacion al usuario. 
                            sMensajeErrorUsuario = Mid(sRespuesta_.Trim, 18, Len(sRespuesta_.Trim))
                            sRespuesta_ = "|\t|"
                        Else
                            If Left(sRespuesta_.Trim, 2) = "AU" Then

                                'Recuperar Data
                                Dim sDisponibleSuperEfectivo As String = ""
                                Dim sFechaVence As String = ""
                                Dim sDato As String = ""
                                Dim sCondicion As String = ""

                                sDato = CLng(Trim(Mid(sRespuesta_, 26, 14))).ToString
                                sDisponibleSuperEfectivo = Format(Val(SET_MONTO_DECIMAL(sDato)), "##,##0.00")

                                sFechaVence = Mid(sRespuesta_, 54, 8)

                                Dim vFecHoy As Date
                                Dim vFecVence As Date

                                vFecHoy = Date.Now
                                sFechaVence = FormatearFecha_DDMMYYYY(sFechaVence.Trim)

                                If sFechaVence.Trim <> "" Then
                                    'VALIDAR FECHAS
                                    If sFechaVence.Trim.Length = 10 Then
                                        vFecVence = CDate(sFechaVence.Trim)
                                        If vFecVence > vFecHoy Then 'SI LA FECHA DE VENCIMIENTO ES MAYOR A LA FECHA ACTUAL ENTONCES MOSTRAR LOS DATOS DEL SUPEREFECTIVO
                                            sXML = sDisponibleSuperEfectivo.Trim & "|\t|" & sFechaVence.Trim
                                        Else
                                            sXML = "|\t|"
                                        End If

                                    End If
                                Else
                                    sXML = "|\t|"
                                End If

                                sRespuesta_ = sXML

                            Else 'Cualquier Otro Caso
                                sRespuesta_ = "|\t|"
                            End If


                        End If

                    Else 'Sino devuelve nada
                        sRespuesta_ = "|\t|"
                    End If

                Else 'si ocurrio un error en la llamada de la consulta del mirapi web
                    sRespuesta_ = "|\t|"
                End If


            Else
                'Mostrar Mensaje de Error
                sRespuesta_ = "|\t|"
            End If

        Catch ex As Exception
            'save log error

            sRespuesta_ = "|\t|"

        End Try

        Return sRespuesta_

    End Function



    'CORTAR LAS COLUMNAS DE LOS MOVIMIENTOS
    Private Function Cortar_Movimientos_Clasica(ByVal sDataTramaMovi As String, ByRef lTotalMovimientos As Long, _
                                                       ByRef g_FlgMasData As Boolean, ByRef g_NumReg As Long) As String


        Dim r_DvCta As String = ""
        Dim r_IdCli As String = ""
        Dim r_DvCli As String = ""
        Dim r_Saldo As String = ""
        Dim r_Aux As String = ""
        Dim r_NumNeg As String = ""
        Dim r_Corte As Long = 0
        Dim r_NumMov As Long = 0
        Dim fila As Long = 0

        Dim r_glodet As String = ""
        Dim r_Tipdet As String = ""
        Dim r_fecdet As String = ""
        Dim r_sucdet As String = ""
        Dim r_cajdet As String = ""
        Dim r_docdet As String = ""

        Dim r_vendet As String = ""
        Dim r_pladet As String = ""
        Dim r_debdet As String = ""
        Dim r_habdet As String = ""
        Dim r_saldet As String = ""
        Dim r_piedet As String = ""
        Dim r_capdet As String = ""
        Dim r_fcdet As String = ""
        Dim r_intdet As String = ""
        Dim r_ffdet As String = ""
        Dim r_gasdet As String = ""
        Dim r_subdet As String = ""
        Dim r_Estab As String = ""

        Dim g_Nombre As String = ""
        Dim r_NumCta As String = ""


        Dim l_FecTran As String = ""                                            '' Variable local para contener la fecha de transaccion.
        Dim ldblMontoHaber As Double                                            '' Variable local para contener el Haber
        Dim ldblMontoDebe As Double                                             '' Variable local para contener el Debe
        Dim ldblMdebdet As Double, ldblMhabdet As Double, ldblMpiedet As Double '' Variables locales auxiliares para el calculo de los montos.
        Dim g_Posicion As Long = 15



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

        Dim strFechaProcesoInicial As String = "00000000"


        g_Nombre = g_ExtraeData(sDataTramaMovi, 35, g_Posicion)

        r_NumCta = g_ExtraeData(sDataTramaMovi, 8, g_Posicion)
        r_DvCta = g_ExtraeData(sDataTramaMovi, 1, g_Posicion)
        r_IdCli = g_ExtraeData(sDataTramaMovi, 12, g_Posicion)
        r_DvCli = g_ExtraeData(sDataTramaMovi, 1, g_Posicion)
        r_Saldo = g_ExtraeData(sDataTramaMovi, 11, g_Posicion)
        r_Aux = g_ExtraeData(sDataTramaMovi, 2, g_Posicion)
        r_Corte = Val(g_ExtraeData(sDataTramaMovi, 2, g_Posicion))
        r_NumMov = Val(g_ExtraeData(sDataTramaMovi, 2, g_Posicion))

        For fila = 1 To 12

            lTotalMovimientos = lTotalMovimientos + 1

            'Limpiar Variables.
            Fecha_mov = ""
            Descripcion_Mov = ""
            Establecimiento = ""
            Ticket = ""
            Monto = ""


            r_glodet = g_ExtraeData(sDataTramaMovi, 2, g_Posicion)
            r_Tipdet = g_ExtraeData(sDataTramaMovi, 1, g_Posicion)
            r_fecdet = g_ExtraeData(sDataTramaMovi, 8, g_Posicion)

            If Trim(r_fecdet) = "" Then ' SI NO TIENE FECHA DE MOV DEBE TERMINAR EL FOR
                g_FlgMasData = False
                Exit For
            End If

            Dim nom_estab As String
            nom_estab = String.Empty

            r_sucdet = g_ExtraeData(sDataTramaMovi, 2, g_Posicion)
            r_cajdet = g_ExtraeData(sDataTramaMovi, 3, g_Posicion)
            r_docdet = g_ExtraeData(sDataTramaMovi, 6, g_Posicion)
            r_vendet = g_ExtraeData(sDataTramaMovi, 8, g_Posicion)
            r_pladet = g_ExtraeData(sDataTramaMovi, 2, g_Posicion)
            r_debdet = g_ExtraeData(sDataTramaMovi, 8, g_Posicion)
            r_habdet = g_ExtraeData(sDataTramaMovi, 8, g_Posicion)
            r_saldet = g_ExtraeData(sDataTramaMovi, 11, g_Posicion)
            r_Aux = g_ExtraeData(sDataTramaMovi, 2, g_Posicion)
            r_piedet = g_ExtraeData(sDataTramaMovi, 8, g_Posicion) 'ITF
            r_capdet = g_ExtraeData(sDataTramaMovi, 8, g_Posicion)
            r_fcdet = g_ExtraeData(sDataTramaMovi, 6, g_Posicion)
            r_intdet = g_ExtraeData(sDataTramaMovi, 8, g_Posicion)
            r_ffdet = g_ExtraeData(sDataTramaMovi, 6, g_Posicion)
            r_gasdet = g_ExtraeData(sDataTramaMovi, 8, g_Posicion)
            r_subdet = g_ExtraeData(sDataTramaMovi, 2, g_Posicion)
            r_Estab = g_ExtraeData(sDataTramaMovi, 5, g_Posicion)
            nom_estab = g_ExtraeData(sDataTramaMovi, 40, g_Posicion)



            If Len(Trim(r_Estab)) = 0 Then
                r_Estab = "00000"
            End If

            Descripcion_Mov = ""

            'BUSCAR LA DESCRIPCION DEL MOVIMIENTO
            Descripcion_Mov = BUSCAR_DESCRIPCION_MOV(Val(r_sucdet), Val(r_glodet), Val(Trim(r_subdet)))

            If Descripcion_Mov <> "*N*" Then 'SOLO SE PINTARA SI EL VALO ES DIFERENTE A  *N*

                'BUSCAR DESCRIPCION DE ESTABLECIMIENTO
                If Val(r_sucdet) = 29 Then
                    'BUSCAR DESC ESTABLECIMIENTO DESDE ORACLE
                    'Cod Depa siempre es 95
                    'Cod Linea son los 2 primeros digitos de la variable r_Estab esta variable contiene el codigo de linea y Articulo

                    sCodigoEstablece = "00000" & r_Estab.Trim

                    sCodLinea = Left(Right(sCodigoEstablece.Trim, 5), 2)
                    sCodArticulo = Right(sCodigoEstablece.Trim, 3)

                    If r_Estab = "00000" Then
                        r_sucdet = BUSCAR_DESC_ESTABLECIMIENTO_ORACLE("95", sCodLinea.Trim, sCodArticulo.Trim)
                    Else
                        r_sucdet = nom_estab.Trim
                    End If


                Else
                    'BUSCAR ESTABLECIMIENTO DE SQL
                    r_sucdet = BUSCAR_DESCRIPCION_ESTABLECIMIENTO(Val(r_sucdet))
                End If



                'FECHA MOVIMIENTO YYYYMMDD CONVERTIR A DD/MM/YYYY
                If r_fecdet = "00000000" Then
                    r_fecdet = ""
                Else
                    r_fecdet = Mid(r_fecdet, 1, 2) & "/" & Mid(r_fecdet, 3, 2) & "/" & Mid(r_fecdet, 5, 4) 'DD/MM/YYYY
                End If

                '' Obtencion de montos en variables auxiliares... r_piedet$= esta variable tiene el ITF
                ldblMpiedet = r_piedet
                ldblMdebdet = r_debdet
                ldblMhabdet = r_habdet

                '' Calculo del Debe considerando el ITF.
                ldblMontoDebe = CDbl(ldblMdebdet + ldblMpiedet)

                '' Calculo del Haber considerando el ITF.
                ldblMontoHaber = CDbl(ldblMhabdet + ldblMpiedet)

                '' Asignacion de los montos calculados... MONTO DEBE Y MONTO HABER
                If Val(r_debdet) = 0 Then
                    r_debdet = "0"
                Else
                    r_debdet = Format(CDbl(ldblMontoDebe), "######0.00")
                End If

                If Val(r_habdet) = 0 Then
                    r_habdet = "0"
                Else
                    r_habdet = Format(CDbl(ldblMontoHaber), "######0.00")
                End If



                'OBTENER EL NUMERO DE TICKET
                r_docdet = Format(Val(r_docdet), "00000000")


                Fecha_mov = r_fecdet
                Establecimiento = r_sucdet
                Ticket = r_docdet

                If Val(r_debdet) > 0 Then
                    Monto = r_debdet
                Else
                    Monto = r_habdet
                End If

                'SI TIENE FECHA DE MOV
                If Fecha_mov.Trim.Length > 0 Then
                    Dim TramaSalida As String
                    Dim longitud As Integer
                    Fecha_mov = Fecha_mov.Trim
                    Descripcion_Mov = Descripcion_Mov.Trim
                    Establecimiento = RTrim(Establecimiento.Trim)
                    longitud = Establecimiento.Length
                    Ticket = Ticket.Trim
                    Monto = Monto.Trim

                    'If (longitud = 40) Then
                    '    Establecimiento = Establecimiento.Substring(0, 16)
                    'Else
                    '    Establecimiento.Trim()
                    'End If

                    Establecimiento = Establecimiento.Trim()


                    TramaSalida = Fecha_mov & "|\t|" & Descripcion_Mov & "|\t|" & Establecimiento & "|\t|" & Ticket & "|\t|" & Monto & "|\n|"
                    sDATA_MOV_RETORNO = sDATA_MOV_RETORNO & Fecha_mov.Trim & "|\t|" & Descripcion_Mov.Trim & "|\t|" & Establecimiento.Trim & "|\t|" & Ticket.Trim & "|\t|" & Monto.Trim & "|\n|"
                End If

            End If 'FIN VALIDA DATO DEL MOV *N* NO SE DEBE MOSTRAR SI BIENE *N*

        Next


        'Carga de la fecha de transaccion al arreglo de movimientos...
        g_Posicion = g_Posicion + 2

        Dim r_I As Long = 1
        Dim lFilasFechas As Long = 0
        Dim sDATA_FECHAS_MOV As String = ""

        g_NumReg = g_NumReg + 15
        g_FlgMasData = True
        If g_NumReg >= r_NumMov Then

            g_FlgMasData = False

        End If

        Return sDATA_MOV_RETORNO.Trim

    End Function


    'Cortar datos de movimientos de tarjetas asociadas.
    'CORTAR LAS COLUMNAS DE LOS MOVIMIENTOS
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


    Private Function g_ExtraeData(ByVal p_data As String, ByVal p_Longitud As Long, ByRef g_Posicion As Long) As String

        g_ExtraeData = Mid(p_data, g_Posicion, p_Longitud)
        g_Posicion = g_Posicion + p_Longitud

        Return g_ExtraeData

    End Function

    Public Function g_PresentStr(ByVal cadena As String) As String

        Dim r_I As Long = 0
        Dim palabra(1, 1) As String
        Dim Nombre As String = ""
        Dim r_Index As Long = 0
        Dim r_Aux As String

        On Error Resume Next

        cadena = Format(cadena, "<")

        If InStr(cadena, " ") > 0 Then
            r_Index = 0
            r_Aux = ""
            Erase palabra

            For r_I = 1 To Len(cadena)
                If Mid(cadena, r_I, 1) = " " Then
                    r_Index = r_Index + 1
                    ReDim Preserve palabra(1, r_Index)
                    palabra(1, r_Index) = r_Aux
                    r_Aux = ""
                Else
                    r_Aux = r_Aux + Mid(cadena, r_I, 1)
                End If
            Next
            r_Index = r_Index + 1
            ReDim Preserve palabra(1, r_Index)
            palabra(1, r_Index) = r_Aux

            cadena = ""
            For r_I = 1 To r_Index
                If Trim(palabra(1, r_I)) <> "" Then
                    palabra(1, r_I) = UCase(Left(palabra(1, r_I), 1)) + Mid(palabra(1, r_I), 2, Len(palabra(1, r_I)) - 1)
                End If
                cadena = cadena + " " + palabra(1, r_I)
            Next

        Else
            If Len(cadena) > 0 Then
                cadena = UCase(Left(cadena, 1)) + Mid(cadena, 2, Len(cadena) - 1)
            Else
                cadena = ""
            End If
        End If

        g_PresentStr = cadena

    End Function


    'FUNCIONES DE SQL BUSCAR DESCRIPCION DEL MOVIMIENTO
    Private Function BUSCAR_DESCRIPCION_MOV(ByVal cod_suc As Long, ByVal cod_trx As Long, ByVal sub_cod As Long) As String

        Try
            'Realizar Conexion a la base de datos
            sMensajeError_SQL = ""

            oConexion = SQL_ConnectionOpen(Get_CadenaConexion(), sMensajeError_SQL)

            If sMensajeError_SQL <> "" Then
                sResultado_SQL = ""
            Else

                If oConexion.State = ConnectionState.Open Then


                    m_ssql = "SELECT DBO.fun_devolver_movimiento(" & cod_suc.ToString & "," & cod_trx.ToString & "," & sub_cod.ToString & ") AS VSTR_MOVIMIENTO"


                    Dim cmd As SqlClient.SqlCommand = oConexion.CreateCommand

                    cmd.CommandTimeout = 360
                    cmd.CommandType = CommandType.Text
                    cmd.CommandText = m_ssql

                    Dim rdx As SqlClient.SqlDataReader = cmd.ExecuteReader
                    sXML_SQL = ""
                    If rdx.Read = True Then
                        'ARMAR CADENA
                        sXML_SQL = IIf(IsDBNull(rdx.Item("VSTR_MOVIMIENTO")), "", Trim(rdx.Item("VSTR_MOVIMIENTO")))
                    Else

                        sXML_SQL = ""
                    End If

                    sResultado_SQL = sXML_SQL

                    rdx.Close()
                    rdx = Nothing
                    cmd.Dispose()
                    cmd = Nothing
                    oConexion.Close()
                    oConexion = Nothing
                Else
                    sResultado_SQL = ""
                    oConexion = Nothing
                End If
            End If


        Catch ex As Exception
            sResultado_SQL = ""
            oConexion = Nothing
        End Try

        Return sResultado_SQL

    End Function

    '<WebMethod(Description:="Servicio devuelve Descripcion RSAT")> _
    Private Function Buscar_Desc_Mov_RSAT(ByVal cod_rsat As String) As String

        Dim Oconex As New SqlConnection(Get_CadenaConexion())
        Dim Descripcion As String
        Descripcion = String.Empty

        Dim comando As New SqlCommand("Sp_Buscar_Desc_Mov_RSAT", Oconex)

        Try

            comando.CommandType = CommandType.StoredProcedure

            comando.Parameters.Add(New SqlParameter("@cod_rsat", SqlDbType.VarChar, 5))
            comando.Parameters("@cod_rsat").Direction = ParameterDirection.Input
            comando.Parameters("@cod_rsat").Value = cod_rsat

            comando.Parameters.Add(New SqlParameter("@DESC_MOV", SqlDbType.VarChar, 17))
            comando.Parameters("@DESC_MOV").Direction = ParameterDirection.Output

            Oconex.Open()

            comando.ExecuteNonQuery()
            Descripcion = comando.Parameters(1).Value
            Oconex.Close()

        Catch ex As Exception
            Oconex.Close()
            Return "ERROR:CONEX SQL"
        End Try

        comando.Dispose()
        Oconex.Dispose()

        Return Descripcion

    End Function



    'FUNCIONES DE SQL BUSCAR DESCRIPCION DEL ESTABLECIMIENTO
    Private Function BUSCAR_DESCRIPCION_ESTABLECIMIENTO(ByVal cod_suc As Long) As String

        Try
            'Realizar Conexion a la base de datos
            sMensajeError_SQL = ""

            oConexion = SQL_ConnectionOpen(Get_CadenaConexion(), sMensajeError_SQL)

            If sMensajeError_SQL <> "" Then
                sResultado_SQL = ""
            Else

                If oConexion.State = ConnectionState.Open Then


                    m_ssql = "SELECT DBO.FUN_DEVOLVER_ESTABLECIMIENTO(" & cod_suc.ToString & ") AS VSTR_SUCURSAL"


                    Dim cmd As SqlClient.SqlCommand = oConexion.CreateCommand

                    cmd.CommandTimeout = 360
                    cmd.CommandType = CommandType.Text
                    cmd.CommandText = m_ssql

                    Dim rdx As SqlClient.SqlDataReader = cmd.ExecuteReader
                    sXML_SQL = ""
                    If rdx.Read = True Then
                        'ARMAR CADENA
                        sXML_SQL = IIf(IsDBNull(rdx.Item("VSTR_SUCURSAL")), "", Trim(rdx.Item("VSTR_SUCURSAL")))
                    Else

                        sXML_SQL = ""
                    End If

                    sResultado_SQL = sXML_SQL

                    rdx.Close()
                    rdx = Nothing
                    cmd.Dispose()
                    cmd = Nothing
                    oConexion.Close()
                    oConexion = Nothing
                Else
                    sResultado_SQL = ""
                    oConexion = Nothing
                End If
            End If


        Catch ex As Exception
            sResultado_SQL = ""
            oConexion = Nothing
        End Try

        Return sResultado_SQL

    End Function


    'BUSCAR ESTABLECIMIENTO DESDE ORACLE
    '<WebMethod(Description:="Buscar Establecimiento Ora.")> _
    Private Function BUSCAR_DESC_ESTABLECIMIENTO_ORACLE(ByVal sCodDepa As String, ByVal sCodLinea As String, ByVal sCodArt As String) As String


        Dim conConexion_Ora_ As OracleConnection


        Try
            sResultado_ORA = ""

            conConexion_Ora_ = New OracleConnection(mCadenaConexion_ORA_XX)

            If Not conConexion_Ora_ Is Nothing Then
                conConexion_Ora_.Open()

                If conConexion_Ora_.State = ConnectionState.Open Then

                    'LLAMAR AL PROCEDIMIENTO DEPOSITO A PLAZO
                    Dim cmd_ora_ As New OracleCommand
                    Dim param1_ As New OracleParameter
                    Dim param2_ As New OracleParameter
                    Dim param3_ As New OracleParameter
                    Dim param4_ As New OracleParameter
                    Dim sDescEstablec As String = ""

                    cmd_ora_.Connection = conConexion_Ora_
                    cmd_ora_.CommandText = "SIS_ADM.RPM_SP_BUSCA_ESTABLECIMIENTO"
                    cmd_ora_.CommandType = CommandType.StoredProcedure

                    param1_.ParameterName = "p_cod_dpt"
                    param1_.OracleType = OracleType.Double
                    param1_.Direction = ParameterDirection.Input
                    param1_.Value = Val(sCodDepa.Trim)
                    cmd_ora_.Parameters.Add(param1_)


                    param2_.ParameterName = "p_cod_lin"
                    param2_.OracleType = OracleType.Double
                    param2_.Direction = ParameterDirection.Input
                    param2_.Value = Val(sCodLinea.Trim)
                    cmd_ora_.Parameters.Add(param2_)


                    param3_.ParameterName = "p_cod_art"
                    param3_.OracleType = OracleType.Double
                    param3_.Direction = ParameterDirection.Input
                    param3_.Value = Val(sCodArt.Trim)
                    cmd_ora_.Parameters.Add(param3_)


                    param4_.ParameterName = "o_establecimient"
                    param4_.OracleType = OracleType.VarChar
                    param4_.Size = 30
                    param4_.Direction = ParameterDirection.Output
                    cmd_ora_.Parameters.Add(param4_)


                    cmd_ora_.ExecuteNonQuery()

                    sDescEstablec = cmd_ora_.Parameters("o_establecimient").Value.ToString

                    sResultado_ORA = sDescEstablec.Trim

                    cmd_ora_.Dispose()
                    cmd_ora_ = Nothing
                    conConexion_Ora_.Close()
                    conConexion_Ora_ = Nothing

                End If

            Else

                sResultado_ORA = ""

            End If

        Catch ex As Exception
            sResultado_ORA = ex.Message.Trim
        End Try


        Return sResultado_ORA

    End Function



    'FUNCION PARA OBTENER EL NUMERO DE OPERACION
    Private Function GET_NUMERO_OPERACION() As String

        Try
            'Realizar Conexion a la base de datos
            sMensajeError_SQL = ""

            oConexion = SQL_ConnectionOpen(Get_CadenaConexion(), sMensajeError_SQL)

            If sMensajeError_SQL <> "" Then
                sResultado_SQL = ""
            Else

                If oConexion.State = ConnectionState.Open Then

                    m_ssql = "SP_NUM_OPERACION"

                    Dim cmd As SqlClient.SqlCommand = oConexion.CreateCommand
                    Dim lTotal As Double = 0

                    cmd.CommandTimeout = 900
                    cmd.CommandType = CommandType.StoredProcedure
                    cmd.CommandText = m_ssql



                    Dim rd_GET As SqlClient.SqlDataReader = cmd.ExecuteReader
                    sXML_SQL = ""

                    If rd_GET.Read = True Then
                        sXML_SQL = rd_GET.GetValue(0)
                    End If

                    sResultado_SQL = sXML_SQL

                    rd_GET.Close()
                    rd_GET = Nothing
                    cmd.Dispose()
                    cmd = Nothing
                    oConexion.Close()
                    oConexion = Nothing
                Else
                    sResultado_SQL = ""
                    oConexion = Nothing
                End If
            End If


        Catch ex As Exception
            sResultado_SQL = ""
            oConexion = Nothing
        End Try

        Return sResultado_SQL

    End Function


    '<WebMethod(Description:="PRUEBA ORDENAR MOV")> _
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


    'Buscar la tasa de interes segun moneda y plazo Deposito a Plazo Fijo
    Private Function GET_TASA_SIMULADOR_DPF(ByVal sMoneda As String, ByVal dMonto As Double, ByVal lPlazoDias As Long) As String

        Try
            'Realizar Conexion a la base de datos
            sMensajeError_SQL = ""

            oConexion = SQL_ConnectionOpen(Get_CadenaConexion(), sMensajeError_SQL)

            If sMensajeError_SQL <> "" Then
                sResultado_SQL = ""
            Else

                If oConexion.State = ConnectionState.Open Then

                    If sMoneda.Trim.ToUpper = "D" Then
                        m_ssql = "SELECT trea,rango1,rango2 FROM KIO_TASAS_SIMULADOR WHERE moneda='" & sMoneda.Trim & "' AND (" & lPlazoDias & ">=rango1 AND " & lPlazoDias & "<=rango2)"
                    End If

                    If sMoneda.Trim.ToUpper = "S" Then
                        'VALIDAR MONTO PARA CONSULTAR EL CAMPO CORRECTO
                        If (dMonto >= 1000 And dMonto <= 10000) Then
                            m_ssql = "SELECT trea,rango1,rango2 FROM KIO_TASAS_SIMULADOR WHERE moneda='" & sMoneda.Trim & "' AND (" & lPlazoDias & ">=rango1 and " & lPlazoDias & "<=rango2)"
                        End If

                        If (dMonto > 10000) Then
                            m_ssql = "SELECT trea2,rango1,rango2 FROM KIO_TASAS_SIMULADOR WHERE moneda='" & sMoneda.Trim & "' AND (" & lPlazoDias & ">=rango1 and " & lPlazoDias & "<=rango2)"
                        End If

                    End If


                    Dim cmd As SqlClient.SqlCommand = oConexion.CreateCommand
                    Dim TREA As String = ""

                    cmd.CommandTimeout = 180
                    cmd.CommandType = CommandType.Text
                    cmd.CommandText = m_ssql



                    Dim rd_tasa As SqlClient.SqlDataReader = cmd.ExecuteReader
                    sXML_SQL = ""

                    If rd_tasa.Read = True Then
                        sXML_SQL = rd_tasa.GetValue(0)
                    End If

                    sResultado_SQL = sXML_SQL

                    rd_tasa.Close()
                    rd_tasa = Nothing
                    cmd.Dispose()
                    cmd = Nothing
                    oConexion.Close()
                    oConexion = Nothing
                Else
                    sResultado_SQL = "0"
                    oConexion = Nothing
                End If
            End If


        Catch ex As Exception
            sResultado_SQL = "0"
            oConexion = Nothing
        End Try

        Return sResultado_SQL

    End Function

    'OBTENER EL FACTOR DE ITF
    <WebMethod(Description:="OBTENER FACTOR ITF")> _
    Public Function OBTENER_FACTOR_ITF() As Double
        Dim sRespuesta_ITF As String = ""
        Dim sParametros_ITF As String = ""
        Dim sMensajeErrorUsuario_ITF As String = ""
        Dim sXML_ITF As String = ""


        'Variables nuevas
        Dim actions_ As Long = 1
        Dim inetputBuff_ As String = ""
        Dim outpputBuff_ As String = ""
        Dim errorMsg_ As String = "" ':Mensaje de Error

        Try

            'Instancia al mirapiweb
            Dim obSendMirror_ As ClsTxMirapi = Nothing
            obSendMirror_ = New ClsTxMirapi()

            sParametros_ITF = "0001" & DateTime.Now.Year.ToString.Trim & DateTime.Now.Month.ToString("00").Trim & DateTime.Now.Day.ToString("00").Trim
            inetputBuff_ = "      " + "V999" + sParametros_ITF

            sRespuesta_ITF = obSendMirror_.ExecuteTX(actions_.ToString, inetputBuff_, outpputBuff_, TIME_OUT_SERVER, SERVIDOR_MIRROR_DESTINO, SERVIDOR_MIRROR_NODE, PUERTO, errorMsg_)

            If sRespuesta_ITF = "0" Then 'EXITO

                If outpputBuff_.Length > 0 Then sRespuesta_ITF = outpputBuff_.Substring(8, outpputBuff_.Length - 8)

                'Evaluar Respuesta si es ERROR,  RU, AU (Respuesta correcta)
                If sRespuesta_ITF.Trim.Length > 0 Then
                    If Left(sRespuesta_ITF.Trim, 2) = "RU" Then 'Mensaje de Validacion al usuario. 
                        sMensajeErrorUsuario_ITF = Mid(sRespuesta_ITF.Trim, 18, Len(sRespuesta_ITF.Trim))
                        sRespuesta_ITF = "0"
                    Else
                        If Left(sRespuesta_ITF.Trim, 2) = "AU" Then
                            'Recuperar Data
                            Dim sITF As String = ""
                            Dim sDato As String = ""
                            Dim sDecimal As String = ""
                            Dim dITF As Double = 0

                            sDato = Trim(Mid(sRespuesta_ITF, 5, 12))
                            sDecimal = Trim(Right(sDato, 6))
                            sDato = Trim(Left(sDato, 6)) & "." & sDecimal.Trim
                            dITF = CDbl(sDato.Trim) + 1


                            sRespuesta_ITF = dITF

                        Else 'Cualquier Otro Caso
                            sRespuesta_ITF = "0"
                        End If

                    End If

                Else 'Sino devuelve nada
                    sRespuesta_ITF = "0"
                End If

            ElseIf sRespuesta_ITF = "-2" Then 'Ocurrio un error Recuperar el error
                'sRespuesta_ITF = "ERROR:" & errorMsg_.Trim
                sRespuesta_ITF = "0"
            Else
                'sRespuesta_ITF = "ERROR:Servicio no disponible."
                sRespuesta_ITF = "0"
            End If



        Catch ex As Exception

            sRespuesta_ITF = "0"

        End Try

        Return sRespuesta_ITF

    End Function


    '    '***********************************  METODOS DE ORACLE ******************************

    <WebMethod(Description:="Prueba de Conexion a la base de datos de Oracle.")> _
    Public Function PRUEBA_CONEXION_ORA() As String
        Dim conConexion_Ora As OracleConnection

        Try
            conConexion_Ora = New OracleConnection(mCadenaConexion_ORA)

            If Not conConexion_Ora Is Nothing Then
                conConexion_Ora.Open()

                If conConexion_Ora.State = ConnectionState.Open Then conConexion_Ora.Close()

                conConexion_Ora.Close()
                conConexion_Ora = Nothing
                sResultado_ORA = "OK"

            Else

                sResultado_ORA = "ERROR:No se pudo conectar al servidor de base de datos Oracle."

            End If


        Catch err As Exception
            sResultado_ORA = "ERROR:No se pudo conectar al servidor de base de datos."

        End Try


        Return sResultado_ORA


    End Function


    <WebMethod(Description:="Prueba de Conexion a la base de datos de Oracle SuperEfectivo.")> _
    Public Function PRUEBA_CONEXION_ORA_SEF() As String
        Dim conConexion_Ora As OracleConnection

        Try
            conConexion_Ora = New OracleConnection(mCadenaConexion_ORA_SEF)

            If Not conConexion_Ora Is Nothing Then
                conConexion_Ora.Open()

                If conConexion_Ora.State = ConnectionState.Open Then conConexion_Ora.Close()

                conConexion_Ora.Close()
                conConexion_Ora = Nothing
                sResultado_ORA = "OK"

            Else

                sResultado_ORA = "ERROR:No se pudo conectar al servidor de base de datos Oracle."

            End If


        Catch err As Exception
            sResultado_ORA = "ERROR:No se pudo conectar al servidor de base de datos."

        End Try


        Return sResultado_ORA


    End Function

    ''' <summary>
    ''' Funciona para probar la conexion hacia la base de datos Oracle de Incremento de Linea
    ''' </summary>
    ''' <returns>Una cadena describiendo si la conexion se hizo con exito, o si ocurrio algun error</returns>
    ''' <remarks></remarks>
    <WebMethod(Description:="Prueba de Conexion a la base de datos de Oracle Incremento de Linea.")> _
    Public Function PRUEBA_CONEXION_ORA_INC() As String
        Dim conexionOracle As OracleConnection
        Dim resultado As String = String.Empty

        Try
            conexionOracle = New OracleConnection(cadenaConexionOraInc)
            If Not conexionOracle Is Nothing Then
                conexionOracle.Open()
                If conexionOracle.State = ConnectionState.Open Then conexionOracle.Close()

                conexionOracle.Close()
                conexionOracle = Nothing
                resultado = "OK"
            Else
                resultado = "ERROR:No se pudo conectar al servidor de base de datos Oracle."
            End If
        Catch err As Exception
            resultado = "ERROR:No se pudo conectar al servidor de base de datos."
        End Try

        Return resultado
    End Function



    <WebMethod(Description:="CONSULTAR CTS")> _
    Public Function CONSULTA_CTS(ByVal sDNI As String, ByVal sDATA_MONITOR_KIOSCO As String) As String

        Dim conConexion_Ora As OracleConnection


        Try
            conConexion_Ora = New OracleConnection(mCadenaConexion_ORA)

            If Not conConexion_Ora Is Nothing Then
                conConexion_Ora.Open()

                If conConexion_Ora.State = ConnectionState.Open Then

                    'LLAMAR AL PROCEDIMIENTO CTS
                    Dim cmd_ora As New OracleCommand
                    Dim param1 As New OracleParameter
                    Dim param2 As New OracleParameter
                    Dim param_cur As New OracleParameter

                    cmd_ora.Connection = conConexion_Ora
                    cmd_ora.CommandText = "RIPLEY_DA.RPM_PKG_CONSULTA.RPM_PRC_CTS"
                    cmd_ora.CommandType = CommandType.StoredProcedure

                    param1.ParameterName = "DNI"
                    param1.OracleType = OracleType.VarChar
                    param1.Direction = ParameterDirection.Input
                    param1.Value = sDNI.Trim
                    cmd_ora.Parameters.Add(param1)

                    param2.ParameterName = "p_FECHA"
                    param2.OracleType = OracleType.DateTime
                    param2.Direction = ParameterDirection.Input

                    Dim vFecHoy As Date
                    Dim sFechaParam As String = ""

                    vFecHoy = Date.Now
                    sFechaParam = Right(Trim("00" & Day(vFecHoy).ToString), 2) & "/" & Right(Trim("00" & Month(vFecHoy).ToString), 2) & "/" & Right(Trim("0000" & Year(vFecHoy).ToString), 4)

                    param2.Value = sFechaParam.Trim
                    cmd_ora.Parameters.Add(param2)

                    param_cur.ParameterName = "V_CURSOR"
                    param_cur.OracleType = OracleType.Cursor
                    param_cur.Direction = ParameterDirection.Output
                    param_cur.Value = vbNull

                    cmd_ora.Parameters.Add(param_cur)

                    Dim rd_cur As OracleDataReader = cmd_ora.ExecuteReader

                    Dim sFechaKiosco As String = DateTime.Now.Day.ToString("00").Trim + "/" + DateTime.Now.Month.ToString("00").Trim + "/" + DateTime.Now.Year.ToString("0000").Trim
                    Dim sHoraKiosco As String = DateTime.Now.ToString("HH:mm:ss")
                    Dim sNOperacion As String = GET_NUMERO_OPERACION()


                    If rd_cur.Read = True Then

                        sResultado_ORA = sFechaKiosco.Trim & "|\t|" & sHoraKiosco.Trim & "|\t|" & sNOperacion.Trim & "|\t|"
                        sResultado_ORA = sResultado_ORA & IIf(IsDBNull(rd_cur.Item("NUMCTA")), "", rd_cur.Item("NUMCTA").ToString) & "|\t|SOLES|\t|"
                        sResultado_ORA = sResultado_ORA & IIf(IsDBNull(rd_cur.Item("SALDOBLQ")), "0.00", Format(CDbl(rd_cur.Item("SALDOBLQ").ToString.Trim), "##,##0.00")) & "|\t|" 'SALDO NO DISPONIBLE 
                        sResultado_ORA = sResultado_ORA & IIf(IsDBNull(rd_cur.Item("SALDOEFE")), "0.00", Format(CDbl(rd_cur.Item("SALDOEFE").ToString.Trim), "##,##0.00")) & "|\t|" 'SALDO DISPONIBLE
                        sResultado_ORA = sResultado_ORA & IIf(IsDBNull(rd_cur.Item("SALDOTOT")), "0.00", Format(CDbl(rd_cur.Item("SALDOTOT").ToString.Trim), "##,##0.00")) 'TOTAL CTS

                    Else
                        sResultado_ORA = "ERROR:NODATA"
                    End If

                    cmd_ora.Dispose()
                    cmd_ora = Nothing

                    conConexion_Ora.Close()
                    conConexion_Ora = Nothing

                End If

            Else

                sResultado_ORA = "ERROR:No se pudo conectar al servidor de base de datos Oracle."

            End If


        Catch ex As Exception
            sResultado_ORA = "ERROR:" & ex.Message.Trim
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
            'SAVE_LOG_CONSULTAS(sIDSucursal.Trim, sIDTerminal.Trim, "44", FUN_BUSCAR_TIPO_TARJETA(sNroTarjetax.Trim))


            sNroCuentax = Right("          " & sNroCuentax.Trim, 10)
            sNroTarjetax = Right(Strings.StrDup(20, " ") & sNroTarjetax.Trim, 20)

            sNombreCliente = Right(Strings.StrDup(26, " ") & sNombreCliente.Trim, 26)

            sSaldoDisponible = Right("      " & sSaldoDisponible.Trim, 6)
            sTotalDeuda = Right("      " & sTotalDeuda.Trim, 6)



            'armar variable desde flash con: numero de cuenta, numero de tarjeta, nombre del cliente, modo de entrada, canal de atencion, id de sucursal,codigo kiosco y codigo de transaccion (opcion)
            If Left(sResultado_ORA.Trim, 5) <> "ERROR" Then
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


        Return sResultado_ORA


    End Function


    'CONSULTAR TIPO DOC Y NRO DOC DE MC O VISA
    <WebMethod(Description:="Consultar Documento del Cliente MC y VS.")> _
    Public Function CONSULTAR_DOC_MC_VISA(ByVal sNroTarjeta As String, _
                                          ByRef Nro_Contrato As String) As String

        Dim conConexion_mc As OracleConnection

        Try
            conConexion_mc = New OracleConnection(mCadenaConexion_ORA_X)
            sResultado_ORA = ""

            If Not conConexion_mc Is Nothing Then
                conConexion_mc.Open()

                If conConexion_mc.State = ConnectionState.Open Then

                    'LLAMAR AL PROCEDIMIENTO PRESTAMO EN EFECTIVO
                    Dim cmd_doc As New OracleCommand
                    Dim param1 As New OracleParameter
                    Dim param2 As New OracleParameter
                    Dim param3 As New OracleParameter
                    Dim param4 As New OracleParameter
                    Dim sNroDoc As String = ""
                    Dim sTipoDoc As String = ""

                    cmd_doc.Connection = conConexion_mc
                    cmd_doc.CommandText = "rpm_sp_busca_documento"
                    cmd_doc.CommandType = CommandType.StoredProcedure

                    param1.ParameterName = "p_tarjeta"
                    param1.OracleType = OracleType.VarChar
                    param1.Size = 16
                    param1.Direction = ParameterDirection.Input
                    param1.Value = sNroTarjeta.Trim
                    cmd_doc.Parameters.Add(param1)


                    param2.ParameterName = "o_numdoc"
                    param2.OracleType = OracleType.VarChar
                    param2.Size = 12
                    param2.Direction = ParameterDirection.Output
                    param2.Value = sNroDoc.Trim
                    cmd_doc.Parameters.Add(param2)

                    param3.ParameterName = "o_tipdoc"
                    param3.OracleType = OracleType.VarChar
                    param3.Size = 1
                    param3.Direction = ParameterDirection.Output
                    cmd_doc.Parameters.Add(param3)

                    param4.ParameterName = "o_nro_cuenta"
                    param4.OracleType = OracleType.VarChar
                    param4.Size = 16
                    param4.Direction = ParameterDirection.Output
                    cmd_doc.Parameters.Add(param4)

                    cmd_doc.ExecuteNonQuery()

                    sNroDoc = cmd_doc.Parameters("o_numdoc").Value.ToString
                    sTipoDoc = cmd_doc.Parameters("o_tipdoc").Value.ToString
                    Nro_Contrato = cmd_doc.Parameters("o_nro_cuenta").Value.ToString

                    If sNroDoc.Length > 0 And sTipoDoc.Length > 0 Then
                        sResultado_ORA = sTipoDoc.Trim & "|\t|" & sNroDoc.Trim
                    Else
                        sResultado_ORA = "|\t|"
                    End If

                    cmd_doc.Dispose()
                    cmd_doc = Nothing
                    conConexion_mc.Close()
                    conConexion_mc = Nothing

                End If

            Else

                sResultado_ORA = "|\t|"

            End If

        Catch ex As Exception
            sResultado_ORA = "|\t||\t|" & ex.Message.Trim
            Nro_Contrato = ex.Message.Trim
        End Try

        Return sResultado_ORA

    End Function

    <WebMethod(Description:="Consultar Documento del Cliente MC y VS.")> _
    Public Function CONSULTAR_DOC_MC_VISA_ORI(ByVal sNroTarjeta As String) As String

        Dim conConexion_mc As OracleConnection

        Try
            conConexion_mc = New OracleConnection(mCadenaConexion_ORA_X)
            sResultado_ORA = ""

            If Not conConexion_mc Is Nothing Then
                conConexion_mc.Open()

                If conConexion_mc.State = ConnectionState.Open Then

                    'LLAMAR AL PROCEDIMIENTO PRESTAMO EN EFECTIVO
                    Dim cmd_doc As New OracleCommand
                    Dim param1 As New OracleParameter
                    Dim param2 As New OracleParameter
                    Dim param3 As New OracleParameter
                    Dim param4 As New OracleParameter
                    Dim sNroDoc As String = ""
                    Dim sTipoDoc As String = ""

                    cmd_doc.Connection = conConexion_mc
                    cmd_doc.CommandText = "rpm_sp_busca_documento"
                    cmd_doc.CommandType = CommandType.StoredProcedure

                    param1.ParameterName = "p_tarjeta"
                    param1.OracleType = OracleType.VarChar
                    param1.Size = 16
                    param1.Direction = ParameterDirection.Input
                    param1.Value = sNroTarjeta.Trim
                    cmd_doc.Parameters.Add(param1)


                    param2.ParameterName = "o_numdoc"
                    param2.OracleType = OracleType.VarChar
                    param2.Size = 12
                    param2.Direction = ParameterDirection.Output
                    param2.Value = sNroDoc.Trim
                    cmd_doc.Parameters.Add(param2)


                    param3.ParameterName = "o_tipdoc"
                    param3.OracleType = OracleType.VarChar
                    param3.Size = 1
                    param3.Direction = ParameterDirection.Output
                    cmd_doc.Parameters.Add(param3)

                    cmd_doc.ExecuteNonQuery()

                    sNroDoc = cmd_doc.Parameters("o_numdoc").Value.ToString
                    sTipoDoc = cmd_doc.Parameters("o_tipdoc").Value.ToString

                    If sNroDoc.Length > 0 And sTipoDoc.Length > 0 Then
                        sResultado_ORA = sTipoDoc.Trim & "|\t|" & sNroDoc.Trim
                    Else
                        sResultado_ORA = "|\t|"
                    End If

                    cmd_doc.Dispose()
                    cmd_doc = Nothing
                    conConexion_mc.Close()
                    conConexion_mc = Nothing

                End If

            Else

                sResultado_ORA = "|\t|"

            End If

        Catch ex As Exception
            sResultado_ORA = "|\t||\t|" & ex.Message.Trim
        End Try

        Return sResultado_ORA

    End Function


    Public Function FUN_BUSCAR_FECHA_CORTE_CLASICA_RSAT(ByVal NroTarjeta As String, ByVal NroCuenta As String) As String


        Dim objMQ As New MQ
        Dim strServicio As String = String.Empty
        Dim strMensaje As String = String.Empty
        Dim strParametros As String = String.Empty
        Dim strRespuesta As String = String.Empty
        Dim strPan As String = String.Empty
        Dim strNroPan As String = String.Empty
        Dim lngLargo As Long = 0
        Dim i As Integer = 0

        NroCuenta = NroCuenta.Trim
        NroTarjeta = NroTarjeta.Trim


        'strServicio = "SFSCANT0114"
        'strMensaje = "0000000000SFSCANT0114                                       000000069000002USUARIO127  SANI0001PE0001" & NroCuenta & "" & NroTarjeta & "      604"
        strServicio = ReadAppConfig("SFSCAN_SALDO_FECHACORTE_CLASICA_RSAT")
        strMensaje = "0000000000" & strServicio & "                                       000000069000002USUARIO127  SANI0001PE0001" & NroCuenta & "" & NroTarjeta & "      604"
        objMQ.Service = strServicio
        objMQ.Message = strMensaje
        objMQ.Execute()

        If objMQ.ReasonMd <> 0 Then
            strRespuesta = "ERROR"
        End If
        If objMQ.ReasonApp = 0 Then
            strRespuesta = objMQ.Response
        End If

        'VARIABLES DEL SALDO DE TARJETA CLASICA
        Dim sLineaCredito As String = "" 'LINEA DE CRÉDITO
        Dim dLineaCreditoUtilizada As Double = 0 'LÍNEA DE CRÈDITO UTILIZADA
        Dim sDisponibleCompras As String = "" 'DISPONIBLE COMPRAS
        Dim sDisponibleEfectivo As String = "" 'DISPONIBLE EFECTIVO EXPRESS
        Dim sPagoTotalMes As String = "" 'PAGO TOTAL DEL MES
        Dim sPagoMinimo As String = "" 'PAGO MINIMO DEL MES
        Dim sPeriodo_Facturacion As String = "" 'PERIODO DE FACTURACION
        Dim sFechaPago As String = "" 'FECHA DE PAGO
        Dim sRipleyPuntos As String = "" 'RIPLEY PUNTOS
        Dim sDisponibleSuperEfectivo As String = ""
        Dim sPagoTotal As String = "" 'PAGO TOTAL

        Dim sFechaKiosco As String = DateTime.Now.Day.ToString("00").Trim + "/" + DateTime.Now.Month.ToString("00").Trim + "/" + DateTime.Now.Year.ToString("0000").Trim
        Dim sHoraKiosco As String = DateTime.Now.ToString("HH:mm:ss")
        Dim sNOperacion As String = GET_NUMERO_OPERACION()


        Dim strResp114, strResp103, strResp113, strResp114_2 As String

        strResp114 = String.Empty
        strResp103 = String.Empty
        strResp113 = String.Empty
        strResp114_2 = String.Empty

        If Val(Mid(strRespuesta, 1, 10)) = 0 Then

            strResp114 = Mid(strRespuesta, 139, 2)
            strResp103 = Mid(strRespuesta, 141, 1096)
            strResp113 = Mid(strRespuesta, 1237, 303)
            strResp114_2 = Mid(strRespuesta, 1540, 17)

        End If


        Dim sMontoLinEE, sMontoLinSE, ws_MontosDisp, sLineaAct, sMontoLin, sMontoLin12 As String
        Dim iNumLin, sInd As Integer

        sMontoLinEE = "0.00"
        sMontoLinSE = "0.00"
        sMontoLin12 = String.Empty

        iNumLin = Val(Mid(strResp103, 35, 2))
        iNumLin = 3

        ws_MontosDisp = Mid(strResp103, 37)

        sInd = 0

        While sInd < iNumLin
            sLineaAct = Mid(ws_MontosDisp, 3 + 53 * sInd, 4)
            sMontoLin = Format(Val(Mid(ws_MontosDisp, 37 + 53 * sInd, 17)) / 100, "###00.00")
            If sLineaAct = "0001" Then
                sMontoLin12 = sMontoLin
            Else
                If sLineaAct = "0002" Then
                    sMontoLinEE = sMontoLin
                Else
                    If sLineaAct = "0003" Then
                        sMontoLinSE = sMontoLin
                    End If
                End If
            End If
            sInd = sInd + 1
        End While

        Dim stdPagoTot, stdPagoMin As String

        stdPagoTot = Mid(strResp113, 1, 17) 'Cuota Mes
        stdPagoMin = Mid(strResp113, 18, 17) 'Pago Minimo

        Dim lstrFechaEvaluacion, lintDiaVencimiento, vDia As String

        lstrFechaEvaluacion = Date.Now().ToString("yyyyMMdd")
        lintDiaVencimiento = Val(strResp114)

        Dim vFecHoy, vFecPago As Date
        Dim vFecCorteIni As Date
        Dim vFecCorteFin As Date
        Dim sFecha1, sFecha2, sFechaPagox As String

        vFecHoy = Date.Now
        vDia = lintDiaVencimiento

        vFecPago = CStr(funDevuelve_FechaPago(vFecHoy, vDia))
        procDevuelve_PeriodoFacturacion(vFecPago, vFecCorteIni, vFecCorteFin)

        'Formatear DDMMYYYY
        sFechaPagox = Right(Trim("00" & Day(vFecPago).ToString), 2) & Right(Trim("00" & Month(vFecPago).ToString), 2) & Right(Trim("0000" & Year(vFecPago).ToString), 4)

        If sFechaPagox = "01010001" Then
            Return "ERROR:NODATA"
        End If


        sFecha1 = Right(Trim("00" & Day(vFecCorteIni).ToString), 2) & Right(Trim("00" & Month(vFecCorteIni).ToString), 2) & Right(Trim("0000" & Year(vFecCorteIni).ToString), 4)
        sFecha2 = Right(Trim("00" & Day(vFecCorteFin).ToString), 2) & Right(Trim("00" & Month(vFecCorteFin).ToString), 2) & Right(Trim("0000" & Year(vFecCorteFin).ToString), 4)

        'PERIODO DE CORTE

        sFechaPago = sFecha2.Trim

        sFechaPago = sFecha2.Substring(0, 2) & "/" & sFecha2.Substring(2, 2) & "/" & sFecha2.Substring(4, 4)

        Return sFechaPago



    End Function


    'funcion que devuelve la fecha final de facturacion para clasica esto para la consulta de puntos
    '<WebMethod(Description:="Consultar Fecha de Corte final de Clasica")> _
    Public Function FUN_BUSCAR_FECHA_CORTE_CLASICA_SICRON(ByVal sNroCuenta As String) As String
        Dim sRespuestax As String = ""
        Dim sParametrosx As String = ""
        Dim sXML_ As String = ""


        'Variables nuevas
        Dim actions As Long = 1
        Dim inetputBuff As String = ""
        Dim outpputBuff As String = ""
        Dim errorMsg As String = ""

        Try

            If sNroCuenta.Trim.Length = 8 Then


                'Instancia al mirapiweb
                Dim obSendMirror As ClsTxMirapi = Nothing
                obSendMirror = New ClsTxMirapi()

                sParametrosx = Strings.StrDup(4, " ") & "245                00REU          1" & sNroCuenta.Trim & "000000000000" & "0" & "10000" & "000000000000" & Strings.StrDup(114, " ") & "00"
                inetputBuff = "      " + "V107" + sParametrosx

                sRespuestax = obSendMirror.ExecuteTX(actions.ToString, inetputBuff, outpputBuff, TIME_OUT_SERVER, SERVIDOR_MIRROR_DESTINO, SERVIDOR_MIRROR_NODE, PUERTO, errorMsg)

                If sRespuestax = "0" Then 'EXITO

                    If outpputBuff.Length > 0 Then sRespuestax = outpputBuff.Substring(8, outpputBuff.Length - 8)

                    'Evaluar Respuesta si es ERROR,  RU, AU (Respuesta correcta)
                    If sRespuestax.Trim.Length > 0 Then
                        If Left(sRespuestax.Trim, 2) = "RU" Then
                            sRespuestax = ""
                        Else
                            If Left(sRespuestax.Trim, 2) = "AU" Then

                                'VARIABLES DEL SALDO DE TARJETA CLASICA
                                Dim lintDiaVencimiento As Long = 0
                                Dim lstrFechaEvaluacion As String = ""
                                Dim lstrFecFactVigente As String = ""
                                Dim lstrFecFactVigenteDDMMYYYY As String = ""

                                'dia de vencimiento de pago
                                lintDiaVencimiento = Val(Mid(sRespuestax, 250, 2))

                                Dim vFecPago As Date
                                Dim vFecHoy As Date
                                Dim vDia As Integer
                                Dim sFechaPagox As String = ""
                                Dim vFecCorteIni As Date
                                Dim vFecCorteFin As Date
                                Dim sFecha1 As String = ""
                                Dim sFecha2 As String = ""

                                vFecHoy = Date.Now
                                vDia = lintDiaVencimiento

                                vFecPago = CStr(funDevuelve_FechaPago(vFecHoy, vDia))

                                Call procDevuelve_PeriodoFacturacion(vFecPago, vFecCorteIni, vFecCorteFin)

                                sFecha2 = Right(Trim("00" & Day(vFecCorteFin).ToString), 2) & "/" & Right(Trim("00" & Month(vFecCorteFin).ToString), 2) & "/" & Right(Trim("0000" & Year(vFecCorteFin).ToString), 4)

                                sFechaPagox = sFecha2.Trim

                                sRespuestax = ""
                                sRespuestax = sFechaPagox.Trim

                            Else 'Cualquier Otro Caso
                                sRespuestax = ""
                            End If


                        End If

                    Else 'Sino devuelve nada
                        sRespuestax = ""
                    End If

                ElseIf sRespuestax = "-2" Then 'Ocurrio un error en mirapiweb
                    'sRespuestax = "ERROR:" & errorMsg.Trim
                    sRespuestax = ""
                Else
                    'sRespuestax = "ERROR:Servicio no disponible."
                    sRespuestax = ""
                End If

            Else
                sRespuestax = ""
            End If

        Catch ex As Exception
            sRespuestax = ""
        End Try


        Return sRespuestax

    End Function


    'funcion que devuelve la fecha de pago para tarjetas asociadas esto para la consulta de puntos
    Private Function FUN_BUSCAR_FECHA_CORTE_ASOCIADA(ByVal sNroTarjeta As String) As String
        Dim sGetTrama_MC As String = ""
        Dim sParametros_MC As String = ""
        Dim sXMLx As String = ""

        Try


            If sNroTarjeta.Trim.Length = 16 Then

                Dim obSendWAS_ As New WSCONSULTAS_MC_VS.defaultService

                sParametros_MC = "HQ000001073RMDQCS" & sNroTarjeta.Trim & Strings.StrDup(16, " ") & "0000000000" & "0" & "000000" & "000073" & "@"

                sGetTrama_MC = obSendWAS_.execute(sParametros_MC.Trim)

                'Evaluar Respuesta si es ERROR
                If sGetTrama_MC.Trim.Length > 0 Then
                    If Left(sGetTrama_MC.Trim, 2) = "HE" Then
                        sGetTrama_MC = ""

                    Else

                        'VARIABLES DEL SALDO DE TARJETA ASOCIADA
                        Dim sFechaPago As String = "" 'FECHA DE PAGO
                        Dim sFechaCorteFinal As String = ""
                        Dim sDia As String = ""
                        Dim sMes As String = ""
                        Dim sAnio As String = ""

                        'Fecha de Pago
                        sDia = Trim(Mid(sGetTrama_MC, 227, 2))
                        sMes = Trim(Mid(sGetTrama_MC, 225, 2))
                        sAnio = Trim(Mid(sGetTrama_MC, 221, 4))

                        sFechaPago = sDia.Trim & "/" & sMes.Trim & "/" & sAnio

                        If Not IsDate(sFechaPago.Trim) Then
                            sFechaPago = ""
                        Else

                            'PERIODO DE FACTURACION
                            Dim sFecha1 As String = ""
                            Dim sFecha2 As String = ""
                            Dim lMes1 As Integer = 0
                            Dim lAnio As Integer = 0
                            Dim lAnio1 As Integer = 0
                            Dim lAnio2 As Integer = 0

                            If sFechaPago.Trim.ToUpper <> "" Then

                                If CInt(sMes.Trim) < 2 Then
                                    lMes1 = 12 + CInt(sMes.Trim)
                                Else
                                    lMes1 = CInt(sMes.Trim)
                                End If

                                lAnio = CInt(sAnio.Trim)

                                'Año1
                                If CInt(sMes.Trim) < 3 Then
                                    lAnio1 = lAnio - 1
                                Else
                                    lAnio1 = lAnio
                                End If

                                'Año2
                                If CInt(sMes.Trim) < 3 Then
                                    Select Case CInt(sMes.Trim)
                                        Case 1
                                            lAnio2 = lAnio - 1
                                        Case 2
                                            lAnio2 = lAnio
                                    End Select

                                Else
                                    lAnio2 = lAnio
                                End If


                                If CInt(sDia.Trim) = 5 Then
                                    If lMes1 = 2 Then
                                        sFecha1 = "21" & "/" & Right("12", 2) & "/" & lAnio1.ToString
                                    Else
                                        sFecha1 = "21" & "/" & Right("00" & (lMes1 - 2).ToString, 2) & "/" & lAnio1.ToString
                                    End If

                                    sFecha2 = "20" & "/" & Right("00" & (lMes1 - 1).ToString, 2) & "/" & lAnio2.ToString
                                End If


                                If CInt(sDia.Trim) = 25 Then
                                    sFecha1 = "11" & "/" & Right("00" & (lMes1 - 1).ToString, 2) & "/" & lAnio2.ToString
                                    sFecha2 = "10" & "/" & Right("00" & sMes.Trim, 2) & "/" & sAnio.Trim
                                End If


                                sFechaCorteFinal = sFecha2.Trim
                            Else
                                sFechaCorteFinal = ""
                            End If

                        End If

                        sGetTrama_MC = sFechaCorteFinal.Trim

                    End If

                Else 'Sino devuelve nada
                    sGetTrama_MC = ""
                End If


            Else
                'Mostrar Mensaje de Error
                sGetTrama_MC = ""
            End If

        Catch ex As Exception
            sGetTrama_MC = ""

        End Try


        Return sGetTrama_MC

    End Function

    Private Function FUN_BUSCAR_LINEA_DISPONIBLE_COMPRAS_CLASICA_RSAT(ByVal NroTarjeta As String, ByVal NroCuenta As String) As String

        Dim SaldoRSAT() As String = SALDO_TARJETA_CLASICA_RSAT(NroTarjeta, NroCuenta, "").Split("|\t|")

        Dim TarjetaBin As New TarjetaBin
        Dim BINN_TARJETA As String = String.Empty

        BINN_TARJETA = NroTarjeta.Substring(1, 6)
        Try
            TarjetaBin = BNTarjetaBin.Instancia.ObtenerTarjetaBin(BINN_TARJETA)
        Catch ex As Exception
            ErrorLog(ex.Message)
        End Try

        Return "TARJETA RIPLEY " & TarjetaBin.Nombre & "|\t|" & NroTarjeta.Trim & "|\t|" & "S/. " & SaldoRSAT(4)

    End Function



    'BUSCAR LA LINEA DISPONIBLE DE COMPRAS
    Private Function FUN_BUSCAR_LINEA_DISPONIBLE_COMPRAS_CLASICA_SICRON(ByVal sNroCuenta As String) As String
        Dim sRespuestax As String = ""
        Dim sParametrosx As String = ""
        Dim sXML_ As String = ""

        'Variables nuevas
        Dim actions_ As Long = 1
        Dim inetputBuff_ As String = ""
        Dim outpputBuff_ As String = ""
        Dim errorMsg_ As String = "" ':Mensaje de Error

        Try

            If sNroCuenta.Trim.Length = 8 Then


                'Instancia al mirapiweb
                Dim obSendMirror_ As ClsTxMirapi = Nothing
                obSendMirror_ = New ClsTxMirapi()

                sParametrosx = Strings.StrDup(4, " ") & "245                00REU          1" & sNroCuenta.Trim & "000000000000" & "0" & "10000" & "000000000000" & Strings.StrDup(114, " ") & "00"
                inetputBuff_ = "      " + "V107" + sParametrosx

                sRespuestax = obSendMirror_.ExecuteTX(actions_.ToString, inetputBuff_, outpputBuff_, TIME_OUT_SERVER, SERVIDOR_MIRROR_DESTINO, SERVIDOR_MIRROR_NODE, PUERTO, errorMsg_)

                If sRespuestax = "0" Then 'EXITO

                    If outpputBuff_.Length > 0 Then sRespuestax = outpputBuff_.Substring(8, outpputBuff_.Length - 8)

                    'Evaluar Respuesta si es  RU, AU (Respuesta correcta)
                    If sRespuestax.Trim.Length > 0 Then
                        If Left(sRespuestax.Trim, 2) = "RU" Then
                            sRespuestax = ""
                        Else
                            If Left(sRespuestax.Trim, 2) = "AU" Then

                                'VARIABLES DEL SALDO DE TARJETA CLASICA
                                Dim sTarjetaRecuperada As String = ""
                                Dim sDisponibleCompras As String = ""

                                sDisponibleCompras = Format(Val(Mid(sRespuestax, 144, 8)), "##,##0.00")
                                sTarjetaRecuperada = Mid(sRespuestax, 28, 16)


                                sRespuestax = ""
                                sRespuestax = "TARJETA RIPLEY CLASICA|\t|" & sTarjetaRecuperada.Trim & "|\t|" & "S/. " & sDisponibleCompras.Trim

                            Else 'Cualquier Otro Caso
                                sRespuestax = ""
                            End If


                        End If

                    Else 'Sino devuelve nada
                        sRespuestax = ""
                    End If


                ElseIf sRespuestax = "-2" Then 'Ocurrio un error en el mirapi web
                    'sRespuestax = "ERROR:" & errorMsg_.Trim
                    sRespuestax = ""
                Else
                    'sRespuestax = "ERROR:Servicio no disponible."
                    sRespuestax = ""
                End If

            Else
                sRespuestax = ""
            End If

        Catch ex As Exception
            sRespuestax = ""
        End Try


        Return sRespuestax

    End Function


    'BUSCAR LINEA DISPONIBLE DE COMPRAS ASOCIADAS
    Private Function FUN_BUSCAR_LINEA_DISPONIBLE_COMPRAS_ASOCIADA(ByVal sNroTarjeta As String) As String
        Dim sGetTrama_MC As String = ""
        Dim sParametros_MC As String = ""
        Dim sXMLx As String = ""

        Try

            If sNroTarjeta.Trim.Length = 16 Then

                Dim obSendWAS_ As New WSCONSULTAS_MC_VS.defaultService

                sParametros_MC = "HQ000001073RMDQCS" & sNroTarjeta.Trim & Strings.StrDup(16, " ") & "0000000000" & "0" & "000000" & "000073" & "@"

                sGetTrama_MC = obSendWAS_.execute(sParametros_MC.Trim)

                'Evaluar Respuesta si es ERROR
                If sGetTrama_MC.Trim.Length > 0 Then
                    If Left(sGetTrama_MC.Trim, 2) = "HE" Then
                        sGetTrama_MC = ""
                    Else

                        'VARIABLES DEL SALDO DE TARJETA ASOCIADA
                        Dim sFechaPago As String = "" 'FECHA DE PAGO
                        Dim sDia As String = ""
                        Dim sMes As String = ""
                        Dim sAnio As String = ""
                        Dim sDisponibleCompras As String = ""

                        sDisponibleCompras = Format(Val(SET_MONTO_DECIMAL(CLng(Mid(sGetTrama_MC, 147, 14)).ToString)), "###,##0.00")

                        'BUSCAR TIPO DE TARJETA NUMERO_TARJETA 
                        Dim SBINN_TARJETA As String = ""
                        Dim TarjetaBin As New TarjetaBin

                        SBINN_TARJETA = Mid(sNroTarjeta, 1, 6)
                        TarjetaBin = BNTarjetaBin.Instancia.ObtenerTarjetaBin(SBINN_TARJETA)


                        sGetTrama_MC = "TARJETA RIPLEY " & TarjetaBin.Nombre.Trim() & "|\t|" & sNroTarjeta.Trim & "|\t|" & "S/. " & sDisponibleCompras.Trim

                    End If

                Else 'Sino devuelve nada
                    sGetTrama_MC = ""
                End If


            Else
                'Mostrar Mensaje de Error
                sGetTrama_MC = ""
            End If

        Catch ex As Exception
            sGetTrama_MC = ""

        End Try


        Return sGetTrama_MC

    End Function





    '********************************METODOS PARA GIF CARD***************************************

    'numero de tarjeta=2121210040000121 y opcion=1
    <WebMethod(Description:="Datos Tarjeta Gif Card.")> _
    Public Function DATOS_GIF_CARD(ByVal sNumero As String, ByVal sOpcion As String, ByVal sDATA_MONITOR_KIOSCO As String) As String
        Dim obGifCard As New WS_GIF_CARD.wsgcpConsultaSaldo
        Dim arrData As Array
        Dim sResultadoGIF As String = ""
        Dim sNroSecuencia As String = ""
        Dim sSaldoDisponibleGIF As String = ""

        Try
            ErrorLog(obGifCard.Url)
            obGifCard.UseDefaultCredentials = True
            arrData = obGifCard.gcpconssaldos(sNumero.Trim, sOpcion)

            If UBound(arrData) > -1 Then
                'Campos gif card
                Dim strNroTarjeta As String = arrData(0).NumeroTarjeta.ToString
                Dim strSecuencia As String = arrData(0).NumeroSecuencia.ToString
                Dim strOpcion As String = arrData(0).Opcion.ToString
                Dim strEstado As String = arrData(0).Estado.trim
                Dim strMonto As String = arrData(0).Monto.ToString
                Dim sFechaKiosco As String = DateTime.Now.Day.ToString("00").Trim + "/" + DateTime.Now.Month.ToString("00").Trim + "/" + DateTime.Now.Year.ToString("0000").Trim
                Dim sHoraKiosco As String = DateTime.Now.ToString("HH:mm:ss")
                Dim sNOperacion As String = GET_NUMERO_OPERACION()

                If strEstado.Trim.ToUpper = "AC" Then
                    sNroSecuencia = strSecuencia.Trim
                    sSaldoDisponibleGIF = strMonto.Trim
                    sResultadoGIF = strNroTarjeta.Trim & "|\t|" & strSecuencia.Trim & "|\t|" & strMonto.Trim & "|\t|" & sFechaKiosco.Trim & "|\t|" & sHoraKiosco.Trim & "|\t|" & sNOperacion.Trim
                Else
                    sResultadoGIF = "ERROR:Su tarjeta no es válida."

                    'Select Case strEstado.Trim.ToUpper

                    '    Case "IN" '
                    '        sResultadoGIF = "Su tarjeta no es válida."
                    '    Case "CA" 'CANCELADA

                    'End Select
                End If
            Else
                sResultadoGIF = "ERROR:Su tarjeta no es válida."
            End If
        Catch ex As Exception
            sResultadoGIF = "ERROR:Su tarjeta no es válida."
        End Try

        'Save Monitor
        Dim sDataMonitor As String = ""
        Dim FECHA_TX As String = DateTime.Now.Year.ToString("0000").Trim & DateTime.Now.Month.ToString("00").Trim + DateTime.Now.Day.ToString("00").Trim
        Dim HORA_TX As String = DateTime.Now.Hour.ToString("00").Trim & DateTime.Now.Minute.ToString("00").Trim & DateTime.Now.Second.ToString("00").Trim
        Dim sModoEntrada As String = ""
        Dim sCanalAtencion As String = "01" 'RipleyMatico
        Dim sSaldoDisponible As String = "      " 'Right("      " & sSaldoDisponibleGIF.Trim, 6)
        Dim sTotalDeuda As String = "      "
        Dim sIDSucursal As String = ""
        Dim sIDTerminal As String = ""
        Dim sEstadoCuenta As String = ""
        Dim sRespuestaServidor As String = ""
        Dim sCodigoTransaccion As String = ""
        Dim sNroCuenta As String = ""
        Dim sNroTarjeta As String = ""
        Dim sCliente As String = ""


        If sDATA_MONITOR_KIOSCO.Length > 0 Then
            'ORDEN DE LAS VARRIABLES
            'DATA_MONITOR=_root.CUENTA_TARJETA_VALIDAR+"|\u005Ct|"+_root.P_NroTarjeta+"|\u005Ct|"+"01"+"|\u005Ct|"+_global.gCODE_SUCURSAL+"|\u005Ct|"+_global.gCODE_KIOSCO+"|\u005Ct|"+"01";
            'RECORRER ARREGLO
            Dim ADATA_MONITOR As Array
            ADATA_MONITOR = Split(sDATA_MONITOR_KIOSCO, "|\t|", , CompareMethod.Text)

            sNroCuenta = sNroSecuencia
            sNroTarjeta = ADATA_MONITOR(1)
            sModoEntrada = ADATA_MONITOR(2)
            sIDSucursal = ADATA_MONITOR(3)
            sIDTerminal = ADATA_MONITOR(4)
            sCodigoTransaccion = ADATA_MONITOR(5) '02= Identificacion y Consulta de datos. Login y mostrar datos de saldos GIF CARD

            If sSaldoDisponibleGIF.Trim.Length > 0 Then
                sSaldoDisponible = sSaldoDisponibleGIF
                sSaldoDisponible = Replace(sSaldoDisponible.Trim, ",", "") 'QUITAR LA COMA
                sSaldoDisponible = Replace(sSaldoDisponible.Trim, ".", "") 'QUITAR EL PUNTO DECIMAL
                sSaldoDisponible = Mid(sSaldoDisponible, 1, sSaldoDisponible.Trim.Length - 2) 'QUITAR LOS DOS ULTIMOS DIGITOS QUE SON DECIMALES
            Else
                sSaldoDisponible = "0"
            End If
            sSaldoDisponible = Right("      " & sSaldoDisponible.Trim, 6)

            'GUADAR LOG CONSULTAS
            SAVE_LOG_CONSULTAS(sIDSucursal.Trim, sIDTerminal.Trim, "58", FUN_BUSCAR_TIPO_TARJETA(sNroTarjeta.Trim))

            'armar variable desde flash con: numero de cuenta, numero de tarjeta, nombre del cliente, modo de entrada, canal de atencion, id de sucursal,codigo kiosco y codigo de transaccion (opcion)
            If Left(sResultadoGIF.Trim, 5) <> "ERROR" Then
                'VALIDACION CORRECTA
                sNroCuenta = Right("          " & sNroCuenta.Trim, 10)
                sNroTarjeta = Right("                    " & sNroTarjeta.Trim, 20)
                sCliente = Right(Strings.StrDup(26, " ") & sCliente.Trim, 26)
                sEstadoCuenta = "01" 'Estado de la cuenta
                sRespuestaServidor = "01" 'Atendido

                'ENVIAR_MONITOR
                sDataMonitor = ""
                sDataMonitor = FECHA_TX & HORA_TX & sNroCuenta & sNroTarjeta & sCliente & sModoEntrada & sCanalAtencion & sSaldoDisponible & sTotalDeuda & sIDSucursal & sIDTerminal & sEstadoCuenta & sRespuestaServidor & sCodigoTransaccion
                ENVIAR_MONITOR(sDataMonitor)
            Else
                'ENVIAR_MONITOR
                sNroCuenta = Right("          " & sNroCuenta.Trim, 10)
                sNroTarjeta = Right("                    " & sNroTarjeta.Trim, 20)
                sCliente = Strings.StrDup(26, " ")
                sEstadoCuenta = "02" 'Estado de la cuenta
                sRespuestaServidor = "02" 'Rechazado

                sDataMonitor = ""
                sDataMonitor = FECHA_TX & HORA_TX & sNroCuenta & sNroTarjeta & sCliente & sModoEntrada & sCanalAtencion & sSaldoDisponible & sTotalDeuda & sIDSucursal & sIDTerminal & sEstadoCuenta & sRespuestaServidor & sCodigoTransaccion
                ENVIAR_MONITOR(sDataMonitor)
            End If
        End If

        Return sResultadoGIF
    End Function



    '    '********************************* FUNCIONES PARA PONER EL SIMBOLO DECIMAL

    'FORMATEAR A DECIMALES LOS DOS ULTIMOS NUMEROS
    '<WebMethod(Description:="FORMATEAR LOS DECIMALES")> _
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


    'Calculo del periodo de facturación
    Private Function funDevuelve_FechaPago(ByVal DatFechaHoy As Date, _
                                    ByVal IntDiaPago As Integer) As Date
        Dim ldatFecActual As Date
        Dim ldatFecPago As Date
        Dim ldif As Long = 0

        Try

            'ldatFecActual = Format(Day(DatFechaHoy), "00") & "/" & _
            '              Format(Month(DatFechaHoy), "00") & "/" & _
            '              Format(Year(DatFechaHoy), "0000")

            ldatFecActual = New Date(Format(Year(DatFechaHoy), "0000"), Format(Month(DatFechaHoy), "00"), Format(Day(DatFechaHoy), "00"))

            'La Fecha de Pago es la del mes actual.
            'ldatFecPago = Format(IntDiaPago, "00") & "/" & _
            '              Format(Month(DatFechaHoy), "00") & "/" & _
            '              Format(Year(DatFechaHoy), "0000")

            ldatFecPago = New Date(Format(Year(DatFechaHoy), "0000"), Format(Month(DatFechaHoy), "00"), Format(IntDiaPago, "00"))

            'Si la Fecha de Pago es Mayor igual a la Fecha Actual
            If ldatFecPago >= ldatFecActual Then
                'Si la diferencia de Fechas es mayor a 15 dias...
                ldif = DateDiff(DateInterval.Day, ldatFecActual, ldatFecPago)
                If ldif >= 15 Then
                    'La Fecha de Pago es la del mes anterior al mes actual
                    ldatFecPago = DateAdd("m", -1, ldatFecPago)
                End If
            Else
                'la Fecha de Pago es la del mes siguiente...
                ldatFecPago = DateAdd("m", 1, ldatFecPago)
                'Si la diferencia de Fechas es mayor a 15 dias...

                ldif = DateDiff(DateInterval.Day, ldatFecActual, ldatFecPago)

                If ldif >= 15 Then
                    'La Fecha de Pago es la del mes anterior al mes actual
                    ldatFecPago = DateAdd("m", -1, ldatFecPago)
                End If


            End If


        Catch ex As Exception

            funDevuelve_FechaPago = ldatFecPago

        End Try

        funDevuelve_FechaPago = ldatFecPago

    End Function


    Private Sub procDevuelve_PeriodoFacturacion(ByVal DatFechaPago As Date, _
                                    ByRef DatFechaCorteIni As Date, ByRef DatFechaCorteFin As Date)

        Try

            'Resto 15 dias a la Fecha de Pago para obtener la Fecha de Corte Final.
            DatFechaCorteFin = DateAdd("d", -15, DatFechaPago)

            'Resto 1 Mes a la Fecha de Pago
            Dim fPagoAux As Date
            fPagoAux = DateAdd("m", -1, DatFechaPago)
            'Resto 15 dias a la Fecha de Pago del Periodo anterior y le sumo 1 dia para obtener la Fecha de Corte Inicial.
            fPagoAux = DateAdd("d", -15, fPagoAux)
            fPagoAux = DateAdd("d", 1, fPagoAux)

            DatFechaCorteIni = fPagoAux

        Catch ex As Exception

        End Try

    End Sub


    'METODOS DE CLIENTE SOCKETS MONITOR
    Private Function ENVIAR_MONITOR(ByVal sDatosEnvio As String) As String

        Dim sEstadoEnvio As String = String.Empty

        Try


            If sDatosEnvio.Trim.Length > 0 Then


                Dim tcpClient As New System.Net.Sockets.TcpClient()
                tcpClient.Connect(ConfigurationManager.AppSettings("IP_SERVER_MONITOR"), Val(ConfigurationManager.AppSettings("PUERTO_SERVER_MONITOR")))
                Dim networkStream As NetworkStream = tcpClient.GetStream()
                If networkStream.CanWrite And networkStream.CanRead Then
                    ' Do a simple write.
                    Dim sendBytes As [Byte]() = Encoding.ASCII.GetBytes(sDatosEnvio.Trim)
                    networkStream.Write(sendBytes, 0, sendBytes.Length)
                    sEstadoEnvio = "ENVIADO"
                Else
                    sEstadoEnvio = "NOENVIADO"
                End If

                tcpClient.Close()

            End If

        Catch ex As Exception

            'OCURRIO UN ERROR
            sEstadoEnvio = "NOENVIADO"
        End Try


        Return sEstadoEnvio.Trim

    End Function


    'DEVULEVE EL TIPO DE TARJETA
    'RIPLEY CLASICA  96041             1
    'RIPLEY MC GOLD  542020            2
    'RIPLEY MC SILVER  525474          3
    'RIPLEY VISA PLATINIUM  450035     4
    'RIPLEY VISA SILVER  450034        5
    'GIF CARD       212121             6

    Private Function FUN_BUSCAR_TIPO_TARJETA(ByVal SNRO_TARJETA As String) As String

        'BUSCAR TIPO DE TARJETA NUMERO_TARJETA 
        Dim TarjetaBin As New TarjetaBin
        Dim SBINN_TARJETA As String = ""
        Dim STIPO_TARJETA As String = ""

        Try
            If SNRO_TARJETA.Trim.Length > 0 Then
                SBINN_TARJETA = Mid(SNRO_TARJETA.Trim, 1, 6)
                TarjetaBin = BNTarjetaBin.Instancia.ObtenerTarjetaBin(SBINN_TARJETA)
            End If

            If TarjetaBin.Tipo = String.Empty Then
                TarjetaBin.Tipo = "0"
            End If
        Catch ex As Exception
            ErrorLog(ex.Message)
        End Try

        Return TarjetaBin.Tipo
    End Function

    'Buscar el codigo del tipo de producto del superefectivo SEF
    Private Function FUN_BUSCAR_TIPO_PRODUCTO_SEF(ByVal SBIN_TARJETA As String) As String

        Dim TarjetaBin As New TarjetaBin
        Try
            If SBIN_TARJETA.Trim.Length > 0 Then
                TarjetaBin = BNTarjetaBin.Instancia.ObtenerTarjetaBin(SBIN_TARJETA)
            End If

            If TarjetaBin.TipoSEF = String.Empty Then
                TarjetaBin.TipoSEF = "0"
            End If
        Catch ex As Exception
            ErrorLog(ex.Message)
        End Try

        Return TarjetaBin.TipoSEF

    End Function


    Private Function FUN_BUSCAR_TIPO_TARJETA_SIMULADOR_OFERTA(ByVal SNRO_TARJETA As String) As String

        'BUSCAR TIPO DE TARJETA NUMERO_TARJETA 
        Dim SBINN_TARJETA As String = ""
        Dim TarjetaBin As New TarjetaBin

        Try
            If SNRO_TARJETA.Trim.Length > 0 Then
                SBINN_TARJETA = Mid(SNRO_TARJETA.Trim, 1, 6)
                TarjetaBin = BNTarjetaBin.Instancia.ObtenerTarjetaBin(SBINN_TARJETA)
            End If

            If TarjetaBin.TipoSimuladorSEF = String.Empty Then
                TarjetaBin.TipoSimuladorSEF = "0"
            End If
        Catch ex As Exception
            ErrorLog(ex.Message)
        End Try

        Return TarjetaBin.TipoSimuladorSEF
    End Function


    '<INI>
    Private Function VALIDAR_TARJETA_CLASICA_RSAT(ByVal NroTarjeta As String, _
                                                  ByVal sDATA_MONITOR_KIOSCO As String, _
                                                  ByRef Nro_Contrato As String, _
                                                  ByRef Tipo_Tarjeta_Ofertas As String, _
                                                  Optional ByRef sDataSuperEfectivoSEF As String = "") As Respuesta

        Dim objMQ As New MQ
        Dim strServicio As String = String.Empty
        Dim strMensaje As String = String.Empty
        Dim strParametros As String = String.Empty
        Dim sRespuesta As String = String.Empty
        Dim oRespuesta As Respuesta = New Respuesta
        Dim strPan As String = String.Empty
        Dim strNroPan As String = String.Empty
        Dim lngLargo As Long = 0
        Dim i As Integer = 0
        Dim titular As String = String.Empty
        NroTarjeta = NroTarjeta.Trim


        'strServicio = "SFSCANC0012"
        'strMensaje = "0000000000SFSCANC0012                                       000000065PE00010000080027RSAT     IVR00           00" + NroTarjeta + "                          0"
        strServicio = ReadAppConfig("SFSCAN_TARJETA_CLASICA_RSAT")
        strMensaje = "0000000000" & strServicio & "                                       000000065PE00010000080027RSAT     IVR00           00" + NroTarjeta + "                          0"

        If NroTarjeta.Substring(0, 6) = "542020" Or NroTarjeta.Substring(0, 6) = "525474" Or _
            NroTarjeta.Substring(0, 6) = "450034" Or NroTarjeta.Substring(0, 6) = "450035" Then
            strMensaje = "0000000000" & strServicio & "                                       000000065PE00010000080027RSAT     IVR00           00" + NroTarjeta + "              9999999999990"
        End If

        objMQ.Service = strServicio
        objMQ.Message = strMensaje
        objMQ.Execute()

        If objMQ.ReasonMd <> 0 Then
            sRespuesta = "ERROR"
            oRespuesta.Estado = "ERROR"
        End If
        If objMQ.ReasonApp = 0 Then
            sRespuesta = objMQ.Response
            oRespuesta.Estado = "EXITO"
            oRespuesta.Cadena = sRespuesta

            titular = Mid(sRespuesta, 261, 2)

            ErrorLog("oRespuesta.Cadena= " & sRespuesta)
        End If
        'sRespuesta = "0000000000SFSCANC0012                                       000002000PE00010000080027RSAT     IVRM00121397    005420200003063168              99999999999900ROSA N. CIURLIZZA K.                                        028957700100010001000002025411050001-01-0102BE0000000000025420200003063168      00                              20                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                      "

        Dim sParametros As String = ""
        Dim sMensajeErrorUsuario As String = ""
        Dim sXML As String = ""

        Dim sNroTarjeta As String = ""
        Dim sNroDocumento As String = ""
        Dim sApellidoParterno As String = ""
        Dim sApellidoMaterno As String = ""
        Dim sNombres As String = ""
        Dim sCliente As String = ""
        Dim sFechaNac As String = ""
        Dim sTipoProducto As String = "" 'Segun el bin de la tarjeta 6 primeros numeros de la tarjeta
        Dim sTipoDocumento As String = ""
        Dim sNroCuenta As String = String.Empty

        Try

            If NroTarjeta.Length = 16 Then

                If oRespuesta.Cadena.Length > 0 Then
                    If oRespuesta.Estado = "ERROR" Then
                        'Save Error
                        sRespuesta = "ERROR:Servicio no disponible."
                        oRespuesta.Cadena = "ERROR:Servicio no Disponible"
                        oRespuesta.Mensaje = "Servicio no Disponible"
                    Else

                        '---- INICIO CORTE ----'
                        Dim fnEstMigra As String
                        fnEstMigra = Mid(sRespuesta, 155, 2)
                        If fnEstMigra = "00" Then
                            sTipoDocumento = sRespuesta.Substring(97, 1) 'C=DNI 2=CEX Tipo de Documento
                            ErrorLog("sTipoDocumento= " & sTipoDocumento)
                            sNroDocumento = sRespuesta.Substring(98, 8) 'Numero de Documento
                            ErrorLog("sTipoDocumento= " & sNroDocumento)
                            sNombres = Mid$(sRespuesta, 197, 20)            '-> Nombres
                            ErrorLog("sTipoDocumento= " & sNombres)
                            sApellidoParterno = Mid$(sRespuesta, 157, 20)   '-> Apellido Paterno
                            sApellidoMaterno = Mid$(sRespuesta, 177, 20)    '-> Apellido Materno
                            Nro_Contrato = Mid$(sRespuesta, 227, 20)
                            Tipo_Tarjeta_Ofertas = "RSAT"
                            sCliente = sApellidoParterno.Trim & " " & sApellidoMaterno.Trim & ", " & sNombres.Trim
                            If (NroTarjeta.Substring(0, 6) = "542020" Or NroTarjeta.Substring(0, 6) = "525474" Or _
                                NroTarjeta.Substring(0, 6) = "450034" Or NroTarjeta.Substring(0, 6) = "450035") And (sApellidoParterno.Trim = "NO MIGRADA") Then

                                Dim obSendWAS As New WSCONSULTAS_MC_VS.defaultService

                                sParametros = "HQ000001073RMDQCS" & NroTarjeta.Trim & Strings.StrDup(16, " ") & "0000000000" & "0" & "000000" & "000073" & "@"
                                ErrorLog("sParametros= " & sParametros)
                                sCliente = obSendWAS.execute(sParametros.Trim)
                                ErrorLog("sCliente= " & sCliente)
                                'sCliente = "                                                                        ESPINOZA CAJAHUARINGA, SANDRO          "
                                sCliente = Trim(Mid(sCliente, 73, 30))
                                sCliente = Replace(sCliente, "#", "Ñ")

                                sCliente = sCliente & "||SI"

                            ElseIf (NroTarjeta.Substring(0, 6) = "542020" Or NroTarjeta.Substring(0, 6) = "525474" Or _
                                    NroTarjeta.Substring(0, 6) = "450034" Or NroTarjeta.Substring(0, 6) = "450035") And (sApellidoParterno.Trim <> "NO MIGRADA") Then

                                sCliente = sCliente & "||NO"

                            End If

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
                                If VALIDAR_BLOQUEDO_TARJETA_RSAT(vg_RSAT_CodBloqueo) = False And vg_RSAT_Sitarj = "05" And vg_RSAT_FecBaj = "0001-01-01" And (vg_RSAT_CodTitularAdicional = "TI" Or vg_RSAT_CodTitularAdicional = "BE") Then
                                    bCondCli = True
                                Else
                                    oRespuesta.Mensaje = ValidarMensajeAMostrarEnRipleymatico(VALIDAR_BLOQUEDO_TARJETA_RSAT(vg_RSAT_CodBloqueo), vg_RSAT_Sitarj, vg_RSAT_FecBaj, "")
                                    oRespuesta.Estado = "ERROR"
                                    oRespuesta.Codigo = "MSGMSG"
                                End If

                                If bCondCli Then
                                    Exit For
                                End If
                            Next iIndEle

                            If bCondCli Then
                                '<INI TCK-563699-01 DHERRERA 20-03-2014>
                                If sTipoDocumento.Trim = "C" Then 'DNI
                                    sTipoDocumento = 1
                                ElseIf sTipoDocumento.Trim = "2" Then
                                    sTipoDocumento = 2
                                Else
                                    sTipoDocumento = 1
                                End If
                                '<FIN TCK-563699-01 DHERRERA 20-03-2014>

                                sTipoProducto = Mid(NroTarjeta.Trim, 1, 6)
                                ErrorLog("sTipoProducto" & sTipoProducto)
                                sDataSuperEfectivoSEF = MOSTRAR_SUPEREFECTIVO_SEF("1", sTipoDocumento.Trim, sNroDocumento.Trim)
                                ErrorLog("sDataSuperEfectivoSEF" & sDataSuperEfectivoSEF)
                                sXML = NroTarjeta.Trim & "|\t|" & sNroDocumento.Trim & "|\t|" & sCliente.Trim & "|\t|" & sFechaNac.Trim & "|\t|" & sTipoProducto.Trim & "|\t|" & MOSTRAR_SUPEREFECTIVO(sNroCuenta.Trim, sNroDocumento.Trim, TServidor.RSAT) & "|\t|" & sNroCuenta.Trim

                                Dim ADATA_MONITOR As Array
                                Dim sIDSucursal, sIDTerminal As String
                                sIDSucursal = ""
                                sIDTerminal = ""

                                If sDATA_MONITOR_KIOSCO.Length > 0 Then
                                    ADATA_MONITOR = Split(sDATA_MONITOR_KIOSCO, "|\t|", , CompareMethod.Text)
                                    sIDSucursal = ADATA_MONITOR(3)
                                    sIDTerminal = ADATA_MONITOR(4)

                                    getRegistrar_Ingreso_Ripleymatico(sIDSucursal, sIDTerminal, sNroDocumento, sTipoDocumento, sNroTarjeta, sTipoProducto)

                                End If



                            Else
                                'sXML = "ERROR:Tarjeta no valida."
                            End If

                            sRespuesta = sXML
                            oRespuesta.Cadena = sXML


                            Select Case titular.ToUpper()
                                Case Constantes.Titular
                                    oRespuesta.EsTitular = "1"
                                Case Constantes.Beneficiario
                                    oRespuesta.EsTitular = "0"
                                Case Else
                                    oRespuesta.EsTitular = "0"
                            End Select


                        Else
                            sRespuesta = "ERROR:Tarjeta no valida."
                            oRespuesta.Cadena = "ERROR:Tarjeta no valida."
                        End If



                    End If

                Else 'Sino devuelve nada
                    sRespuesta = "ERROR:Servicio no disponible"
                    oRespuesta.Cadena = "ERROR:Tarjeta no valida."
                End If


            Else
                'Mostrar Mensaje de Error
                sRespuesta = "ERROR:Tarjeta no valida."
                oRespuesta.Cadena = "ERROR:Tarjeta no valida."
            End If

        Catch ex As Exception

            sRespuesta = "ERROR:Tarjeta no valida."
            oRespuesta.Cadena = "ERROR:Tarjeta no valida."

        End Try


        'Return sRespuesta
        Return oRespuesta



    End Function

    Private Function SALDO_TARJETA_CLASICA_RSAT(ByVal nrotarjeta As String, _
                                                ByVal nrocuenta As String, _
                                                ByVal sMigrado As String) As String

        Dim objMQ As New MQ
        Dim strServicio As String = String.Empty
        Dim strMensaje As String = String.Empty
        Dim strParametros As String = String.Empty
        Dim strRespuesta As String = String.Empty
        Dim strPan As String = String.Empty
        Dim strNroPan As String = String.Empty
        Dim lngLargo As Long = 0
        Dim i As Integer = 0

        nrocuenta = nrocuenta.Trim
        nrotarjeta = nrotarjeta.Trim

        If (nrotarjeta.Substring(0, 6) = "542020" Or nrotarjeta.Substring(0, 6) = "525474" Or _
            nrotarjeta.Substring(0, 6) = "450034" Or nrotarjeta.Substring(0, 6) = "450035") And sMigrado = "SI" Then
            nrocuenta = "000000000000"
        End If


        'strServicio = "SFSCANT0114"
        'strMensaje = "0000000000SFSCANT0114                                       000000069000002USUARIO127  SANI0001PE0001" & nrocuenta & "" & nrotarjeta & "      604"
        strServicio = ReadAppConfig("SFSCAN_SALDO_FECHACORTE_CLASICA_RSAT")
        strMensaje = "0000000000" & strServicio & "                                       000000069000002USUARIO127  SANI0001PE0001" & nrocuenta & "" & nrotarjeta & "      604"
        objMQ.Service = strServicio
        objMQ.Message = strMensaje
        objMQ.Execute()

        If objMQ.ReasonMd <> 0 Then
            strRespuesta = "ERROR"
        End If
        If objMQ.ReasonApp = 0 Then
            strRespuesta = objMQ.Response
        End If

        'strRespuesta = "0000000000SFSCANT0114                                       000005895000002USUARIO127  SANI0001PE00010000000007049604100079592514      6040100000000000300000 000000000000000001120001COMPRAS                       00000000000292620120002EFECTIVO EXPRESS              00000000000292620120003NO EXISTE                     00000000000000000240001COMPRAS                       00000000000292620360001COMPRAS                       0000000000029262000                                  0000000000000000000                                  0000000000000000000                                  0000000000000000000                                  0000000000000000000                                  0000000000000000000                                  0000000000000000000                                  0000000000000000000                                  0000000000000000000                                  0000000000000000000                                  0000000000000000000                                  0000000000000000000                                  0000000000000000000                                  0000000000000000000                                  0000000000000000000                                  00000000000000000 0000000000000000 000000000000000000 000000000000000000000000 000000000000000000000000 000000000000000000000000 000000000000000000000000 000000000000000000000000 000000000000000000000000 000000000000000000000000 000000000000000000000000 000000000000000000000000 000000000000000000000000 0000000000000000 0000000000000000 00000000000000000000000                                                                                             0000000000000000  0000                                                                                             0000000000000000  0000                                                                                             0000000000000000  0000                                                                                             0000000000000000  0000                                                                                             0000000000000000  0000                                                                                             0000000000000000  0000                                                                                             0000000000000000  0000                                                                                             0000000000000000  0000                                                                                             0000000000000000  0000                                                                                             0000000000000000  0000                                                                                             0000000000000000  0000                                                                                             0000000000000000  0000                                                                                             0000000000000000  0000                                                                                             0000000000000000  0000                                                                                             0000000000000000  0000                                                                                             0000000000000000  0000                                                                                             0000000000000000  0000                                                                                             0000000000000000  0000                                                                                             0000000000000000  0000                                                                                             0000000000000000  0000                                                                                             0000000000000000  0000                                                                                             0000000000000000  0000                                                                                             0000000000000000  0000                                                                                             0000000000000000  0000                                                                                             0000000000000000                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                      "

        'VARIABLES DEL SALDO DE TARJETA CLASICA
        Dim sLineaCredito As String = "" 'LINEA DE CRÉDITO
        Dim dLineaCreditoUtilizada As String = "" 'LÍNEA DE CRÈDITO UTILIZADA
        Dim sDisponibleCompras As String = "" 'DISPONIBLE COMPRAS
        Dim sDisponibleEfectivo As String = "" 'DISPONIBLE EFECTIVO EXPRESS
        Dim sPagoTotalMes As String = "" 'PAGO TOTAL DEL MES
        Dim sPagoMinimo As String = "" 'PAGO MINIMO DEL MES
        Dim sPeriodo_Facturacion As String = "" 'PERIODO DE FACTURACION
        Dim sFechaPago As String = "" 'FECHA DE PAGO
        Dim sRipleyPuntos As String = "" 'RIPLEY PUNTOS
        Dim sDisponibleSuperEfectivo As String = ""
        Dim sPagoTotal As String = "" 'PAGO TOTAL

        Dim sXML As String

        Dim sFechaKiosco As String = DateTime.Now.Day.ToString("00").Trim + "/" + DateTime.Now.Month.ToString("00").Trim + "/" + DateTime.Now.Year.ToString("0000").Trim
        Dim sHoraKiosco As String = DateTime.Now.ToString("HH:mm:ss")
        Dim sNOperacion As String = GET_NUMERO_OPERACION()


        Dim strResp114, strResp103, strResp113, strResp114_2 As String

        strResp114 = String.Empty
        strResp103 = String.Empty
        strResp113 = String.Empty
        strResp114_2 = String.Empty

        If Val(Mid(strRespuesta, 1, 10)) = 0 Then

            strResp114 = Mid(strRespuesta, 139, 2)
            strResp103 = Mid(strRespuesta, 141, 1096)
            strResp113 = Mid(strRespuesta, 1237, 303)
            strResp114_2 = Mid(strRespuesta, 1540, 17)

        End If


        Dim sMontoLinEE, sMontoLinSE, ws_MontosDisp, sLineaAct, sMontoLin, sMontoLin12 As String
        Dim iNumLin, sInd As Integer

        sMontoLinEE = "0.00"
        sMontoLinSE = "0.00"
        sMontoLin12 = String.Empty

        iNumLin = Val(Mid(strResp103, 35, 2))
        iNumLin = 3

        ws_MontosDisp = Mid(strResp103, 37)

        sInd = 0

        While sInd < iNumLin
            sLineaAct = Mid(ws_MontosDisp, 3 + 53 * sInd, 4)
            sMontoLin = Format(Val(Mid(ws_MontosDisp, 37 + 53 * sInd, 17)) / 100, "###00.00")
            If sLineaAct = "0001" Then
                sMontoLin12 = sMontoLin
            Else
                If sLineaAct = "0002" Then
                    sMontoLinEE = sMontoLin
                Else
                    If sLineaAct = "0003" Then
                        sMontoLinSE = sMontoLin
                    End If
                End If
            End If
            sInd = sInd + 1
        End While

        Dim stdPagoTot, stdPagoMin As String

        stdPagoTot = Mid(strResp113, 1, 17) 'Cuota Mes
        stdPagoMin = Mid(strResp113, 18, 17) 'Pago Minimo

        Dim lstrFechaEvaluacion, lintDiaVencimiento, vDia As String

        lstrFechaEvaluacion = Date.Now().ToString("yyyyMMdd")
        lintDiaVencimiento = Val(strResp114)

        Dim vFecHoy, vFecPago As Date
        Dim vFecCorteIni As Date
        Dim vFecCorteFin As Date
        Dim sFecha1, sFecha2, sFechaPagox As String

        vFecHoy = Date.Now
        vDia = lintDiaVencimiento

        vFecPago = CStr(funDevuelve_FechaPago(vFecHoy, vDia))
        procDevuelve_PeriodoFacturacion(vFecPago, vFecCorteIni, vFecCorteFin)

        'Formatear DDMMYYYY
        sFechaPagox = Right(Trim("00" & Day(vFecPago).ToString), 2) & Right(Trim("00" & Month(vFecPago).ToString), 2) & Right(Trim("0000" & Year(vFecPago).ToString), 4)

        If sFechaPagox = "01010001" And (sMigrado = "" Or sMigrado = "NO") Then
            Return "ERROR:NODATA"
        End If


        sFecha1 = Right(Trim("00" & Day(vFecCorteIni).ToString), 2) & Right(Trim("00" & Month(vFecCorteIni).ToString), 2) & Right(Trim("0000" & Year(vFecCorteIni).ToString), 4)
        sFecha2 = Right(Trim("00" & Day(vFecCorteFin).ToString), 2) & Right(Trim("00" & Month(vFecCorteFin).ToString), 2) & Right(Trim("0000" & Year(vFecCorteFin).ToString), 4)

        'PERIODO DE FACTURACION
        If sMigrado <> "SI" Then
            sFechaPago = Cambia_Fecha(sFechaPagox.Trim)
        End If

        Dim BINN_ASOCIADA As String = nrotarjeta.Substring(0, 6)

        If BINN_ASOCIADA = BINN_TARJETA.RSAT_MC_GOLD Or _
                BINN_ASOCIADA = BINN_TARJETA.RSAT_MC_SILVER Or _
                BINN_ASOCIADA = BINN_TARJETA.RSAT_VISA_SILVER Or _
                BINN_ASOCIADA = BINN_TARJETA.RSAT_VISA_PLATINUM Or _
                (BINN_ASOCIADA = BINN_TARJETA.MC_GOLD And sMigrado = "NO") Or _
                (BINN_ASOCIADA = BINN_TARJETA.VISA_PLATINUM And sMigrado = "NO") Or _
                (BINN_ASOCIADA = BINN_TARJETA.VISA_SILVER And sMigrado = "NO") Or _
                (BINN_ASOCIADA = BINN_TARJETA.MC_SILVER And sMigrado = "NO") Or _
                (BINN_TARJETA.CLASICA = nrotarjeta.Substring(0, 5)) Then

            sPeriodo_Facturacion = getPeriodo_Facturacion_RSAT(nrotarjeta, nrocuenta)
        Else
            sPeriodo_Facturacion = Cambia_Fecha(sFecha1.Trim) & " - " & Cambia_Fecha(sFecha2.Trim)

        End If

        If (BINN_ASOCIADA = BINN_TARJETA.MC_GOLD And sMigrado = "SI") Or _
           (BINN_ASOCIADA = BINN_TARJETA.VISA_PLATINUM And sMigrado = "SI") Or _
           (BINN_ASOCIADA = BINN_TARJETA.VISA_SILVER And sMigrado = "SI") Or _
           (BINN_ASOCIADA = BINN_TARJETA.MC_SILVER And sMigrado = "SI") Then

            Dim sTrama_MC As String
            sTrama_MC = SALDO_TARJETA_ASOCIADA_OLD(nrotarjeta, "")
            'sTrama_MC = "2,220.00|\t|1,357.56|\t|862.44|\t|862.00|\t|481.63|\t|0.00|\t|11/09/2013 - 10/10/2013|\t|25/10/2013|\t|2728|\t|0.00|\t|1,357.63|\t|07/11/2013|\t|17:27:17|\t|005138*$¿*5420200004815764|\t|41690197|\t|ALTAMIRANO LEIVA, JUAN ALBERTO|\t||\t|542020|\t||\t||\t|5420200004815764|\t||\t|1|\t|1|\t|41690197|\t|Super Efectivo GOLD MC|\t|JUAN ALBERTO ALTAMIRANO LEIVA|\t|5420200004812308|\t|4000|\t|36|\t|66.59|\t|218.15|\t|18/11/2013|\t|9737965|\t|3.99|\t||\t|5420200004812308|\t|MC"

            Dim aTrama1 As Array
            Dim aTrama2 As Array

            aTrama1 = Split(sTrama_MC, "*$¿*", , CompareMethod.Text)
            aTrama2 = Split(aTrama1(0), "|\t|", , CompareMethod.Text)

            sFechaPago = aTrama2(7)
            sPeriodo_Facturacion = aTrama2(6)
            sDisponibleCompras = Format(Val(sMontoLin12), "##,##0.00")
            sDisponibleEfectivo = aTrama2(3)
            sDisponibleSuperEfectivo = aTrama2(9)
            sPagoTotalMes = aTrama2(4)
            sPagoMinimo = aTrama2(5)
            sPagoTotal = aTrama2(10)
            sLineaCredito = aTrama2(0)
            dLineaCreditoUtilizada = aTrama2(1)
            sRipleyPuntos = aTrama2(8)
        Else
            sDisponibleCompras = Format(Val(sMontoLin12), "##,##0.00")
            sDisponibleEfectivo = Format(Val(sMontoLinEE), "##,##0.00")
            sDisponibleSuperEfectivo = Format(Val(sMontoLinSE), "##,##0.00")

            sPagoTotalMes = Format(CDbl(Val(stdPagoTot)) / 100, "##,##0.00")
            sPagoMinimo = Format(CDbl(Val(stdPagoMin)) / 100, "##,##0.00")
            sPagoTotal = Format(CDbl(Val(strResp114_2)) / 100, "##,##0.00")

            sLineaCredito = Format(CDbl(Val(Mid(strResp103, 1, 17))) / 100, "###,##0.00")
            dLineaCreditoUtilizada = Format(CDbl(Val(Mid(strResp103, 18, 17))) / 100, "###,##0.00")
            sRipleyPuntos = PUNTOS_RIPLEY_ACUMULADOS(nrotarjeta.Trim)
        End If

        sXML = sLineaCredito.Trim & "|\t|" & dLineaCreditoUtilizada.Trim & "|\t|" & sDisponibleCompras.Trim & "|\t|"
        sXML = sXML & sDisponibleEfectivo.Trim & "|\t|" & sPagoTotalMes.Trim & "|\t|" & sPagoMinimo.Trim & "|\t|"
        sXML = sXML & sPeriodo_Facturacion.Trim & "|\t|" & sFechaPago.Trim & "|\t|"
        sXML = sXML & sRipleyPuntos.Trim & "|\t|" & sDisponibleSuperEfectivo.Trim & "|\t|" & sPagoTotal.Trim & "|\t|"
        sXML = sXML & sFechaKiosco.Trim & "|\t|" & sHoraKiosco.Trim & "|\t|" & sNOperacion.Trim


        strRespuesta = sXML


        Return strRespuesta

    End Function

    Enum BINN_TARJETA

        CLASICA = 96041

        VISA_SILVER = 450034
        VISA_PLATINUM = 450035
        MC_GOLD = 542020
        MC_SILVER = 525474

        RSAT_VISA_SILVER = 450000
        RSAT_VISA_PLATINUM = 450007
        RSAT_MC_GOLD = 542070
        RSAT_MC_SILVER = 525435

    End Enum

    Private Function MOVIMIENTOS_CLASICA_RSAT(ByVal nrotarjeta As String, ByVal nrocuenta As String, ByVal sMigrado As String) As String

        Dim objMQ As New MQ
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
        strServicio = ReadAppConfig("SFSCAN_MOVIMIENTOS_CLASICA_RSAT")
        strMensaje = "0000000000" & strServicio & "                                       000000043" + nrotarjeta + "      00010001" + nrocuenta + "M"
        objMQ.Service = strServicio
        objMQ.Message = strMensaje
        objMQ.Execute()

        If objMQ.ReasonMd <> 0 Then
            strRespuesta = "ERROR"
        End If
        If objMQ.ReasonApp = 0 Then
            strRespuesta = objMQ.Response
        End If

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

            Dim oResultas As New List(Of Service.Movimiento)

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


    Enum TServidor

        SICRON = 1
        RSAT = 2

    End Enum

    Private Function ValidarTipFacturaMOV(ByVal sCodigoTipFac As String, ByVal vg_ArrTipFacError As String(,)) As Boolean

        Dim ind As Byte
        Dim bTFEncontrada As Boolean
        bTFEncontrada = False
        If vg_ArrTipFacError.Length > 0 Then
            For ind = 0 To vg_ArrTipFacError.GetLength(0) - 1
                If sCodigoTipFac = Mid(vg_ArrTipFacError(ind, 0), 5, 4) Then
                    bTFEncontrada = True
                    Exit For
                End If
            Next
        End If

        Return bTFEncontrada
    End Function


    Private Function VALIDAR_DNI_ABIERTA_RSAT(ByVal TipProducto As String, _
                                              ByVal TipDocumento As String, _
                                              ByVal NroDocumento As String, _
                                              ByRef Nro_Contrato As String, _
                                              ByRef Tipo_Tarjeta_Ofertas As String) As String

        Dim objMQ As New MQ
        Dim strServicio As String = String.Empty
        Dim strMensaje As String = String.Empty
        Dim strParametros As String = String.Empty
        Dim sRespuesta As String = String.Empty
        Dim strPan As String = String.Empty
        Dim strNroPan As String = String.Empty
        Dim lngLargo As Long = 0
        Dim i As Integer = 0


        'TipDocumento = C : DNI or TipDocumento = 2 : CEX


        If TipDocumento = "1" Then
            TipDocumento = "C"
        End If

        'strServicio = "SFSCANC0040"
        ''strMensaje = "0000000000SFSCANC0012                                       000000065PE00010000020027X8  0       C" & NroDocumento & "    1                       "
        'strMensaje = "0000000000SFSCANC0040                                       000000065PE00010000020027X8  0       " & TipDocumento & NroDocumento & "    1                                           00                                                                    00                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                            "
        strServicio = ReadAppConfig("SFSCAN_DNI_CLASICAABIERTA_RSAT")
        strMensaje = "0000000000" & strServicio & "                                       000000065PE00010000020027X8  0       " & TipDocumento & NroDocumento & "    1                                           00                                                                    00                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                            "

        Try

            objMQ.Service = strServicio
            objMQ.Message = strMensaje
            objMQ.Execute()
            ErrorLog("objMQ.ReasonMd=" & objMQ.ReasonMd)
            If objMQ.ReasonMd <> 0 Then
                sRespuesta = "ERROR: DNI no encontrado"
            End If

            If objMQ.ReasonApp = 0 Then
                ErrorLog("objMQ.Response=" & objMQ.Response)
                sRespuesta = objMQ.Response

                Dim sNroTarjeta As String = String.Empty
                Dim sNroDocumento As String = NroDocumento
                Dim sApellidoParterno As String = String.Empty
                Dim sApellidoMaterno As String = String.Empty
                Dim sNombres As String = String.Empty
                Dim sCliente As String = String.Empty
                Dim sFechaNac As String = String.Empty
                Dim sTipoProducto As String = String.Empty 'Segun el bin de la tarjeta 6 primeros numeros de la tarjeta
                Dim sTipoDocumento As String = String.Empty
                Dim sNroCuenta As String = String.Empty
                Dim sXMLRespuesta As String = String.Empty

                If sRespuesta.Trim.Length > 0 Then
                    If Left(sRespuesta.Trim, 5) = "ERROR" Then
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
                                If bCondCli = True Then
                                    '<FIN TCK-563699-01 DHERRERA 20-03-2014>
                                    If getTipProducto_AbiertaRSAT(sNroTarjeta.Substring(0, 6)).ToString = TipProducto Then

                                        If TipDocumento = "C" Then 'DNI
                                            sTipoDocumento = "1"
                                        Else
                                            sTipoDocumento = "2" 'Carnet de Extranjeria
                                        End If

                                        sTipoProducto = Mid(sNroTarjeta.Trim, 1, 6)
                                        sRespuesta = sNroTarjeta.Trim & "|\t|" & NroDocumento.Trim & "|\t|" & sCliente.Trim & "|\t|" & sFechaNac.Trim & "|\t|" & sTipoProducto.Trim & "|\t|" & MOSTRAR_SUPEREFECTIVO(sNroCuenta.Trim, sNroDocumento.Trim, TServidor.RSAT) & "|\t|" & sNroCuenta.Trim & "|\t|RSAT|\t|" & MOSTRAR_SUPEREFECTIVO_SEF("1", sTipoDocumento.Trim, NroDocumento.Trim) & "|\t|RSAT"
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
            Else
                sRespuesta = "ERROR: DNI no encontrado"
            End If

        Catch ex As Exception
            sRespuesta = "ERROR: Servicio no Disponible"
        End Try




        Return sRespuesta

    End Function


    Private Function getTipProducto_AbiertaRSAT(ByVal BINN As String) As Integer
        Dim TarjetaBin As New TarjetaBin
        Try
            TarjetaBin = BNTarjetaBin.Instancia.ObtenerTarjetaBin(BINN)
        Catch ex As Exception
            ErrorLog(ex.Message)
        End Try

        Return TarjetaBin.TipoAbiertoRsat

    End Function

    Private Function VALIDAR_DNI_CLASICA_RSAT(ByVal TipDocumento As String, _
                                               ByVal NroDocumento As String, _
                                               ByVal sDATA_MONITOR_KIOSCO As String, _
                                               ByRef Nro_Contrato As String, _
                                               ByRef Tipo_Tarjeta_Ofertas As String, _
                                               Optional ByRef sDataSuperEfectivo As String = "") As String

        ErrorLog("Entro a metodo VALIDAR_DNI_CLASICA_RSAT(" & TipDocumento & "," & NroDocumento & "," & sDATA_MONITOR_KIOSCO & "," & Tipo_Tarjeta_Ofertas & "," & sDataSuperEfectivo & ")")

        Dim objMQ As New MQ
        Dim strServicio As String = String.Empty
        Dim strMensaje As String = String.Empty
        Dim strParametros As String = String.Empty
        Dim sRespuesta As String = String.Empty
        Dim strPan As String = String.Empty
        Dim strNroPan As String = String.Empty
        Dim lngLargo As Long = 0
        Dim i As Integer = 0


        'TipDocumento = C : DNI or TipDocumento = 2 : CEX



        'strServicio = "SFSCANC0040"
        ''strMensaje = "0000000000SFSCANC0012                                       000000065PE00010000020027X8  0       C" & NroDocumento & "    1                       "
        'strMensaje = "0000000000SFSCANC0040                                       000000065PE00010000020027X8  0       " & TipDocumento & NroDocumento & "    1                                           00                                                                    00                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                            "
        strServicio = ReadAppConfig("SFSCAN_DNI_CLASICAABIERTA_RSAT")
        strMensaje = "0000000000" & strServicio & "                                       000000065PE00010000020027X8  0       " & TipDocumento & NroDocumento & "    1                                           00                                                                    00                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                            "
        ErrorLog("strMensaje=" & strMensaje)
        Try

            objMQ.Service = strServicio
            objMQ.Message = strMensaje
            objMQ.Execute()

            If objMQ.ReasonMd <> 0 Then
                sRespuesta = "ERROR: DNI no encontrado"
            End If

            If objMQ.ReasonApp = 0 Then
                sRespuesta = objMQ.Response
                ErrorLog("sRespuesta=" & sRespuesta)
                Dim sNroTarjeta As String = String.Empty
                Dim sNroDocumento As String = NroDocumento
                Dim sApellidoParterno As String = String.Empty
                Dim sApellidoMaterno As String = String.Empty
                Dim sNombres As String = String.Empty
                Dim sCliente As String = String.Empty
                Dim sFechaNac As String = String.Empty
                Dim sTipoProducto As String = String.Empty 'Segun el bin de la tarjeta 6 primeros numeros de la tarjeta
                Dim sTipoDocumento As String = String.Empty
                Dim sNroCuenta As String = String.Empty

                If sRespuesta.Trim.Length > 0 Then
                    If Left(sRespuesta.Trim, 5) = "ERROR" Then
                        'Save Error
                        sRespuesta = "ERROR:Servicio no disponible."

                    Else

                        '---- INICIO CORTE ----'
                        ErrorLog("Inicio de Corte")
                        Dim fnEstMigra As String
                        fnEstMigra = Mid(sRespuesta, 155, 2)
                        ErrorLog("fnEstMigra=" & fnEstMigra)
                        If fnEstMigra = "00" Then

                            sNombres = Mid$(sRespuesta, 197, 20)            '-> Nombres
                            sApellidoParterno = Mid$(sRespuesta, 157, 20)   '-> Apellido Paterno
                            sApellidoMaterno = Mid$(sRespuesta, 177, 20)    '-> Apellido Materno
                            Nro_Contrato = Mid$(sRespuesta, 227, 20)
                            Tipo_Tarjeta_Ofertas = "RSAT"
                            sCliente = sApellidoParterno.Trim & " " & sApellidoMaterno.Trim & ", " & sNombres.Trim
                            Dim int_cont, iIndEle, iLonBucle As Integer
                            Dim bCondCli As Boolean

                            Dim vg_RSAT_Sitarj, vg_RSAT_CodBloqueo, vg_RSAT_FecBaj, vg_RSAT_CodTitularAdicional As String

                            int_cont = Val(Mid(sRespuesta, 225, 2))
                            iLonBucle = 104
                            bCondCli = False

                            Dim Binn_Tarjeta As String

                            For iIndEle = 1 To int_cont
                                sNroCuenta = Mid(sRespuesta, 235 + iLonBucle * (iIndEle - 1), 12)
                                vg_RSAT_Sitarj = Mid(sRespuesta, 247 + iLonBucle * (iIndEle - 1), 2) 'SITUACION TARJETA
                                vg_RSAT_FecBaj = Mid(sRespuesta, 249 + iLonBucle * (iIndEle - 1), 10)
                                vg_RSAT_CodTitularAdicional = Mid(sRespuesta, 261 + iLonBucle * (iIndEle - 1), 2)
                                sNroTarjeta = Mid(sRespuesta, 275 + iLonBucle * (iIndEle - 1), 22)
                                vg_RSAT_CodBloqueo = Mid(sRespuesta, 297 + iLonBucle * (iIndEle - 1), 2)
                                Binn_Tarjeta = sNroTarjeta.Substring(0, 5)

                                ErrorLog("Binn_Tarjeta = " & Binn_Tarjeta)
                                ErrorLog("vg_RSAT_CodBloqueo = " & vg_RSAT_CodBloqueo)
                                ErrorLog("vg_RSAT_Sitarj = " & vg_RSAT_Sitarj)
                                ErrorLog("vg_RSAT_FecBaj = " & vg_RSAT_FecBaj)
                                ErrorLog("vg_RSAT_CodTitularAdicional = " & vg_RSAT_CodTitularAdicional)

                                If (Binn_Tarjeta = "96041" Or Binn_Tarjeta = "54202" Or Binn_Tarjeta = "54207") And (vg_RSAT_CodTitularAdicional = "TI" Or vg_RSAT_CodTitularAdicional = "BE") And
                                    VALIDAR_BLOQUEDO_TARJETA_RSAT(vg_RSAT_CodBloqueo) = False And vg_RSAT_Sitarj = "05" And vg_RSAT_FecBaj = "0001-01-01" Then
                                    bCondCli = True
                                End If

                                If bCondCli Then
                                    Exit For
                                End If
                            Next iIndEle

                            ErrorLog("bCondCli=" & bCondCli)
                            If bCondCli Then
                                sTipoProducto = Mid(sNroTarjeta.Trim, 1, 6)
                                'TipDocumento = C : DNI or TipDocumento = 2 : CEX
                                If TipDocumento.Trim = "C" Then
                                    TipDocumento = "1"
                                End If

                                sDataSuperEfectivo = MOSTRAR_SUPEREFECTIVO_SEF("1", TipDocumento.Trim, NroDocumento.Trim)
                                sRespuesta = sNroTarjeta.Trim & "|\t|" & NroDocumento.Trim & "|\t|" & sCliente.Trim & "|\t|" & sFechaNac.Trim & "|\t|" & sTipoProducto.Trim & "|\t|" & MOSTRAR_SUPEREFECTIVO(sNroCuenta.Trim, sNroDocumento.Trim, TServidor.RSAT) & "|\t|" & sNroCuenta.Trim
                            Else
                                sRespuesta = "ERROR:Servicio no disponible."
                            End If
                        Else
                            sRespuesta = "ERROR: DNI no encontrado"
                        End If
                    End If

                Else
                    'Sino Encontro DNI en RSAT
                    sRespuesta = "ERROR: DNI no encontrado"
                End If
            Else
                sRespuesta = "ERROR: DNI no encontrado"
            End If
        Catch ex As Exception
            ErrorLog("Excepción=" & ex.Message)
            sRespuesta = "ERROR: Servicio no Disponible"
        End Try

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


    Private Function MOSTRAR_SUPEREFECTIVO_RSAT(ByVal NroDocumento As String) As String

        Dim objMQ As New MQ
        Dim strServicio As String = String.Empty
        Dim strMensaje As String = String.Empty
        Dim strParametros As String = String.Empty
        Dim sRespuesta As String = String.Empty
        Dim strPan As String = String.Empty
        Dim strNroPan As String = String.Empty
        Dim lngLargo As Long = 0
        Dim i As Integer = 0

        '<INI TCK-XXX DHERRERA 07-05-2014>
        'strServicio = "SFSCANC0018"
        'strMensaje = "0000000000SFSCANC0018                                       00000095500010000020       0036K4  " & Format(Date.Now, "yyyy-mm-dd") & "  PEN                                                                                                                                                                                                                                                                                                                                                                                                                 0                                                                                             0000000000SFSCANC0018                                       000000367PE00010000020036K4  0000000002C" & NroDocumento & "0001000005005030                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 "
        Dim arrSef() As String
        Dim pCodigo_Oferta As String
        Dim pSecuencia As String
        Dim pFec_Ini As String
        Dim pFec_Fin As String
        Dim pGrupo_SEF As String
        Dim pPlazo As String
        Dim pTaza As String
        Dim Monto_SEF As String

        'strServicio = "SFSCANC0041"
        'strMensaje = "0000000000SFSCANC0041                                       00000095500010000020       0080C6  " & Date.Now.Year.ToString() & "-" & Date.Now.Month.ToString("00") & "-" & Date.Now.Day.ToString("00") & "  PEN                                                                                                                                                                                                                                                                                                                                                                                                                 0                                                                                             0000000000SFSCANC0041                                       000000367PE00010000020080C6  0000000002C" & NroDocumento & "    0001000005206150                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 "
        strServicio = ReadAppConfig("SFSCAN_SUPEREFECTIVO_RSAT")
        strMensaje = "0000000000" & strServicio & "                                       00000095500010000020       0080C6  " & Date.Now.Year.ToString() & "-" & Date.Now.Month.ToString("00") & "-" & Date.Now.Day.ToString("00") & "  PEN                                                                                                                                                                                                                                                                                                                                                                                                                 0                                                                                             0000000000" & strServicio & "                                       000000367PE00010000020080C6  0000000002C" & NroDocumento & "    0001000005206150                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 "
        '<FIN TCK-XXX DHERRERA 07-05-2014>

        Try

            objMQ.Service = strServicio
            objMQ.Message = strMensaje
            objMQ.Execute()

            If objMQ.ReasonMd <> 0 Then
                sRespuesta = "ERROR:NODATA"
            End If

            '<INI TCK-XXX DHERRERA 07-05-2014>
            'If objMQ.ReasonApp = 0 Then
            '    sRespuesta = objMQ.Response

            '    Dim SuperEfectivo, FechaVencimiento As String

            '    SuperEfectivo = sRespuesta.Substring(1026, 1300)
            '    FechaVencimiento = sRespuesta.Substring(0, 10)

            '    Dim vFecHoy As Date
            '    Dim vFecVence As Date

            '    vFecHoy = Date.Now
            '    FechaVencimiento = FormatearFecha_DDMMYYYY(FechaVencimiento.Trim)

            '    If FechaVencimiento.Trim <> "" Then

            '        If FechaVencimiento.Trim.Length = 10 Then
            '            vFecVence = CDate(FechaVencimiento.Trim)
            '            If vFecVence > vFecHoy Then
            '                sRespuesta = SuperEfectivo.Trim & "|\t|" & FechaVencimiento.Trim
            '            Else
            '                sRespuesta = "|\t|"
            '            End If

            '        End If
            '    Else
            '        sRespuesta = "|\t|"
            '    End If
            'Else
            '    sRespuesta = "|\t|"
            'End If

            If objMQ.ReasonApp = 0 Then
                sRespuesta = objMQ.Response

                sRespuesta = sRespuesta.Substring(605 + 451)

                arrSef = sRespuesta.Split(",")

                pCodigo_Oferta = arrSef(0)
                pSecuencia = arrSef(1)
                pFec_Ini = arrSef(2)
                pFec_Fin = arrSef(3)
                pGrupo_SEF = arrSef(4)
                pPlazo = arrSef(5)
                pTaza = arrSef(6)
                Monto_SEF = arrSef(7)

                sRespuesta = Monto_SEF.Trim & "|\t|" & pFec_Fin.TrimEnd
            Else
                sRespuesta = "|\t|"
            End If
            '<FIN TCK-XXX DHERRERA 07-05-2014>


        Catch ex As Exception
            sRespuesta = "|\t|"
        End Try

        Return sRespuesta

    End Function


    Private Function getPeriodo_Facturacion_RSAT(ByVal sNroTarjeta As String, ByVal sNroCuenta As String) As String
        Dim sRespuesta As String = ""
        Dim sParametros As String = ""
        Dim sMensajeErrorUsuario As String = ""
        Dim sXML As String = ""
        Dim sAnio As String = ""
        Dim sMes As String = ""
        Dim lContador As Long = 0
        Dim sPagina As String = ""

        'Variables nuevas
        Dim actions As Long = 1
        Dim inetputBuff As String = ""
        Dim outpputBuff As String = ""
        Dim errorMsg As String = ""

        Try

            sAnio = Format(DateAdd("m", 1, Now), "yyyy")
            sMes = Format(DateAdd("m", 1, Now), "mm")

            sNroCuenta = Right(sNroCuenta.Trim, 8)

            If sNroCuenta.Trim.Length = 8 And sNroTarjeta.Length = 16 Then

                'Instancia al mirapiweb
                Dim obSendMirror As ClsTxMirapi = Nothing
                obSendMirror = New ClsTxMirapi()

                Dim sTrama As String = ""
                Dim sXMLMOV As String = ""
                Dim sXMLMOV_FINAL As String = ""
                Dim sXMLCAB As String = ""
                Dim sXMLPIE As String = ""
                Dim sDataMov As String = ""
                Dim sDataMovAux As String = ""

                Dim lFila As Long = 0
                Dim Incrementa As Long = 1

                'VARIABLES RESUMEN DE ESTADO DE CUENTA
                Dim sPeriodoFacturacion As String = ""


                Dim sPeriodoFinal As String = ""
                Dim sPeriodo1 As String = ""
                Dim sPeriodo2 As String = ""
                Dim sPeriodo3 As String = ""
                Dim vFechaHoy As Date
                Dim fFechaHoyAux1 As Date
                Dim lEstadoLlamadaEECC As Long = 9 'Variable para iniciar llamadas al EECC 3 veces


                vFechaHoy = Date.Now
                fFechaHoyAux1 = vFechaHoy

                fFechaHoyAux1 = DateAdd("m", 1, fFechaHoyAux1)

                sPeriodo1 = Right(Trim("0000" & Year(fFechaHoyAux1).ToString), 4) & Right(Trim("00" & Month(fFechaHoyAux1).ToString), 2)

                sPeriodo2 = Right(Trim("0000" & Year(vFechaHoy).ToString), 4) & Right(Trim("00" & Month(vFechaHoy).ToString), 2)

                vFechaHoy = DateAdd("m", -1, vFechaHoy)

                sPeriodo3 = Right(Trim("0000" & Year(vFechaHoy).ToString), 4) & Right(Trim("00" & Month(vFechaHoy).ToString), 2)

                sPeriodoFinal = sPeriodo1.Trim

                Do
                    lContador = lContador + 1

                    If lContador > 9 Then
                        sPagina = lContador.ToString
                    Else
                        sPagina = "0" & lContador.ToString
                    End If

                    Dim CServidor As String
                    CServidor = "S"



                    sParametros = "0000000000" & CServidor & sNroTarjeta & sNroCuenta.Trim & sPeriodoFinal.Trim & sPagina
                    inetputBuff = "      " + "R192" + sParametros

                    sTrama = ""
                    sTrama = obSendMirror.ExecuteTX(actions.ToString, inetputBuff, outpputBuff, TIME_OUT_SERVER, SERVIDOR_MIRROR_DESTINO, SERVIDOR_MIRROR_NODE, PUERTO, errorMsg)

                    outpputBuff = RTrim(outpputBuff)

                    If sTrama = "0" Then 'EXITO
                        If outpputBuff.Length > 0 Then sTrama = outpputBuff.Substring(8, outpputBuff.Length - 8)

                        'Intentos de call EECC
                        If lEstadoLlamadaEECC = 9 Then
                            lEstadoLlamadaEECC = 0 'Para que no vuelva a ingresar a esta logica

                            If Left(sTrama, 2) = "RU" Then
                                'Segunda LLamada segundo periodo
                                sPeriodoFinal = sPeriodo2.Trim
                                sParametros = "0000000000" & CServidor & sNroTarjeta.Trim & sNroCuenta.Trim & sPeriodoFinal.Trim & sPagina
                                inetputBuff = "      " + "R192" + sParametros

                                sTrama = ""
                                sTrama = obSendMirror.ExecuteTX(actions.ToString, inetputBuff, outpputBuff, TIME_OUT_SERVER, SERVIDOR_MIRROR_DESTINO, SERVIDOR_MIRROR_NODE, PUERTO, errorMsg)
                                outpputBuff = RTrim(outpputBuff)
                                If sTrama = "0" Then 'EXITO
                                    If outpputBuff.Length > 0 Then sTrama = outpputBuff.Substring(8, outpputBuff.Length - 8)

                                ElseIf sTrama = "-2" Then 'Ocurrio un error Recuperar el error
                                    'sTrama = "ERROR:" & errorMsg.Trim
                                    sTrama = ""
                                    Exit Do
                                Else
                                    'sTrama = "ERROR:Servicio no disponible."
                                    sTrama = ""
                                    Exit Do
                                End If


                            End If

                            If Left(sTrama, 2) = "RU" Then
                                'Tercera LLamada Tercer periodo
                                sPeriodoFinal = sPeriodo3.Trim
                                sParametros = "0000000000" & CServidor & sNroTarjeta.Trim & sNroCuenta.Trim & sPeriodoFinal.Trim & sPagina
                                inetputBuff = "      " + "R192" + sParametros

                                sTrama = ""
                                sTrama = obSendMirror.ExecuteTX(actions.ToString, inetputBuff, outpputBuff, TIME_OUT_SERVER, SERVIDOR_MIRROR_DESTINO, SERVIDOR_MIRROR_NODE, PUERTO, errorMsg)
                                outpputBuff = RTrim(outpputBuff)
                                If sTrama = "0" Then 'EXITO
                                    If outpputBuff.Length > 0 Then sTrama = outpputBuff.Substring(8, outpputBuff.Length - 8)

                                ElseIf sTrama = "-2" Then 'Ocurrio un error Recuperar el error
                                    'sTrama = "ERROR:" & errorMsg.Trim
                                    sTrama = ""
                                    Exit Do
                                Else
                                    'sTrama = "ERROR:Servicio no disponible."
                                    sTrama = ""
                                    Exit Do
                                End If

                            End If

                        End If 'fin de los intentos en periodos diferentes EECC

                    ElseIf sTrama = "-2" Then 'Ocurrio un error en la consulta de la transaccion
                        'sTrama = "ERROR:" & errorMsg.Trim
                        sTrama = ""
                        Exit Do
                    Else
                        'sTrama = "ERROR:Servicio no disponible."
                        sTrama = ""
                        Exit Do
                    End If


                    If Left(sTrama, 2) <> "RU" Then
                        'CONSTRUIR CADENA DE LA CABECERA Y PIE DE PAGINA
                        If lContador = 1 Then

                            sPeriodoFacturacion = Mid(sTrama, 293, 23)
                            sRespuesta = sPeriodoFacturacion.Trim

                            Return sRespuesta

                        End If 'FIN DE VALIDACION DEL CONTADOR

                    End If

                    If lContador > 20 Then
                        Exit Do
                    End If

                Loop Until Right(sTrama, 1) <> "C"


                'Evaluar Respuesta si es ERROR,  RU, AU (Respuesta correcta)
                If sTrama.Trim.Length > 0 Then
                    If Left(sTrama.Trim, 2) = "RU" Then 'Mensaje de Validacion al usuario. RU            000CUENTA BLOQUEADA
                        sMensajeErrorUsuario = Mid(sTrama.Trim, 18, Len(sTrama.Trim))
                        sRespuesta = "ERROR:NODATA-" & sMensajeErrorUsuario.Trim
                    End If
                Else 'Sino devuelve nada
                    sRespuesta = "ERROR:Servicio no disponible."
                End If
            Else
                'Mostrar Mensaje de Error
                sRespuesta = "ERROR:Parametros Incompletos"
            End If

        Catch ex As Exception
            'save log error

            sRespuesta = "ERROR:" & ex.Message.Trim

        End Try


        If sRespuesta.Substring(0, 5) = "ERROR" Then

            sRespuesta = "<EECC/Pend>"

        End If

        Return sRespuesta

    End Function


    Private _Detalle_Movimientos As New List(Of Movimiento)


    Sub Agregar_Movimiento(ByVal oMovimiento As Movimiento)
        _Detalle_Movimientos.Add(oMovimiento)
    End Sub

    Class Movimiento

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


    '<FIN>
    <WebMethod(Description:="Funcion para registrar los ingresos al Ripleymatico")> _
    Public Function getRegistrar_Ingreso_Ripleymatico(ByVal id_sucursal As String, _
                                                    ByVal id_kiosco As String, _
                                                    ByVal nro_documento As String, _
                                                    ByVal tip_documento As Integer, _
                                                    ByVal nro_tarjeta As String, _
                                                    ByVal tip_tarjeta As Integer) As String


        Dim msgError As String = "OK"

        Try
            If OpLog_Registro = "ON" Then
                oConexion = SQL_ConnectionOpen(Get_CadenaConexion(), sMensajeError_SQL)
                Dim cmd As New SqlCommand("SP_KIO_REGISTRAR_INGRESO_RM", oConexion)
                cmd.CommandType = CommandType.StoredProcedure

                cmd.Parameters.Add(New SqlParameter("@id_sucursal", SqlDbType.VarChar, 20))
                cmd.Parameters("@id_sucursal").Direction = ParameterDirection.Input
                cmd.Parameters("@id_sucursal").Value = id_sucursal

                cmd.Parameters.Add(New SqlParameter("@id_kiosco", SqlDbType.VarChar, 20))
                cmd.Parameters("@id_kiosco").Direction = ParameterDirection.Input
                cmd.Parameters("@id_kiosco").Value = id_kiosco

                cmd.Parameters.Add(New SqlParameter("@nro_documento", SqlDbType.VarChar, 20))
                cmd.Parameters("@nro_documento").Direction = ParameterDirection.Input
                cmd.Parameters("@nro_documento").Value = nro_documento

                cmd.Parameters.Add(New SqlParameter("@tip_documento", SqlDbType.Int))
                cmd.Parameters("@tip_documento").Direction = ParameterDirection.Input
                cmd.Parameters("@tip_documento").Value = tip_documento

                cmd.Parameters.Add(New SqlParameter("@nro_tarjeta", SqlDbType.VarChar, 16))
                cmd.Parameters("@nro_tarjeta").Direction = ParameterDirection.Input
                cmd.Parameters("@nro_tarjeta").Value = nro_tarjeta

                cmd.Parameters.Add(New SqlParameter("@tip_tarjeta", SqlDbType.Int))
                cmd.Parameters("@tip_tarjeta").Direction = ParameterDirection.Input
                cmd.Parameters("@tip_tarjeta").Value = tip_tarjeta

                cmd.ExecuteNonQuery()
                oConexion.Close()
                cmd.Dispose()
            End If
        Catch ex As Exception
            msgError = ex.Message
        End Try

        Return msgError


    End Function

    '<INI DHM
    '<WebMethod(Description:="Funcion para RSAT")> _
    'Public Function ActualizarLinea(ByVal objActualizarLinea As BEInIncrementoLinea) As BEOutIncrementoLinea

    '    Dim objBL As New BLIncrementoLinea()

    '    Return objBL.ActualizarLinea(objActualizarLinea)

    'End Function
    '<FIN DHM

    <WebMethod(Description:="Validar todas las tarjetas")> _
    Public Function VALIDAR_TARJETA(ByVal sNumeroTarjeta As String, ByVal sCodigoKiosko As String) As Object()
        Dim esRipleymatico As Boolean = True 'Esta ws es llamado sólo desde un Ripleymatico registrado
        Dim respuesta As Object() = New Object() {}
        ReDim respuesta(4)
        Dim consultar As Boolean = False
        'Asignamos valores por defecto a la respuesta
        respuesta(0) = "Exito"      'tipoRetorno
        respuesta(1) = ""           'sRetorno
        respuesta(2) = "Cliente"    'tipoConsulta
        respuesta(3) = Nothing      'opcionesTarjeta
        respuesta(4) = ""           'mensaje

        Dim CODIGO_EMP As String
        Dim NRO_TARJETA_CLIENTE As String
        Dim BINN_TARJETA As String
        Dim P_NroTarjeta_GIFCARD As String
        Dim BINN_TARJETA_GIFCARD As String
        Dim sucursal As Sucursal
        Dim msg As String = ""
        Dim opcionesTarjeta As OpcionesTarjeta
        Dim configuracionKiosko As ConfiguracionKiosko

        Try

            '--Llamar una funcion para que retorne la configuracion del kiosko de donde se hace la consulta
            configuracionKiosko = BNConfiguracionKiosco.Instancia.BuscarConfiguracionKioskoPorCodigoKiosco(sCodigoKiosko)
            If configuracionKiosko Is Nothing Then
                respuesta(0) = Constantes.BINN_VALIDADO_NO '--"sBinValidadoNO"
                respuesta(4) = "ErrorConfig"
                Return respuesta
            End If
            BINN_TARJETA = sNumeroTarjeta.Substring(configuracionKiosko.Posini_Bin1, configuracionKiosko.Longitud_Bin_Bin1)
            Dim SSBINN As String : Dim P_NroTarjeta As String
            SSBINN = sNumeroTarjeta.Substring(15, 2)
            sucursal = BuscarSucursal(sCodigoKiosko)


            '--Ahora por Binn vamos a consultar los datos a la base de datos.
            opcionesTarjeta = BuscarTipoTarjetaPorBinn(BINN_TARJETA)
            '--si es nulo buscar con una longitud de 5 dígitos
            If opcionesTarjeta Is Nothing Then
                BINN_TARJETA = sNumeroTarjeta.Substring(configuracionKiosko.Posini_Bin1, 5)
                opcionesTarjeta = BuscarTipoTarjetaPorBinn(BINN_TARJETA)
                If opcionesTarjeta Is Nothing Then
                    If SSBINN = "01" Then
                        If VerificarHorario(sucursal, "ValidaEmpleado", msg) = False Then
                            respuesta(0) = "HORARIO_NO"
                            respuesta(4) = msg
                            Return respuesta
                        End If
                        P_NroTarjeta = sNumeroTarjeta.Substring(15, 11)
                        CODIGO_EMP = P_NroTarjeta.Substring(5, 5)
                        NRO_TARJETA_CLIENTE = P_NroTarjeta

                        respuesta(0) = ""
                        respuesta(1) = CODIGO_EMP
                        respuesta(2) = "ValidaEmpleado"
                        Return respuesta
                    Else
                        respuesta(0) = Constantes.BINN_VALIDADO_NO '-- "sBinValidadoNO"
                        respuesta(4) = "ErrorBinn"
                        Return respuesta
                    End If
                Else
                    consultar = True
                End If
            Else
                consultar = True
            End If

            If consultar Then
                If VerificarHorario(sucursal, "", msg) = False Then
                    respuesta(0) = "HORARIO_NO"
                    respuesta(4) = msg
                    Return respuesta
                End If
                P_NroTarjeta = sNumeroTarjeta.Substring(configuracionKiosko.Posini_Bin1, configuracionKiosko.Longitud_Tarjeta_Bin1)
                'BINN_TARJETA = sNumeroTarjeta.Substring(configuracionKiosko.Posini_Bin1, configuracionKiosko.Longitud_Bin_Bin1)

                P_NroTarjeta_GIFCARD = sNumeroTarjeta.Substring(configuracionKiosko.Posini_Bin2, configuracionKiosko.Longitud_Tarjeta_Bin2)
                BINN_TARJETA_GIFCARD = sNumeroTarjeta.Substring(configuracionKiosko.Posini_Bin2, configuracionKiosko.Longitud_Bin_Bin2)

                Dim CUENTA_TARJETA As String
                If P_NroTarjeta.Length = configuracionKiosko.Longitud_Tarjeta_Bin1 Then
                    If opcionesTarjeta.TipoTarjetaPasada = "C" Then
                        CUENTA_TARJETA = "5" + P_NroTarjeta.Substring(8, 7)
                        '--ESTRAER DATO PARA VALIDAR SI ES TITULAR 9604110122818162
                        If opcionesTarjeta.MetodoWsCall = "VALIDAR_TARJETA_ABIERTA_RSAT" Then
                            opcionesTarjeta.EsTitular = "0"
                        Else
                            opcionesTarjeta.EsTitular = P_NroTarjeta.Substring(5, 1)
                        End If
                    Else
                        '--si son asociadas
                        CUENTA_TARJETA = P_NroTarjeta.Substring(7, 8)
                        opcionesTarjeta.EsTitular = "0" 'TODOS SON TITULARES
                    End If


                    If String.IsNullOrEmpty(CUENTA_TARJETA) = False Then
                        If IsNumeric(CUENTA_TARJETA) Then
                            '--NUMERO DE CUENTA OK
                            Dim CUENTA_TARJETA_VALIDAR As String = ""
                            CUENTA_TARJETA_VALIDAR = CUENTA_TARJETA
                            opcionesTarjeta.NroCuenta = CUENTA_TARJETA_VALIDAR
                            Dim DATA_MONITOR As String = ""
                            '--01=Lectura banda
                            DATA_MONITOR = CUENTA_TARJETA_VALIDAR + "|\u005Ct|" + P_NroTarjeta + "|\u005Ct|" + "01" + "|\u005Ct|" + configuracionKiosko.Codigo_Sucursal + "|\u005Ct|" + configuracionKiosko.Codigo_Kiosko + "|\u005Ct|" + "01"
                            ErrorLog("DATA_MONITOR=" & DATA_MONITOR)
                            '--Funcionando
                            respuesta(0) = "Exito"
                            'respuesta(1) = SSBINN + "!" + P_NroTarjeta + "!" + opcionesTarjeta.MetodoWsCall + "!" + DATA_MONITOR
                            respuesta(1) = LLamarMetodoParaValidarTarjeta(P_NroTarjeta, opcionesTarjeta.MetodoWsCall, DATA_MONITOR, esRipleymatico)
                            respuesta(2) = "Cliente"
                            respuesta(3) = opcionesTarjeta
                        Else 'SI NO ES NUMERICO
                            respuesta(0) = "Error"
                            respuesta(4) = "TarjetaNoValida"
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            oConexion = Nothing
            respuesta(0) = "Error"
            respuesta(4) = ex.Message
            Return respuesta
        End Try
        Return respuesta
    End Function

    <WebMethod(Description:="Obtiene la Configuracion de un Kiosko por su IP")> _
    <Xml.Serialization.XmlInclude(GetType(ConfiguracionKiosko))> _
    Public Function OBTENER_CONFIGURACION_KIOSKO(ByVal sIP As String) As Object()
        ErrorLog("OBTENER_CONFIGURACION_KIOSKO(" & sIP & ")")
        Dim rpta As Object() = New Object() {}
        ReDim rpta(0)
        Dim config As ConfiguracionKiosko
        config = BNConfiguracionKiosco.Instancia.BuscarConfiguracionKioskoPorIP(sIP)
        rpta(0) = config
        Return rpta

    End Function

    '<WebMethod(Description:="Obtiene la Configuracion de un Kiosko por su IP, sirve Internamente.")> _
    'Public Function Obtener_ConfiguracionKiosko_Por_IP(ByVal sIP As String) As ConfiguracionKiosko
    '    Dim config = New ConfiguracionKiosko
    '    Try
    '        'Realizar Conexion a la base de datos
    '        sMensajeError_SQL = ""

    '        oConexion = SQL_ConnectionOpen(Get_CadenaConexion(), sMensajeError_SQL)

    '        If sMensajeError_SQL <> "" Then
    '            config = Nothing
    '        Else

    '            If oConexion.State = ConnectionState.Open Then

    '                m_ssql = "USP_GET_CONFIGURACION_POR_IP"

    '                Dim cmd As SqlClient.SqlCommand = oConexion.CreateCommand
    '                Dim lTotal As Double = 0

    '                cmd.CommandTimeout = 600
    '                cmd.CommandType = CommandType.StoredProcedure
    '                cmd.CommandText = m_ssql

    '                cmd.Parameters.AddWithValue("@IP", sIP)

    '                Dim rd_cod As SqlClient.SqlDataReader = cmd.ExecuteReader
    '                If rd_cod.Read = True Then
    '                    config = New ConfiguracionKiosko

    '                    config.ID = rd_cod.GetInt32(rd_cod.GetOrdinal("ID"))
    '                    config.Nombre = rd_cod.GetString(rd_cod.GetOrdinal("NOMBRE"))
    '                    config.Server = rd_cod.GetString(rd_cod.GetOrdinal("SERVER"))
    '                    config.Server_Simulador = rd_cod.GetString(rd_cod.GetOrdinal("SERVER_SIMULADOR"))
    '                    config.Server_Com = rd_cod.GetString(rd_cod.GetOrdinal("SERVER_COM"))
    '                    config.Server_Print = rd_cod.GetString(rd_cod.GetOrdinal("SERVER_PRINT"))

    '                    config.Codigo_Kiosko = rd_cod.GetString(rd_cod.GetOrdinal("Codigo_Kiosko"))
    '                    config.Codigo_Sucursal = rd_cod.GetInt32(rd_cod.GetOrdinal("Codigo_Sucursal")).ToString()
    '                    config.ID_Kiosko = rd_cod.GetInt32(rd_cod.GetOrdinal("ID_Kiosko")).ToString()

    '                    config.Bin1 = rd_cod.GetString(rd_cod.GetOrdinal("BIN1"))
    '                    config.Longitud_Tarjeta_Bin1 = rd_cod.GetInt32(rd_cod.GetOrdinal("LONGITUD_TARJETA_BIN1"))
    '                    config.Posini_Bin1 = rd_cod.GetInt32(rd_cod.GetOrdinal("POSINI_BIN1"))
    '                    config.Longitud_Bin_Bin1 = rd_cod.GetInt32(rd_cod.GetOrdinal("LONGITUD_BIN_BIN1"))

    '                    config.Bin2 = rd_cod.GetString(rd_cod.GetOrdinal("BIN2"))
    '                    config.Longitud_Tarjeta_Bin2 = rd_cod.GetInt32(rd_cod.GetOrdinal("LONGITUD_TARJETA_BIN2"))
    '                    config.Posini_Bin2 = rd_cod.GetInt32(rd_cod.GetOrdinal("POSINI_BIN2"))
    '                    config.Longitud_Bin_Bin2 = rd_cod.GetInt32(rd_cod.GetOrdinal("LONGITUD_BIN_BIN2"))
    '                End If

    '                rd_cod.Close()
    '                rd_cod = Nothing
    '                cmd.Dispose()
    '                cmd = Nothing
    '                oConexion.Close()
    '                oConexion = Nothing
    '            Else
    '                config = Nothing
    '                oConexion = Nothing
    '            End If
    '        End If

    '    Catch ex As Exception
    '        config = Nothing
    '        oConexion = Nothing
    '    End Try

    '    Return config

    'End Function

    <WebMethod(Description:="Sirve Internamente para llamar a varios Metodos")> _
    Public Function LLamarMetodoParaValidarTarjeta(ByVal sCuenta As String, ByVal sMetodo As String, ByVal DataMonitor As String, ByVal esRipleymatico As Boolean) As String
        Dim sRespuesta As String = ""
        Dim oObjeto As Object = New Object
        Try
            DataMonitor = DataMonitor.Replace("|\u005Ct|", "|\t|")
            ErrorLog("LLamarMetodoParaValidarTarjeta(" & sCuenta & "," & sMetodo & "," & DataMonitor & "," & esRipleymatico & ")")

            oObjeto = CallByName(Me, sMetodo, CallType.Method, sCuenta, DataMonitor)

            If oObjeto.GetType() Is GetType(String) Then
                sRespuesta = oObjeto
            Else
                Dim rpta As Respuesta
                rpta = oObjeto
                If rpta.Estado = "ERROR" Then
                    sRespuesta = rpta.Estado + ":" + "MSGMSG" + ":" + rpta.Mensaje
                Else
                    sRespuesta = rpta.Cadena
                End If

            End If


        Catch ex As Exception
            '--save log error
        End Try

        Return sRespuesta
    End Function

    'Private Function BuscarConfiguracionKiosko(ByVal sCodigoKiosko As String) As ConfiguracionKiosko
    '    Dim config As New ConfiguracionKiosko
    '    Try
    '        'Realizar Conexion a la base de datos
    '        sMensajeError_SQL = ""

    '        Using oConexion As SqlConnection = SQL_ConnectionOpen(Get_CadenaConexion(), sMensajeError_SQL)

    '            If sMensajeError_SQL <> "" Then
    '                config = New ConfiguracionKiosko
    '                config.ID = 100
    '                config.Nombre = sCodigoKiosko
    '            Else

    '                If oConexion.State = ConnectionState.Open Then

    '                    m_ssql = "USP_GET_CONFIGURACION_POR_CODIGOKIOSKO"

    '                    Using cmd As SqlClient.SqlCommand = oConexion.CreateCommand


    '                        Dim lTotal As Double = 0

    '                        cmd.CommandTimeout = 600
    '                        cmd.CommandType = CommandType.StoredProcedure
    '                        cmd.CommandText = m_ssql

    '                        cmd.Parameters.AddWithValue("@COD_KIOSKO", sCodigoKiosko)
    '                        config = Nothing
    '                        Using rd_cod As SqlClient.SqlDataReader = cmd.ExecuteReader
    '                            If rd_cod.Read = True Then
    '                                config = New ConfiguracionKiosko

    '                                config.ID = rd_cod.GetInt32(rd_cod.GetOrdinal("ID"))
    '                                config.Nombre = rd_cod.GetString(rd_cod.GetOrdinal("NOMBRE"))
    '                                config.Server = rd_cod.GetString(rd_cod.GetOrdinal("SERVER"))
    '                                config.Server_Simulador = rd_cod.GetString(rd_cod.GetOrdinal("SERVER_SIMULADOR"))
    '                                config.Server_Com = rd_cod.GetString(rd_cod.GetOrdinal("SERVER_COM"))
    '                                config.Server_Print = rd_cod.GetString(rd_cod.GetOrdinal("SERVER_PRINT"))

    '                                config.Codigo_Kiosko = rd_cod.GetString(rd_cod.GetOrdinal("Codigo_Kiosko"))
    '                                config.Codigo_Sucursal = rd_cod.GetInt32(rd_cod.GetOrdinal("Codigo_Sucursal")).ToString()
    '                                config.ID_Kiosko = rd_cod.GetInt32(rd_cod.GetOrdinal("ID_Kiosko")).ToString()

    '                                config.Bin1 = rd_cod.GetString(rd_cod.GetOrdinal("BIN1"))
    '                                config.Longitud_Tarjeta_Bin1 = rd_cod.GetInt32(rd_cod.GetOrdinal("LONGITUD_TARJETA_BIN1"))
    '                                config.Posini_Bin1 = rd_cod.GetInt32(rd_cod.GetOrdinal("POSINI_BIN1"))
    '                                config.Longitud_Bin_Bin1 = rd_cod.GetInt32(rd_cod.GetOrdinal("LONGITUD_BIN_BIN1"))

    '                                config.Bin2 = rd_cod.GetString(rd_cod.GetOrdinal("BIN2"))
    '                                config.Longitud_Tarjeta_Bin2 = rd_cod.GetInt32(rd_cod.GetOrdinal("LONGITUD_TARJETA_BIN2"))
    '                                config.Posini_Bin2 = rd_cod.GetInt32(rd_cod.GetOrdinal("POSINI_BIN2"))
    '                                config.Longitud_Bin_Bin2 = rd_cod.GetInt32(rd_cod.GetOrdinal("LONGITUD_BIN_BIN2"))
    '                            End If
    '                        End Using
    '                    End Using
    '                Else
    '                    config = Nothing
    '                End If
    '            End If
    '        End Using
    '    Catch ex As Exception
    '        config = Nothing
    '    End Try

    '    Return config

    'End Function

    <WebMethod(Description:="busca tarjetas por bin")> _
    Public Function BuscarTipoTarjetaPorBinn(ByVal sBinnTarjeta As String) As OpcionesTarjeta
        Dim tipoTarjeta As New OpcionesTarjeta
        Try
            'Realizar Conexion a la base de datos
            sMensajeError_SQL = ""

            oConexion = SQL_ConnectionOpen(Get_CadenaConexion(), sMensajeError_SQL)

            If sMensajeError_SQL <> "" Then
                tipoTarjeta = Nothing
            Else

                If oConexion.State = ConnectionState.Open Then

                    m_ssql = "USP_GET_TIPOTARJETA_POR_BINN"

                    Dim cmd As SqlClient.SqlCommand = oConexion.CreateCommand
                    Dim lTotal As Double = 0

                    cmd.CommandTimeout = 600
                    cmd.CommandType = CommandType.StoredProcedure
                    cmd.CommandText = m_ssql

                    cmd.Parameters.AddWithValue("@BIN", sBinnTarjeta)

                    Dim rd_cod As SqlClient.SqlDataReader = cmd.ExecuteReader
                    tipoTarjeta = Nothing
                    If rd_cod.Read = True Then
                        tipoTarjeta = New OpcionesTarjeta

                        tipoTarjeta.Binn = rd_cod.GetString(rd_cod.GetOrdinal("BIN"))
                        tipoTarjeta.Tipo = rd_cod.GetString(rd_cod.GetOrdinal("TIPO"))
                        tipoTarjeta.NombreTarjeta = rd_cod.GetString(rd_cod.GetOrdinal("NOMBRE_TARJETA"))
                        tipoTarjeta.MetodoWsCall = rd_cod.GetString(rd_cod.GetOrdinal("METODO_WS_CALL"))
                        tipoTarjeta.TipoTarjetaPasada = rd_cod.GetString(rd_cod.GetOrdinal("TIPO_TARJETA_PASADA"))
                        ErrorLog("METODO_WS_CALL=" & tipoTarjeta.MetodoWsCall)
                    End If

                    rd_cod.Close()
                    rd_cod = Nothing
                    cmd.Dispose()
                    cmd = Nothing
                    oConexion.Close()
                    oConexion = Nothing
                Else
                    tipoTarjeta = Nothing
                    oConexion = Nothing
                End If
            End If

        Catch ex As Exception
            tipoTarjeta = Nothing
            oConexion = Nothing
        End Try

        Return tipoTarjeta

    End Function

    Private Function BuscarSucursal(ByVal sCodigoKiosko As String) As Sucursal
        Dim sucursal As New Sucursal
        sucursal = BNSucursal.Instancia.BuscarSucursalPorCodigo(sCodigoKiosko)
        Return sucursal
        'Dim sucursal As New Sucursal
        'Try
        '    'Realizar Conexion a la base de datos
        '    sMensajeError_SQL = ""

        '    Using oConexion As SqlConnection = SQL_ConnectionOpen(Get_CadenaConexion(), sMensajeError_SQL)
        '        If sMensajeError_SQL <> "" Then
        '            sucursal = Nothing
        '        Else
        '            If oConexion.State = ConnectionState.Open Then
        '                m_ssql = "USP_GET_SUCURSALPORKIOSKO"
        '                Using cmd As SqlClient.SqlCommand = oConexion.CreateCommand
        '                    Dim lTotal As Double = 0
        '                    cmd.CommandTimeout = 600
        '                    cmd.CommandType = CommandType.StoredProcedure
        '                    cmd.CommandText = m_ssql
        '                    cmd.Parameters.AddWithValue("@COD_KIOSKO", sCodigoKiosko)
        '                    Using rd_cod As SqlClient.SqlDataReader = cmd.ExecuteReader
        '                        sucursal = Nothing
        '                        If rd_cod.Read = True Then
        '                            sucursal = New Sucursal

        '                            sucursal.ID = rd_cod.GetInt32(rd_cod.GetOrdinal("ID"))
        '                            sucursal.Ciudad = rd_cod.GetString(rd_cod.GetOrdinal("CIUDAD"))
        '                            sucursal.Sucursal = rd_cod.GetString(rd_cod.GetOrdinal("SUCURSAL"))
        '                            sucursal.IdUbigeo = rd_cod.GetInt32(rd_cod.GetOrdinal("IdUbigeo"))
        '                            sucursal.EstTienda = rd_cod.GetInt32(rd_cod.GetOrdinal("esttienda"))
        '                            sucursal.Direccion = rd_cod.GetString(rd_cod.GetOrdinal("direccion"))
        '                            sucursal.Hini_com1 = rd_cod.GetString(rd_cod.GetOrdinal("hini_com1"))
        '                            sucursal.Hfin_com1 = rd_cod.GetString(rd_cod.GetOrdinal("hfin_com1"))
        '                            sucursal.Hini_com2 = rd_cod.GetString(rd_cod.GetOrdinal("hini_com2"))
        '                            sucursal.Hfin_com2 = rd_cod.GetString(rd_cod.GetOrdinal("hfin_com2"))
        '                            sucursal.Cod_Sucursal_Banco = rd_cod.GetString(rd_cod.GetOrdinal("cod_sucursal_banco"))
        '                            If rd_cod.IsDBNull(rd_cod.GetOrdinal("hini_cli")) Then
        '                                sucursal.Hini_Cli = ""
        '                            Else
        '                                sucursal.Hini_Cli = rd_cod.GetString(rd_cod.GetOrdinal("hini_cli"))
        '                            End If
        '                            If rd_cod.IsDBNull(rd_cod.GetOrdinal("hfin_cli")) Then
        '                                sucursal.Hfin_Cli = ""
        '                            Else
        '                                sucursal.Hfin_Cli = rd_cod.GetString(rd_cod.GetOrdinal("hfin_cli"))
        '                            End If
        '                        End If
        '                    End Using
        '                End Using
        '            Else
        '                sucursal = Nothing
        '            End If
        '        End If
        '    End Using
        'Catch ex As Exception
        '    sucursal = Nothing
        'End Try
        'Return sucursal

    End Function

    Private Function VerificarHorario(ByVal sucursal As Sucursal, ByVal tipo As String, ByRef msg As String) As Boolean
        Dim bool As Boolean = False

        Dim h As String = DateTime.Now.Hour.ToString()
        h = IIf(DateTime.Now.Hour < 10, "0" + h, h)
        Dim m As String = DateTime.Now.Minute.ToString()
        m = IIf(DateTime.Now.Minute < 10, "0" + m, m)
        msg = "Usted no puede realizar consultas en este momento, "
        h = String.Format("{0}:{1}", h, m)
        Dim horacero = "00:00"
        If tipo = "ValidaEmpleado" Then ' verificar horario de comisiones
            If (String.IsNullOrEmpty(sucursal.Hini_com1) Or String.IsNullOrEmpty(sucursal.Hini_com2) Or String.IsNullOrEmpty(sucursal.Hfin_com1) Or String.IsNullOrEmpty(sucursal.Hfin_com2) Or (sucursal.Hfin_com1 = horacero And sucursal.Hini_com1 = horacero)) Then
                msg += "no existe horario configurado para hacer consultas."
                Return False
            End If
            If (h < sucursal.Hini_com1 Or h > sucursal.Hfin_com2) Then
                bool = False
                msg += "puede hacer consultas de " + sucursal.Hini_com1 + " a " + sucursal.Hfin_com2 + ", Gracias."
            Else
                If (h > sucursal.Hfin_com1 And h < sucursal.Hini_com2) Then
                    bool = False
                    msg += "puede hacer consultas a partir de las " + sucursal.Hini_com2 + " hasta " + sucursal.Hfin_com2 + ", Gracias."
                Else
                    bool = True
                End If
            End If
        Else 'verificar horario cliente
            If (String.IsNullOrEmpty(sucursal.Hini_Cli) Or String.IsNullOrEmpty(sucursal.Hfin_Cli) Or sucursal.Hini_Cli = horacero) Then
                msg += "no existe horario configurado para hacer consultas."
                Return False
            End If
            If (h < sucursal.Hini_Cli Or h > sucursal.Hfin_Cli) Then
                bool = False
                msg += "puede hacer consultas de " + sucursal.Hini_Cli + " a " + sucursal.Hfin_Cli + ", Gracias."
            Else
                bool = True
            End If
        End If

        Return bool
    End Function

    Private Function ValidarMensajeAMostrarEnRipleymatico(ByVal bloqueo As Boolean, ByVal vg_RSAT_Sitarj As String, ByVal vg_RSAT_FecBaj As String, ByVal code As String) As String
        Dim rpta As String = ""
        Dim codigo As String = ""
        Dim situacion As String = ""
        If String.IsNullOrEmpty(code) Then
            If bloqueo Then
                codigo = "Bloqueo"
            Else
                If vg_RSAT_FecBaj <> "0001-01-01" Then
                    codigo = "Baja"
                Else
                    codigo = vg_RSAT_Sitarj
                End If
            End If
        Else
            codigo = code
        End If

        Try
            'Realizar Conexion a la base de datos
            sMensajeError_SQL = ""
            Using oConexion As SqlConnection = SQL_ConnectionOpen(Get_CadenaConexion(), sMensajeError_SQL)
                If sMensajeError_SQL <> "" Then
                    rpta = Constantes.MSG_PROBLEMA_CONECCION
                Else
                    If oConexion.State = ConnectionState.Open Then
                        m_ssql = "USP_GET_MENSAJE_BLOQUEO_TARJETA"
                        Using cmd As SqlClient.SqlCommand = oConexion.CreateCommand
                            Dim lTotal As Double = 0
                            cmd.CommandTimeout = 600
                            cmd.CommandType = CommandType.StoredProcedure
                            cmd.CommandText = m_ssql
                            cmd.Parameters.AddWithValue("@codigo", codigo)

                            rpta = cmd.ExecuteScalar
                            If String.IsNullOrEmpty(rpta) Then
                                rpta = Constantes.MSG_ERROR_CODIGO + "No se encuentra el código de bloqueo."
                            End If
                        End Using
                    Else
                        rpta = Constantes.MSG_ERROR_CODIGO
                    End If
                End If
            End Using
        Catch ex As Exception
            rpta = Constantes.MSG_ERROR_CODIGO + "!"
        End Try
        Return rpta
    End Function

    'GUARDAR LOG DE CONSULTAS DE COMISIONES EN EL KIOSCO RIPLEYMATICO
    <WebMethod(Description:="GUARDAR LOG DE CONSULTAS DE COMISIONES DEL KIOSCO.")> _
    Public Function SAVE_LOG_CONSULTAS_COMISIONES( _
     ByVal ID_SUCURSAL As String, ByVal NOMBRE_KIOSCO As String, ByVal OPCION As String, ByVal CODIGO_EMPLEADO As String) As String

        Dim sEstado As String
        Dim sInsert As String

        If OpLog_Consulta = "ON" Then
            Try
                sEstado = "NO"
                Dim oCN As SqlConnection
                Dim oCMD As SqlCommand

                'Conexion a la base de datos
                oCN = SQL_ConnectionOpen(Get_CadenaConexion(), sMensajeError_SQL)

                'Armar cadena Insert
                sInsert = "INSERT INTO KIO_CONSULTAS_COMISIONES (ID_SUCURSAL,NOMBRE_KIOSCO,OPCION,CODIGO_EMPLEADO)"
                sInsert = sInsert & " VALUES(" & ID_SUCURSAL.Trim & ",'" & NOMBRE_KIOSCO.Trim & "'," & OPCION.Trim & "," & CODIGO_EMPLEADO.Trim & ")"

                'Ejecutar el Stored Procedure
                oCMD = New SqlClient.SqlCommand
                oCMD.CommandText = sInsert
                oCMD.CommandType = CommandType.Text
                oCMD.Connection = oCN

                'Ejecutar oCMD
                oCMD.ExecuteNonQuery()
                'operacion con exito
                sEstado = "SI"

                oCMD.Dispose()
                oCN.Close()
                oCMD = Nothing
                oCN = Nothing
            Catch ex As Exception
                sEstado = "NO"
            End Try
        Else
            sEstado = "NO"
        End If
        Return sEstado
    End Function

    <WebMethod(Description:="return super efectivo")> _
    Public Function EjemploSuperEfectivo(ByVal sBinnTarjeta As String) As SuperEfectivo
        Dim sef As New SuperEfectivo

        Return sef

    End Function

    <WebMethod(Description:="return super efectivo")> _
    Public Function EjemploSuperEfectivoParametros(ByVal sBinnTarjeta As String) As SuperEfectivo.Parametros
        Dim sef As New SuperEfectivo.Parametros

        Return sef

    End Function

#Region "PopUp Invasivo"

    <WebMethod(Description:="ObtenerPopUpInvasivo")> _
    <Xml.Serialization.XmlInclude(GetType(Response))> _
    <Xml.Serialization.XmlInclude(GetType(PopUpInvasivo))> _
    Public Function ObtenerPopUpInvasivo(ByVal tipoOperacion As String _
                                         , ByVal tipoDocumento As String _
                                         , ByVal numeroDocumento As String _
                                         , ByVal numeroTarjeta As String) As Response
        Dim loadLPDP As LoadLPDP = Nothing
        Dim popUpInvasivo As New PopUpInvasivo
        Dim response As New Response
        response.Success = True
        response.Warning = False
        response.Message = String.Empty
        Dim nombreServicio As String = String.Empty
        Dim mensajeServicio As String = String.Empty
        Dim objMQ As New MQ
        Dim fechaActual As String = Date.Now.ToShortDateString()
        Dim fechaResultante As String = ""
        Dim dia As String = ""
        Dim mes As String = ""
        Dim anio As String = ""
        Dim fechaCaducidad As String = ""
        Dim terminosCondicionesLPDP As New TerminosCondicionesLPDP

        Try
            ErrorLog("ObtenerPopUpInvasivo")
            If fechaActual.Trim.Length > 0 Then
                dia = Left(fechaActual.Trim, 2)
                mes = Mid(fechaActual.Trim, 4, 2)
                anio = Right(fechaActual.Trim, 4)

                fechaResultante = anio.Trim() & "/" & mes.Trim() & "/" & dia.Trim()

            End If

            terminosCondicionesLPDP = BNConsultaLPDP.Instancia.GetTerminosCondicionesLPDP()            

            If fechaResultante <= terminosCondicionesLPDP.Caducidad Then

                loadLPDP = BNConsultaLPDP.Instancia.GetLoadLPDP(numeroDocumento)

                If loadLPDP Is Nothing Then
                    Select Case tipoOperacion
                        Case Constantes.OperacionTarjeta

                            nombreServicio = ReadAppConfig("SFSCAN_TARJETA_CLASICA_RSAT")
                            mensajeServicio = "0000000000" & nombreServicio & "                                       000000065PE00010000080027RSAT     IVR00           00" + numeroTarjeta + "                          0"

                            If numeroTarjeta.Substring(0, 6) = "542020" Or numeroTarjeta.Substring(0, 6) = "525474" Or _
                                numeroTarjeta.Substring(0, 6) = "450034" Or numeroTarjeta.Substring(0, 6) = "450035" Then
                                mensajeServicio = "0000000000" & nombreServicio & "                                       000000065PE00010000080027RSAT     IVR00           00" + numeroTarjeta + "              9999999999990"
                            End If

                            objMQ.Service = nombreServicio
                            objMQ.Message = mensajeServicio
                            objMQ.Execute()

                            If objMQ.ReasonMd <> 0 Then
                                response.Warning = True
                                response.Message = "El servicio de Validación de tarjeta devolvió un valor diferente de cero"
                            ElseIf objMQ.ReasonApp = 0 Then
                                popUpInvasivo.Codigo = Mid(objMQ.Response, 261, 2)
                                response.Data = popUpInvasivo
                            End If

                        Case Constantes.OperacionDNI

                            If tipoDocumento = "1" Then
                                tipoDocumento = "C"
                            End If

                            nombreServicio = ReadAppConfig("SFSCAN_DNI_CLASICAABIERTA_RSAT")
                            mensajeServicio = "0000000000" & nombreServicio & "                                       000000065PE00010000020027X8  0       " & tipoDocumento & numeroDocumento & "    1                                           00                                                                    00                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                    00          00  000000000000                      00                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                            "

                            objMQ.Service = nombreServicio
                            objMQ.Message = mensajeServicio
                            objMQ.Execute()

                            If objMQ.ReasonMd <> 0 Then
                                response.Warning = True
                                response.Message = "El servicio de Validación de DNI devolvió un valor diferente de cero"
                            ElseIf objMQ.ReasonApp = 0 Then
                                popUpInvasivo.Codigo = Mid(objMQ.Response, 261, 2)
                                response.Data = popUpInvasivo
                            End If

                        Case Else
                            response.Warning = True
                            response.Message = "No se ha ingresado ningún tipo de operación permitida (1 o 2). Se ingresó " & tipoOperacion

                    End Select
                Else
                    response.Warning = True
                    response.Message = "Cliente " & numeroDocumento & " ya ha aceptado TC"
                End If

            Else
                response.Warning = True
                response.Message = "La fecha Actual sobrepasa la fecha de caducidad = " & terminosCondicionesLPDP.Caducidad
            End If

        Catch ex As Exception
            response.Success = False
            response.Message = "En estos momentos no podemos atenderle. Por favor inténtelo más tarde"
            ErrorLog("ObtenerPopUpInvasivo Error =" & ex.Message)
        End Try
        ErrorLog("ObtenerPopUpInvasivo.message = " & response.Message)
        ErrorLog("Fin ObtenerPopUpInvasivo")
        Return response
    End Function

    <WebMethod(Description:="Guarda el log de Aceptación de LPDP")> _
    Public Function SAVE_LOG_CONSULTAS_LPDP(ByVal codigoSucursalBanco As Double, _
                                        ByVal codigoKiosko As String, _
                                        ByVal tipoDocumento As Integer, _
                                        ByVal nroDocumento As String, _
                                        ByVal opcion As String, _
                                        ByVal numeroTarjeta As String, _
                                        ByVal nroCuenta As String, _
                                        ByVal idSucursal As Integer) As Boolean

        Dim estado As Boolean
        Dim oConsultaLPDP As New ConsultaLPDP
        Try
            ErrorLog("Entro al método SAVE_LOG_CONSULTAS_LPDP")

            oConsultaLPDP.Cod_Sucursal_Ban = Trim(codigoSucursalBanco)
            oConsultaLPDP.Nombre_Kiosko = Trim(codigoKiosko)
            oConsultaLPDP.Tipo_Doc = Trim(tipoDocumento)
            oConsultaLPDP.Nro_Doc = Trim(nroDocumento)
            oConsultaLPDP.Opcion = Trim(opcion)
            oConsultaLPDP.Numero_Tarjeta = Trim(numeroTarjeta)
            oConsultaLPDP.Numero_Cuenta = Trim(nroCuenta)
            oConsultaLPDP.Id_Sucursal = idSucursal

            estado = BNConsultaLPDP.Instancia.InsertConsultaLPDP(oConsultaLPDP)
        Catch ex As Exception
            ErrorLog("Error InsertConsultaLPDP = " & ex.Message)
            estado = False
        End Try


        Return estado
    End Function

    <WebMethod(Description:="Obtener Términos y Condiciones LPDP")> _
    <Xml.Serialization.XmlInclude(GetType(Response))> _
    <Xml.Serialization.XmlInclude(GetType(TerminosCondicionesLPDP))> _
    Public Function GetTerminosCondicionesLPDP() As Response
        Dim respuesta As New Response
        respuesta.Success = True
        respuesta.Warning = False
        Dim terminos As TerminosCondicionesLPDP = Nothing

        Try
            terminos = BNConsultaLPDP.Instancia.GetTerminosCondicionesLPDP()
            If terminos Is Nothing Then
                respuesta.Warning = True
                respuesta.Message = "No existen términos y condiciones"
            Else
                respuesta.Data = terminos
            End If

        Catch ex As Exception
            respuesta.Success = False
            respuesta.Message = "En estos momentos no podemos atenderle. Por favor inténtelo más tarde"
            ErrorLog("GetTerminosCondicionesLPDP Error =" & ex.Message)
        End Try
        Return respuesta
    End Function

#End Region

#Region "Métodos de Cambio de Producto"

    <WebMethod(Description:="Guarda el log de Cambio Producto")> _
    Public Function SAVE_LOG_CONSULTAS_CAMBIOPRODUCTO(ByVal codigoSucursalBanco As Double, _
                                        ByVal codigoKiosko As String, _
                                        ByVal tipoDocumento As Integer, _
                                        ByVal nroDocumento As String, _
                                        ByVal opcion As String, _
                                        ByVal numeroTarjeta As String, _
                                        ByVal tipoTarjetaInicial As String, _
                                        ByVal tipoTarjetaFinal As String, _
                                        ByVal fechaOfertaInicial As String, _
                                        ByVal fechaOfertaFinal As String, _
                                        ByVal nroCuenta As String, _
                                        ByVal idSucursal As Integer) As Boolean

        Dim estado As Boolean
        Dim oConsultaCambioProducto As New ConsultaCambioProducto
        Try
            ErrorLog("Entro al método SAVE_LOG_CONSULTAS_CAMBIOPRODUCTO")
            Dim dFechaIni As Date = New Date(fechaOfertaInicial.Substring(6, 4), fechaOfertaInicial.Substring(3, 2), fechaOfertaInicial.Substring(0, 2))
            Dim dFechaFin As Date = New Date(fechaOfertaFinal.Substring(6, 4), fechaOfertaFinal.Substring(3, 2), fechaOfertaFinal.Substring(0, 2))

            oConsultaCambioProducto.Cod_Sucursal_Ban = Trim(codigoSucursalBanco)
            oConsultaCambioProducto.Nombre_Kiosko = Trim(codigoKiosko)
            oConsultaCambioProducto.Tipo_Doc = Trim(tipoDocumento)
            oConsultaCambioProducto.Nro_Doc = Trim(nroDocumento)
            oConsultaCambioProducto.Opcion = Trim(opcion)
            oConsultaCambioProducto.Numero_Tarjeta = Trim(numeroTarjeta)
            oConsultaCambioProducto.Tipo_Tarjeta_Inicial = Trim(tipoTarjetaInicial)
            oConsultaCambioProducto.Tipo_Tarjeta_Final = Trim(tipoTarjetaFinal)
            oConsultaCambioProducto.Fecha_Inicial_Oferta = dFechaIni
            oConsultaCambioProducto.Fecha_Final_Oferta = dFechaFin
            oConsultaCambioProducto.Numero_Cuenta = Trim(nroCuenta)
            oConsultaCambioProducto.Id_Sucursal = idSucursal

            estado = BNConsultaCambioProducto.Instancia.InsertConsultaCambioProducto(oConsultaCambioProducto)
        Catch ex As Exception
            estado = False
        End Try


        Return estado
    End Function

    <WebMethod(Description:="Consulta Oferta Cambio Producto en Plataforma Comercial")> _
    <Xml.Serialization.XmlInclude(GetType(Response))> _
    <Xml.Serialization.XmlInclude(GetType(OfertaCambioProducto))> _
    Public Function ObtenerOfertaCambioProductoComercial(ByVal contrato As String) As Response
        Dim respuesta As New Response
        respuesta.Success = False
        respuesta.Warning = False
        respuesta.Message = String.Empty
        Dim oferta As New OfertaCambioProducto

        Try

            oferta = BNOfertaCambioProducto.Instancia.ObtenerOferta(contrato)


            If oferta.ContratoTarjeta <> String.Empty Then
                Dim validaTarjeta As String = String.Empty
                Dim productoOrigenTarjeta As String
                Dim productoDestinoTarjeta As String
                Dim descripcionProductoOrigenTarjeta As String
                Dim descripcionProductoDestinoTarjeta As String

                Dim validaSEF As String = String.Empty
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

                    oferta.ProductoOrigenTarjeta = productoOrigenTarjeta
                    oferta.ProductoDestinoTarjeta = productoDestinoTarjeta
                    oferta.DescripcionProductoOrigenTarjeta = descripcionProductoOrigenTarjeta
                    oferta.DescripcionProductoDestinoTarjeta = descripcionProductoDestinoTarjeta

                    oferta.CarpetaDestino = AsignarCarpetaDestino(oferta.ProductoDestinoTarjeta)


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


                If validaTarjeta <> Constantes.TarjetaSEFValido Then
                    respuesta.Warning = True
                    respuesta.Message &= Constantes.ErrorCodigoTarjetaInvalido
                End If

                If validaSEF <> Constantes.TarjetaSEFValido Then
                    respuesta.Warning = True
                    respuesta.Message &= Constantes.ErrorCodigoSEFInvalido
                End If

                If oferta.OfertaVigente <> Constantes.OfertaVigente Then
                    respuesta.Warning = True
                    respuesta.Message &= Constantes.ErrorPromocionNoVigente
                End If
                ErrorLog(respuesta.Message)

            Else
                respuesta.Warning = True
                respuesta.Message = Constantes.SinPromocion
                ErrorLog(respuesta.Message)
            End If

            respuesta.Data = New OfertaCambioProducto
            respuesta.Data = oferta
            respuesta.Success = True
        Catch ex As Exception
            respuesta.Message = Constantes.ErrorServidor
            ErrorLog(String.Format(respuesta.Message & " Excepcion = {0}", ex.Message))
        End Try

        Return respuesta
    End Function

    <WebMethod(Description:="Consulta Oferta Cambio Producto.")> _
    <Xml.Serialization.XmlInclude(GetType(Response))> _
    <Xml.Serialization.XmlInclude(GetType(OfertaCambioProducto))> _
    Public Function ObtenerOfertaCambioProductoValidado(ByVal contrato As String, ByVal codUsuario As String, ByVal ip As String, ByVal codOficina As String) As Response
        Dim respuesta As New Response
        respuesta.Success = False
        respuesta.Warning = False
        respuesta.Message = String.Empty
        Dim oferta As New OfertaCambioProducto

        Try

            oferta = BNOfertaCambioProducto.Instancia.ObtenerOferta(contrato)
            oferta.CodSucursalBanco = codOficina

            If oferta.ContratoTarjeta <> String.Empty Then
                Dim validaTarjeta As String = String.Empty
                Dim productoOrigenTarjeta As String
                Dim productoDestinoTarjeta As String
                Dim descripcionProductoOrigenTarjeta As String
                Dim descripcionProductoDestinoTarjeta As String

                Dim validaSEF As String = String.Empty
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

                    oferta.ProductoOrigenTarjeta = productoOrigenTarjeta
                    oferta.ProductoDestinoTarjeta = productoDestinoTarjeta
                    oferta.DescripcionProductoOrigenTarjeta = descripcionProductoOrigenTarjeta
                    oferta.DescripcionProductoDestinoTarjeta = descripcionProductoDestinoTarjeta

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

                If validaTarjeta = Constantes.TarjetaSEFValido And validaSEF = Constantes.TarjetaSEFValido And oferta.OfertaVigente = Constantes.OfertaVigente Then
                    Dim tipoDocumento As String = String.Empty
                    Dim tipoContrato As String = String.Empty

                    Select Case oferta.TipoDocumento
                        Case 1
                            tipoDocumento = Constantes.TipoDocumentoDNI
                        Case Else
                            tipoDocumento = Convert.ToChar(oferta.TipoDocumento)
                    End Select
                    'Validar Contrato Tarjeta
                    tipoContrato = Constantes.ServicioVTAAUTO5001TIPOCONTRATOTARJETA

                    Dim parameterIn As New VTAAUTO5001ParameterIN
                    Dim parameterOut As New VTAAUTO5001ParameterOUT
                    parameterIn.TipoTransaccion = Constantes.TipoConsulta
                    parameterIn.CodEntidadTarjeta = oferta.ContratoTarjeta.Substring(0, 4)
                    parameterIn.CentroAltaTarjeta = oferta.ContratoTarjeta.Substring(4, 4)
                    parameterIn.CuentaTarjeta = oferta.ContratoTarjeta.Substring(8, 12)
                    'Falta saber los codigos de producto y subproducto para tarjeta.
                    parameterIn.CodCambioTarjeta = oferta.CodCambioTarjeta

                    parameterIn.CodEntidadSEF = oferta.ContratoSEF.Substring(0, 4)
                    parameterIn.CentroAltaSEF = oferta.ContratoSEF.Substring(4, 4)
                    parameterIn.CuentaSEF = oferta.ContratoSEF.Substring(8, 12)
                    'Falta saber los codigos de producto y subproducto para SEF.
                    parameterIn.CodCambioSEF = oferta.CodCambioSEF

                    parameterIn.TipoDocumento = tipoDocumento
                    parameterIn.NumeroDocumento = oferta.NumeroDocumento
                    parameterIn.SystemSource = Constantes.TipoSistema
                    parameterIn.ModuleId = Constantes.ModuleId
                    parameterIn.TipoTerminal = Constantes.STipoTerminal
                    parameterIn.CodUsuario = String.Empty
                    parameterIn.MotImpresion = codOficina

                    parameterIn.ModoEstampacion = String.Empty
                    parameterIn.SEnvioOficina = String.Empty
                    parameterIn.CodDestino = String.Empty
                    parameterIn.EnvEmailPers = String.Empty
                    parameterIn.EnvEmailLab = String.Empty
                    parameterIn.Envtarjeta = String.Empty

                    parameterIn.CodUsrApr = codUsuario
                    parameterIn.NumeroIp = ip

                    parameterOut = BNOfertaCambioProducto.Instancia.ValidarContrato(parameterIn, codUsuario, codOficina)

                    If parameterOut.Paso1 <> Constantes.ServicioVTAAUTO5000PASO01 Or parameterOut.Paso2 <> Constantes.ServicioVTAAUTO5000PASO02 Or parameterOut.CodRpta <> Constantes.ServicioVTAAUTO5000RPTA Then
                        respuesta.Warning = True
                        respuesta.Message = parameterOut.DesRpta
                        ErrorLog(respuesta.Message)
                        InsertarLogIncidenciaCambioProducto(contrato, codOficina, codUsuario, Constantes.SistemaRM, Constantes.ErrorRM, "01", respuesta.Message)
                    End If

                Else
                    If validaTarjeta <> Constantes.TarjetaSEFValido Then
                        respuesta.Warning = True
                        respuesta.Message &= Constantes.ErrorCodigoTarjetaInvalido
                    End If

                    If validaSEF <> Constantes.TarjetaSEFValido Then
                        respuesta.Warning = True
                        respuesta.Message &= Constantes.ErrorCodigoSEFInvalido
                    End If

                    If oferta.OfertaVigente <> Constantes.OfertaVigente Then
                        respuesta.Warning = True
                        respuesta.Message &= Constantes.ErrorPromocionNoVigente
                    End If
                    ErrorLog(respuesta.Message)
                    InsertarLogIncidenciaCambioProducto(contrato, codOficina, codUsuario, Constantes.SistemaRM, Constantes.ErrorRM, "01", respuesta.Message)
                End If

            Else
                respuesta.Warning = True
                respuesta.Message = Constantes.SinPromocion
                ErrorLog(respuesta.Message)
            End If

            respuesta.Data = New OfertaCambioProducto
            respuesta.Data = oferta
            respuesta.Success = True
        Catch ex As Exception
            respuesta.Message = Constantes.ErrorServidor
            ErrorLog(String.Format(respuesta.Message & " Excepcion = {0}", ex.Message))
        End Try

        Return respuesta
    End Function

    Private Sub InsertarLogIncidenciaCambioProducto(ByVal contrato As String, _
                                                    ByVal codOficina As String, _
                                                    ByVal codUsuario As String, _
                                                    ByVal sistema As String, _
                                                    ByVal subSistema As String, _
                                                    ByVal codRespuesta As String, _
                                                    ByVal desRespuesta As String)
        Try
            ErrorLog(desRespuesta)
            BNOfertaCambioProducto.Instancia.InsertarLogIncidenciaCambioProducto(contrato, codOficina, codUsuario, sistema, subSistema, codRespuesta, desRespuesta)
        Catch ex As Exception
            ErrorLog("Exception InsertarLogIncidenciaCambioProducto = " & ex.Message)
        End Try
    End Sub

    <WebMethod(Description:="Actualizar Cambio Producto en Plataforma Comercial")> _
    Public Function ActualizarCambioProducto(ByVal contrato As String, _
                                        ByVal codVendedor As Integer, _
                                        ByVal codSucursal As Integer, _
                                        ByVal numeroCaja As String, _
                                        ByVal numeroTicket As Integer) As Response

        Dim respuesta As New Response
        respuesta.Data = "0"
        Try
            ErrorLog("ActualizarCambioProducto " & contrato)
            respuesta.Data = BNOfertaCambioProducto.Instancia.ActualizarCambioProducto(contrato, codVendedor, codSucursal, numeroCaja, numeroTicket)
            respuesta.Success = True
            If respuesta.Data <> "1" Then
                respuesta.Warning = True
                respuesta.Message = "Ocurrio un error en ActualizarCambioProducto, devolvió = " & respuesta.Data
            Else
                respuesta.Warning = False
            End If

        Catch ex As Exception
            ErrorLog("Exception ActualizarCambioProducto = " & ex.Message)
            respuesta.Data = "0"
            respuesta.Success = False
            respuesta.Warning = True
            respuesta.Message = "En estos momentos no podemos atenderlo, inténtelo más tarde"
        End Try

        Return respuesta
    End Function

    Private Function AsignarCarpetaDestino(ByVal codigoTarjeta As String) As String
        Dim carpetaDestino As String = String.Empty
        Select Case codigoTarjeta
            Case "01"
                carpetaDestino = Constantes.PLAMC
            Case "02"
                carpetaDestino = Constantes.CLASICA
            Case "03"
                carpetaDestino = Constantes.GOLDMC
            Case "04"
                carpetaDestino = Constantes.PLAVI
            Case "05"
                carpetaDestino = Constantes.SILVERMC
            Case "06"
                carpetaDestino = Constantes.GOLDMCR
            Case "07"
                carpetaDestino = Constantes.PLAVISAR
            Case "08"
                carpetaDestino = Constantes.SILVERMCR
            Case "09"
                carpetaDestino = Constantes.SILVERVISA
            Case "10"
                carpetaDestino = Constantes.SILVERVISAR
            Case Else
                carpetaDestino = String.Empty
        End Select
        Return carpetaDestino
    End Function

#End Region

#Region "Métodos de Validación de Login Service"


    <WebMethod(Description:="Obtener Email dado DNI en FISA")> _
    Public Function ObtenerEmailFISA(ByVal tipoDocumento As String, ByVal numeroDocumento As String) As String
        Dim respuesta As String = String.Empty
        Try
            ErrorLog("Entro al método ObtenerEmailFISA(" & tipoDocumento & ", " & numeroDocumento & ",)")
            Select Case tipoDocumento
                Case 1
                    numeroDocumento = "C-" & numeroDocumento
                Case 2
                    numeroDocumento = "P-" & numeroDocumento
                Case Else
                    numeroDocumento = "C-" & numeroDocumento
            End Select


            ErrorLog("Entro al método ObtenerEmailFISA(" & numeroDocumento & ")")
            respuesta = BNBRPFISA.Instancia.ObtenerEmailFISA(numeroDocumento)
            ErrorLog("Salidó al método ObtenerEmailFISA(" & numeroDocumento & ")")
        Catch ex As Exception
            respuesta = "Hubo error en llamada a método ObtenerEmailFISA " & ex.Message
        End Try

        Return respuesta
    End Function

    <WebMethod(Description:="Validación de Login Service")> _
    Public Function ValidarLoginService(ByVal canal As String, ByVal tipoDocumento As String, ByVal numeroDocumento As String, ByVal clave As String, ByVal email As String, _
                                        ByVal codigoSucursalBanco As String, _
                                        ByVal codigoKiosko As String, _
                                        ByVal idSucursal As Integer, _
                                        ByVal numeroTarjeta As String, _
                                        ByVal nombreCliente As String) As Response
        Dim respuesta As New Response
        Dim mensaje As String = String.Empty
        Dim detalle As String = String.Empty
        Dim label As String = String.Empty
        Dim nivel As String = String.Empty
        Dim resultado As String = String.Empty
        Dim resultadoCorreo As String = String.Empty
        Dim listaEmail As New List(Of String)
        Dim claveValida As Integer = 0

        Dim loginService As New LoginService.loginService
        loginService.Url = ReadAppConfig("LoginService.loginService")
        ErrorLog("ValidarLoginUrl=" & loginService.Url)

        Try
            ErrorLog("ValidarLoginService")

            If BNConsultaClave6.Instancia.ValidarClave(clave) = 1 Then
                ErrorLog("canal=" & canal)
                ErrorLog("tipoDocumento=" & tipoDocumento)
                ErrorLog("numeroDocumento=" & numeroDocumento)
                ErrorLog("clave=" & clave)
                ErrorLog("email=" & email)

                resultado = loginService.executeTransfer(canal, tipoDocumento, numeroDocumento, clave, email, mensaje, detalle, label, nivel)

                ErrorLog("resultado=" & resultado)
                ErrorLog("mensaje=" & mensaje)
                ErrorLog("detalle=" & detalle)
                ErrorLog("label=" & label)
                ErrorLog("nivel=" & nivel)
                respuesta.Message = mensaje
                respuesta.Success = (resultado = "1")

                Try
                    Dim tipoDocumentoInteger As Integer = 1
                    Select Case tipoDocumento
                        Case "C"
                            tipoDocumentoInteger = 1
                        Case "P"
                            tipoDocumentoInteger = 2
                        Case Else
                            tipoDocumentoInteger = 1
                    End Select

                    SAVE_LOG_CLAVE6(Trim(codigoSucursalBanco), Trim(codigoKiosko), idSucursal, tipoDocumentoInteger, Trim(numeroDocumento), Trim(numeroTarjeta), Trim(resultado), Trim(email), Trim(mensaje))
                Catch ex As Exception
                    ErrorLog("No se pudo guardar en Log CLAVE 06= " & ex.Message)
                End Try

                If respuesta.Success Then
                    ErrorLog("Si genero la clave")

                    Dim envioData As New EnvioData
                    Dim threadEnvioCorreo As New Thread(New ThreadStart(AddressOf envioData.EnviarEmailMasivoFisa))
                    listaEmail.Add(email)
                    envioData.Addresses = listaEmail
                    envioData.Subject = "Generación de Clave Internet"

                    Dim fecha As Date = DateTime.Now
                    Dim fechaActual As String = Format(fecha, "g")

                    envioData.Body = "<p style='text-align:right;'>" & fechaActual & "</p><p style='text-align:left;'>Estimado(a) " & nombreCliente & "</p>" & Constantes.MensajeEmailClave6
                    Dim rutaTemporal As String = String.Format("{0}{1}{2}", System.Web.HttpContext.Current.Request.PhysicalApplicationPath, "Archivos\", "TYCHB.pdf")
                    ErrorLog("rutaTemporal de TYC= " & rutaTemporal)
                    Dim listaArchivo As New List(Of String)
                    listaArchivo.Add(rutaTemporal)
                    envioData.FileAttach = listaArchivo
                    envioData.NumeroDocumento = numeroDocumento
                    envioData.TipoDocumento = tipoDocumento

                    threadEnvioCorreo.Start()
                End If
            Else
                respuesta.Success = False
                respuesta.Message = "La clave ingresada no es segura, por favor ingrese otra clave"
                ErrorLog(respuesta.Message)
            End If
        Catch ex As Exception
            respuesta.Success = False
            respuesta.Message = "En estos momentos no podemos atenderle. Por favor inténtelo más tarde"
            ErrorLog("ValidarLoginService Error =" & ex.Message)
        End Try
        ErrorLog("Fin ValidarLoginService")
        Return respuesta
    End Function

    <WebMethod(Description:="Guarda el log de Clave6")> _
    Public Function SAVE_LOG_CLAVE6(ByVal codigoSucursalBanco As String, _
                                    ByVal codigoKiosko As String, _
                                    ByVal idSucursal As Integer, _
                                    ByVal tipoDocumento As Integer, _
                                    ByVal nroDocumento As String, _
                                    ByVal numeroTarjeta As String, _
                                    ByVal opcion As String, _
                                    ByVal email As String, _
                                    ByVal mensajeError As String) As Boolean

        Dim estado As Boolean
        Dim oconsultaClave6 As New ConsultaClave6
        Try
            ErrorLog("Entro al método SAVE_LOG_CLAVE6")

            oconsultaClave6.Cod_Sucursal_Ban = Trim(codigoSucursalBanco)
            oconsultaClave6.Nombre_Kiosko = Trim(codigoKiosko)
            oconsultaClave6.Id_Kiosko = idSucursal
            oconsultaClave6.Tipo_Doc = Trim(tipoDocumento)
            oconsultaClave6.Nro_Doc = Trim(nroDocumento)
            oconsultaClave6.Numero_Tarjeta = Trim(numeroTarjeta)
            oconsultaClave6.Opcion = Trim(opcion)
            oconsultaClave6.Email = Trim(email)
            oconsultaClave6.Mensaje_Error = Trim(mensajeError)

            estado = BNConsultaClave6.Instancia.InsertConsultaClave6(oconsultaClave6)
        Catch ex As Exception
            estado = False
        End Try


        Return estado
    End Function

    <WebMethod(Description:="Obtener Términos y Condiciones")> _
    <Xml.Serialization.XmlInclude(GetType(Response))> _
    <Xml.Serialization.XmlInclude(GetType(TerminosCondiciones))> _
    Public Function GetTerminosCondiciones() As Response
        Dim respuesta As New Response
        respuesta.Success = True
        respuesta.Warning = False
        Dim terminos As TerminosCondiciones = Nothing

        Try
            terminos = BNConsultaClave6.Instancia.GetTerminosCondiciones()
            If terminos Is Nothing Then
                respuesta.Warning = True
                respuesta.Message = "No existen términos y condiciones"
            Else
                respuesta.Data = terminos
            End If

        Catch ex As Exception
            respuesta.Success = False
            respuesta.Message = "En estos momentos no podemos atenderle. Por favor inténtelo más tarde"
            ErrorLog("GetTerminosCondiciones Error =" & ex.Message)
        End Try
        Return respuesta
    End Function

#End Region

#Region "Validacion de DNI"

#Region "Métodos Públicos"

    <WebMethod(Description:="Validar Dni")> _
    Public Function VALIDAR_DNI(ByVal idTipoTarjeta As String, ByVal tipoDocumento As String, ByVal numeroDocumento As String, ByVal dataMonitor As String) As String
        Dim respuesta As String = String.Empty
        Dim aDataMonitor As String() = Split(dataMonitor, "|\t|", , CompareMethod.Text)
        Dim sCodigoKiosko As String
        Dim sucursal As Sucursal
        Dim idTipoTarjetaInt = Convert.ToInt32(idTipoTarjeta)
        Dim tipoDocumentoInt = Convert.ToInt32(tipoDocumento)

        Try
            sCodigoKiosko = aDataMonitor(4)
            sucursal = BuscarSucursal(sCodigoKiosko)
            If VerificarHorario(sucursal, "ValidaCliente", respuesta) = False Then
                Return "ERROR:" + respuesta
            End If

            ErrorLog("Entro al método VALIDAR_DNI(" & idTipoTarjetaInt & "," & tipoDocumentoInt & "," & numeroDocumento & "," & dataMonitor & ")")
            Dim metodoTipoDocumento As New MetodoTipoDocumento
            metodoTipoDocumento = BuscarMetodoPorDocumentoYTarjeta(tipoDocumentoInt, idTipoTarjetaInt)
            respuesta = LLamarMetodoParaValidarDni(metodoTipoDocumento.Metodo, numeroDocumento, dataMonitor, idTipoTarjetaInt, tipoDocumentoInt)
            ErrorLog("Respuesta de método VALIDAR_DNI = " & respuesta)
        Catch ex As Exception
            respuesta = "ERROR:Hubo error en llamada a método VALIDAR_DNI " & ex.Message
        End Try

        Return respuesta
    End Function

    <WebMethod(Description:="ObtenerTarjetasParaDocumento")> _
    Public Function ObtenerTarjetasParaDocumento() As List(Of TipoTarjeta)
        ErrorLog("Entro al método ObtenerTarjetasParaDocumento()")
        Dim tipoTarjetas As New List(Of TipoTarjeta)
        tipoTarjetas = BNTipoTarjeta.Instancia.ObtenerTarjetasParaDocumento()
        Return tipoTarjetas
    End Function

#End Region

#Region "Métodos Privados"

    <WebMethod(Description:="BuscarMetodoPorDocumentoYTarjeta")> _
    Private Function BuscarMetodoPorDocumentoYTarjeta(ByVal tipoDocumento As Integer, ByVal idTipoTarjeta As Integer) As MetodoTipoDocumento
        ErrorLog("Entro al Método BuscarMetodoPorDocumentoYTarjeta(" & tipoDocumento & "," & idTipoTarjeta & ")")
        Dim metodoTipoDocumento As New MetodoTipoDocumento
        metodoTipoDocumento = BNMetodoTipoDocumento.Instancia.BuscarMetodoPorDocumentoYTarjeta(tipoDocumento, idTipoTarjeta)
        Return metodoTipoDocumento
    End Function

    <WebMethod(Description:="LLamarMetodoParaValidarDni")> _
    Private Function LLamarMetodoParaValidarDni(ByVal metodo As String, ByVal numeroDocumento As String, ByVal dataMonitor As String, ByVal idTipoTarjeta As Integer, ByVal tipoDocumento As Integer) As String
        ErrorLog("Entro al método LLamarMetodoParaValidarDni(" & metodo & "," & numeroDocumento & "," & dataMonitor & "," & idTipoTarjeta & "," & tipoDocumento & ")")
        Dim respuesta As String = String.Empty
        Dim objeto As Object = New Object

        Try
            objeto = CallByName(Me, metodo, CallType.Method, numeroDocumento, dataMonitor, idTipoTarjeta.ToString(), tipoDocumento.ToString())
            ErrorLog("Respuesta del Método CallByName(" & metodo & "," & numeroDocumento & "," & dataMonitor & "," & idTipoTarjeta & "," & tipoDocumento & ")")
            If objeto.GetType() Is GetType(String) Then
                respuesta = objeto
            End If
        Catch ex As Exception
        End Try

        Return respuesta
    End Function

#End Region

#End Region

#Region "Reenganche"
#Region "Métodos Públicos"
    <WebMethod(Description:="Consulta Reenganche de SuperEfectivo SEF.")> _
    Public Function VerificarReenganche(ByVal tipoProducto As String, ByVal tipoDocumento As String, ByVal numeroDocumento As String) As String
        '17/12/2014 Retornar 0 para evitar Reenganche
        Dim respuesta As String = "0"

        'Dim respuesta As String = String.Empty
        'Dim oferta As New OfertaCliente
        'Try
        '    ErrorLog("Entro al método VerificarReenganche(" & tipoProducto & "," & tipoDocumento & "," & numeroDocumento & ")")
        '    oferta = BNOfertaCliente.Instancia.VerificarReenganche(tipoProducto, tipoDocumento, numeroDocumento)
        '    respuesta = oferta.Reenganche
        '    ErrorLog("La respuesta es " & oferta.Reenganche)
        'Catch ex As Exception
        '    respuesta = "Hubo error en llamada a método VerificarReenganche " & ex.Message
        'End Try

        Return respuesta
    End Function
#End Region
#End Region

#Region "SuperEfectivo"
#Region "Métodos Públicos"
    <WebMethod(Description:="Consulta Desembolso de SuperEfectivo SEF.")> _
    Public Function VerificarDesembolso(ByVal tipoProducto As String, ByVal tipoDocumento As String, ByVal numeroDocumento As String) As String
        Dim respuesta As String = String.Empty
        Dim oferta As New OfertaCliente
        Try
            ErrorLog("Entro al método VerificarDesembolso(" & tipoProducto & "," & tipoDocumento & "," & numeroDocumento & ")")
            oferta = BNOfertaCliente.Instancia.VerificarDesembolso(tipoProducto, tipoDocumento, numeroDocumento)
            respuesta = oferta.Estado
            ErrorLog("La respuesta es " & oferta.Estado)
        Catch ex As Exception
            respuesta = "Hubo error en llamada a método VerificarDesembolso " & ex.Message
        End Try

        Return respuesta
    End Function

    <WebMethod(Description:="Consulta SEF.")> _
    Public Function ObtenerOfertaSEF(ByVal tipoProducto As String, ByVal tipoDocumento As String, ByVal numeroDocumento As String) As String
        Dim respuesta As String = String.Empty
        Dim fechaActual As Date
        Dim oferta As New OfertaCliente
        Try
            ErrorLog("Entro al método ObtenerOfertaSEF(" & tipoProducto & "," & tipoDocumento & "," & numeroDocumento & ")")
            oferta = BNOfertaCliente.Instancia.ObtenerOferta(tipoProducto, tipoDocumento, numeroDocumento)

            respuesta = tipoProducto.Trim() & "|\t|" & tipoDocumento & "|\t|" & numeroDocumento.Trim() & "|\t|" & tipoProducto.Trim() & "|\t|" & oferta.NombreCliente.Trim()
            respuesta = respuesta & "|\t|" & oferta.NumeroCuenta.Trim() & "|\t|" & Format(CDbl(oferta.LineaOferta.Trim()), "##,##0.00") & "|\t|" & oferta.Plazo.Trim() & "|\t|" & Format(CDbl(oferta.Tasa.Trim()), "##,##0.00") & "|\t|" & Format(CDbl(oferta.Cuota.Trim()), "##,##0.00") & "|\t|" & oferta.FechaFinVigencia.Trim() & "|\t|" & GET_CODIGO_ATENCION_SEF() & "|\t|" & oferta.Tem.Trim()
            respuesta = respuesta & "|\t|" & Format(CDbl(oferta.LineaOfertaInicial.Trim()), "##,##0.00") & "|\t|" & oferta.PlazoOfertaInicial.Trim() & "|\t|" & Format(CDbl(oferta.TasaOfertaInicial.Trim()), "##,##0.00") & "|\t|" & Format(CDbl(oferta.CuotaOfertaInicial.Trim()), "##,##0.00") & "|\t|" & oferta.CodigoOferta.Trim()
            respuesta = respuesta & "|\t|" & fechaActual.Now.ToShortDateString()
            'Demo'
            'respuesta = "1|\t|1|\t|09924784|\t|1|\t|RUBEN EMILIO PARRA WILLIAMS|\t|5254740045393940|\t|4500.00|\t|36|\t|55.93|\t|21.19|\t|31/12/2014|\t|001|\t|1.99|\t|4500.00|\t|12|\t|33.5|\t|170|\t|3000|\t|30/09/2014"
            Dim vFecVence As Date
            Dim vFechaActual As Date

            vFechaActual = New Date(Now.Year, Now.Month, Now.Day)

            If oferta.FechaFinVigencia.Trim() <> "" Then
                'VALIDAR FECHAS
                If oferta.FechaFinVigencia.Trim().Length = 10 Then

                    Dim aFecha() As String = oferta.FechaFinVigencia.Trim().Split("/")
                    vFecVence = New Date(aFecha(2), aFecha(1), aFecha(0))

                    If vFecVence < vFechaActual Then 'SI LA FECHA DE VENCIMIENTO YA VENCIO NO MOSTRAR NINGUN DATO
                        respuesta = "|\t||\t||\t||\t||\t||\t||\t||\t||\t||\t||\t||\t||\t||\t||\t||\t||\t||\t|"
                    End If
                End If
            Else
                respuesta = "|\t||\t||\t||\t||\t||\t||\t||\t||\t||\t||\t||\t||\t||\t||\t||\t||\t||\t|"
            End If
        Catch ex As Exception
            respuesta = "|\t||\t||\t||\t||\t||\t||\t||\t||\t||\t||\t||\t||\t||\t||\t||\t||\t|"
        End Try

        Return respuesta
    End Function

    <WebMethod(Description:="Obtener Dirección del Banco en un Sucursal")> _
    Public Function ObtenerSucursalBanco(ByVal codigoKiosko As String) As String
        Dim respuesta As String = String.Empty
        Try
            ErrorLog("Entro al método ObtenerSucursalBanco(" & codigoKiosko & ")")
            Dim sucursal As New Sucursal
            sucursal = BNSucursal.Instancia.ObtenerSucursalBanco(codigoKiosko)
            ErrorLog("Salidó al método ObtenerSucursalBanco(" & codigoKiosko & ")")
            respuesta = sucursal.Banco
        Catch ex As Exception
            respuesta = "Hubo error en llamada a método ObtenerSucursalBanco " & ex.Message
        End Try

        Return respuesta
    End Function

    <WebMethod(Description:="Obtener el factor de oferta por plazo")> _
    Public Function ObtenerFactorPorPlazo(ByVal plazo As String) As String
        Dim respuesta As String = String.Empty
        Try
            ErrorLog("Entro al método ObtenerFactorPorPlazo(" & plazo & ")")
            Dim factor As New Factor
            factor = BNFactor.Instancia.ObtenerFactorPorPlazo(plazo)
            ErrorLog("Salió al método ObtenerFactorPorPlazo(" & plazo & ")")
            If factor.FactorPlazo = "" Then
                factor.FactorPlazo = "1"
            End If
            respuesta = factor.FactorPlazo
        Catch ex As Exception
            respuesta = "Hubo error en llamada a método ObtenerSucursalBanco " & ex.Message
        End Try

        Return respuesta
    End Function

    <WebMethod(Description:="Obtener Fecha Facturación")> _
    Public Function ObtenerFechaFacturacion(ByVal ultimoDiaPago As String, ByVal mes As String, ByVal anio As String) As String
        Dim diasMes As Integer = 0
        diasMes = Date.DaysInMonth(CInt(anio), CInt(mes))
        Dim respuesta As String = String.Empty
        Try

            Select Case ultimoDiaPago
                Case "01"
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
                Case "05"
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
                Case "10"
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
                Case "15"
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
                Case "20"
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
                Case "25"
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

    <WebMethod(Description:="Obtener Cronograma de Pagos")> _
    Public Function ObtenerCronogramaPagos(ByVal sMonto As String, ByVal sPlazoMeses As String, ByVal sdiferido As String, ByVal sTasaTEA As String, ByVal sfechaConsumo As String, ByVal sfechaFacturacion As String) As String
        Dim respuesta As String = String.Empty
        Try

            Dim simuladorSEF As New SimuladorSEF
            simuladorSEF.LNK_FINANCIADO = sMonto.Trim()
            simuladorSEF.LNK_CUOTAS = sPlazoMeses.Trim()
            simuladorSEF.LNK_DIFERIDO = sdiferido.Trim()
            simuladorSEF.LNK_TASA_ANUAL = sTasaTEA
            simuladorSEF.LNK_FECHA_CONSUMO = sfechaConsumo
            simuladorSEF.LNK_FECHA_FACTURACION = sfechaFacturacion
            ErrorLog(" simuladorSEF.LNK_FINANCIADO: " & simuladorSEF.LNK_FINANCIADO)
            ErrorLog(" simuladorSEF.LNK_CUOTAS: " & simuladorSEF.LNK_CUOTAS)
            ErrorLog(" simuladorSEF.LNK_DIFERIDO: " & simuladorSEF.LNK_DIFERIDO)
            ErrorLog(" simuladorSEF.LNK_TASA_ANUAL: " & simuladorSEF.LNK_TASA_ANUAL)
            ErrorLog(" simuladorSEF.LNK_FECHA_CONSUMO: " & simuladorSEF.LNK_FECHA_CONSUMO)
            ErrorLog(" simuladorSEF.LNK_FECHA_FACTURACION: " & simuladorSEF.LNK_FECHA_FACTURACION)
            simuladorSEF = BNSimuladoSEF.Instancia.ObtenerCronogramaPagos(simuladorSEF)
            ErrorLog(" simuladorSEF.LNK_ESTADO: " & simuladorSEF.LNK_ESTADO)
            ErrorLog(" simuladorSEF.LNK_TABLA: " & simuladorSEF.LNK_TABLA)

            If simuladorSEF.Resultado = 1 Then
                If simuladorSEF.LNK_ESTADO = Constantes.ServicioSFSCANT0013ESTADOEXITOSO Or simuladorSEF.LNK_ESTADO = Constantes.ServicioSFSCANT0013ESTADOEXITOSOVACIO Then
                    respuesta = simuladorSEF.LNK_TABLA
                    ErrorLog("EXITOSO: simuladorSEF.LNK_TABLA: " & respuesta)

                Else
                    respuesta = "ERROR:NODATA-No se pudo realizar el cálculo de Cuota Mes. " & simuladorSEF.Mensaje
                    ErrorLog(respuesta)
                End If
            Else
                respuesta = "ERROR:NODATA-No se pudo realizar el cálculo de Cuota Mes. " & simuladorSEF.Mensaje
                ErrorLog(respuesta)
            End If

        Catch ex As Exception
            respuesta = "ERROR:NODATA-Hubo error en llamada a método ObtenerCronogramaPagos " & ex.Message
        End Try

        Return respuesta
    End Function

    <WebMethod(Description:="Obtener Cuota Mes")> _
    Public Function ObtenerCuotaMes(ByVal sMonto As String, ByVal sPlazoMeses As String, ByVal sdiferido As String, ByVal sTasaTEA As String, ByVal sfechaConsumo As String, ByVal sfechaFacturacion As String) As String
        Dim respuesta As String = String.Empty
        Dim cronograma As String = String.Empty
        Try
            cronograma = ObtenerCronogramaPagos(sMonto.Trim, sPlazoMeses.Trim, sdiferido.Trim, sTasaTEA.Trim, sfechaConsumo.Trim, sfechaFacturacion.Trim)

            If cronograma.Substring(0, 5) = "ERROR" Then
                respuesta = cronograma
            Else
                respuesta = cronograma.Substring(0, 51)
                ErrorLog("Primera Fila: " & respuesta)
                respuesta = cronograma.Substring(43, 8)
                ErrorLog("CuotaMes: " & respuesta)
            End If

        Catch ex As Exception
            respuesta = "ERROR:NODATA-Hubo error en llamada a método ObtenerCuotaMes " & ex.Message
        End Try

        Return respuesta
    End Function

#End Region
#End Region

#Region "Clave 6"

    <WebMethod(Description:="Validar por NroTarjeta, si un cliente puede acceder a la opción Generar Clave 6 de Ripleymático.")> _
    Public Function PuedeGenerarClave6(ByVal nroTarjeta As String) As Boolean
        Log.ErrorLog("PuedeGenerarClave6 Inicio")
        Dim Respuesta As Boolean
        Dim CountClientes As Integer

        Try
            CountClientes = BNConsultaClave6.Instancia.CountClientesByNroTarjeta(nroTarjeta)
            Log.ErrorLog("PuedeGenerarClave6 CountClientes: " + CountClientes.ToString())
            Respuesta = (CountClientes > 0)
        Catch ex As Exception
            Log.ErrorLog("PuedeGenerarClave6 Exception: " + ex.Message)
            Respuesta = False
        End Try
        Return Respuesta
    End Function

#End Region

    ''' <summary>
    ''' Valida la fecha de desembolso, que sea una fecha válida y que sea mayor o igual a la actual
    ''' </summary>
    ''' <param name="valor">Valor de fecha a validar</param>
    ''' <returns>Verdadero si es válida, caso contrario devuelve falso</returns>
    ''' <remarks></remarks>
    Private Function ValidarFechaDesembolso(ByVal valor As DateTime) As Boolean
        Dim hoy As DateTime = Date.Today
        If Not (valor = Date.MinValue) And valor >= hoy Then
            Return True
        End If
        Return False
    End Function
End Class