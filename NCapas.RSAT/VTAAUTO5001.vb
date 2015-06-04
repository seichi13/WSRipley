Imports NCapas.Utility.Funciones
Imports NCapas.Utility.Tipos
Imports NCapas.Utility.Log
Imports MQCOMLib
Imports NCapas.Entity
Imports NCapas.Utility

Public Class VTAAUTO5001
    Private header As New VTAAUTO5001Header
    Private parameterIn As New VTAAUTO5001ParameterIN
    Private parameterOut As New VTAAUTO5001ParameterOUT
    Private serviceName As String = String.Empty

    Sub New(ByVal codRetorno As String,
            ByVal nombreServicio As String,
            ByVal largoMensaje As String,
            ByVal codKiosko As String,
            ByVal idSucursal As Integer)

        Dim fechaActual As String = Date.Now.ToShortDateString()

        Dim dia As String = fechaActual.Substring(0, 2)
        Dim mes As String = fechaActual.Substring(3, 2)
        Dim anio As String = fechaActual.Substring(6, 4)

        serviceName = nombreServicio
        header.CodRetorno = fx_Completar_Campo("0", 10, codRetorno, TYPE_ALINEAR.DERECHA)
        header.NombreServicio = fx_Completar_Campo(" ", 50, nombreServicio.ToUpper(), TYPE_ALINEAR.IZQUIERDA)
        header.LargoMensaje = fx_Completar_Campo("0", 9, largoMensaje, TYPE_ALINEAR.DERECHA)

        header.CodEntidad = fx_Completar_Campo("0", 4, Constantes.ServicioVTAAUTO5001CODENTIDAD, TYPE_ALINEAR.IZQUIERDA)
        header.TipoTerminal = fx_Completar_Campo("0", 6, Constantes.ServicioVTAAUTO5001TIPOTERMINAL, TYPE_ALINEAR.IZQUIERDA)
        header.CodUsuario = fx_Completar_Campo(" ", 8, codKiosko.ToUpper(), TYPE_ALINEAR.IZQUIERDA)
        header.CodOficina = fx_Completar_Campo("0", 4, idSucursal.ToString(), TYPE_ALINEAR.DERECHA)
        header.TerminalFisico = fx_Completar_Campo(" ", 4, Constantes.ServicioVTAAUTO5001TERMINALFISICO, TYPE_ALINEAR.IZQUIERDA)
        header.FechaConta = fx_Completar_Campo(" ", 10, anio & "-" & mes & "-" & dia, TYPE_ALINEAR.IZQUIERDA)
        header.CodAplicacion = fx_Completar_Campo(" ", 2, Constantes.SistemaRM, TYPE_ALINEAR.IZQUIERDA)
        header.CodPais = fx_Completar_Campo(" ", 2, Constantes.ServicioVTAAUTO5001PAIS, TYPE_ALINEAR.IZQUIERDA)
        header.Commit = fx_Completar_Campo(" ", 1, Constantes.ServicioVTAAUTO5001COMMIT, TYPE_ALINEAR.IZQUIERDA)
        header.Paginacion = fx_Completar_Campo(" ", 1, " ", TYPE_ALINEAR.IZQUIERDA)
        header.ClaveInicio = fx_Completar_Campo(" ", 200, " ", TYPE_ALINEAR.IZQUIERDA)
        header.ClaveFin = fx_Completar_Campo(" ", 200, " ", TYPE_ALINEAR.IZQUIERDA)
        header.Pantalla = fx_Completar_Campo("0", 3, Constantes.ServicioVTAAUTO5001PANTALLA, TYPE_ALINEAR.IZQUIERDA)
        header.MasDatos = fx_Completar_Campo(" ", 1, " ", TYPE_ALINEAR.IZQUIERDA)
        header.OtrosDatos = fx_Completar_Campo(" ", 90, Constantes.ServicioVTAAUTO5001OTROSDATOS, TYPE_ALINEAR.IZQUIERDA)

    End Sub

    Sub AddINParameters(ByVal parametrosIn As VTAAUTO5001ParameterIN)

        parametrosIn.TipoTransaccion = fx_Completar_Campo(" ", 1, parametrosIn.TipoTransaccion, TYPE_ALINEAR.IZQUIERDA)

        parameterIn.CodEntidadTarjeta = fx_Completar_Campo("0", 4, parametrosIn.CodEntidadTarjeta, TYPE_ALINEAR.DERECHA)
        parameterIn.CentroAltaTarjeta = fx_Completar_Campo("0", 4, parametrosIn.CentroAltaTarjeta, TYPE_ALINEAR.DERECHA)
        parameterIn.CuentaTarjeta = fx_Completar_Campo("0", 12, parametrosIn.CuentaTarjeta, TYPE_ALINEAR.DERECHA)
        parameterIn.CodProductoTarjeta = fx_Completar_Campo("0", 10, parametrosIn.CodProductoTarjeta, TYPE_ALINEAR.DERECHA)
        parameterIn.CodProductoTarjeta = fx_Completar_Campo("0", 10, parametrosIn.CodProductoTarjeta, TYPE_ALINEAR.DERECHA)
        parameterIn.CodCambioTarjeta = fx_Completar_Campo("0", 4, parametrosIn.CodCambioTarjeta, TYPE_ALINEAR.DERECHA)

        parameterIn.CodEntidadSEF = fx_Completar_Campo("0", 4, parametrosIn.CodEntidadSEF, TYPE_ALINEAR.DERECHA)
        parameterIn.CentroAltaSEF = fx_Completar_Campo("0", 4, parametrosIn.CentroAltaSEF, TYPE_ALINEAR.DERECHA)
        parameterIn.CuentaSEF = fx_Completar_Campo("0", 12, parametrosIn.CuentaSEF, TYPE_ALINEAR.DERECHA)
        parameterIn.CodProductoSEF = fx_Completar_Campo("0", 10, parametrosIn.CodProductoSEF, TYPE_ALINEAR.DERECHA)
        parameterIn.CodProductoSEF = fx_Completar_Campo("0", 10, parametrosIn.CodProductoSEF, TYPE_ALINEAR.DERECHA)
        parameterIn.CodCambioSEF = fx_Completar_Campo("0", 4, parametrosIn.CodCambioSEF, TYPE_ALINEAR.DERECHA)

        parameterIn.TipoDocumento = fx_Completar_Campo(" ", 2, parametrosIn.TipoDocumento, TYPE_ALINEAR.IZQUIERDA)
        parameterIn.NumeroDocumento = fx_Completar_Campo(" ", 12, parametrosIn.NumeroDocumento, TYPE_ALINEAR.IZQUIERDA)

        parameterIn.SystemSource = fx_Completar_Campo(" ", 10, parametrosIn.SystemSource, TYPE_ALINEAR.IZQUIERDA)
        parameterIn.ModuleId = fx_Completar_Campo("0", 2, parametrosIn.ModuleId, TYPE_ALINEAR.DERECHA)
        parameterIn.TipoTerminal = fx_Completar_Campo(" ", 10, parametrosIn.TipoTerminal, TYPE_ALINEAR.IZQUIERDA)
        parameterIn.CodUsuario = fx_Completar_Campo(" ", 10, parametrosIn.CodUsuario, TYPE_ALINEAR.IZQUIERDA)
        parameterIn.CodOficina = fx_Completar_Campo(" ", 10, parametrosIn.CodOficina, TYPE_ALINEAR.IZQUIERDA)
        parameterIn.MotImpresion = fx_Completar_Campo(" ", 10, parametrosIn.MotImpresion, TYPE_ALINEAR.IZQUIERDA)

        parameterIn.ModoEstampacion = fx_Completar_Campo(" ", 10, parametrosIn.ModoEstampacion, TYPE_ALINEAR.IZQUIERDA)
        parameterIn.SEnvioOficina = fx_Completar_Campo(" ", 10, parametrosIn.SEnvioOficina, TYPE_ALINEAR.IZQUIERDA)
        parameterIn.CodDestino = fx_Completar_Campo(" ", 1, parametrosIn.CodDestino, TYPE_ALINEAR.IZQUIERDA)
        parameterIn.EnvEmailPers = fx_Completar_Campo(" ", 1, parametrosIn.EnvEmailPers, TYPE_ALINEAR.IZQUIERDA)
        parameterIn.EnvEmailLab = fx_Completar_Campo(" ", 1, parametrosIn.EnvEmailLab, TYPE_ALINEAR.IZQUIERDA)
        parameterIn.Envtarjeta = fx_Completar_Campo(" ", 1, parametrosIn.Envtarjeta, TYPE_ALINEAR.IZQUIERDA)
        parameterIn.CodUsrApr = fx_Completar_Campo(" ", 12, parametrosIn.CodUsrApr, TYPE_ALINEAR.IZQUIERDA)
        parameterIn.NumeroIp = fx_Completar_Campo(" ", 15, parametrosIn.NumeroIp, TYPE_ALINEAR.IZQUIERDA)

    End Sub

    Sub AddOUTParameters(ByVal parametrosOUT As VTAAUTO5001ParameterOUT)

        parameterOut.ProdTarjetaOld = fx_Completar_Campo(" ", 2, parametrosOUT.ProdTarjetaOld, TYPE_ALINEAR.IZQUIERDA)
        parameterOut.SubProdTarjetaOld = fx_Completar_Campo(" ", 4, parametrosOUT.SubProdTarjetaOld, TYPE_ALINEAR.IZQUIERDA)
        parameterOut.ConProdTarjetaOld = fx_Completar_Campo(" ", 3, parametrosOUT.ConProdTarjetaOld, TYPE_ALINEAR.IZQUIERDA)

        parameterOut.ProdSEFOld = fx_Completar_Campo(" ", 2, parametrosOUT.ProdSEFOld, TYPE_ALINEAR.IZQUIERDA)
        parameterOut.SubProdSEFOld = fx_Completar_Campo(" ", 4, parametrosOUT.SubProdSEFOld, TYPE_ALINEAR.IZQUIERDA)
        parameterOut.ConProdSEFOld = fx_Completar_Campo(" ", 3, parametrosOUT.ConProdSEFOld, TYPE_ALINEAR.IZQUIERDA)

        parameterOut.ProdTarjetaNew = fx_Completar_Campo(" ", 2, parametrosOUT.ProdTarjetaNew, TYPE_ALINEAR.IZQUIERDA)
        parameterOut.SubProdTarjetaNew = fx_Completar_Campo(" ", 4, parametrosOUT.SubProdTarjetaNew, TYPE_ALINEAR.IZQUIERDA)
        parameterOut.ConProdTarjetaNew = fx_Completar_Campo(" ", 3, parametrosOUT.ConProdTarjetaNew, TYPE_ALINEAR.IZQUIERDA)
        parameterOut.CodMarTarjeta = fx_Completar_Campo("0", 2, parametrosOUT.CodMarTarjeta, TYPE_ALINEAR.DERECHA)
        parameterOut.IndTiPtTarjeta = fx_Completar_Campo("0", 2, parametrosOUT.IndTiPtTarjeta, TYPE_ALINEAR.DERECHA)
        parameterOut.DesProductoTarjeta = fx_Completar_Campo(" ", 50, parametrosOUT.DesProductoTarjeta, TYPE_ALINEAR.IZQUIERDA)
        parameterOut.DesMarcaTarjeta = fx_Completar_Campo(" ", 50, parametrosOUT.DesMarcaTarjeta, TYPE_ALINEAR.IZQUIERDA)
        parameterOut.DesTipoTarjeta = fx_Completar_Campo(" ", 50, parametrosOUT.DesTipoTarjeta, TYPE_ALINEAR.IZQUIERDA)
        parameterOut.CodClienteTarjeta = fx_Completar_Campo("0", 10, parametrosOUT.CodClienteTarjeta, TYPE_ALINEAR.DERECHA)
        parameterOut.NumTarjetas = fx_Completar_Campo("0", 2, parametrosOUT.NumTarjetas, TYPE_ALINEAR.DERECHA)

        parameterOut.ProdSEFNew = fx_Completar_Campo(" ", 2, parametrosOUT.ProdSEFNew, TYPE_ALINEAR.IZQUIERDA)
        parameterOut.SubProdSEFNew = fx_Completar_Campo(" ", 4, parametrosOUT.SubProdSEFNew, TYPE_ALINEAR.IZQUIERDA)
        parameterOut.ConProdSEFNew = fx_Completar_Campo(" ", 3, parametrosOUT.ConProdSEFNew, TYPE_ALINEAR.IZQUIERDA)
        parameterOut.CodMarSEF = fx_Completar_Campo("0", 2, parametrosOUT.CodMarSEF, TYPE_ALINEAR.DERECHA)
        parameterOut.IndTiPtSEF = fx_Completar_Campo("0", 2, parametrosOUT.IndTiPtSEF, TYPE_ALINEAR.DERECHA)
        parameterOut.DesProductoSEF = fx_Completar_Campo(" ", 50, parametrosOUT.DesProductoSEF, TYPE_ALINEAR.IZQUIERDA)
        parameterOut.DesMarcaSEF = fx_Completar_Campo(" ", 50, parametrosOUT.DesMarcaSEF, TYPE_ALINEAR.IZQUIERDA)
        parameterOut.DesTipoSEF = fx_Completar_Campo(" ", 50, parametrosOUT.DesTipoSEF, TYPE_ALINEAR.IZQUIERDA)
        parameterOut.CodClienteSEF = fx_Completar_Campo("0", 10, parametrosOUT.CodClienteSEF, TYPE_ALINEAR.DERECHA)
        parameterOut.NumSEFs = fx_Completar_Campo("0", 2, parametrosOUT.NumSEFs, TYPE_ALINEAR.DERECHA)

        parameterOut.CuentaTarjetaPTNEnc = fx_Completar_Campo("0", 1, parametrosOUT.CuentaTarjetaPTNEnc, TYPE_ALINEAR.DERECHA)
        parameterOut.TieneSEFPTN = fx_Completar_Campo("0", 1, parametrosOUT.TieneSEFPTN, TYPE_ALINEAR.DERECHA)
        parameterOut.ContratoSEFPTN = fx_Completar_Campo(" ", 20, parametrosOUT.ContratoSEFPTN, TYPE_ALINEAR.IZQUIERDA)
        parameterOut.CuentaTarjetaEnc = fx_Completar_Campo("0", 1, parametrosOUT.CuentaTarjetaEnc, TYPE_ALINEAR.DERECHA)
        parameterOut.CodCambioTarjetaEnc = fx_Completar_Campo("0", 1, parametrosOUT.CodCambioTarjetaEnc, TYPE_ALINEAR.DERECHA)
        parameterOut.CambiotarjetaOK = fx_Completar_Campo("0", 1, parametrosOUT.CambiotarjetaOK, TYPE_ALINEAR.DERECHA)

        parameterOut.CuentaSEFPTNEnc = fx_Completar_Campo("0", 1, parametrosOUT.CuentaSEFPTNEnc, TYPE_ALINEAR.DERECHA)
        parameterOut.CuentaSEFEnc = fx_Completar_Campo("0", 1, parametrosOUT.CuentaSEFEnc, TYPE_ALINEAR.DERECHA)
        parameterOut.CodCambioSEFEnc = fx_Completar_Campo("0", 1, parametrosOUT.CodCambioSEFEnc, TYPE_ALINEAR.DERECHA)
        parameterOut.CambioSEFOK = fx_Completar_Campo("0", 1, parametrosOUT.CambioSEFOK, TYPE_ALINEAR.DERECHA)

        parameterOut.Paso1 = fx_Completar_Campo("0", 2, parametrosOUT.Paso1, TYPE_ALINEAR.DERECHA)
        parameterOut.Paso2 = fx_Completar_Campo(" ", 1, parametrosOUT.Paso2, TYPE_ALINEAR.IZQUIERDA)
        parameterOut.CodRpta = fx_Completar_Campo(" ", 2, parametrosOUT.CodRpta, TYPE_ALINEAR.IZQUIERDA)
        parameterOut.DesRpta = fx_Completar_Campo(" ", 200, parametrosOUT.DesRpta, TYPE_ALINEAR.IZQUIERDA)

    End Sub

    Public Function Execute() As VTAAUTO5001ParameterOUT
        Dim oMQCOM As New MQ
        Dim tramaRespuesta As String = String.Empty
        Dim tramaEnvio As String = String.Empty

        Try

            tramaEnvio = header.ToString() & parameterIn.ToString() & parameterOut.ToString()
            Registrar_Log(3, serviceName, "IN: " & tramaEnvio)

            oMQCOM.Service = serviceName
            oMQCOM.Message = tramaEnvio.ToUpper()
            oMQCOM.Execute()
            tramaRespuesta = oMQCOM.Response

            'If tramaEnvio.Substring(639, 1) = "1" Then
            '    'tramaRespuesta = "0000000000VTAAUTO5001                                       0000010240001000004SIRIP-990188VACP2015-02-11RMPES                                                                                                                                                                                                                                                                                                                                                                                                                 000 ES                                                                                        00010001000005351359C 10085380    1SAT       188       SIRIP-99    10.25.4.6      04801110001000100000001212110301EL CONTRATO CUMPLE CON LAS CONDICIONES DE CAMBIO DE PRODUCTO. CODIGO DE CAMBIO DE PRODUCYO (0480) ENCONTRADO                                                                                                                                                                                                                                                 "
            '    tramaRespuesta = "0000000000VTAAUTO5001                                       0000010240001000004SIRIP-990188VACP2015-02-11RMPES                                                                                                                                                                                                                                                                                                                                                                                                                 000 ES                                                                                        00010001000005305448C 10085380    1SAT       188       SIRIP-99    127.0.0.1      00871110001000100000535135910301EL CONTRATO CUMPLE CON LAS CONDICIONES DE CAMBIO DE PRODUCTO. CLIENTE CON UN CONTRATO SEF VIGENTE.                                                                                                                                                                                                                                                           "
            'Else
            '    'tramaRespuesta = "0000000000VTAAUTO5001                                       0000010240001000004SIRIP-990188VACP2015-02-11RMPES                                                                                                                                                                                                                                                                                                                                                                                                                 000 ES                                                                                        00010001000005351359C 10085380    2SAT       188       SIRIP-99    10.25.4.6      04801110001000100000001212110200EL CONTRATO CUMPLE CON LAS CONDICIONES DE CAMBIO DE PRODUCTO. CODIGO DE CAMBIO DE PRODUCYO (0480) ENCONTRADO                                                                                                                                                                                                                                               "
            '    tramaRespuesta = "0000000000VTAAUTO5001                                       0000010240001000004SIRIP-990188VACP2015-02-11RMPES                                                                                                                                                                                                                                                                                                                                                                                                                 000 ES                                                                                        00010001000005351359C 10085380    2SAT       188       SIRIP-99    127.0.0.1      04801100000000000000000000000200EL CONTRATO CUMPLE CON LAS CONDICIONES DE CAMBIO DE PRODUCTO. CODIGO DE CAMBIO DE PRODUCTO (0480) ENCONTRADO.                                                                                                                                                                                                                                              "
            'End If
            'oMQCOM.ReasonMd = 0
            Registrar_Log(3, serviceName, "OUT:" & tramaRespuesta)

            If oMQCOM.ReasonMd = 0 Then
                Registrar_Log(3, serviceName, "oMQCOM.ReasonMd = 0 " & oMQCOM.ReasonMd)

                parameterOut = New VTAAUTO5001ParameterOUT(tramaRespuesta)
                parameterOut.Resultado = 1
            Else
                parameterOut.Resultado = 0
                parameterOut.Mensaje = "Hubo un Error en la respuesta Servicio VTAAUTO5000. Inténtelo más tarde"
                Registrar_Log(3, serviceName, "oMQCOM.ReasonMd != 0 " & oMQCOM.ReasonMd)
            End If

        Catch ex As Exception
            Registrar_Log(3, serviceName, "Exception:" & ex.Message)
            parameterOut.Resultado = 0
            parameterOut.Mensaje = "Hubo un Error en la invocación de Servicio VTAAUTO5000. Inténtelo más tarde"
        End Try

        Return parameterOut

    End Function
End Class
