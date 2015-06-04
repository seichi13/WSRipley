Imports NCapas.Utility.Funciones
Imports NCapas.Utility.Tipos
Imports NCapas.Utility.Log
Imports MQCOMLib
Imports NCapas.Entity

Public Class SFSCANT0013

    Private header As New SFSCANT0013Header
    Private parameterIn As New SFSCANT0013ParameterIN
    Private parameterOut As New SFSCANT0013ParameterOUT
    Private serviceName As String = String.Empty

    Sub New(ByVal lnkCodRetorno As String,
            ByVal lnkNombreServicio As String,
            ByVal lnkLargoMensaje As String)

        serviceName = lnkNombreServicio
        header.LNK_COD_RETORNO = fx_Completar_Campo("0", 10, lnkCodRetorno, TYPE_ALINEAR.DERECHA)
        header.LNK_NOMBRE_SERVICIO = fx_Completar_Campo(" ", 50, lnkNombreServicio, TYPE_ALINEAR.IZQUIERDA)
        header.LNK_LARGO_MENSAJE = fx_Completar_Campo("0", 9, lnkLargoMensaje, TYPE_ALINEAR.DERECHA)

    End Sub

    Sub AddINParameters(ByVal parametroIN As SimuladorSEF)

        parameterIn.LNK_FINANCIADO = fx_Completar_Campo("0", 8, parametroIN.LNK_FINANCIADO, TYPE_ALINEAR.DERECHA)
        parameterIn.LNK_CUOTAS = fx_Completar_Campo("0", 2, parametroIN.LNK_CUOTAS, TYPE_ALINEAR.DERECHA)
        parameterIn.LNK_DIFERIDO = fx_Completar_Campo("0", 1, parametroIN.LNK_DIFERIDO, TYPE_ALINEAR.DERECHA)
        parameterIn.LNK_TASA_ANUAL = fx_Completar_Campo("0", 5, parametroIN.LNK_TASA_ANUAL, TYPE_ALINEAR.DERECHA)
        parameterIn.LNK_FECHA_CONSUMO = fx_Completar_Campo("0", 8, parametroIN.LNK_FECHA_CONSUMO, TYPE_ALINEAR.DERECHA)
        parameterIn.LNK_FECHA_FACTURACION = fx_Completar_Campo("0", 8, parametroIN.LNK_FECHA_FACTURACION, TYPE_ALINEAR.DERECHA)

    End Sub

    Public Function Execute() As SimuladorSEF

        Dim simuladorSEF As New SimuladorSEF
        Dim oMQCOM As New MQ
        Dim tramaRespuesta As String = String.Empty
        Dim tramaEnvio As String = String.Empty

        Try

            tramaEnvio = header.ToString() & parameterIn.ToString()
            Registrar_Log(3, serviceName, "IN:" & tramaEnvio)

            oMQCOM.Service = serviceName
            oMQCOM.Message = tramaEnvio
            oMQCOM.Execute()
            tramaRespuesta = oMQCOM.Response
            Registrar_Log(3, serviceName, "OUT:" & tramaRespuesta)
            If oMQCOM.ReasonMd = 0 Then
                Registrar_Log(3, serviceName, "oMQCOM.ReasonMd = 0 " & oMQCOM.ReasonMd)
                'OK
                Registrar_Log(3, serviceName, "tramaRespuesta.Substring(101, 1) = " & tramaRespuesta.Substring(101, 1))
                Registrar_Log(3, serviceName, "tramaRespuesta.Substring(101, 1) = " & tramaRespuesta.Substring(102, tramaRespuesta.Length - 102))
                parameterOut = New SFSCANT0013ParameterOUT(tramaRespuesta)
                simuladorSEF.LNK_ESTADO = parameterOut.LNK_ESTADO
                simuladorSEF.LNK_TABLA = parameterOut.LNK_TABLA
                simuladorSEF.Resultado = 1
                simuladorSEF.Mensaje = parameterOut.LNK_TABLA
            Else
                simuladorSEF.Resultado = 0
                simuladorSEF.Mensaje = "Hubo un Error en la respuesta de Cuota Mes. Inténtelo más tarde"
                Registrar_Log(3, serviceName, "oMQCOM.ReasonMd != 0 " & oMQCOM.ReasonMd)
            End If

        Catch ex As Exception
            Registrar_Log(3, serviceName, "Exception:" & ex.Message)
            simuladorSEF.Resultado = 0
            simuladorSEF.Mensaje = "Hubo un Error en la invocación de Servicio Cuota Mes. Inténtelo más tarde"
        End Try

        Return simuladorSEF

    End Function

End Class
