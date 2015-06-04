Imports NCapas.Entity.IncrementoLinea
Imports NCapas.RSAT.VTAAUTO0004
Imports NCapas.RSAT

Public Class BN_Incremento_Linea

    Private oRSAT_Inc_Linea As VTAAUTO0004

    Public Function Buscar_Oferta(ByVal pNro_Contrato As String) As Oferta
        Dim oOferta As New Oferta
        Return oOferta
    End Function

    Public Function Aprobar_Oferta(ByVal oParametros As Parametros) As Aprobacion
        Dim oAprobacion As New Aprobacion
        Dim oResultado As New PARAMETROS_OUT
        oRSAT_Inc_Linea = New VTAAUTO0004(oParametros.Cod_Usuario_Plataforma, oParametros.Cod_Oficina, oParametros.Fecha_Conta)

        oRSAT_Inc_Linea.Add_Parametros(oParametros.Nro_Contrato, _
                                       oParametros.Linea_EGM, _
                                       oParametros.Linea_1, _
                                       oParametros.Linea_2, _
                                       oParametros.Cod_Sucursal, _
                                       oParametros.Nro_Caja, _
                                       oParametros.Nro_Transaccion)

        oResultado = oRSAT_Inc_Linea.Ejecutar_Servicio()

        If oResultado.COD_RESPUESTA = "00" Then
            oAprobacion.Cod_Respuesta = oResultado.COD_RESPUESTA
            oAprobacion.Msj_Respuesta = oResultado.MSJ_RESPUESTA
            oAprobacion.Paso = oResultado.PASO
            oAprobacion.Upd_Linea_EGM_OK = oResultado.LIMCRECTA_EXI
            oAprobacion.Upd_Linea_1_OK = oResultado.LIMCRELNA_EXI_1
            oAprobacion.Upd_Linea_2_OK = oResultado.LIMCRELNA_EXI_2
        Else
            oAprobacion.Cod_Respuesta = oResultado.COD_RESPUESTA
            oAprobacion.Msj_Respuesta = oResultado.MSJ_RESPUESTA
            oAprobacion.Paso = String.Empty
            oAprobacion.Upd_Linea_EGM_OK = String.Empty
            oAprobacion.Upd_Linea_1_OK = String.Empty
            oAprobacion.Upd_Linea_2_OK = String.Empty
        End If


        Return oAprobacion


    End Function


End Class
