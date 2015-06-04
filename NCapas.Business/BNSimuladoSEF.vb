Imports NCapas.Entity
Imports NCapas.Data
Imports NCapas.Utility
Imports NCapas.Utility.Log
Imports NCapas.RSAT.Servicio
Imports NCapas.RSAT

Public Class BNSimuladoSEF
    Inherits Singleton(Of BNSimuladoSEF)

    Private servicioSFSCANT0013 As SFSCANT0013

    Public Function ObtenerCronogramaPagos(ByVal parameterIN As SimuladorSEF) As SimuladorSEF

        Dim simuladorSEF As New SimuladorSEF
        Registrar_Log(3, Constantes.ServicioSFSCANT0013NOMBRE, "Antes de llamar a ValidarParametrosEntradaSEF(parameterIN)")
        ValidarParametrosEntradaSEF(parameterIN)
        Registrar_Log(3, Constantes.ServicioSFSCANT0013NOMBRE, "Inicializar Service()")
        servicioSFSCANT0013 = New SFSCANT0013(Constantes.ServicioSFSCANT0013CODRETORNO, Constantes.ServicioSFSCANT0013NOMBRE, Constantes.ServicioSFSCANT0013LONGMENSAJE)
        servicioSFSCANT0013.AddINParameters(parameterIN)
        Registrar_Log(3, Constantes.ServicioSFSCANT0013NOMBRE, "llamar a servicioSFSCANT0013.Execute()")
        simuladorSEF = servicioSFSCANT0013.Execute()
        Registrar_Log(3, "", "simuladorSEF.Mensaje=" & simuladorSEF.Mensaje)

        Return simuladorSEF
    End Function

    Private Sub ValidarParametrosEntradaSEF(ByVal parameterIN As SimuladorSEF)
        Registrar_Log(3, "", "Antes de llamar fx_Preparar_Decimal " & parameterIN.LNK_FINANCIADO)
        parameterIN.LNK_FINANCIADO = Utility.Funciones.fx_Preparar_Decimal(parameterIN.LNK_FINANCIADO)
        Registrar_Log(3, "", "Antes de llamar fx_Preparar_Decimal " & parameterIN.LNK_TASA_ANUAL)
        parameterIN.LNK_TASA_ANUAL = Utility.Funciones.fx_Preparar_Decimal(parameterIN.LNK_TASA_ANUAL)
    End Sub

End Class
