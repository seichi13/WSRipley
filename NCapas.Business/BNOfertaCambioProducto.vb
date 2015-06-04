Imports NCapas.Entity
Imports NCapas.Data
Imports NCapas.Utility
Imports NCapas.Utility.Log
Imports NCapas.RSAT

Public Class BNOfertaCambioProducto
    Inherits Singleton(Of BNOfertaCambioProducto)

    Private servicioVTAAUTO5001 As VTAAUTO5001

    Public Function ObtenerOferta(ByVal contrato As String) As OfertaCambioProducto
        Dim oferta As New OfertaCambioProducto
        Try
            oferta = DAOfertaCambioProducto.Instancia.ObtenerOferta(contrato)
        Catch ex As Exception
            ErrorLog(ex.Message)
            oferta.MensajeError = "Hubo un error al invocar servicio"
        End Try
        Return oferta
    End Function

    Public Function ValidarContrato(ByVal parameterIN As VTAAUTO5001ParameterIN, ByVal codKiosko As String, ByVal idSucursal As Integer) As VTAAUTO5001ParameterOUT

        Dim respuesta As New VTAAUTO5001ParameterOUT
        Registrar_Log(3, Constantes.ServicioVTAAUTO5001NOMBRE, "Inicializar Service()")
        servicioVTAAUTO5001 = New VTAAUTO5001(Constantes.ServicioVTAAUTO5001CODRETORNO, Constantes.ServicioVTAAUTO5001NOMBRE, Constantes.ServicioVTAAUTO5001LONGMENSAJE, codKiosko, idSucursal)
        servicioVTAAUTO5001.AddINParameters(parameterIN)
        servicioVTAAUTO5001.AddOUTParameters(New VTAAUTO5001ParameterOUT)
        Registrar_Log(3, Constantes.ServicioVTAAUTO5001NOMBRE, "llamar a servicioVTAAUTO5000.Execute()")
        respuesta = servicioVTAAUTO5001.Execute()
        Registrar_Log(3, Constantes.ServicioVTAAUTO5001NOMBRE, "respuesta.Mensaje=" & respuesta.Mensaje)

        Return respuesta
    End Function

    Public Function InsertarLogIncidenciaCambioProducto(ByVal contrato As String, _
                                                    ByVal codOficina As String, _
                                                    ByVal codUsuario As String, _
                                                    ByVal sistema As String, _
                                                    ByVal subSistema As String, _
                                                    ByVal codRespuesta As String, _
                                                    ByVal desRespuesta As String) As String
        Dim estado As String = "0"
        Try
            estado = DAOfertaCambioProducto.Instancia.InsertarLogIncidenciaCambioProducto(contrato, codOficina, codUsuario, sistema, subSistema, codRespuesta, desRespuesta)
        Catch ex As Exception
            ErrorLog(ex.Message)
            estado = "0"
        End Try
        Return estado
    End Function

    Public Function ActualizarCambioProducto(ByVal contrato As String, _
                                        ByVal codVendedor As Integer, _
                                        ByVal codSucursal As Integer, _
                                        ByVal numeroCaja As String, _
                                        ByVal numeroTicket As Integer) As String
        Dim estado As String = "0"
        Try
            estado = DAOfertaCambioProducto.Instancia.ActualizarCambioProducto(contrato, codVendedor, codSucursal, numeroCaja, numeroTicket)
        Catch ex As Exception
            ErrorLog(ex.Message)
            estado = "0"
        End Try
        Return estado
    End Function

End Class
