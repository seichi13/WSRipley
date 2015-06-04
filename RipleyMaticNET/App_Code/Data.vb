Imports Microsoft.VisualBasic
Imports MQCOMLib
Imports NCapas.Utility.Log

Public Class Data

    Private _prioridad_clasica As String = String.Empty
    Private _prioridad_asociada As String = String.Empty
    Private _strServicio As String = String.Empty
    Private _strMensaje As String = String.Empty
    Private _strRespuesta As String = String.Empty

    Private _TablaBloqueo As String(,)
    Private _TipoFactura As String(,)
    Private _OrdenCore As String(,)

    Public Sub New()
        sg_RSAT_IniciarRequerimientos()
    End Sub

    Public ReadOnly Property prioridad_clasica() As String
        Get
            Return _prioridad_clasica
        End Get
    End Property

    Public ReadOnly Property prioridad_asociada() As String
        Get
            Return _prioridad_asociada
        End Get
    End Property

    Public ReadOnly Property TablaBloqueo() As String(,)
        Get
            Return _TablaBloqueo
        End Get
    End Property

    Public ReadOnly Property TipoFactura() As String(,)
        Get
            Return _TipoFactura
        End Get
    End Property

    Public ReadOnly Property OrdenCore() As String(,)
        Get
            Return _OrdenCore
        End Get
    End Property

    Private Function sg_RSAT_CargaOrdenCore(ByVal strRespOrdCore As String) As String(,)
        Dim vls_Respuesta As String
        Dim vli_Pos, vli_I, vli_j As Integer
        Dim vgi_CntCore As Integer
        Dim vg_ArrOrdenCore(,) As String
        Dim vg_RSAT_NumArrCore As Integer

        vls_Respuesta = strRespOrdCore
        vgi_CntCore = Val(Mid(vls_Respuesta, 1, 4)) / 62
        vg_RSAT_NumArrCore = vgi_CntCore
        vls_Respuesta = Trim(Mid(vls_Respuesta, 5, Len(vls_Respuesta)))

        ReDim vg_ArrOrdenCore(vgi_CntCore - 1, 3)

        If Len(Trim(vls_Respuesta)) > 0 Then
            vli_Pos = 1
            For vli_I = 0 To vgi_CntCore - 1
                For vli_j = 0 To 3
                    If vli_j = 0 Or vli_j = 1 Then
                        vg_ArrOrdenCore(vli_I, vli_j) = CStr(Val(Mid(vls_Respuesta, vli_Pos, 20)))
                        vli_Pos = vli_Pos + 20
                    Else
                        If vli_j = 2 Then
                            vg_ArrOrdenCore(vli_I, vli_j) = New String("0", 2 - Len(CStr(Val(Mid(vls_Respuesta, vli_Pos, 20))))) & CStr(Val(Mid(vls_Respuesta, vli_Pos, 20)))
                            vli_Pos = vli_Pos + 20
                        Else
                            vg_ArrOrdenCore(vli_I, vli_j) = New String("0", 2 - Len(CStr(Val(Mid(vls_Respuesta, vli_Pos, 2))))) & CStr(Val(Mid(vls_Respuesta, vli_Pos, 2)))
                            vli_Pos = vli_Pos + 2
                        End If
                    End If

                Next
            Next
        End If

        Return vg_ArrOrdenCore
    End Function

    Private Function sg_RSAT_CargaTablaBloqueos(ByVal strRespTabBloq As String) As String(,)
        Dim vls_Respuesta As String
        Dim vli_Pos, vli_I As Integer
        Dim vgi_CntCore As Integer
        Dim vg_RSAT_NumArrError As String
        Dim vg_ArrTablaError(,) As String

        vls_Respuesta = strRespTabBloq

        vgi_CntCore = Val(Mid(vls_Respuesta, 1, 4)) / 34

        vg_RSAT_NumArrError = vgi_CntCore
        vls_Respuesta = Trim(Mid(vls_Respuesta, 5, Len(vls_Respuesta)))

        ReDim vg_ArrTablaError(vgi_CntCore - 1, 3)

        If Len(Trim(vls_Respuesta)) > 0 Then
            vli_Pos = 1
            For vli_I = 0 To vgi_CntCore - 1
                vg_ArrTablaError(vli_I, 0) = Mid(vls_Respuesta, vli_Pos, 2)
                vli_Pos = vli_Pos + 2
                vg_ArrTablaError(vli_I, 1) = Mid(vls_Respuesta, vli_Pos, 30)
                vli_Pos = vli_Pos + 30
                vg_ArrTablaError(vli_I, 2) = Mid(vls_Respuesta, vli_Pos, 1)
                vli_Pos = vli_Pos + 1
                vg_ArrTablaError(vli_I, 3) = Mid(vls_Respuesta, vli_Pos, 1)
                vli_Pos = vli_Pos + 1
            Next
        End If

        Return vg_ArrTablaError
    End Function

    Private Function sg_RSAT_CargaTipFactura(ByVal strDatos As String) As String(,)
        Dim vtf_Respuesta As String
        Dim vtf_Pos, vtf_I As Integer
        Dim vtf_CntCore As Integer
        Dim vg_RSAT_TipArrError As String
        Dim vg_ArrTipFacError(,) As String

        vtf_Respuesta = strDatos

        vtf_CntCore = Val(Mid(vtf_Respuesta, 1, 4)) / 23
        '
        vtf_Respuesta = Mid(vtf_Respuesta, 5)
        '
        vg_RSAT_TipArrError = vtf_CntCore
        vtf_Respuesta = Trim(Mid(vtf_Respuesta, 1, Len(vtf_Respuesta)))

        ReDim vg_ArrTipFacError(vtf_CntCore - 1, 1)

        If Len(Trim(vtf_Respuesta)) > 0 Then
            vtf_Pos = 16
            For vtf_I = 0 To vtf_CntCore - 1
                vg_ArrTipFacError(vtf_I, 0) = Mid(vtf_Respuesta, vtf_Pos, 8)
                vtf_Pos = vtf_Pos + 23
            Next
        End If

        Return vg_ArrTipFacError

    End Function

    Private Sub sg_RSAT_Obtener_PrioridadTarjetas()


        Dim vg_ArrOrdenCore(,) As String
        vg_ArrOrdenCore = _OrdenCore

        Dim xMax As Integer = vg_ArrOrdenCore.GetUpperBound(0)
        Dim v_prioridad_tarjeta As String = String.Empty
        Dim v_binn_tarjeta As String = String.Empty

        _prioridad_clasica = String.Empty
        _prioridad_asociada = String.Empty

        For x As Integer = 0 To xMax

            v_prioridad_tarjeta = vg_ArrOrdenCore(x, 2)
            v_binn_tarjeta = vg_ArrOrdenCore(x, 0)

            'SICRON = 02 | las tarjetas asociadas apuntan por defecto a SICRON
            'RSAT   = 01

            If v_prioridad_tarjeta <> "01" And v_prioridad_tarjeta <> "02" Then
                v_prioridad_tarjeta = "02"
            End If

            Select Case v_binn_tarjeta
                Case "960410"
                    If _prioridad_clasica = String.Empty Then
                        _prioridad_clasica = v_prioridad_tarjeta
                    End If

                Case Else
                    If _prioridad_asociada = String.Empty Then
                        _prioridad_asociada = v_prioridad_tarjeta
                    End If
            End Select

        Next

    End Sub

    Public Sub sg_RSAT_IniciarRequerimientos()
        ErrorLog("Inicia sg_RSAT_IniciarRequerimientos")
        Try
            Dim objMQ As New MQ

            _strServicio = "SFSCANC0017"
            _strMensaje = "0000000000SFSCANC0017                                       000000104PE00010000020027SANIUSUARIO19900040000050005010001000001010004000001040004000001050004000001070004000005"
            objMQ.Service = _strServicio
            objMQ.Message = _strMensaje
            ErrorLog("objMQ.Service = " & _strServicio)
            ErrorLog("objMQ.Message = " & _strMensaje)
            ErrorLog("INICIA objMQ.Execute()")
            objMQ.Execute()
            ErrorLog("TERMINA objMQ.Execute()")
            ErrorLog("objMQ.Response = " & objMQ.Response)
            If objMQ.ReasonMd <> 0 Then
                _strRespuesta = "ERROR"
            End If

            If objMQ.ReasonApp = 0 Then
                _strRespuesta = objMQ.Response
            End If

            _strRespuesta = Mid(_strRespuesta, 112)

            Dim intLon As Long = 0
            Dim intPos As Integer = 1
            Dim strRespOrdCore As String = String.Empty
            Dim strRespITF As String = String.Empty
            Dim strRespPreced As String = String.Empty
            Dim strRespTabBloq As String = String.Empty
            Dim strRespTipFact As String = String.Empty

            'ITF
            intLon = Val(Mid(_strRespuesta, 7, 4))
            strRespITF = Mid(_strRespuesta, 11, intLon)
            'PRECEDENCIA 
            intPos = 11 + intLon
            intLon = Val(Mid(_strRespuesta, intPos + 2, 4))
            strRespPreced = Mid(_strRespuesta, intPos + 6, intLon)

            'ORDEN DEL CORE        
            intPos = intPos + 6 + intLon
            intLon = Val(Mid(_strRespuesta, intPos + 2, 4))
            strRespOrdCore = Mid(_strRespuesta, intPos + 2, intLon)

            'TABLA DE BLOQUEOS
            intPos = intPos + 6 + intLon
            intLon = Val(Mid(_strRespuesta, intPos + 2, 4))
            strRespTabBloq = Mid(_strRespuesta, intPos + 2, intLon + 4)

            'TIPO DE FACTURA
            intPos = intPos + 6 + intLon
            intLon = Val(Mid(_strRespuesta, intPos + 2, 4))
            strRespTipFact = Mid(_strRespuesta, intPos + 2, intLon)

            'Se carga tabla de Tipos de Factura
            _TipoFactura = sg_RSAT_CargaTipFactura(strRespTipFact)
            'Se carga tabla de codigos de bloqueos
            _TablaBloqueo = sg_RSAT_CargaTablaBloqueos(strRespTabBloq)
            ' Se Carga Tabla de Orden de Acceso al Core
            _OrdenCore = sg_RSAT_CargaOrdenCore(strRespOrdCore)

            sg_RSAT_Obtener_PrioridadTarjetas()
        Catch ex As Exception
            ErrorLog("Catch sg_RSAT_IniciarRequerimientos " & ex.Message)
        End Try

        ErrorLog("Termina sg_RSAT_IniciarRequerimientos")
    End Sub

End Class
