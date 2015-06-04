Public Class IncrementoLinea

    Class Oferta

    End Class

    Public Class Aprobacion

        Private Cod_Respuesta_ As String
        Private Msj_Respuesta_ As String
        Private Paso_ As String

        Private Upd_Linea_EGM_OK_ As String
        Private Upd_Linea_1_OK_ As String
        Private Upd_Linea_2_OK_ As String

        Public Property Cod_Respuesta As String
            Get
                Return Cod_Respuesta_
            End Get
            Set(ByVal value As String)
                Cod_Respuesta_ = value
            End Set
        End Property

        Public Property Msj_Respuesta As String
            Get
                Return Msj_Respuesta_
            End Get
            Set(ByVal value As String)
                Msj_Respuesta_ = value
            End Set
        End Property

        Public Property Paso As String
            Get
                Return Paso_
            End Get
            Set(ByVal value As String)
                Paso_ = value
            End Set
        End Property

        Public Property Upd_Linea_EGM_OK As String
            Get
                Return Upd_Linea_EGM_OK_
            End Get
            Set(ByVal value As String)
                Upd_Linea_EGM_OK_ = value
            End Set
        End Property

        Public Property Upd_Linea_1_OK As String
            Get
                Return Upd_Linea_1_OK_
            End Get
            Set(ByVal value As String)
                Upd_Linea_1_OK_ = value
            End Set
        End Property

        Public Property Upd_Linea_2_OK As String
            Get
                Return Upd_Linea_2_OK_
            End Get
            Set(ByVal value As String)
                Upd_Linea_2_OK_ = value
            End Set
        End Property

    End Class

    Public Class Parametros

        Private Cod_Usuario_Plataforma_ As String
        Private Cod_Oficina_ As String
        Private Fecha_Conta_ As Date
        Private Nro_Contrato_ As String
        Private Linea_EGM_ As String
        Private Linea_1_ As String
        Private Linea_2_ As String
        Private Cod_Sucursal_ As String
        Private Nro_Caja_ As String
        Private Nro_Transaccion_ As String

        Public Property Cod_Usuario_Plataforma As String
            Get
                Return Cod_Usuario_Plataforma_
            End Get
            Set(ByVal value As String)
                Cod_Usuario_Plataforma_ = value
            End Set
        End Property

        Public Property Cod_Oficina As String
            Get
                Return Cod_Oficina_
            End Get
            Set(ByVal value As String)
                Cod_Oficina_ = value
            End Set
        End Property

        Public Property Fecha_Conta As Date
            Get
                Return Fecha_Conta_
            End Get
            Set(ByVal value As Date)
                Fecha_Conta_ = value
            End Set
        End Property


        Public Property Nro_Contrato As String
            Get
                Return Nro_Contrato_
            End Get
            Set(ByVal value As String)
                Nro_Contrato_ = value
            End Set
        End Property

        Public Property Linea_EGM As String
            Get
                Return Linea_EGM_
            End Get
            Set(ByVal value As String)
                Linea_EGM_ = value
            End Set
        End Property

        Public Property Linea_1 As String
            Get
                Return Linea_1_
            End Get
            Set(ByVal value As String)
                Linea_1_ = value
            End Set
        End Property

        Public Property Linea_2 As String
            Get
                Return Linea_2_
            End Get
            Set(ByVal value As String)
                Linea_2_ = value
            End Set
        End Property

        Public Property Cod_Sucursal As String
            Get
                Return Cod_Sucursal_
            End Get
            Set(ByVal value As String)
                Cod_Sucursal_ = value
            End Set
        End Property

        Public Property Nro_Caja As String
            Get
                Return Nro_Caja_
            End Get
            Set(ByVal value As String)
                Nro_Caja_ = value
            End Set
        End Property

        Public Property Nro_Transaccion As String
            Get
                Return Nro_Transaccion_
            End Get
            Set(ByVal value As String)
                Nro_Transaccion_ = value
            End Set
        End Property

    End Class


End Class
