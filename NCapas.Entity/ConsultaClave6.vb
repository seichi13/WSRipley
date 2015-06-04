Public Class ConsultaClave6

    Private _id As Integer
    Public Property Id() As Integer
        Get
            Return _id
        End Get
        Set(ByVal value As Integer)
            _id = value
        End Set
    End Property

    Private _cod_sucursal_ban As String
    Public Property Cod_Sucursal_Ban() As String
        Get
            Return _cod_sucursal_ban
        End Get
        Set(ByVal value As String)
            _cod_sucursal_ban = value
        End Set
    End Property

    Private _nombre_Kiosko As String
    Public Property Nombre_Kiosko() As String
        Get
            Return _nombre_Kiosko
        End Get
        Set(ByVal value As String)
            _nombre_Kiosko = value
        End Set
    End Property

    Private _id_Kiosko As Integer
    Public Property Id_Kiosko() As Integer
        Get
            Return _id_Kiosko
        End Get
        Set(ByVal value As Integer)
            _id_Kiosko = value
        End Set

    End Property

    Private _tipo_Doc As String
    Public Property Tipo_Doc() As String
        Get
            Return _tipo_Doc
        End Get
        Set(ByVal value As String)
            _tipo_Doc = value
        End Set
    End Property

    Private _numero_Tarjeta As String
    Public Property Numero_Tarjeta() As String
        Get
            Return _numero_Tarjeta
        End Get
        Set(ByVal value As String)
            _numero_Tarjeta = value
        End Set
    End Property

    Private _nro_Doc As String
    Public Property Nro_Doc() As String
        Get
            Return _nro_Doc
        End Get
        Set(ByVal value As String)
            _nro_Doc = value
        End Set
    End Property

    Private _opcion As String
    Public Property Opcion() As String
        Get
            Return _opcion
        End Get
        Set(ByVal value As String)
            _opcion = value
        End Set
    End Property

    Private _mensaje_Error As String
    Public Property Mensaje_Error() As String
        Get
            Return _mensaje_Error
        End Get
        Set(ByVal value As String)
            _mensaje_Error = value
        End Set
    End Property

    Private _email As String
    Public Property Email() As String
        Get
            Return _email
        End Get
        Set(ByVal value As String)
            _email = value
        End Set
    End Property

    Private _envio As Boolean
    Public Property Envio() As Boolean
        Get
            Return _envio
        End Get
        Set(ByVal value As Boolean)
            _envio = value
        End Set
    End Property

    Private _contador As Integer
    Public Property Contador() As Integer
        Get
            Return _contador
        End Get
        Set(ByVal value As Integer)
            _contador = value
        End Set
    End Property



End Class
