Public Class ConsultaLPDP
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

    Private _tipo_Doc As String
    Public Property Tipo_Doc() As String
        Get
            Return _tipo_Doc
        End Get
        Set(ByVal value As String)
            _tipo_Doc = value
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

    Private _fecha_Hora As DateTime
    Public Property Fecha_Hora() As DateTime
        Get
            Return _fecha_Hora
        End Get
        Set(ByVal value As DateTime)
            _fecha_Hora = value
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


    Private _terminal_Banco As String
    Public Property Terminal_Banco() As String
        Get
            Return _terminal_Banco
        End Get
        Set(ByVal value As String)
            _terminal_Banco = value
        End Set
    End Property

    Private _numero_Cuenta As String
    Public Property Numero_Cuenta() As String
        Get
            Return _numero_Cuenta
        End Get
        Set(ByVal value As String)
            _numero_Cuenta = value
        End Set
    End Property

    Private _fecha_Aceptacion As DateTime
    Public Property Fecha_Aceptacion() As DateTime
        Get
            Return _fecha_Aceptacion
        End Get
        Set(ByVal value As DateTime)
            _fecha_Aceptacion = value
        End Set
    End Property

    Private _id_Sucursal As Integer
    Public Property Id_Sucursal() As Integer
        Get
            Return _id_Sucursal
        End Get
        Set(ByVal value As Integer)
            _id_Sucursal = value
        End Set
    End Property


End Class
