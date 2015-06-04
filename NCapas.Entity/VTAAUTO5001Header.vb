Public Class VTAAUTO5001Header
    Private _codRetorno As String
    Public Property CodRetorno() As String
        Get
            Return _codRetorno
        End Get
        Set(ByVal value As String)
            _codRetorno = value
        End Set
    End Property

    Private _nombreServicio As String
    Public Property NombreServicio() As String
        Get
            Return _nombreServicio
        End Get
        Set(ByVal value As String)
            _nombreServicio = value
        End Set
    End Property

    Private _largoMensaje As String
    Public Property LargoMensaje() As String
        Get
            Return _largoMensaje
        End Get
        Set(ByVal value As String)
            _largoMensaje = value
        End Set
    End Property

    Private _codEntidad As String
    Public Property CodEntidad() As String
        Get
            Return _codEntidad
        End Get
        Set(ByVal value As String)
            _codEntidad = value
        End Set
    End Property

    Private _tipoTerminal As String
    Public Property TipoTerminal() As String
        Get
            Return _tipoTerminal
        End Get
        Set(ByVal value As String)
            _tipoTerminal = value
        End Set
    End Property

    Private _codUsuario As String
    Public Property CodUsuario() As String
        Get
            Return _codUsuario
        End Get
        Set(ByVal value As String)
            _codUsuario = value
        End Set
    End Property

    Private _codOficina As String
    Public Property CodOficina() As String
        Get
            Return _codOficina
        End Get
        Set(ByVal value As String)
            _codOficina = value
        End Set
    End Property

    Private _terminalFisico As String
    Public Property TerminalFisico() As String
        Get
            Return _terminalFisico
        End Get
        Set(ByVal value As String)
            _terminalFisico = value
        End Set
    End Property

    Private _fechaConta As String
    Public Property FechaConta() As String
        Get
            Return _fechaConta
        End Get
        Set(ByVal value As String)
            _fechaConta = value
        End Set
    End Property

    Private _codAplicacion As String
    Public Property CodAplicacion() As String
        Get
            Return _codAplicacion
        End Get
        Set(ByVal value As String)
            _codAplicacion = value
        End Set
    End Property

    Private _codPais As String
    Public Property CodPais() As String
        Get
            Return _codPais
        End Get
        Set(ByVal value As String)
            _codPais = value
        End Set
    End Property

    Private _commit As String
    Public Property Commit() As String
        Get
            Return _commit
        End Get
        Set(ByVal value As String)
            _commit = value
        End Set
    End Property

    Private _paginacion As String
    Public Property Paginacion() As String
        Get
            Return _paginacion
        End Get
        Set(ByVal value As String)
            _paginacion = value
        End Set
    End Property

    Private _claveInicio As String
    Public Property ClaveInicio() As String
        Get
            Return _claveInicio
        End Get
        Set(ByVal value As String)
            _claveInicio = value
        End Set
    End Property

    Private _claveFin As String
    Public Property ClaveFin() As String
        Get
            Return _claveFin
        End Get
        Set(ByVal value As String)
            _claveFin = value
        End Set
    End Property

    Private _pantalla As String
    Public Property Pantalla() As String
        Get
            Return _pantalla
        End Get
        Set(ByVal value As String)
            _pantalla = value
        End Set
    End Property

    Private _masDatos As String
    Public Property MasDatos() As String
        Get
            Return _masDatos
        End Get
        Set(ByVal value As String)
            _masDatos = value
        End Set
    End Property

    Private _otrosDatos As String
    Public Property OtrosDatos() As String
        Get
            Return _otrosDatos
        End Get
        Set(ByVal value As String)
            _otrosDatos = value
        End Set
    End Property

    Private _mensaje As String
    Public Property Mensaje() As String
        Get
            Return _mensaje
        End Get
        Set(ByVal value As String)
            _mensaje = value
        End Set
    End Property



    Public Overrides Function ToString() As String

        Dim tramaHeader As String = String.Empty

        tramaHeader &= CodRetorno
        tramaHeader &= NombreServicio
        tramaHeader &= LargoMensaje
        tramaHeader &= CodEntidad
        tramaHeader &= TipoTerminal
        tramaHeader &= CodUsuario
        tramaHeader &= CodOficina
        tramaHeader &= TerminalFisico
        tramaHeader &= FechaConta
        tramaHeader &= CodAplicacion
        tramaHeader &= CodPais
        tramaHeader &= Commit
        tramaHeader &= Paginacion
        tramaHeader &= ClaveInicio
        tramaHeader &= ClaveFin
        tramaHeader &= Pantalla
        tramaHeader &= MasDatos
        tramaHeader &= OtrosDatos
        tramaHeader &= Mensaje

        Return tramaHeader

    End Function

End Class
