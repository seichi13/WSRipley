Public Class SFSCANT0013Header

    Private _LNK_COD_RETORNO As String
    Public Property LNK_COD_RETORNO() As String
        Get
            Return _LNK_COD_RETORNO
        End Get
        Set(ByVal value As String)
            _LNK_COD_RETORNO = value
        End Set
    End Property

    Private _LNK_NOMBRE_SERVICIO As String
    Public Property LNK_NOMBRE_SERVICIO() As String
        Get
            Return _LNK_NOMBRE_SERVICIO
        End Get
        Set(ByVal value As String)
            _LNK_NOMBRE_SERVICIO = value
        End Set
    End Property

    Private _LNK_LARGO_MENSAJE As String
    Public Property LNK_LARGO_MENSAJE() As String
        Get
            Return _LNK_LARGO_MENSAJE
        End Get
        Set(ByVal value As String)
            _LNK_LARGO_MENSAJE = value
        End Set
    End Property

    Public Overrides Function ToString() As String

        Dim tramaHeader As String = String.Empty

        tramaHeader &= LNK_COD_RETORNO
        tramaHeader &= LNK_NOMBRE_SERVICIO
        tramaHeader &= LNK_LARGO_MENSAJE

        Return tramaHeader

    End Function


End Class
