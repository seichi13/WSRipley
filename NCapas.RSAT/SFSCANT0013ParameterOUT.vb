Public Class SFSCANT0013ParameterOUT

    Private _LNK_ESTADO As String
    Public Property LNK_ESTADO() As String
        Get
            Return _LNK_ESTADO
        End Get
        Set(ByVal value As String)
            _LNK_ESTADO = value
        End Set
    End Property

    Private _LNK_TABLA As String
    Public Property LNK_TABLA() As String
        Get
            Return _LNK_TABLA
        End Get
        Set(ByVal value As String)
            _LNK_TABLA = value
        End Set
    End Property

    Sub New(ByVal tramaRespuesta As String)

        LNK_ESTADO = tramaRespuesta.Substring(101, 1)
        LNK_TABLA = tramaRespuesta.Substring(102, tramaRespuesta.Length - 102)

    End Sub

    Sub New()
    End Sub

    Public Overrides Function ToString() As String

        Dim tramaParameterOUT As String = String.Empty

        tramaParameterOUT &= LNK_ESTADO
        tramaParameterOUT &= LNK_TABLA

        Return tramaParameterOUT

    End Function

End Class
