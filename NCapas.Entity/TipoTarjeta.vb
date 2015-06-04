Public Class TipoTarjeta

    Private _id As Integer
    Public Property Id() As Integer
        Get
            Return _id
        End Get
        Set(ByVal value As Integer)
            _id = value
        End Set
    End Property

    Private _nombreTarjeta As String
    Public Property NombreTarjeta() As String
        Get
            Return _nombreTarjeta
        End Get
        Set(ByVal value As String)
            _nombreTarjeta = value
        End Set
    End Property

    Private _tipo As Integer
    Public Property Tipo() As Integer
        Get
            Return _tipo
        End Get
        Set(ByVal value As Integer)
            _tipo = value
        End Set
    End Property


End Class
