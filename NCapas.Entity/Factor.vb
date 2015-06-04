Public Class Factor
    Private _id As Integer
    Public Property Id() As Integer
        Get
            Return _id
        End Get
        Set(ByVal value As Integer)
            _id = value
        End Set
    End Property

    Private _factorPlazo As String
    Public Property FactorPlazo() As String
        Get
            Return _factorPlazo
        End Get
        Set(ByVal value As String)
            _factorPlazo = value
        End Set
    End Property
End Class
