Public Class Kiosco
    Private _idKiosco As Integer
    Public Property IdKiosco() As Integer
        Get
            Return _idKiosco
        End Get
        Set(ByVal value As Integer)
            _idKiosco = value
        End Set
    End Property

    Private _nombreKiosco As String
    Public Property NombreKiosco() As String
        Get
            Return _nombreKiosco
        End Get
        Set(ByVal value As String)
            _nombreKiosco = value
        End Set
    End Property

End Class
