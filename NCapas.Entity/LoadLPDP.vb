<Serializable()>
Public Class LoadLPDP
    Private _id As Integer
    Public Property Id() As Integer
        Get
            Return _id
        End Get
        Set(ByVal value As Integer)
            _id = value
        End Set
    End Property


    Private _nroDocumento As String
    Public Property NroDocumento() As String
        Get
            Return _nroDocumento
        End Get
        Set(ByVal value As String)
            _nroDocumento = value
        End Set
    End Property
End Class
