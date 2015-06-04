Public Class Response

    Private _message As String
    Public Property Message() As String
        Get
            Return _message
        End Get
        Set(ByVal value As String)
            _message = value
        End Set
    End Property

    Private _success As Boolean
    Public Property Success() As Boolean
        Get
            Return _success
        End Get
        Set(ByVal value As Boolean)
            _success = value
        End Set
    End Property

    Private _warning As Boolean
    Public Property Warning() As Boolean
        Get
            Return _warning
        End Get
        Set(ByVal value As Boolean)
            _warning = value
        End Set
    End Property

    Private _data As Object
    Public Property Data() As Object
        Get
            Return _data
        End Get
        Set(ByVal value As Object)
            _data = value
        End Set
    End Property

End Class
