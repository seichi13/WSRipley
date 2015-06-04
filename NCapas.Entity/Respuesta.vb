Imports Microsoft.VisualBasic

Public Class Respuesta
    Private _estado As String
    Public Property Estado() As String
        Get
            Return _estado
        End Get
        Set(ByVal value As String)
            _estado = value
        End Set
    End Property

    Private _cadena As String
    Public Property Cadena() As String
        Get
            Return _cadena
        End Get
        Set(ByVal value As String)
            _cadena = value
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

    Private _codigo As String
    Public Property Codigo() As String
        Get
            Return _codigo
        End Get
        Set(ByVal value As String)
            _codigo = value
        End Set
    End Property

    Private _EsTitular As String
    Public Property EsTitular() As String
        Get
            Return _EsTitular
        End Get
        Set(ByVal value As String)
            _EsTitular = value
        End Set
    End Property


End Class
