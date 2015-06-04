Imports Microsoft.VisualBasic
Imports System.Collections.Generic

Public Class OpcionesTarjeta
    Private _nroTarjeta As String
    Private _nombreTarjeta As String
    Private _binn As String
    Private _tipo As String

    Private _tipoTarjetaPasada As String
    Private _metodoWsCall As String


    Public Property NroTarjeta() As String
        Get
            Return _nroTarjeta
        End Get
        Set(value As String)
            _nroTarjeta = value
        End Set
    End Property

    Public Property NombreTarjeta() As String
        Get
            Return _nombreTarjeta
        End Get
        Set(value As String)
            _nombreTarjeta = value
        End Set
    End Property

    Public Property Binn() As Integer
        Get
            Return _binn
        End Get
        Set(value As Integer)
            _binn = value
        End Set
    End Property

    Public Property Tipo() As String
        Get
            Return _tipo
        End Get
        Set(value As String)
            _tipo = value
        End Set
    End Property


    Public Property TipoTarjetaPasada() As String
        Get
            Return _tipoTarjetaPasada
        End Get
        Set(ByVal value As String)
            _tipoTarjetaPasada = value
        End Set
    End Property

    Public Property MetodoWsCall() As String
        Get
            Return _metodoWsCall
        End Get
        Set(ByVal value As String)
            _metodoWsCall = value
        End Set
    End Property

    Private _esTitular As String
    Public Property EsTitular() As String
        Get
            Return _esTitular
        End Get
        Set(ByVal value As String)
            _esTitular = value
        End Set
    End Property

    Private _nroCuenta As String
    Public Property NroCuenta() As String
        Get
            Return _nroCuenta
        End Get
        Set(ByVal value As String)
            _nroCuenta = value
        End Set
    End Property

End Class
