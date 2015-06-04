Imports Microsoft.VisualBasic
Imports System.Collections.Generic

Public Class Salida
    Private _code As String
    Private _mensaje As String
    Private _mensajeImpresion As String
    Private _promocion As String
    Private _existe As Integer
    Private _nroTarjeta As String
    Private _nroDocumento As String
    Private _nombreCliente As String
    Private _lineaInicial As Double
    Private _lineaEgm As Double
    Private _lineaCons As Double
    Private _lineaEfex As Double
    Private _fechaInicioVigencia As String
    Private _fechaFinVigencia As String
    Private _ok As Double
    Private _codigoAtencion As Integer

    Public Property Existe() As Integer
        Get
            Return _existe
        End Get
        Set(value As Integer)
            _existe = value
        End Set
    End Property

    Public Property Promocion() As String
        Get
            Return _promocion
        End Get
        Set(value As String)
            _promocion = value
        End Set
    End Property

    Public Property NroTarjeta() As String
        Get
            Return _nroTarjeta
        End Get
        Set(value As String)
            _nroTarjeta = value
        End Set
    End Property

    Public Property NroDocumento() As String
        Get
            Return _nroDocumento
        End Get
        Set(value As String)
            _nroDocumento = value
        End Set
    End Property

    Public Property NombreCliente() As String
        Get
            Return _nombreCliente
        End Get
        Set(value As String)
            _nombreCliente = value
        End Set
    End Property

    Public Property LineaInicial() As Double
        Get
            Return _lineaInicial
        End Get
        Set(value As Double)
            _lineaInicial = value
        End Set
    End Property

    Public Property LineaEgm() As Double
        Get
            Return _lineaEgm
        End Get
        Set(value As Double)
            _lineaEgm = value
        End Set
    End Property

    Public Property LineaCons() As Double
        Get
            Return _lineaCons
        End Get
        Set(value As Double)
            _lineaCons = value
        End Set
    End Property

    Public Property LineaEfex() As Double
        Get
            Return _lineaEfex
        End Get
        Set(value As Double)
            _lineaEfex = value
        End Set
    End Property

    Public Property FechaInicioVigencia() As String
        Get
            Return _fechaInicioVigencia
        End Get
        Set(value As String)
            _fechaInicioVigencia = value
        End Set
    End Property

    Public Property FechaFinVigencia() As String
        Get
            Return _fechaFinVigencia
        End Get
        Set(value As String)
            _fechaFinVigencia = value
        End Set
    End Property

    Public Property Ok() As Double
        Get
            Return _ok
        End Get
        Set(value As Double)
            _ok = value
        End Set
    End Property

    Public Property CodigoAtencion() As Integer
        Get
            Return _codigoAtencion
        End Get
        Set(value As Integer)
            _codigoAtencion = value
        End Set
    End Property

    Public Property Code() As String
        Get
            Return _code
        End Get
        Set(value As String)
            _code = value
        End Set
    End Property

    Public Property Mensaje() As String
        Get
            Return _mensaje
        End Get
        Set(value As String)
            _mensaje = value
        End Set
    End Property

    Public Property MensajeImpresion() As String
        Get
            Return _mensajeImpresion
        End Get
        Set(value As String)
            _mensajeImpresion = value
        End Set
    End Property
End Class
