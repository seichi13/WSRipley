Imports Microsoft.VisualBasic
Imports System.Collections.Generic

<Serializable()>
Public Class ConfiguracionKiosko
    Private _id As Integer
    Private _nombre As String
    Private _server As String
    Private _server_simulador As String
    Private _server_print As String
    Private _server_com As String
    Private _bin1 As String
    Private _longitud_tarjeta_bin1 As Integer
    Private _posini_bin1 As Integer
    Private _longitud_bin_bin1 As Integer
    Private _bin2 As String
    Private _longitud_tarjeta_bin2 As Integer
    Private _posini_bin2 As Integer
    Private _longitud_bin_bin2 As Integer


    Public Property ID() As Integer
        Get
            Return _id
        End Get
        Set(ByVal value As Integer)
            _id = value
        End Set
    End Property

    Public Property Nombre() As String
        Get
            Return _nombre
        End Get
        Set(ByVal value As String)
            _nombre = value
        End Set
    End Property

    Public Property Server() As String
        Get
            Return _server
        End Get
        Set(ByVal value As String)
            _server = value
        End Set
    End Property

    Public Property Server_Simulador() As String
        Get
            Return _server_simulador
        End Get
        Set(ByVal value As String)
            _server_simulador = value
        End Set
    End Property

    Public Property Server_Print() As String
        Get
            Return _server_print
        End Get
        Set(ByVal value As String)
            _server_print = value
        End Set
    End Property

    Public Property Server_Com() As String
        Get
            Return _server_com
        End Get
        Set(ByVal value As String)
            _server_com = value
        End Set
    End Property

    Public Property Bin1() As String
        Get
            Return _bin1
        End Get
        Set(ByVal value As String)
            _bin1 = value
        End Set
    End Property

    Public Property Longitud_Tarjeta_Bin1() As Integer
        Get
            Return _longitud_tarjeta_bin1
        End Get
        Set(ByVal value As Integer)
            _longitud_tarjeta_bin1 = value
        End Set
    End Property

    Public Property Posini_Bin1() As Integer
        Get
            Return _posini_bin1
        End Get
        Set(ByVal value As Integer)
            _posini_bin1 = value
        End Set
    End Property

    Public Property Longitud_Bin_Bin1() As Integer
        Get
            Return _longitud_bin_bin1
        End Get
        Set(ByVal value As Integer)
            _longitud_bin_bin1 = value
        End Set
    End Property

    Public Property Bin2() As String
        Get
            Return _bin2
        End Get
        Set(ByVal value As String)
            _bin2 = value
        End Set
    End Property

    Public Property Longitud_Tarjeta_Bin2() As Integer
        Get
            Return _longitud_tarjeta_bin2
        End Get
        Set(ByVal value As Integer)
            _longitud_tarjeta_bin2 = value
        End Set
    End Property

    Public Property Posini_Bin2() As Integer
        Get
            Return _posini_bin2
        End Get
        Set(ByVal value As Integer)
            _posini_bin2 = value
        End Set
    End Property

    Public Property Longitud_Bin_Bin2() As Integer
        Get
            Return _longitud_bin_bin2
        End Get
        Set(ByVal value As Integer)
            _longitud_bin_bin2 = value
        End Set
    End Property


    Private _codigo_sucursal As String
    Public Property Codigo_Sucursal() As String
        Get
            Return _codigo_sucursal
        End Get
        Set(ByVal value As String)
            _codigo_sucursal = value
        End Set
    End Property

    Private _codigo_kiosko As String
    Public Property Codigo_Kiosko() As String
        Get
            Return _codigo_kiosko
        End Get
        Set(ByVal value As String)
            _codigo_kiosko = value
        End Set
    End Property

    Private _id_kiosko As String
    Public Property ID_Kiosko() As String
        Get
            Return _id_kiosko
        End Get
        Set(ByVal value As String)
            _id_kiosko = value
        End Set
    End Property

End Class
