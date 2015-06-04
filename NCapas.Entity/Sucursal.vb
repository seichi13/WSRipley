Imports Microsoft.VisualBasic
Imports System.Collections.Generic

Public Class Sucursal
    Private _id As Integer
    Public Property ID() As Integer
        Get
            Return _id
        End Get
        Set(ByVal value As Integer)
            _id = value
        End Set
    End Property

    Private _ciudad As String
    Public Property Ciudad() As String
        Get
            Return _ciudad
        End Get
        Set(ByVal value As String)
            _ciudad = value
        End Set
    End Property

    Private _sucursal As String
    Public Property Sucursal() As String
        Get
            Return _sucursal
        End Get
        Set(ByVal value As String)
            _sucursal = value
        End Set
    End Property

    Private _idUbigeo As Integer
    Public Property IdUbigeo() As Integer
        Get
            Return _idUbigeo
        End Get
        Set(ByVal value As Integer)
            _idUbigeo = value
        End Set
    End Property

    Private _estTienda As Integer
    Public Property EstTienda() As Integer
        Get
            Return _estTienda
        End Get
        Set(ByVal value As Integer)
            _estTienda = value
        End Set
    End Property

    Private _direccion As String
    Public Property Direccion() As String
        Get
            Return _direccion
        End Get
        Set(ByVal value As String)
            _direccion = value
        End Set
    End Property

    Private _hini_com1 As String
    Public Property Hini_com1() As String
        Get
            Return _hini_com1
        End Get
        Set(ByVal value As String)
            _hini_com1 = value
        End Set
    End Property

    Private _hfin_com1 As String
    Public Property Hfin_com1() As String
        Get
            Return _hfin_com1
        End Get
        Set(ByVal value As String)
            _hfin_com1 = value
        End Set
    End Property

    Private _hini_com2 As String
    Public Property Hini_com2() As String
        Get
            Return _hini_com2
        End Get
        Set(ByVal value As String)
            _hini_com2 = value
        End Set
    End Property

    Private _hfin_com2 As String
    Public Property Hfin_com2() As String
        Get
            Return _hfin_com2
        End Get
        Set(ByVal value As String)
            _hfin_com2 = value
        End Set
    End Property

    Private _cod_sucursal_banco As String
    Public Property Cod_Sucursal_Banco() As String
        Get
            Return _cod_sucursal_banco
        End Get
        Set(ByVal value As String)
            _cod_sucursal_banco = value
        End Set
    End Property

    Private _hini_cli As String
    Public Property Hini_Cli() As String
        Get
            Return _hini_cli
        End Get
        Set(ByVal value As String)
            _hini_cli = value
        End Set
    End Property

    Private _hfin_cli As String
    Public Property Hfin_Cli() As String
        Get
            Return _hfin_cli
        End Get
        Set(ByVal value As String)
            _hfin_cli = value
        End Set
    End Property

    Private _banco As String
    Public Property Banco() As String
        Get
            Return _banco
        End Get
        Set(ByVal value As String)
            _banco = value
        End Set
    End Property

End Class
