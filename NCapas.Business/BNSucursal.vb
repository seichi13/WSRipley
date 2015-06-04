Imports NCapas.Entity
Imports NCapas.Data
Imports NCapas.Utility.Log

Public Class BNSucursal
    Inherits Singleton(Of BNSucursal)
    Public Function BuscarSucursalPorCodigo(ByVal sCodigoKiosko As String) As Sucursal
        Dim sucursal As Sucursal
        Try
            sucursal = DASucursal.Instancia.BuscarSucursalPorCodigo(sCodigoKiosko)
        Catch ex As Exception
            sucursal = Nothing
        End Try
        Return sucursal
    End Function

    Public Function BuscarSucursalPorId(ByVal sIDSucursalKiosco As String) As Sucursal
        Dim sucursal As Sucursal
        Try
            sucursal = DASucursal.Instancia.BuscarSucursalPorId(sIDSucursalKiosco)
        Catch ex As Exception
            sucursal = Nothing
        End Try
        Return sucursal
    End Function

    Public Function ObtenerSucursalBanco(ByVal codigoKiosko As String) As Sucursal
        Dim sucursal As Sucursal
        Try
            sucursal = DASucursal.Instancia.ObtenerSucursalBanco(codigoKiosko)
            ErrorLog("Entro a BN  ObtenerSucursalBanco(" & codigoKiosko & ")")
        Catch ex As Exception
            sucursal = New Sucursal
        End Try
        Return sucursal
    End Function
End Class
