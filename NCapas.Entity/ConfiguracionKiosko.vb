<Serializable()>
Public Class ConfiguracionKiosko
    Public Property ID As Integer
    Public Property Nombre As String
    Public Property Server As String
    Public Property Server_Simulador As String
    Public Property Server_Print As String
    Public Property Server_Com As String
    Public Property Bin1 As String
    Public Property Longitud_Tarjeta_Bin1 As Integer
    Public Property Posini_Bin1 As Integer
    Public Property Longitud_Bin_Bin1 As Integer
    Public Property Bin2 As String
    Public Property Longitud_Tarjeta_Bin2 As Integer
    Public Property Posini_Bin2 As Integer
    Public Property Longitud_Bin_Bin2 As Integer
    Public Property Pin4Intentos As Short = 0
    Public Property Pin4HorasBloqueo As Short = 0

    Public Property Codigo_Sucursal As String
    Public Property Codigo_Kiosko As String
    Public Property ID_Kiosko As String
End Class
