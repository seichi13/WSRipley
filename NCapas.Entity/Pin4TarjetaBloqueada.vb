Public Class Pin4TarjetaBloqueada
    Public Property Id As Integer
    Public Property NroTarjeta As String
    Public Property EstaBloqueada As Boolean = False
    Public Property NroIntentos As Short = 0
    Public Property FechaBloqueo As DateTime = Nothing
    Public Property FechaActual As DateTime = Nothing
End Class
