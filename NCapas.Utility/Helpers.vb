Imports System.Collections.Generic
Imports System


Public NotInheritable Class Helpers
    ''' <summary>
    ''' Transforma un cadena en formato fecha, a un objeto DateTime de acuerdo a un formato de fecha especifico
    ''' </summary>
    ''' <param name="format">Formato de fecha</param>
    ''' <param name="valor">Fecha a transformar</param>
    ''' <returns>Fecha equivalente en DateTime</returns>
    ''' <remarks>Ejemplos de formatos: "yyyy-MM-dd", "dd/MM/yyyy"</remarks>
    Public Shared Function GetDatetimeFromString(ByVal format As String, ByVal valor As String) As DateTime
        If String.IsNullOrEmpty(format) Or String.IsNullOrEmpty(valor) Then
            Return Nothing
        Else
            Dim fecha As DateTime

            If (DateTime.TryParseExact(valor, format, System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, fecha)) Then
                Return fecha
            Else
                Return Nothing
            End If
        End If
    End Function

    ''' <summary>
    ''' Transforma una fecha, agregandole una determinada cantidad de meses, y asignandole un dia especifico
    ''' </summary>
    ''' <param name="addMes">Cantidad de mese a agregar</param>
    ''' <param name="dia">Dia del mes a establecer</param>
    ''' <param name="fecha">Fecha base que será transformada</param>
    ''' <returns>Fecha transformada</returns>
    ''' <remarks>Si la fecha es 10/01/2012 e invoco a esta funcion con addMes: 1 y dia: 27, el valor de retorno será 27/02/2012. Para el caso que el dia sea mayor a 29 y el mes retorno este en febrero, éste retornará el dia 28</remarks>
    Public Shared Function DateAddMesFecha(ByVal addMes As Integer, ByVal dia As Integer, ByVal fecha As DateTime) As DateTime
        Dim fechaAux As DateTime
        Dim fechaRes As DateTime

        fechaAux = DateAdd(DateInterval.Month, addMes, fecha)

        Try
            fechaRes = New DateTime(fechaAux.Year, fechaAux.Month, IIf(dia > 28, 30, dia))
        Catch ex As Exception
            fechaRes = New DateTime(fechaAux.Year, fechaAux.Month, 28)
        End Try

        Return fechaRes
    End Function
End Class