Imports TestVB.Tipos


Public Class Funciones

    Public Shared Function fx_Completar_Campo(ByVal pCaracter As String, _
                                    ByVal pLongitud As Integer, _
                                    ByVal pTexto As String, _
                                    ByVal pAlinear As TYPE_ALINEAR)

        Dim Resultado As String = String.Empty

        Select Case pAlinear

            Case TYPE_ALINEAR.DERECHA

                Resultado = New String(pCaracter, pLongitud - pTexto.Length) & pTexto

            Case TYPE_ALINEAR.IZQUIERDA

                Resultado = pTexto & New String(pCaracter, pLongitud - pTexto.Length)

        End Select

        Return Resultado


    End Function

    Public Shared Function fx_Preparar_Decimal(ByVal textoDec As String) As String
        Dim result As String = String.Empty
        textoDec = textoDec.Replace(".", ",")
        Dim arr As String() = Split(textoDec, ",", , CompareMethod.Text)
        If arr.Length = 1 Then
            result = textoDec + "00"
        Else
            result = arr(0) + arr(1) + New String("0", 2 - arr(1).Length)
        End If
        Return result
    End Function
End Class
