Imports NCapas.Entity
Public Class DAKiosco
    Inherits Singleton(Of DAKiosco)

    Public Function BuscarKioscoPorNombre(ByVal sNombreKiosco As String) As Kiosco
        Dim kiosco As Kiosco
        Dim sMensajeError_SQL As String = ""
        Using oConexion = DAConexion.Instancia.SQL_ConnectionOpen(DAConexion.Instancia.Get_CadenaConexion(), sMensajeError_SQL)
            If oConexion.State = ConnectionState.Open Then
                Dim m_ssql As String = "SELECT IDKIOSKO FROM KIO_NOMBRE_KIOSKO WHERE NOMBRE_KIOSKO='" & sNombreKiosco.Trim & "'"

                Using cmd_kio As SqlClient.SqlCommand = oConexion.CreateCommand
                    cmd_kio.CommandTimeout = 600
                    cmd_kio.CommandType = CommandType.Text
                    cmd_kio.CommandText = m_ssql.Trim

                    Using rd_cod As SqlClient.SqlDataReader = cmd_kio.ExecuteReader
                        kiosco = New Kiosco

                        If rd_cod.Read = True Then
                            'kiosco.IdKiosco = rd_cod.GetInt32(rd_cod.GetOrdinal("IDKIOSKO"))
                            kiosco.IdKiosco = IIf(IsDBNull(rd_cod.Item("IDKIOSKO")), 0, rd_cod.Item("IDKIOSKO"))
                        Else
                            kiosco.IdKiosco = 0
                        End If
                    End Using
                End Using
            Else
                kiosco = Nothing
            End If
        End Using
        Return kiosco
    End Function

End Class
