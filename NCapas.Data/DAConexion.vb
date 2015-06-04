Imports NCapas.Entity
Imports System.Configuration
Imports NCapas.Utility.Log

Public Class DAConexion
    Inherits Singleton(Of DAConexion)

    Public Function ReadAppConfig(ByVal llave As String) As String
        Dim reader As New AppSettingsReader()
        Return reader.GetValue(llave, Type.GetType("System.String")).ToString
    End Function

    Public Function Get_CadenaConexion() As String
        Dim sCadena As String = ""
        Try
            sCadena = "Data Source=" & ReadAppConfig("IPSERVER_SQL") & ";Initial Catalog=" & ReadAppConfig("BD_SQL") & ";User ID=" & ReadAppConfig("USER_SQL") & ";PASSWORD=" & ReadAppConfig("PASSWORD_SQL") & ";MultipleActiveResultSets=True"
        Catch ex As Exception
            sCadena = ""
        End Try
        Return sCadena
    End Function

    Public Function SQL_ConnectionOpen(ByVal sCadenaConexion As String, ByRef msgError As String) As SqlClient.SqlConnection
        Dim cnnConnection As SqlClient.SqlConnection
        Try
            cnnConnection = New SqlClient.SqlConnection
            cnnConnection.ConnectionString = sCadenaConexion
            cnnConnection.Open()
        Catch ex As Exception
            msgError = "La conexión a la base de datos No fue satisfactoria, verificar (configuración del web.config)."
            cnnConnection = Nothing
            SQL_ConnectionOpen = Nothing
            Exit Function
        End Try

        Return cnnConnection
    End Function

    Public Function PRUEBA_CONEXION_SQL() As Boolean
        Dim result As Boolean
        Dim sMensajeError_SQL As String = ""
        Using oConexion As SqlClient.SqlConnection = SQL_ConnectionOpen(Get_CadenaConexion(), sMensajeError_SQL)
            If sMensajeError_SQL <> "" Then
                result = False
            Else
                result = True
            End If
        End Using
        Return result
    End Function

    Public Function OBTENER_VARIABLES_TIMEOUT() As String
        Dim sMensajeError_SQL As String = ""
        Dim result As String = ""
            Using oConexion As SqlClient.SqlConnection = SQL_ConnectionOpen(Get_CadenaConexion(), sMensajeError_SQL)
                If sMensajeError_SQL <> "" Then
                    result = "ERROR:No se pudo conectar al servidor de base de datos."
                Else
                    Using cmdx As SqlClient.SqlCommand = oConexion.CreateCommand
                    Dim m_ssql As String = "SELECT TIEMPO_DOC,TIEMPO_OPCIONES,NRO_ERROR_TARJETA, TIEMPO_OFERTAS FROM TIEMPO"
                        cmdx.CommandTimeout = 900
                        cmdx.CommandType = CommandType.Text
                        cmdx.CommandText = m_ssql
                        Using rd_time As SqlClient.SqlDataReader = cmdx.ExecuteReader
                            If rd_time.Read = True Then
                            result = IIf(IsDBNull(rd_time.Item("TIEMPO_DOC")), "0", rd_time.Item("TIEMPO_DOC").ToString) & "|\t|" & IIf(IsDBNull(rd_time.Item("TIEMPO_OPCIONES")), "0", rd_time.Item("TIEMPO_OPCIONES").ToString) & "|\t|" & IIf(IsDBNull(rd_time.Item("NRO_ERROR_TARJETA")), "0", rd_time.Item("NRO_ERROR_TARJETA").ToString) & "|\t|" & IIf(IsDBNull(rd_time.Item("TIEMPO_OFERTAS")), "0", rd_time.Item("TIEMPO_OFERTAS").ToString)
                            Else
                            result = "5" & "|\t|" & "10" & "|\t|" & "5" & "|\t|" & "30"
                            End If
                        End Using
                    End Using
                End If
            End Using
        Return result
    End Function

    Public Function Get_CadenaConexionOracle() As String

        Dim sCadena As String = ""
        Try
            'sCadena = "User ID=" & ReadAppConfig("USER_NAME_ORA") & ";Data Source=" + ReadAppConfig("SERVICE_NAME_ORA") & ";Password=" & ReadAppConfig("PASSWORD_ORA")
            sCadena = "User ID=" & ReadAppConfig("USER_NAME_ORA_SEF") & ";Data Source=" + ReadAppConfig("SERVICE_NAME_ORA_SEF") & ";Password=" & ReadAppConfig("PASSWORD_ORA_SEF")
        Catch ex As Exception
            sCadena = ""
        End Try
        Return sCadena

    End Function

    Public Function Get_CadenaConexionOracleFINX7Q2() As String

        Dim sCadena As String = ""
        Try
            'sCadena = "User ID=" & ReadAppConfig("USER_NAME_ORA") & ";Data Source=" + ReadAppConfig("SERVICE_NAME_ORA") & ";Password=" & ReadAppConfig("PASSWORD_ORA")
            sCadena = "User ID=" & ReadAppConfig("USER_NAME_ORA_SEF_FISA") & ";Data Source=" + ReadAppConfig("SERVICE_NAME_ORA_SEF_FISA") & ";Password=" & ReadAppConfig("PASSWORD_ORA_SEF_FISA")
        Catch ex As Exception
            sCadena = ""
        End Try
        Return sCadena

    End Function

    Public Function Get_CadenaConexionOracleBRP_TEST() As String

        Dim sCadena As String = ""
        Try
            sCadena = "User ID=" & ReadAppConfig("USER_NAME_BRP_FISA") & ";Data Source=" + ReadAppConfig("SERVICE_BRP_FISA") & ";Password=" & ReadAppConfig("PASSWORD_BRP_FISA")
        Catch ex As Exception
            sCadena = ""
        End Try
        Return sCadena

    End Function

End Class
