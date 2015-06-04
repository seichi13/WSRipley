Imports NCapas.Entity
Imports System.Data.SqlClient
Imports NCapas.Data.DAConexion
Imports NCapas.Utility.Log
Imports System.Data.OracleClient

Public Class DABRPFISA
    Inherits Singleton(Of DABRPFISA)

    Public Function ObtenerEmailFISA(ByVal numeroDocumento As String) As String
        Dim respuesta As String = ""
        Using oConexion = New OracleConnection(DAConexion.Instancia.Get_CadenaConexionOracleBRP_TEST())
            If Not oConexion Is Nothing Then
                ErrorLog("antes oConexion.Open()")
                oConexion.Open()
                ErrorLog("después oConexion.Open()")
                If oConexion.State = ConnectionState.Open Then
                    ErrorLog("ConnectionState.Open")
                    Using cmd As OracleCommand = oConexion.CreateCommand()

                        cmd.CommandTimeout = 600
                        cmd.CommandType = CommandType.StoredProcedure
                        cmd.CommandText = "PKG_EMAIL_FISA.PRC_OBTENER_EMAIL"

                        Dim parametro1 As New OracleParameter
                        Dim parametro2 As New OracleParameter

                        parametro1.ParameterName = "I_USERNAME"
                        parametro1.OracleType = OracleType.VarChar
                        parametro1.Size = 64
                        parametro1.Direction = ParameterDirection.Input
                        parametro1.Value = numeroDocumento.Trim()
                        cmd.Parameters.Add(parametro1)

                        parametro2.ParameterName = "O_EMAIL"
                        parametro2.OracleType = OracleType.VarChar
                        parametro2.Size = 1000
                        parametro2.Direction = ParameterDirection.Output
                        cmd.Parameters.Add(parametro2)
                        ErrorLog("ANTES cmd.ExecuteReader()")
                        cmd.ExecuteNonQuery()
                        ErrorLog("despues cmd.ExecuteReader()")
                        respuesta = cmd.Parameters("O_EMAIL").Value.ToString()
                        ErrorLog("respuesta=" & respuesta)

                    End Using
                End If
            End If
        End Using
        Return respuesta
    End Function

    Public Function ObtenerEmailPersonal(ByVal tipoDocumento As String, ByVal numeroDocumento As String) As String
        Dim respuesta As String = ""
        Using oConexion = New OracleConnection(DAConexion.Instancia.Get_CadenaConexionOracleBRP_TEST())
            If Not oConexion Is Nothing Then
                ErrorLog("antes oConexion.Open()")
                oConexion.Open()
                ErrorLog("después oConexion.Open()")
                If oConexion.State = ConnectionState.Open Then
                    ErrorLog("ConnectionState.Open")
                    Using cmd As OracleCommand = oConexion.CreateCommand()
                        cmd.CommandTimeout = 600
                        cmd.CommandType = CommandType.StoredProcedure
                        cmd.CommandText = "PKG_DA_CLIENTE.CORREO_PER"

                        Dim parametro1 As New OracleParameter
                        Dim parametro2 As New OracleParameter
                        Dim parametro3 As New OracleParameter

                        parametro1.ParameterName = "pi_tipoid"
                        parametro1.OracleType = OracleType.VarChar
                        parametro1.Size = 64
                        parametro1.Direction = ParameterDirection.Input
                        parametro1.Value = tipoDocumento.Trim()
                        cmd.Parameters.Add(parametro1)

                        parametro2.ParameterName = "pi_numdoc"
                        parametro2.OracleType = OracleType.VarChar
                        parametro2.Size = 64
                        parametro2.Direction = ParameterDirection.Input
                        parametro2.Value = numeroDocumento.Trim()
                        cmd.Parameters.Add(parametro2)

                        parametro3.ParameterName = "O_EMAIL"
                        parametro3.OracleType = OracleType.VarChar
                        parametro3.Size = 1000
                        parametro3.Direction = ParameterDirection.Output
                        cmd.Parameters.Add(parametro3)
                        ErrorLog("ANTES cmd.ExecuteReader()")
                        cmd.ExecuteNonQuery()
                        ErrorLog("despues cmd.ExecuteReader()")
                        respuesta = cmd.Parameters("O_EMAIL").Value.ToString()
                        ErrorLog("respuesta=" & respuesta)

                    End Using
                End If
            End If
        End Using
        Return respuesta
    End Function

    Public Function ObtenerEmailLaboral(ByVal tipoDocumento As String, ByVal numeroDocumento As String) As String
        Dim respuesta As String = ""
        Using oConexion = New OracleConnection(DAConexion.Instancia.Get_CadenaConexionOracleBRP_TEST())
            If Not oConexion Is Nothing Then
                ErrorLog("antes oConexion.Open()")
                oConexion.Open()
                ErrorLog("después oConexion.Open()")
                If oConexion.State = ConnectionState.Open Then
                    ErrorLog("ConnectionState.Open")
                    Using cmd As OracleCommand = oConexion.CreateCommand()
                        cmd.CommandTimeout = 600
                        cmd.CommandType = CommandType.StoredProcedure
                        cmd.CommandText = "PKG_DA_CLIENTE.CORREO_LAB"

                        Dim parametro1 As New OracleParameter
                        Dim parametro2 As New OracleParameter
                        Dim parametro3 As New OracleParameter

                        parametro1.ParameterName = "pi_tipoid"
                        parametro1.OracleType = OracleType.VarChar
                        parametro1.Size = 64
                        parametro1.Direction = ParameterDirection.Input
                        parametro1.Value = tipoDocumento.Trim()
                        cmd.Parameters.Add(parametro1)

                        parametro2.ParameterName = "pi_numdoc"
                        parametro2.OracleType = OracleType.VarChar
                        parametro2.Size = 64
                        parametro2.Direction = ParameterDirection.Input
                        parametro2.Value = numeroDocumento.Trim()
                        cmd.Parameters.Add(parametro2)

                        parametro3.ParameterName = "O_EMAIL"
                        parametro3.OracleType = OracleType.VarChar
                        parametro3.Size = 1000
                        parametro3.Direction = ParameterDirection.Output
                        cmd.Parameters.Add(parametro3)
                        ErrorLog("ANTES cmd.ExecuteReader()")
                        cmd.ExecuteNonQuery()
                        ErrorLog("despues cmd.ExecuteReader()")
                        respuesta = cmd.Parameters("O_EMAIL").Value.ToString()
                        ErrorLog("respuesta=" & respuesta)

                    End Using
                End If
            End If
        End Using
        Return respuesta
    End Function

    Public Function ObtenerOtrosEmail(ByVal tipoDocumento As String, ByVal numeroDocumento As String) As Cliente
        Dim oCliente As New Cliente
        Using oConexion = New OracleConnection(DAConexion.Instancia.Get_CadenaConexionOracleFINX7Q2())
            If Not oConexion Is Nothing Then
                ErrorLog("antes oConexion.Open()")
                oConexion.Open()
                ErrorLog("después oConexion.Open()")
                If oConexion.State = ConnectionState.Open Then
                    ErrorLog("ConnectionState.Open")
                    Using cmd As OracleCommand = oConexion.CreateCommand()
                        cmd.CommandTimeout = 600
                        cmd.CommandType = CommandType.StoredProcedure
                        cmd.CommandText = "PKG_RIPLEYMATICO_EMAIL.PRC_OBTENER_EMAIL"

                        Dim parametro1 As New OracleParameter
                        Dim parametro2 As New OracleParameter
                        Dim parametro3 As New OracleParameter
                        Dim parametro4 As New OracleParameter

                        parametro1.ParameterName = "I_TIPODOCUMENTO"
                        parametro1.OracleType = OracleType.VarChar
                        parametro1.Size = 64
                        parametro1.Direction = ParameterDirection.Input
                        parametro1.Value = tipoDocumento.Trim()
                        cmd.Parameters.Add(parametro1)

                        parametro2.ParameterName = "I_NUMERODOCUMENTO"
                        parametro2.OracleType = OracleType.VarChar
                        parametro2.Size = 64
                        parametro2.Direction = ParameterDirection.Input
                        parametro2.Value = numeroDocumento.Trim()
                        cmd.Parameters.Add(parametro2)

                        parametro3.ParameterName = "O_EMAILLABORAL"
                        parametro3.OracleType = OracleType.VarChar
                        parametro3.Size = 100
                        parametro3.Direction = ParameterDirection.Output
                        cmd.Parameters.Add(parametro3)

                        parametro4.ParameterName = "O_EMAILPERSONAL"
                        parametro4.OracleType = OracleType.VarChar
                        parametro4.Size = 100
                        parametro4.Direction = ParameterDirection.Output
                        cmd.Parameters.Add(parametro4)

                        ErrorLog("ANTES cmd.ExecuteReader() PKG_RIPLEYMATICO_EMAIL.PRC_OBTENER_EMAIL")
                        cmd.ExecuteNonQuery()
                        ErrorLog("despues cmd.ExecuteReader()")
                        oCliente.CorreoLaboral = cmd.Parameters("O_EMAILLABORAL").Value.ToString()
                        oCliente.CorreoPersonal = cmd.Parameters("O_EMAILPERSONAL").Value.ToString()
                        ErrorLog("oCliente.CorreoLaboral=" & oCliente.CorreoLaboral)
                        ErrorLog("oCliente.CorreoPersonal=" & oCliente.CorreoPersonal)

                    End Using
                End If
            End If
        End Using
        Return oCliente
    End Function

End Class
