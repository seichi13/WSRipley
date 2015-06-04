Imports NCapas.Entity
Imports System.Data.SqlClient
Imports NCapas.Data.DAConexion
Imports NCapas.Utility.Log

Public Class DAConsultaCambioProducto
    Inherits Singleton(Of DAConsultaCambioProducto)

    Public Function InsertConsultaCambioProducto(ByVal consultaCambioProducto As ConsultaCambioProducto) As Boolean

        Dim sMensajeError_SQL As String = ""
        Using oConexion As SqlConnection = DAConexion.Instancia.SQL_ConnectionOpen(DAConexion.Instancia.Get_CadenaConexion(), sMensajeError_SQL)
            If Not oConexion Is Nothing Then
                If oConexion.State = ConnectionState.Open Then
                    Using comando As SqlCommand = oConexion.CreateCommand()
                        comando.CommandTimeout = 600
                        comando.CommandType = CommandType.StoredProcedure
                        comando.CommandText = "Usp_Insert_ConsultaCambioProducto"

                        comando.Parameters.AddWithValue("@COD_SUCURSAL_BAN", consultaCambioProducto.Cod_Sucursal_Ban)
                        comando.Parameters.AddWithValue("@NOMBRE_KIOSCO", consultaCambioProducto.Nombre_Kiosko)
                        comando.Parameters.AddWithValue("@TIPO_DOC", consultaCambioProducto.Tipo_Doc)
                        comando.Parameters.AddWithValue("@NRO_DOC", consultaCambioProducto.Nro_Doc)
                        comando.Parameters.AddWithValue("@OPCION", consultaCambioProducto.Opcion)
                        comando.Parameters.AddWithValue("@NUMERO_TARJETA", consultaCambioProducto.Numero_Tarjeta)
                        comando.Parameters.AddWithValue("@TIPO_TARJETA_INICIAL", consultaCambioProducto.Tipo_Tarjeta_Inicial)
                        comando.Parameters.AddWithValue("@TIPO_TARJETA_FINAL", consultaCambioProducto.Tipo_Tarjeta_Final)
                        comando.Parameters.AddWithValue("@FECHA_INICIAL_OFERTA", consultaCambioProducto.Fecha_Inicial_Oferta)
                        comando.Parameters.AddWithValue("@FECHA_FINAL_OFERTA", consultaCambioProducto.Fecha_Final_Oferta)
                        comando.Parameters.AddWithValue("@NRO_CUENTA", consultaCambioProducto.Numero_Cuenta)
                        comando.Parameters.AddWithValue("@ID_KIOSKO", consultaCambioProducto.Id_Sucursal)

                        comando.ExecuteNonQuery()
                    End Using
                End If
            End If
        End Using
        Return True
    End Function

End Class
