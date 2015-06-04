Imports Microsoft.VisualBasic
Imports System.Net.Mail
Imports System.Collections.Generic
Imports NCapas.Utility.Log
Imports NCapas.Entity
Imports NCapas.Business

Public Class EnvioData

    Private _addresses As List(Of String)
    Public Property Addresses() As List(Of String)
        Get
            Return _addresses
        End Get
        Set(ByVal value As List(Of String))
            _addresses = value
        End Set
    End Property


    Private _subject As String
    Public Property Subject() As String
        Get
            Return _subject
        End Get
        Set(ByVal value As String)
            _subject = value
        End Set
    End Property

    Private _body As String
    Public Property Body() As String
        Get
            Return _body
        End Get
        Set(ByVal value As String)
            _body = value
        End Set
    End Property

    Private _fileAttach As List(Of String)
    Public Property FileAttach() As List(Of String)
        Get
            Return _fileAttach
        End Get
        Set(ByVal value As List(Of String))
            _fileAttach = value
        End Set
    End Property

    Private _tipoDocumento As String
    Public Property TipoDocumento() As String
        Get
            Return _tipoDocumento
        End Get
        Set(ByVal value As String)
            _tipoDocumento = value
        End Set
    End Property

    Private _numeroDocumento As String
    Public Property NumeroDocumento() As String
        Get
            Return _numeroDocumento
        End Get
        Set(ByVal value As String)
            _numeroDocumento = value
        End Set
    End Property

    Private _nroTarjeta As String
    Public Property NroTarjeta() As String
        Get
            Return _nroTarjeta
        End Get
        Set(ByVal value As String)
            _nroTarjeta = value
        End Set
    End Property



    Private Sub AdjuntarFile(ByRef message As MailMessage, ByVal files As IList(Of String))
        ErrorLog("Entro a AdjuntarFile")
        If files IsNot Nothing Then
            For Each file As String In files
                If Not String.IsNullOrEmpty(file) Then
                    Dim attach As New Attachment(file)
                    message.Attachments.Add(attach)
                End If
            Next
        End If
    End Sub

    Public Sub EnviarEmailMasivoFisa()
        SyncLock GetType(EnvioData)

            ErrorLog("Entro a EnviarEmail")
            Dim message As New MailMessage()
            Dim oCliente As New Cliente
            Dim contador As Integer = 0
            Dim consultaClave6 As New ConsultaClave6

            Try
                oCliente = BNBRPFISA.Instancia.ObtenerOtrosEmail(TipoDocumento, NumeroDocumento)
                If oCliente.CorreoLaboral <> "" Then
                    Addresses.Add(oCliente.CorreoLaboral)
                End If

                If oCliente.CorreoPersonal <> "" Then
                    Addresses.Add(oCliente.CorreoPersonal)
                End If

                For Each address As String In Addresses
                    contador = contador + 1
                    ErrorLog("Para=" & address)
                    message.To.Add(New MailAddress(address))
                    ErrorLog("Subject=" & Subject)
                    message.Subject = Subject
                    message.SubjectEncoding = Encoding.UTF8
                    message.Body = Body
                    ErrorLog("Body=" & Body)
                    message.BodyEncoding = Encoding.UTF8
                    message.IsBodyHtml = True
                    message.Priority = MailPriority.Normal
                    message.DeliveryNotificationOptions = DeliveryNotificationOptions.OnFailure

                    'adjuntar archivo
                    If contador = 1 Then
                        AdjuntarFile(message, FileAttach)
                    End If

                    Try
                        Using smtp As New SmtpClient()
                            smtp.Send(message)
                        End Using
                        ErrorLog("Envio de Correo Correctamente a " & address)
                        consultaClave6.Nro_Doc = NumeroDocumento
                        consultaClave6.Email = address
                        consultaClave6.Envio = True
                        consultaClave6.Contador = contador
                        BNConsultaClave6.Instancia.UpdateConsultaClave6(consultaClave6)
                    Catch ex As Exception
                        ErrorLog("No se envio de Correo Correctamente a " & address)
                        consultaClave6.Nro_Doc = NumeroDocumento
                        consultaClave6.Email = address
                        consultaClave6.Envio = False
                        consultaClave6.Contador = contador
                        BNConsultaClave6.Instancia.UpdateConsultaClave6(consultaClave6)
                    End Try
                Next
            Catch exs As SmtpException
                ErrorLog("Error SmtpException= " & exs.Message)
            Catch ex As Exception
                ErrorLog("Error Exception= " & ex.Message)
            Finally
                message.Dispose()
            End Try

        End SyncLock
    End Sub

    Public Sub EnviarEmailMasivo()
        SyncLock GetType(EnvioData)

            ErrorLog("Entro a EnviarEmail")
            Dim msj As String = String.Empty
            Dim message As New MailMessage()
            Try
                For Each address As String In Addresses
                    ErrorLog("Para=" & address)
                    message.To.Add(New MailAddress(address))
                    ErrorLog("Subject=" & Subject)
                    message.Subject = Subject
                    message.SubjectEncoding = Encoding.UTF8
                    message.Body = Body
                    ErrorLog("Body=" & Body)
                    message.BodyEncoding = Encoding.UTF8
                    message.IsBodyHtml = True
                    message.Priority = MailPriority.Normal
                    message.DeliveryNotificationOptions = DeliveryNotificationOptions.OnFailure

                    'adjuntar archivo
                    AdjuntarFile(message, FileAttach)

                    Using smtp As New SmtpClient()
                        smtp.Send(message)
                    End Using
                    ErrorLog("Envio de Correo Correctamente a " & address)
                    msj = "ok"
                Next
            Catch exs As SmtpException
                msj = If(exs.InnerException IsNot Nothing, exs.InnerException.Message, exs.Message)
                ErrorLog("Error SmtpException= " & exs.Message)
            Catch ex As Exception
                msj = If(ex.InnerException IsNot Nothing, ex.InnerException.Message, ex.Message)
                ErrorLog("Error Exception= " & ex.Message)
            Finally
                message.Dispose()
            End Try

        End SyncLock
    End Sub

End Class
