Imports System.Threading

Public MustInherit Class Singleton(Of T As {Class, New})
    Private Shared _instancia As T
    Private Shared ReadOnly _mutex As New Mutex()

    Public Shared ReadOnly Property Instancia() As T
        Get
            _mutex.WaitOne()
            If _instancia Is Nothing Then
                _instancia = New T()
            End If
            _mutex.ReleaseMutex()
            Return _instancia
        End Get
    End Property
End Class