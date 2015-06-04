
Imports System
Imports System.ComponentModel
Imports NCapas.Utility

Public Class BNContrato
    Implements IDisposable

    Private handle As IntPtr
    Private component As New Component()
    Private disposed As Boolean = False

    Public Sub New()

    End Sub

    Public Sub Dispose() Implements System.IDisposable.Dispose
        Dispose(True)
        GC.SuppressFinalize(Me)

    End Sub

    Private Sub Dispose(disposing As Boolean)
        If Not Me.disposed Then
            If disposing Then
                component.Dispose()

            End If

            CloseHandle(handle)
            handle = IntPtr.Zero
        End If
        disposed = True

    End Sub

    <System.Runtime.InteropServices.DllImport("Kernel32")>
    Private Shared Function CloseHandle(handle As IntPtr) As Boolean
    End Function

    Sub BNContrato()
        Dispose(False)
    End Sub

End Class



