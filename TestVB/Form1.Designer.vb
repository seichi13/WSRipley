<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form reemplaza a Dispose para limpiar la lista de componentes.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Requerido por el Diseñador de Windows Forms
    Private components As System.ComponentModel.IContainer

    'NOTA: el Diseñador de Windows Forms necesita el siguiente procedimiento
    'Se puede modificar usando el Diseñador de Windows Forms.  
    'No lo modifique con el editor de código.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.btnCalcular = New System.Windows.Forms.Button()
        Me.txtDia = New System.Windows.Forms.TextBox()
        Me.txtMes = New System.Windows.Forms.TextBox()
        Me.txtAnio = New System.Windows.Forms.TextBox()
        Me.lblFecha = New System.Windows.Forms.Label()
        Me.lblRespuesta = New System.Windows.Forms.Label()
        Me.ObtenerTramaEECC = New System.Windows.Forms.Button()
        Me.txtSalidaEECC = New System.Windows.Forms.TextBox()
        Me.btnCambioProducto = New System.Windows.Forms.Button()
        Me.BtnTitular = New System.Windows.Forms.Button()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'btnCalcular
        '
        Me.btnCalcular.Location = New System.Drawing.Point(23, 102)
        Me.btnCalcular.Name = "btnCalcular"
        Me.btnCalcular.Size = New System.Drawing.Size(100, 23)
        Me.btnCalcular.TabIndex = 3
        Me.btnCalcular.Text = "Calcular"
        Me.btnCalcular.UseVisualStyleBackColor = True
        '
        'txtDia
        '
        Me.txtDia.Location = New System.Drawing.Point(23, 24)
        Me.txtDia.Name = "txtDia"
        Me.txtDia.Size = New System.Drawing.Size(100, 20)
        Me.txtDia.TabIndex = 4
        '
        'txtMes
        '
        Me.txtMes.Location = New System.Drawing.Point(23, 50)
        Me.txtMes.Name = "txtMes"
        Me.txtMes.Size = New System.Drawing.Size(100, 20)
        Me.txtMes.TabIndex = 5
        Me.txtMes.Text = "11"
        '
        'txtAnio
        '
        Me.txtAnio.Location = New System.Drawing.Point(23, 76)
        Me.txtAnio.Name = "txtAnio"
        Me.txtAnio.Size = New System.Drawing.Size(100, 20)
        Me.txtAnio.TabIndex = 6
        Me.txtAnio.Text = "2014"
        '
        'lblFecha
        '
        Me.lblFecha.AutoSize = True
        Me.lblFecha.Location = New System.Drawing.Point(142, 24)
        Me.lblFecha.Name = "lblFecha"
        Me.lblFecha.Size = New System.Drawing.Size(0, 13)
        Me.lblFecha.TabIndex = 7
        '
        'lblRespuesta
        '
        Me.lblRespuesta.Location = New System.Drawing.Point(20, 271)
        Me.lblRespuesta.Name = "lblRespuesta"
        Me.lblRespuesta.Size = New System.Drawing.Size(465, 150)
        Me.lblRespuesta.TabIndex = 8
        '
        'ObtenerTramaEECC
        '
        Me.ObtenerTramaEECC.Location = New System.Drawing.Point(766, 385)
        Me.ObtenerTramaEECC.Name = "ObtenerTramaEECC"
        Me.ObtenerTramaEECC.Size = New System.Drawing.Size(141, 23)
        Me.ObtenerTramaEECC.TabIndex = 9
        Me.ObtenerTramaEECC.Text = "Obtener Trama EECC"
        Me.ObtenerTramaEECC.UseVisualStyleBackColor = True
        '
        'txtSalidaEECC
        '
        Me.txtSalidaEECC.Location = New System.Drawing.Point(129, 26)
        Me.txtSalidaEECC.Multiline = True
        Me.txtSalidaEECC.Name = "txtSalidaEECC"
        Me.txtSalidaEECC.Size = New System.Drawing.Size(778, 353)
        Me.txtSalidaEECC.TabIndex = 10
        '
        'btnCambioProducto
        '
        Me.btnCambioProducto.Location = New System.Drawing.Point(23, 131)
        Me.btnCambioProducto.Name = "btnCambioProducto"
        Me.btnCambioProducto.Size = New System.Drawing.Size(100, 23)
        Me.btnCambioProducto.TabIndex = 11
        Me.btnCambioProducto.Text = "Cambio Producto"
        Me.btnCambioProducto.UseVisualStyleBackColor = True
        '
        'BtnTitular
        '
        Me.BtnTitular.Location = New System.Drawing.Point(23, 160)
        Me.BtnTitular.Name = "BtnTitular"
        Me.BtnTitular.Size = New System.Drawing.Size(100, 23)
        Me.BtnTitular.TabIndex = 12
        Me.BtnTitular.Text = "Titular"
        Me.BtnTitular.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(418, 384)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 13
        Me.Button1.Text = "Button1"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(919, 420)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.BtnTitular)
        Me.Controls.Add(Me.btnCambioProducto)
        Me.Controls.Add(Me.txtSalidaEECC)
        Me.Controls.Add(Me.ObtenerTramaEECC)
        Me.Controls.Add(Me.lblRespuesta)
        Me.Controls.Add(Me.lblFecha)
        Me.Controls.Add(Me.txtAnio)
        Me.Controls.Add(Me.txtMes)
        Me.Controls.Add(Me.txtDia)
        Me.Controls.Add(Me.btnCalcular)
        Me.Name = "Form1"
        Me.Text = "Form1"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Private WithEvents btnCalcular As System.Windows.Forms.Button
    Friend WithEvents txtDia As System.Windows.Forms.TextBox
    Friend WithEvents txtMes As System.Windows.Forms.TextBox
    Friend WithEvents txtAnio As System.Windows.Forms.TextBox
    Friend WithEvents lblFecha As System.Windows.Forms.Label
    Friend WithEvents lblRespuesta As System.Windows.Forms.Label
    Friend WithEvents ObtenerTramaEECC As System.Windows.Forms.Button
    Friend WithEvents txtSalidaEECC As System.Windows.Forms.TextBox
    Friend WithEvents btnCambioProducto As System.Windows.Forms.Button
    Friend WithEvents BtnTitular As System.Windows.Forms.Button
    Friend WithEvents Button1 As System.Windows.Forms.Button

End Class
