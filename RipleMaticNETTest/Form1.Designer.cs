namespace RipleMaticNETTest
{
    partial class Form1
    {
        /// <summary>
        /// Variable del diseñador requerida.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Limpiar los recursos que se estén utilizando.
        /// </summary>
        /// <param name="disposing">true si los recursos administrados se deben eliminar; false en caso contrario, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Código generado por el Diseñador de Windows Forms

        /// <summary>
        /// Método necesario para admitir el Diseñador. No se puede modificar
        /// el contenido del método con el editor de código.
        /// </summary>
        private void InitializeComponent()
        {
            this.txtResultado = new System.Windows.Forms.TextBox();
            this.btnInvocar = new System.Windows.Forms.Button();
            this.txtTarjeta = new System.Windows.Forms.TextBox();
            this.txtTem = new System.Windows.Forms.TextBox();
            this.txtMonto = new System.Windows.Forms.TextBox();
            this.txtPlazo = new System.Windows.Forms.TextBox();
            this.txtMontoOfertado = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // txtResultado
            // 
            this.txtResultado.Location = new System.Drawing.Point(39, 253);
            this.txtResultado.Name = "txtResultado";
            this.txtResultado.Size = new System.Drawing.Size(512, 20);
            this.txtResultado.TabIndex = 0;
            // 
            // btnInvocar
            // 
            this.btnInvocar.Location = new System.Drawing.Point(39, 279);
            this.btnInvocar.Name = "btnInvocar";
            this.btnInvocar.Size = new System.Drawing.Size(121, 23);
            this.btnInvocar.TabIndex = 1;
            this.btnInvocar.Text = "Invocar Servicio";
            this.btnInvocar.UseVisualStyleBackColor = true;
            this.btnInvocar.Click += new System.EventHandler(this.btnInvocar_Click);
            // 
            // txtTarjeta
            // 
            this.txtTarjeta.Location = new System.Drawing.Point(47, 28);
            this.txtTarjeta.Name = "txtTarjeta";
            this.txtTarjeta.Size = new System.Drawing.Size(100, 20);
            this.txtTarjeta.TabIndex = 2;
            this.txtTarjeta.Text = "5254740005648129";
            // 
            // txtTem
            // 
            this.txtTem.Location = new System.Drawing.Point(46, 67);
            this.txtTem.Name = "txtTem";
            this.txtTem.Size = new System.Drawing.Size(100, 20);
            this.txtTem.TabIndex = 3;
            this.txtTem.Text = "1.99";
            // 
            // txtMonto
            // 
            this.txtMonto.Location = new System.Drawing.Point(48, 110);
            this.txtMonto.Name = "txtMonto";
            this.txtMonto.Size = new System.Drawing.Size(100, 20);
            this.txtMonto.TabIndex = 4;
            this.txtMonto.Text = "3000";
            // 
            // txtPlazo
            // 
            this.txtPlazo.Location = new System.Drawing.Point(50, 153);
            this.txtPlazo.Name = "txtPlazo";
            this.txtPlazo.Size = new System.Drawing.Size(100, 20);
            this.txtPlazo.TabIndex = 5;
            this.txtPlazo.Text = "36";
            // 
            // txtMontoOfertado
            // 
            this.txtMontoOfertado.Location = new System.Drawing.Point(46, 194);
            this.txtMontoOfertado.Name = "txtMontoOfertado";
            this.txtMontoOfertado.Size = new System.Drawing.Size(100, 20);
            this.txtMontoOfertado.TabIndex = 6;
            this.txtMontoOfertado.Text = "4500";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(586, 353);
            this.Controls.Add(this.txtMontoOfertado);
            this.Controls.Add(this.txtPlazo);
            this.Controls.Add(this.txtMonto);
            this.Controls.Add(this.txtTem);
            this.Controls.Add(this.txtTarjeta);
            this.Controls.Add(this.btnInvocar);
            this.Controls.Add(this.txtResultado);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox txtResultado;
        private System.Windows.Forms.Button btnInvocar;
        private System.Windows.Forms.TextBox txtTarjeta;
        private System.Windows.Forms.TextBox txtTem;
        private System.Windows.Forms.TextBox txtMonto;
        private System.Windows.Forms.TextBox txtPlazo;
        private System.Windows.Forms.TextBox txtMontoOfertado;
    }
}

