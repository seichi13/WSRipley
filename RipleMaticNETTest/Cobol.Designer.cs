namespace RipleMaticNETTest
{
    partial class Cobol
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Cobol));
            this.txtOUT = new System.Windows.Forms.TextBox();
            this.btnCalcular = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // txtOUT
            // 
            this.txtOUT.Location = new System.Drawing.Point(44, 39);
            this.txtOUT.Multiline = true;
            this.txtOUT.Name = "txtOUT";
            this.txtOUT.Size = new System.Drawing.Size(599, 303);
            this.txtOUT.TabIndex = 0;
            this.txtOUT.Text = resources.GetString("txtOUT.Text");
            // 
            // btnCalcular
            // 
            this.btnCalcular.Location = new System.Drawing.Point(75, 375);
            this.btnCalcular.Name = "btnCalcular";
            this.btnCalcular.Size = new System.Drawing.Size(75, 23);
            this.btnCalcular.TabIndex = 1;
            this.btnCalcular.Text = "Calcular";
            this.btnCalcular.UseVisualStyleBackColor = true;
            this.btnCalcular.Click += new System.EventHandler(this.btnCalcular_Click);
            // 
            // Cobol
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(936, 411);
            this.Controls.Add(this.btnCalcular);
            this.Controls.Add(this.txtOUT);
            this.Name = "Cobol";
            this.Text = "Cobol";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox txtOUT;
        private System.Windows.Forms.Button btnCalcular;
    }
}