using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace RipleMaticNETTest
{
    public partial class Cobol : Form
    {
        public Cobol()
        {
            InitializeComponent();
        }

        private void btnCalcular_Click(object sender, EventArgs e)
        {
            var respuesta = string.Empty;
            var cantidad = txtOUT.Text.Length;
            var a = txtOUT.Text.Substring(101,1);
            var cronograma = txtOUT.Text.Substring(102, cantidad - 102);

            respuesta = cronograma.Substring(0, 51);
            respuesta = cronograma.Substring(43, 8);

            MessageBox.Show("Longitud :" + cantidad);
            MessageBox.Show("estado :" + a);
            MessageBox.Show("tabla :" + cronograma);
        }
    }
}
