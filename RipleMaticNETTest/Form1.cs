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
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnInvocar_Click(object sender, EventArgs e)
        {
            string resultado = "";
            Servicio.ServiceSoapClient servicio = new Servicio.ServiceSoapClient();
            resultado = servicio.SIMULADOR_OFERTA_SUPEREFECTIVO_TEST(txtTarjeta.Text, txtTem.Text, txtMonto.Text, txtPlazo.Text, txtMontoOfertado.Text);
            if (resultado != null)
                txtResultado.Text = resultado;
        }
    }
}
