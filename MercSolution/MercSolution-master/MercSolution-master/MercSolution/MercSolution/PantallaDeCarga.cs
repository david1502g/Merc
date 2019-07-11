using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MercSolution
{
    public partial class PantallaDeCarga : Form
    {
        int progreso = 0;
        public PantallaDeCarga()
        {
            InitializeComponent();
            progressBar1.Value = 0;
            
        }

        private void progressBar1_Click(object sender, EventArgs e)
        {

        }

        public void aumentarCarga(int aumento) {
            progreso += aumento;
            progressBar1.Value = progreso;
        }
    }
}
