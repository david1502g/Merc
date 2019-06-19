using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MercSolution
{
    public class Comisiones
    {
        public string tablasComisiones { get; set; }
        public float comisionTotal { get; set; }
        public float comisionPromotor { get; set; }
        public float comisionGerente { get; set; }
        public string paquetes { get; set; }

        public Comisiones()
        {
            tablasComisiones = "";
            comisionTotal = 0;
            comisionGerente = 0;
            comisionPromotor = 0;

            paquetes = "";

        }
    }
}
