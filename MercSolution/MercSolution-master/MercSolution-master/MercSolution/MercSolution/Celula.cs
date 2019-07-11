using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MercSolution
{
    public class Celula
    {

        public int claveCelula { get; set; }
        public string nombreCelula { get; set; }
        public int Ncelula { get; set; }
        public Celula(int a, string b, int c)
        {
            claveCelula = a;
            nombreCelula = b;
            Ncelula = c;
        }

        public Celula()
        {

        }

    }
}
