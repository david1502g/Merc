using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MercSolution
{
   public class Promotor
    {

        public int clavePromotor { get; set; }
        public string nombrePromotor { get; set; }
        public string celula { get; set; }
        public double nominaFija { get; set; }
        public float nominaVariable { get; set; }
        public int estrategia { get; set; }
        public string claveInter { get; set; }
        public string banco { get; set; }
        public int Ncelula { get; set; }

        public int ingresos { get; set; }
        public Promotor()
        {
            clavePromotor = 0;
            nombrePromotor = "";
            celula = "";
            banco = "";
            nominaFija = 0;
            nominaVariable = 0;
            claveInter = "";
            estrategia = 0;
            Ncelula = 0;
            ingresos = 0;
        }
        
        public Promotor(int cp)
        {
            clavePromotor = cp;
            nombrePromotor = "";
            
            estrategia = 0;
        }
    }
}
