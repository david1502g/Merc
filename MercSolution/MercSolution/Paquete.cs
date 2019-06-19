using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MercSolution
{
    public class Paquete
    {
        public int paquetesTotales { set; get; }
        public Dictionary<string, Dictionary<string, double>> packs { get; set; }
        public Dictionary<string, int> diccionarioPaquetes { set; get; }
        public string nombreP { set; get; }

        public Paquete()
        {
            packs = new Dictionary<string, Dictionary<string, double>>();
            /*diccionarioPaquetes = new Dictionary<string, int>();
            diccionarioPaquetes.Add("PAQUETE $289", 0);
            diccionarioPaquetes.Add("PAQUETE $333", 0);
            diccionarioPaquetes.Add("PAQUETE $389", 0);
            diccionarioPaquetes.Add("PAQUETE $435", 0);
            diccionarioPaquetes.Add("PAQUETE $499", 0);
            diccionarioPaquetes.Add("PAQUETE $599", 0);
            diccionarioPaquetes.Add("PAQUETE $999", 0);
            diccionarioPaquetes.Add("PAQUETE $1,499", 0);
            diccionarioPaquetes.Add("INFINITUM 10 MBPS (SOLO INTERNET)", 0);
            diccionarioPaquetes.Add("INFINITUM 20 MBPS (SOLO INTERNET)", 0);
            diccionarioPaquetes.Add("INFINITUM 50 MBPS (SOLO INTERNET)", 0);
            diccionarioPaquetes.Add("INFINITUM 100 MBPS (SOLO INTERNET)", 0);
            diccionarioPaquetes.Add("PAQUETE INFINITUM NEGOCIO $399", 0);
            diccionarioPaquetes.Add("PAQUETE INFINITUM NEGOCIO $549", 0);
            diccionarioPaquetes.Add("PAQUETE INFINITUM NEGOCIO $799", 0);
            diccionarioPaquetes.Add("PAQUETE INFINITUM NEGOCIO $1,499", 0);
            diccionarioPaquetes.Add("PAQUETE INFINITUM NEGOCIO $1,789", 0);
            diccionarioPaquetes.Add("PAQUETE INFINITUM NEGOCIO $2,289", 0);
            diccionarioPaquetes.Add("INFINITUM HASTA 10 MBPS (SOLO INTERNET)", 0);
            diccionarioPaquetes.Add("LINEA SIN PAQUETE", 0);*/
            ExtraerPaquetesAccess();
            Console.WriteLine("Agregue los packs " + packs.Count + " " + packs["2 PLAY"].Count);
        }

        public Paquete(Dictionary<string, Dictionary<string, double>> algo)
        {
            this.packs = algo;
        }
        public Paquete(Dictionary<string, int> algo)
        {
            diccionarioPaquetes = algo;
        }

        public Paquete(List<string> listN)
        {
            diccionarioPaquetes = new Dictionary<string, int>();
            foreach (string i in listN)
            {
                diccionarioPaquetes.Add(i, 0);
            }
        }

        public Paquete(List<string> listN, string nombreP)
        {
            diccionarioPaquetes = new Dictionary<string, int>();
            foreach (string i in listN)
            {
                diccionarioPaquetes.Add(i, 0);
            }
            this.nombreP = nombreP;
        }

        public void aumentar(string llave, string tipoPack)
        {
            string best = Mejor(tipoPack);


            string otherPack = llave.Replace("  ", " ");
            try
            {
                packs[best][otherPack]++;
            }
            catch (System.Collections.Generic.KeyNotFoundException)
            {
                string bestPack;
                if (best == "SOLO LINEA")
                {
                    bestPack = MejorPaquete("PORTABI" + otherPack, packs[best]);
                }
                else
                {
                    bestPack = MejorPaquete(otherPack, packs[best]);
                }

                Console.WriteLine("No encontre el paquete " + otherPack + " de tipo " + best + " y lo insertare en " + bestPack);
                packs[best][bestPack]++;
            }
            /*try
            {
                diccionarioPaquetes[llave]++;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }*/
        }

        public double ObtenerComision(string llave, string tipoPack)
        {
            string best = Mejor(tipoPack);
            double comision = 0;

            string otherPack = llave.Replace("  ", " ");
            try
            {
                comision = packs[best][otherPack];
            }
            catch (System.Collections.Generic.KeyNotFoundException)
            {
                string bestPack;
                if (best == "SOLO LINEA")
                {
                    bestPack = MejorPaquete("PORTABI" + otherPack, packs[best]);
                }
                else
                {
                    bestPack = MejorPaquete(otherPack, packs[best]);
                }

                Console.WriteLine("No encontre el paquete " + otherPack + " de tipo " + best + " y lo insertare en " + bestPack);
                comision = packs[best][bestPack];
            }

            return comision;
        }

        public void aumentar(string llave)
        {
            try
            {
                diccionarioPaquetes[llave]++;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }
        }

        public void ImprimirPaquetes(string tipo)
        {
            foreach (var i in packs[tipo])
            {
                Console.WriteLine(" Paquete[" + i.Key + "] hay: " + i.Value);
            }
        }

        public int totalPaquetes()
        {
            int total = 0;

            foreach (var ir in diccionarioPaquetes.Values)
            {
                total += ir;
            }

            return total;
        }

        public int contarPaquetes()
        {
            int cont = 0;
            foreach (KeyValuePair<string, int> i in diccionarioPaquetes)
            {
                Console.WriteLine("VALOR " + i.Value);
                cont += i.Value;
            }
            return cont;
        }

        public void ExtraerPaquetesAccess()
        {
            string connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source="+System.IO.Directory.GetCurrentDirectory()+"\\comisiones1.accdb";
            string strSQL = "SELECT DISTINCT(TABLACOMISIONES) FROM Paquetes";
            Console.WriteLine("DIRECTORIO"+System.IO.Directory.GetCurrentDirectory());
            // Create a connection  
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                // Create a command and set its connection  
                OleDbCommand command = new OleDbCommand(strSQL, connection);
                // Open the connection and execute the select command.  
                try
                {
                    // Open connecton  
                    connection.Open();
                    // Execute command  
                    using (OleDbDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            packs.Add(reader["TABLACOMISIONES"].ToString().ToUpper(), new Dictionary<string, double>());
                        }
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.ToString());
                }
                foreach (KeyValuePair<string, Dictionary<string, double>> entry in packs)
                {
                    command = new OleDbCommand("SELECT * FROM Paquetes WHERE TABLACOMISIONES = '" + entry.Key + "'", connection);
                    // Open the connection and execute the select command.  
                    try
                    {
                        // Open connecton  
                        // Execute command  
                        using (OleDbDataReader reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                Console.WriteLine("Agregado " + reader["COMISIONPROMOTOR"].ToString());
                                packs[entry.Key].Add(reader["PAQUETES"].ToString().ToUpper(), Double.Parse(reader["COMISIONPROMOTOR"].ToString()));
                            }
                        }
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e.ToString());
                    }
                }
            }
        }

        public string Mejor(string xlPack)
        {
            int aux = 0;
            string mejor = packs.Keys.ElementAt(0).ToUpper();
            int dis = CalcLevenshteinDistance(mejor, xlPack);
            aux = dis;
            foreach (KeyValuePair<string, Dictionary<string, double>> entry in packs)
            {
                aux = CalcLevenshteinDistance(entry.Key.ToUpper(), xlPack);
                //Console.WriteLine("Comparando " + entry.Key + " con " + xlPack);
                if (aux < dis)
                {
                    mejor = entry.Key;
                    dis = aux;
                }
                // do something with entry.Value or entry.Key
            }

            return mejor;
        }

        public string MejorPaquete(string pack, Dictionary<string, double> dicPacks)
        {
            int aux = 0;
            string mejor = dicPacks.Keys.ElementAt(0).ToUpper();
            int dis = CalcLevenshteinDistance(mejor, pack);
            aux = dis;
            foreach (KeyValuePair<string, double> entry in dicPacks)
            {
                aux = CalcLevenshteinDistance(pack, entry.Key.ToUpper());
                //Console.WriteLine("Distancia entre " + pack + " y " + entry.Key.ToUpper() + " es " + aux);
                if (aux < dis)
                {
                    mejor = entry.Key;
                    dis = aux;
                }
            }

            return mejor;
        }

        private int CalcLevenshteinDistance(string a, string b)
        {
            if (String.IsNullOrEmpty(a) && String.IsNullOrEmpty(b))
            {
                return 0;
            }
            if (String.IsNullOrEmpty(a))
            {
                return b.Length;
            }
            if (String.IsNullOrEmpty(b))
            {
                return a.Length;
            }
            int lengthA = a.Length;
            int lengthB = b.Length;
            var distances = new int[lengthA + 1, lengthB + 1];
            for (int i = 0; i <= lengthA; distances[i, 0] = i++) ;
            for (int j = 0; j <= lengthB; distances[0, j] = j++) ;

            for (int i = 1; i <= lengthA; i++)
                for (int j = 1; j <= lengthB; j++)
                {
                    int cost = b[j - 1] == a[i - 1] ? 0 : 1;
                    distances[i, j] = Math.Min
                        (
                        Math.Min(distances[i - 1, j] + 1, distances[i, j - 1] + 1),
                        distances[i - 1, j - 1] + cost
                        );
                }
            return distances[lengthA, lengthB];
        }
    }
}
