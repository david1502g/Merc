using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace MercSolution
{
    class ManipuladorAcces
    {
        string path;
        public ManipuladorAcces(string path)
        {
            this.path = path;
        }

          
        public void nombresPromotor() {

            Excel.Application excelApplication = new Excel.Application();
            Excel.Workbook destworkBook;
            Excel.Worksheet destworkSheet;
            //Opening of first worksheet and copying
           // Console.WriteLine("weeeea" + eliminarSlash(path) + "MercTest2.xlsx");
            destworkBook = excelApplication.Workbooks.Open(eliminarSlash(path) + "MercTest2.xlsx");
            destworkSheet = (Microsoft.Office.Interop.Excel.Worksheet)destworkBook.Sheets.get_Item(1);

            int k = 1;
            destworkSheet.Cells[k, 4].Value = "Nombre Promotor";
            k++;

            Boolean bandera = true;
            int h = 2;
            List<string> cveProm = new List<string>();
            while (bandera == true)
            {
                if(destworkSheet.Cells[h, 3].Value != null)
                {
                    cveProm.Add(destworkSheet.Cells[h, 3].Value.ToString());
                    h++;
                }
                else
                {
                    bandera = false;
                }
                
            }

            List<string> nombre = new List<string>();

           String stringConexion = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path;
            using (OleDbConnection connection = new OleDbConnection(stringConexion))
            {
                connection.Open();
                foreach (string i in cveProm)
                {
                    OleDbDataReader reader = null;
                    OleDbCommand command = new OleDbCommand("select * FROM Promotores WHERE Cve_prom = " + i, connection);
                    reader = command.ExecuteReader();
                    if (reader.Read())
                    {
                        destworkSheet.Cells[k, 4].Value = reader[2].ToString();
                        k++;
                    }
                }
            }

            
            

            destworkBook.Save();
            destworkBook.Close();
            excelApplication.Quit();

        }

        public void DatosPromotor(List<Promotor>[] lista)
        {
            //OleDbConnection con;//Representa una conexión abierta a un origen de datos
            string stringConexion = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path;
            using (OleDbConnection connection = new OleDbConnection(stringConexion))
            {
                connection.Open();
                for (int j=0; j < lista.Length; j++)
                {
                    List<Promotor> list = lista[j];

                    foreach (Promotor i in list)
                    {
                        OleDbDataReader reader = null;
                        OleDbCommand command = new OleDbCommand("SELECT * FROM PROMOTORES WHERE Cve_prom = " + i.clavePromotor, connection);
                        reader = command.ExecuteReader();
                        if (reader.Read())
                        {
                            i.estrategia = int.Parse(reader[4].ToString());
                            i.nombrePromotor = reader[2].ToString();
                        }

                        reader.Close();
                    }
                }
            }
        }

        public static string eliminarSlash(string str)
        {
            int i = str.Length - 1;
            int j = 0;
            while (str[i] != '\\')
            {
                i--;
                j++;
            }

            return str.Remove(i + 1);
        }

        public void DatosPromotorN(List<Promotor> list)
        {
            String stringConexion = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path;
            using (OleDbConnection connection = new OleDbConnection(stringConexion))
            {
                connection.Open();
                OleDbDataReader reader = null;
                OleDbCommand command = new OleDbCommand("select * FROM PROMOTORES", connection);
                reader = command.ExecuteReader();
                Promotor aux;
                while (reader.Read())
                {
                    aux = new Promotor();
                    aux.nombrePromotor= reader["nombre_promotor"].ToString();
                    aux.clavePromotor = int.Parse(reader["Cve_prom"].ToString());
                    aux.estrategia = int.Parse(reader["Estrategia"].ToString());
                    aux.celula = reader["Celula"].ToString();
                    aux.claveInter = reader["CLAVE_INTERBANCARIA"].ToString();
                    aux.banco = reader["Banco"].ToString();
                    aux.Ncelula = int.Parse(reader["Ncelula"].ToString());
                    list.Add(aux);
                }
                reader.Close();
                
                connection.Close();
            }
        }

        public void celulasN(List<Celula> list, string path)
        {
            String stringConexion = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path;
            using (OleDbConnection connection = new OleDbConnection(stringConexion))
            {
                connection.Open();
                OleDbDataReader reader = null;
                OleDbCommand command = new OleDbCommand("select * FROM CELULAS", connection);
                reader = command.ExecuteReader();
                Celula aux;
                while (reader.Read())
                {
                    aux = new Celula();
                    aux.claveCelula = int.Parse(reader["cveCelula"].ToString());
                    aux.nombreCelula = reader["nomCelula"].ToString();
                    aux.Ncelula = int.Parse(reader["numCelula"].ToString());
                    list.Add(aux);
                }
                reader.Close();
                connection.Close();
            }
        }

        public void DatosComisiones(List<Comisiones> list, string path)
        {

            String stringConexion = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path;
            Console.WriteLine("EL PATH ES !!!! " + path);
            using (OleDbConnection connection = new OleDbConnection(stringConexion))
            {
                connection.Open();
                OleDbDataReader reader = null;
                OleDbCommand command = new OleDbCommand("select * FROM PAQUETES", connection);
                reader = command.ExecuteReader();
                Comisiones aux;
                while (reader.Read())
                {
                    aux = new Comisiones();
                    aux.tablasComisiones = reader["TABLACOMISIONES"].ToString();

                    aux.comisionTotal = float.Parse(reader["COMISIONTOTAL"].ToString());

                    aux.comisionPromotor = float.Parse(reader["COMISIONPROMOTOR"].ToString());
                    aux.comisionGerente = float.Parse(reader["COMISIONGERENTE"].ToString());
                    aux.paquetes = reader["PAQUETES"].ToString();

                    list.Add(aux);
                }
                reader.Close();

                connection.Close();
            }
        }

    }
}

