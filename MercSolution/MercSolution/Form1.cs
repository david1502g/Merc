using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace MercSolution
{
    public partial class Form1 : Form
    {
        public int controlDeCasos;
        public PantallaDeCarga pantallaCarga = new PantallaDeCarga();
        public bool listo;
        string path,pathV,pathA, pathGlobaV, pathGlobaF;
        
        string destPath, destPath1;
        ManipuladorExcel manipulador;
        ManipuladorExcel manipuladorV;// nueva variable
        ManipuladorAcces manipuladorA;
        VentanaPromotores VentanaPromotores;
        VentanaCelulas VentanaCelulas;
        VentanaComisiones VentanaComisiones;
        VentanaBusqueda ventanaBus;

        private System.Data.DataTable dt;

        public Dictionary<int, Paquete> promotores { get; set; }
        public Dictionary<int, Paquete> promotoresV { get; set; }
        public Dictionary<string, int> nombrePaquetes { get; set; }
        public Dictionary<string, int> nombrePaquetesV { get; set; }
        public List<Promotor> promotoresN { get; set; }
        public List<Promotor> promotoresNV { get; set; }
        public List<Celula> celulasN { get; set; }//Parece que sirve para la nomina fija.
        public List<Celula> celulasNV { get; set; }
        public List<Comisiones> ComisionesN { get; set; }
        public List<Comisiones> ComisisonesNV { get; set; }
        public List<Promotor> promotoresNaux;
        public List<Promotor> promotoresNVaux;
        public Dictionary<int, int> diccionarioNomFija { set; get; }
        public Dictionary<int, int> diccionarioNomFijaV { set; get; }
        public double totalNominaFIja = 0;

        //Constructor posteos
        public Form1(string path)
        {
            pathGlobaV = eliminarSlash(path);
            controlDeCasos = 1;
            pantallaCarga.Show();
            pantallaCarga.aumentarCarga(10);

            InitializeComponent();
            listo = false;

            promotores = new Dictionary<int, Paquete>();
            diccionarioNomFija = new Dictionary<int, int>();
            this.path = path;
            pathA = Directory.GetCurrentDirectory() + "\\PROMOTORES.accdb";
            //this.pathA = pathA;

            ComisionesN = new List<Comisiones>();
            Console.WriteLine("VOY A AUMENTAR LA BARRA");
            pantallaCarga.aumentarCarga(10);
            string pathComisiones = Directory.GetCurrentDirectory() + "\\comisiones1.accdb";

            promotoresN = new List<Promotor>();
            celulasN = new List<Celula>();

            manipuladorA = new ManipuladorAcces(pathA);
            manipuladorA.DatosPromotorN(promotoresN);
            manipuladorA.DatosComisiones(ComisionesN, pathComisiones);
            Console.WriteLine("Este: " + pathA);
            manipuladorA.celulasN(celulasN, Directory.GetCurrentDirectory() + "\\CELULASN.accdb");
            pantallaCarga.aumentarCarga(10);
            manipulador = new ManipuladorExcel(path, promotoresN);
            pantallaCarga.aumentarCarga(30);
            manipulador.SepararEstrategias("test");
            foreach(string i in manipulador.Estrageias)
            {
                comboBox2.Items.Add(i);
            }
            comboBox2.SelectedItem = manipulador.Estrageias.ElementAt(0);
            string pathAux = manipulador.Copiar("F");

            //tabla.DataSource = manipuladorA.conexion();
            manipuladorA.DatosPromotor(manipulador.promotores);
            //manipuladorA.DatosPromotor(manipulador.promotoresVariable);

            //manipulador.ImprimirPromo();
            //CHECAR
            manipulador.AñadirDatos(manipulador.promotores[comboBox2.SelectedIndex]);
            //manipulador.crearRespaldo(2);
            manipulador.generaRespaldo1();

            //manipulador.AñadirDatos(promotoresN);
            //manipuladorA.nombresPromotor();
            //
            pantallaCarga.aumentarCarga(10);
            destPath = pathAux;
            destPath1 = pathAux;
            manipulador.nominaFijaExcel(destPath, manipulador.promotores);
            string excelConectionConfig;
            excelConectionConfig = "Provider=Microsoft.ACE.OLEDB.12.0; ";
            excelConectionConfig += "Data Source =" + pathAux + "; ";
            excelConectionConfig += "Extended Properties=\"Excel 12.0; HDR=YES\" ";

            tabla.Columns.Clear();
            OleDbConnection excelConnection = default(OleDbConnection);
            excelConnection = new OleDbConnection(excelConectionConfig);
            OleDbCommand filterRows = default(OleDbCommand);
            filterRows = new OleDbCommand("Select * From [Hoja1$]", excelConnection);
            excelConnection.Open();
            pantallaCarga.aumentarCarga(10);
            DataSet ds = new DataSet();

            try
            {
                OleDbDataAdapter adaptador = default(OleDbDataAdapter);
                adaptador = new OleDbDataAdapter();
                adaptador.SelectCommand = filterRows;
                adaptador.Fill(ds);

                tabla.DataSource = ds.Tables[0];
                excelConnection.Close();


            }
            catch (Exception e2)
            {
                Console.WriteLine(e2.ToString());
            }
            pantallaCarga.aumentarCarga(10);
            //
            /*
            List<Promotor> lista = (from p in promotoresN
                                    group p by new { p.claveInter, p.clavePromotor, p.estrategia, p.nombrePromotor, p.celula, p.banco, p.nominaFija, p.nominaVariable } into grupo
                                    where grupo.Count() > 1
                                    select new Promotor()
                                    {
                                        claveInter = grupo.Key.claveInter,
                                        clavePromotor = grupo.Key.clavePromotor,
                                        estrategia = grupo.Key.estrategia,
                                        nombrePromotor = grupo.Key.nombrePromotor,
                                        celula = grupo.Key.celula,
                                        banco = grupo.Key.banco,
                                        nominaFija = grupo.Key.nominaFija,
                                        nominaVariable = grupo.Key.nominaVariable
                                    }).ToList();
            promotoresN = lista;
            */
            /*for (int i = 0; i < tabla.Rows.Count - 2; i++)
            {

                //Console.WriteLine("WEA:  " + tabla[2, i].Value.ToString());
                if (tabla[2, i].Value.ToString().Length>0)
                { 
                    int val = int.Parse(tabla[2, i].Value.ToString());
                    if (!promotores.ContainsKey(val))
                    {
                        promotores.Add(val, new Paquete(manipulador.nomPaquetes, tabla[3, i].Value.ToString()));
                        promotores[val].aumentar(tabla[5, i].Value.ToString());
                    }
                    else
                    {
                    promotores[val].aumentar(tabla[5, i].Value.ToString());
                    }
                }
            }*/

            promotoresNaux = new List<Promotor>(promotoresN);
            pantallaCarga.aumentarCarga(10);
            for (int i = 0; i < tabla.Rows.Count - 2; i++)
            {
                int contadorDeIngresos = 0;
                //Console.WriteLine("WEA:  " + tabla[2, i].Value.ToString());
                if (tabla[2, i].Value.ToString().Length > 0)
                {

                    int val = int.Parse(tabla[2, i].Value.ToString());
                    bool banderaDePosteo = false;
                    foreach (Promotor j in promotoresN)
                    {
                        if (j.clavePromotor == val)
                        {
                            banderaDePosteo = true;
                            string statusPisa = tabla[8, i].Value.ToString();
                            if (statusPisa.Equals("ABIERTA") || statusPisa.Equals("POSTEADA") || statusPisa.Equals(" ABIERTA") || statusPisa.Equals(" POSTEADA") || statusPisa.Equals("ABIERTA ") || statusPisa.Equals("POSTEADA "))
                            {
                                contadorDeIngresos += 1;
                                // Console.WriteLine("El estatus pisa de " + val + " es " + tabla[8, i].Value.ToString() +" Y lleva "+contadorDeIngresos+ " ingresos");
                                int x = promotoresNaux.IndexOf(j);
                                promotoresNaux.ElementAt(x).ingresos += 1;
                            }
                        }

                    }

                    if (!banderaDePosteo)
                    {
                        Console.WriteLine("Un promotor en el excel, no se encuentra en el acces");
                    }


                }
            }
            double nFija;
            //Lenado de la nomina fija
            foreach (Promotor i in promotoresNaux)
            {
                if (i.ingresos > 0)
                {
                    int x = promotoresN.IndexOf(i);
                    promotoresN.ElementAt(x).ingresos = i.ingresos;
                    if (i.estrategia >= 5)
                    {
                        nFija =  i.ingresos * 138.64;
                        promotoresN.ElementAt(x).nominaFija = nFija;


                    }
                    else
                    {
                        nFija = i.ingresos * 115.38;
                        promotoresN.ElementAt(x).nominaFija = nFija;
                    }
                    totalNominaFIja += nFija;
                    Console.WriteLine("El promotor" + i.nombrePromotor + " Realizó " + i.ingresos + " ingresos");
                }


            }

            //Console.WriteLine("Para el promotor MORALES PEREZ SAMUEL");
            //GenerarFormato();

            Paquetes.Columns.Clear();
            
            string excelConectionConfig1;
            excelConectionConfig1 = "Provider=Microsoft.ACE.OLEDB.12.0; ";
            excelConectionConfig1 += "Data Source =" + pathAux + "; ";
            excelConectionConfig1 += "Extended Properties=\"Excel 12.0; HDR=YES\" ";

            OleDbConnection excelConnection1 = default(OleDbConnection);
            excelConnection1 = new OleDbConnection(excelConectionConfig1);
            OleDbCommand filterRows1 = default(OleDbCommand);
            filterRows1 = new OleDbCommand("Select * From [Hoja2$]", excelConnection1);
            excelConnection1.Open();

            DataSet ds1 = new DataSet();

            try
            {
                OleDbDataAdapter adaptador1 = default(OleDbDataAdapter);
                adaptador1 = new OleDbDataAdapter();
                adaptador1.SelectCommand = filterRows1;
                adaptador1.Fill(ds1);

                Paquetes.DataSource = ds1.Tables[0];
                excelConnection1.Close();
            }
            catch (Exception e2)
            {
                Console.WriteLine(e2.ToString());
            }
            listo = true;
            pantallaCarga.Close();
        }
        
        //Constructor posteos e ingresos
        public Form1(string pathI, string pathE)
        {
            pathGlobaF = eliminarSlash(pathI);
            pathGlobaV = eliminarSlash(pathE);
            controlDeCasos = 2;
            pantallaCarga.Show();
            pantallaCarga.aumentarCarga(10);
            listo = false;
            InitializeComponent();
            Console.WriteLine("VOY A AUMENTAR LA BARRA");
            pathA = Directory.GetCurrentDirectory() + "\\PROMOTORES.accdb";
            //this.pathA = pathA;
            
            promotores = new Dictionary<int, Paquete>();
            promotoresV = new Dictionary<int, Paquete>();
            diccionarioNomFija = new Dictionary<int, int>();
            diccionarioNomFijaV = new Dictionary<int, int>();
            this.path = pathI;
            this.pathV = pathE;

            pantallaCarga.aumentarCarga(10);
            ComisionesN = new List<Comisiones>();
            ComisisonesNV = new List<Comisiones>();

            string pathComisiones = Directory.GetCurrentDirectory() + "\\comisiones1.accdb";

            promotoresN = new List<Promotor>();
            promotoresNV = new List<Promotor>();
            celulasN = new List<Celula>();
            celulasNV = new List<Celula>();

            manipuladorA = new ManipuladorAcces(pathA);
            manipuladorA.DatosPromotorN(promotoresN);
            manipuladorA.DatosPromotorN(promotoresNV);
            manipuladorA.DatosComisiones(ComisionesN, pathComisiones);
            manipuladorA.DatosComisiones(ComisisonesNV, pathComisiones);
            Console.WriteLine("Este: " + pathA);
            manipuladorA.celulasN(celulasN, Directory.GetCurrentDirectory() + "\\CELULASN.accdb");

            manipuladorV = new ManipuladorExcel(pathV, promotoresNV);
            pantallaCarga.aumentarCarga(20);
            manipuladorV.SepararEstrategias("testV");
            manipulador = new ManipuladorExcel(path, promotoresN);
            manipulador.SepararEstrategias("test");
            pantallaCarga.aumentarCarga(20);

            foreach (string i in manipulador.Estrageias)
            {
                comboBox2.Items.Add(i);
            }

            foreach (string i in manipuladorV.Estrageias)
            {
                comboPosteos.Items.Add(i);
            }
            
            comboBox2.SelectedItem = manipulador.Estrageias.ElementAt(0);
            comboPosteos.SelectedItem = manipuladorV.Estrageias.ElementAt(0);
            string pathAux = manipulador.Copiar("F");
            string pathAuxV = manipuladorV.Copiar("V");

            pantallaCarga.aumentarCarga(10);

            //tabla.DataSource = manipuladorA.conexion();
            manipuladorA.DatosPromotor(manipulador.promotores);
            manipuladorA.DatosPromotor(manipuladorV.promotores);
            //manipuladorA.DatosPromotor(manipulador.promotoresVariable);
            //manipuladorA.DatosPromotor(manipuladorV.promotoresVariable);

            //manipulador.ImprimirPromo();
            //CHECAR
            manipuladorV.AñadirDatos(manipuladorV.promotores[comboBox2.SelectedIndex]);
            //manipuladorV.crearRespaldo(2);
            manipuladorV.generaRespaldo1();
            manipulador.AñadirDatos(manipulador.promotores[comboBox2.SelectedIndex]);
            //manipulador.crearRespaldo(1);
            manipulador.generaRespaldo1();

            //manipulador.AñadirDatos(promotoresN);
            //manipuladorA.nombresPromotor();

            manipulador.AñadirMontoFijo(pathA);
            //manipuladorV.AñadirMontoFijo(pathA);
            manipuladorV.CalcularNominaVariable();
            manipuladorV.calcularVariable();
            manipulador.calcularNominaFija();

            destPath = pathAux;
            destPath1 = pathAuxV;
            string excelConectionConfig;
            excelConectionConfig = "Provider=Microsoft.ACE.OLEDB.12.0; ";
            excelConectionConfig += "Data Source =" + pathAux + "; ";
            excelConectionConfig += "Extended Properties=\"Excel 12.0; HDR=YES\" ";

            tabla.Columns.Clear();
            OleDbConnection excelConnection = default(OleDbConnection);
            excelConnection = new OleDbConnection(excelConectionConfig);
            OleDbCommand filterRows = default(OleDbCommand);
            filterRows = new OleDbCommand("Select * From [Hoja" + (comboBox2.SelectedIndex+1) + "$]", excelConnection);
            excelConnection.Open();

            DataSet ds = new DataSet();
            pantallaCarga.aumentarCarga(20);
            try
            {
                OleDbDataAdapter adaptador = default(OleDbDataAdapter);
                adaptador = new OleDbDataAdapter();
                adaptador.SelectCommand = filterRows;
                adaptador.Fill(ds);

                tabla.DataSource = ds.Tables[0];
                excelConnection.Close();


            }
            catch (Exception e2)
            {
                Console.WriteLine(e2.ToString());
            }

            promotoresNaux = new List<Promotor>(promotoresN);

            for (int i = 0; i < tabla.Rows.Count - 2; i++)
            {
                int contadorDeIngresos = 0;
                //Console.WriteLine("WEA:  " + tabla[2, i].Value.ToString());
                if (tabla[2, i].Value.ToString().Length > 0)
                {

                    int val = int.Parse(tabla[2, i].Value.ToString());
                    bool banderaDePosteo = false;
                    foreach (Promotor j in promotoresN)
                    {
                        if (j.clavePromotor == val)
                        {
                            banderaDePosteo = true;
                            string statusPisa = tabla[8, i].Value.ToString();
                            if (statusPisa.Equals("ABIERTA") || statusPisa.Equals("POSTEADA") || statusPisa.Equals(" ABIERTA") || statusPisa.Equals(" POSTEADA") || statusPisa.Equals("ABIERTA ") || statusPisa.Equals("POSTEADA "))
                            {
                                contadorDeIngresos += 1;
                                // Console.WriteLine("El estatus pisa de " + val + " es " + tabla[8, i].Value.ToString() +" Y lleva "+contadorDeIngresos+ " ingresos");
                                int x = promotoresNaux.IndexOf(j);
                                promotoresNaux.ElementAt(x).ingresos += 1;
                            }
                        }

                    }

                    if (!banderaDePosteo)
                    {
                        Console.WriteLine("Un promotor en el excel, no se encuentra en el acces");
                    }


                }
            }
            double nFija;
            //Lenado de la nomina fija
            foreach (Promotor i in promotoresNaux)
            {
                if (i.ingresos > 0)
                {
                    int x = promotoresN.IndexOf(i);
                    promotoresN.ElementAt(x).ingresos = i.ingresos;
                    if (i.estrategia >= 5)
                    {
                        nFija = i.ingresos * 138.64;
                        promotoresN.ElementAt(x).nominaFija = nFija;


                    }
                    else
                    {
                        nFija = i.ingresos * 115.38;
                        promotoresN.ElementAt(x).nominaFija = nFija;
                    }
                    totalNominaFIja += nFija;
                    Console.WriteLine("El promotor" + i.nombrePromotor + " Realizó " + i.ingresos + " ingresos");
                }


            }

            //Console.WriteLine("Para el promotor MORALES PEREZ SAMUEL");
            //GenerarFormato();

            /*
            Paquetes.Columns.Clear();

            string excelConectionConfig1;
            excelConectionConfig1 = "Provider=Microsoft.ACE.OLEDB.12.0; ";
            excelConectionConfig1 += "Data Source =" + pathAux + "; ";
            excelConectionConfig1 += "Extended Properties=\"Excel 12.0; HDR=YES\" ";

            OleDbConnection excelConnection1 = default(OleDbConnection);
            excelConnection1 = new OleDbConnection(excelConectionConfig1);
            OleDbCommand filterRows1 = default(OleDbCommand);
            filterRows1 = new OleDbCommand("Select * From [Hoja2$]", excelConnection1);
            excelConnection1.Open();

            DataSet ds1 = new DataSet();

            try
            {
                OleDbDataAdapter adaptador1 = default(OleDbDataAdapter);
                adaptador1 = new OleDbDataAdapter();
                adaptador1.SelectCommand = filterRows1;
                adaptador1.Fill(ds1);

                Paquetes.DataSource = ds1.Tables[0];
                excelConnection1.Close();
            }
            catch (Exception e2)
            {
                Console.WriteLine(e2.ToString());
            }*/

            llenaPaquetes(comboBox2.Text);


            excelConectionConfig = "Provider=Microsoft.ACE.OLEDB.12.0; ";
            excelConectionConfig += "Data Source =" + pathAuxV + "; ";
            excelConectionConfig += "Extended Properties=\"Excel 12.0; HDR=YES\" ";

            tablaPosteos.Columns.Clear();
            excelConnection = new OleDbConnection(excelConectionConfig);
            filterRows = new OleDbCommand("Select * From [Hoja" + (comboBox2.SelectedIndex + 1) + "$]", excelConnection);
            excelConnection.Open();

            DataSet dsV = new DataSet();

            try
            {
                OleDbDataAdapter adaptador = default(OleDbDataAdapter);
                adaptador = new OleDbDataAdapter();
                adaptador.SelectCommand = filterRows;
                adaptador.Fill(dsV);

                tablaPosteos.DataSource = dsV.Tables[0];
                excelConnection.Close();


            }
            catch (Exception e2)
            {
                Console.WriteLine(e2.ToString());
            }
            listo = true;
            pantallaCarga.Close();
        }
        
        //Constructor ingresos
        public Form1(string pathE, int a)
        {
            pathA = Directory.GetCurrentDirectory() + "\\PROMOTORES.accdb";
            pathGlobaF = eliminarSlash(pathE);
            InitializeComponent();
            controlDeCasos = 3;
            pantallaCarga.Close();
        }

        public void CambiarTabla(string id, DataGridView tabla, string hoja)
        {
            string realPath = "";
            switch (id)
            {
                case "F":
                    realPath = destPath;
                    break;
                case "V":
                    realPath = destPath1;
                    break;
            }

            string excelConectionConfig;
            excelConectionConfig = "Provider=Microsoft.ACE.OLEDB.12.0; ";
            excelConectionConfig += "Data Source =" + realPath + "; ";
            excelConectionConfig += "Extended Properties=\"Excel 12.0; HDR=YES\" ";

            tabla.Columns.Clear();
            OleDbConnection excelConnection = default(OleDbConnection);
            excelConnection = new OleDbConnection(excelConectionConfig);
            OleDbCommand filterRows = default(OleDbCommand);
            filterRows = new OleDbCommand("Select * From [" + hoja + "]", excelConnection);
            excelConnection.Open();

            DataSet ds = new DataSet();

            try
            {
                OleDbDataAdapter adaptador = default(OleDbDataAdapter);
                adaptador = new OleDbDataAdapter();
                adaptador.SelectCommand = filterRows;
                adaptador.Fill(ds);

                tabla.DataSource = ds.Tables[0];
                excelConnection.Close();


            }
            catch (Exception e2)
            {
                Console.WriteLine(e2.ToString());
            }
        }

        public void llenaPaquetes(string nombre)
        {

            Excel.Application excelApplication = new Excel.Application();
            excelApplication.DisplayAlerts = false;
            Excel.Workbook destworkBook;
            Excel.Worksheet destworkSheet;
            destworkBook = excelApplication.Workbooks.Open(pathGlobaV + nombre + ".xlsx");
            
            destworkSheet = (Microsoft.Office.Interop.Excel.Worksheet)destworkBook.Sheets.get_Item(1);
            
            int n = destworkSheet.Cells[1, 1].End(Excel.XlDirection.xlDown).Row;
            Console.WriteLine("MEGA FURRO: "+ n);

            dt = new System.Data.DataTable();
            Console.WriteLine("Mega Furro 4: " + destworkSheet.Cells[n + 6, 4].Value.ToString());
            dt.Columns.Add(destworkSheet.Cells[n + 6, 4].Value.ToString());
            Console.WriteLine("Mega Furro 5: " + destworkSheet.Cells[n + 6, 5].Value.ToString());
            dt.Columns.Add(destworkSheet.Cells[n + 6, 5].Value.ToString());
            Console.WriteLine("Mega Furro 6: " + destworkSheet.Cells[n + 6, 6].Value.ToString());
            dt.Columns.Add(destworkSheet.Cells[n + 6, 6].Value.ToString());
            Paquetes.DataSource = dt;
            DataRow row;
            for (int i = n+7; i <= destworkSheet.UsedRange.Rows.Count; i++)
            {
                row = dt.NewRow();
                row[destworkSheet.Cells[n + 6, 4].Value.ToString()] = destworkSheet.Cells[i, 4].Value.ToString();
                row[destworkSheet.Cells[n + 6, 5].Value.ToString()] = destworkSheet.Cells[i, 5].Value.ToString();
                row[destworkSheet.Cells[n + 6, 6].Value.ToString()] = destworkSheet.Cells[i, 6].Value.ToString();
                dt.Rows.Add(row);
            }
           
            destworkBook.Close(true);
            excelApplication.Quit();

        }

        public void escribirTablas(string paths)
        {
            destPath = paths;

            string excelConectionConfig;
            excelConectionConfig = "Provider=Microsoft.ACE.OLEDB.12.0; ";
            excelConectionConfig += "Data Source =" + destPath + "; ";
            excelConectionConfig += "Extended Properties=\"Excel 12.0; HDR=YES\" ";

            tabla.Columns.Clear();
            OleDbConnection excelConnection = default(OleDbConnection);
            excelConnection = new OleDbConnection(excelConectionConfig);
            OleDbCommand filterRows = default(OleDbCommand);
            filterRows = new OleDbCommand("Select * From [Hoja"+ (comboBox2.SelectedIndex + 1) + "$]", excelConnection);
            excelConnection.Open();

            DataSet ds = new DataSet();

            try
            {
                OleDbDataAdapter adaptador = default(OleDbDataAdapter);
                adaptador = new OleDbDataAdapter();
                adaptador.SelectCommand = filterRows;
                adaptador.Fill(ds);

                tabla.DataSource = ds.Tables[0];
                excelConnection.Close();


            }
            catch (Exception e2)
            {
                Console.WriteLine(e2.ToString());
            }
        }

        public void GenerarFormato()
        {
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook book = xlApp.Workbooks.Open(destPath);
            Excel.Worksheet sheet = book.Worksheets.Add(After: book.Sheets[book.Sheets.Count]);

            int j = 1;
            int columnNum = 2;

            sheet.Cells[3, 1].Value = "PROMOTOR";
            foreach (string i in manipulador.nomPaquetes)
            {
                sheet.Cells[2, columnNum].Value = i;
                sheet.Cells[3, columnNum].Value = "PAQ,";
                sheet.Cells[3, columnNum + 1].Value = "PAGO";

                columnNum += 2;
            }

            sheet.Cells[2, columnNum].Value = "Total PAQ,";
            sheet.Cells[2, columnNum + 1].Value = "Total PAGO";

            int row = 4;
            int col = 0;
            foreach (KeyValuePair<int, Paquete> i in promotores)
            {
                // do something with entry.Value or entry.Key
                sheet.Cells[row, 1].Value = i.Value.nombreP;
                col = 2;
                foreach (KeyValuePair<string, int> k in i.Value.diccionarioPaquetes)
                {
                    sheet.Cells[row, col].Value = k.Value;
                    col += 2;
                }
                sheet.Cells[row, col].Value = i.Value.totalPaquetes();
                row++;
            }
            /*foreach (Promotor i in manipulador.promotores)
            {
                sheet.Cells[row, 1].Value = i.nombrePromotor;
                row++;
            }*/


            book.Save();
            book.Close();
            xlApp.Quit();
            
        }
        private void guardar_Click(object sender, EventArgs e)
        {
            
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            Application.Exit();

        }

        

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (listo)
            {
                string strg = "Hoja" + (comboBox2.SelectedIndex + 1) + "$";
                CambiarTabla("F", tabla, strg);
                llenaPaquetes(comboBox2.Text);
                //CambiarTabla("V", tablaPosteos);

                /*int aux = 0;
                foreach (string i in manipulador.Estrageias)
                {
                    if (comboBox2.SelectedItem.Equals(i))
                    {
                        //Console.WriteLine("Libro a escribir: "+ manipulador.Copiar());
                        escribirTablas(manipulador.pathEstrageias.ElementAt(aux));
                    }
                    aux++;
                }*/
            }    
        }
            

        private void tabla_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void abrirVentanaPromotor_Click(object sender, EventArgs e)
        {
            VentanaPromotores = new VentanaPromotores(promotoresN, pathA);
            VentanaPromotores.Show();
        }

        private void celulas_Click(object sender, EventArgs e)
        {
            
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {

        }

        private void toolStripMenuItem2_Click(object sender, EventArgs e)
        {
            VentanaPromotores = new VentanaPromotores(promotoresN, pathA);
            VentanaPromotores.Show();
        }

        private void celulasToolStripMenuItem_Click(object sender, EventArgs e)
        {
            VentanaCelulas = new VentanaCelulas(celulasN, Directory.GetCurrentDirectory() + "\\CELULASN.accdb", pathA, promotoresN);
            VentanaCelulas.Show();
        }

        private void promotoresToolStripMenuItem_Click(object sender, EventArgs e)
        {
            VentanaPromotores = new VentanaPromotores(promotoresN, pathA);
            VentanaPromotores.Show();
        }

        private void celulasToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            VentanaCelulas = new VentanaCelulas(celulasN, Directory.GetCurrentDirectory() + "\\CELULASN.accdb", pathA, promotoresN);
            VentanaCelulas.Show();
        }

        private void guardarToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.SaveFileDialog saveDlg = new System.Windows.Forms.SaveFileDialog();
            saveDlg.InitialDirectory = @"C:\";
            saveDlg.Filter = "Excel files (*.xlsx)|*.xlsx";
            saveDlg.FilterIndex = 0;
            saveDlg.RestoreDirectory = true;
            saveDlg.Title = "Export Excel File To";
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(destPath);

            if (saveDlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string path = saveDlg.FileName;
                xlWorkBook.SaveCopyAs(saveDlg.FileName);
                xlWorkBook.Saved = true;
                xlWorkBook.Close(true, Type.Missing, Type.Missing);
                xlApp.Quit();
            }
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }

        private void archivoToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void buscarPorFolioSIACToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ventanaBus = new VentanaBusqueda(Directory.GetCurrentDirectory() + "\\respaldos.accdb");
            ventanaBus.Show();
        }

        private void paquetesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            VentanaComisiones = new VentanaComisiones(ComisionesN, pathA);
            VentanaComisiones.Show();
        }

        private void comboPosteos_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (listo)
            {
                string strg = "Hoja" + (comboPosteos.SelectedIndex + 1) + "$";
                CambiarTabla("V", tablaPosteos, strg);
            }
        }

        private void paquetesToolStripMenuItem1_Click(object sender, EventArgs e)
        {

        }

        private void Recalcular_Click(object sender, EventArgs e)
        {
            this.Close();
            if (controlDeCasos == 1)
            {
                Form1 form1 = new Form1(path);
                form1.Show();
            }
            else if (controlDeCasos == 2)
            {
                Form1 form1 = new Form1(pathV, path);
                form1.Show();
            }
            else
            {
                Form1 form1 = new Form1(pathV, 1);
                form1.Show();
            }


        }

        private void Paquetes_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void tablaAcces_Click(object sender, EventArgs e)
        {

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
    }
}
