using Microsoft.Office.Interop.Excel;
using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Data.OleDb;

namespace MercSolution
{
    class ManipuladorExcel
    {
        string path;
        string destPath;
        public string[] ModelID { get; }
        string[] Celda { get; }
        int Count;
        public List<Promotor>[] promotores { get; set; }

        public List<Promotor> promotoresN { get; set; }
        public List<string> nomPaquetes { get; set; }
        public List<Promotor>[] promotoresVariable { get; set; }
        public List<string> pathEstrageias { get; set; }
        public List<string> Estrageias { get; set; }
        Paquete packs;

        //Inicio de variables para Nomina Variable
        Dictionary<string, Dictionary<string, double>> dic;
        Dictionary<string, Dictionary<string, int>> dicAux;
        Dictionary<string, int> bestos;
        List<string> fullPacks;
        int actMin;
        string globalBestType;
        string globalBestPack;

        int totalRows;
        int aciertos;
        //Fin de Variable para nomina variable.
        public ManipuladorExcel(string nombreArchivo, List<Promotor> promotores)
        {
            nomPaquetes = new List<string>();
            pathEstrageias = new List<string>();
            Estrageias = new List<string>();

            dic = new Dictionary<string, Dictionary<string, double>>();
            dicAux = new Dictionary<string, Dictionary<string, int>>();
            bestos = new Dictionary<string, int>();
            fullPacks = new List<string>();
            actMin = 99;
            totalRows = 0;
            aciertos = 0;
            //this.promotores = new List<Promotor>();
            //promotoresVariable = new List<Promotor>();
            promotoresN = promotores;
            path = nombreArchivo;
            packs = new Paquete();
            Count = 0;
        }

        public string Copiar(string nom)
        {
            /*string test = "";
            switch (nom)
            {
                case "F":
                    test = "test";
                    break;
                case "V":
                    test = "testV";
                    break;
            }*/
            string pathdes = eliminarSlash(path);

            Excel.Workbook srcworkBook;
            Excel.Worksheet srcworkSheet;
            Excel.Workbook destworkBook;
            Excel.Worksheet destworkSheet;
            Excel.Worksheet destworkSheet2;
            string srcPath;
            Excel.Application excelApplication = new Excel.Application();
            excelApplication.DisplayAlerts = false;
            //Opening of first worksheet and copying
            srcPath = path;
            srcworkBook = excelApplication.Workbooks.Open(srcPath);

            srcworkSheet = (Microsoft.Office.Interop.Excel.Worksheet)srcworkBook.Sheets.get_Item(1);

            destPath = pathdes + "MercTest2" + nom + ".xlsx";
            destworkBook = excelApplication.Workbooks.Add();
            destworkBook.Worksheets[1].Cells[1, 1].Value = "100";
            destworkBook.SaveAs(pathdes + "MercTest2" + nom + ".xlsx", Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlExclusive,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            destworkBook.Close(true);
            destworkBook = excelApplication.Workbooks.Open(destPath);
            for (int i = 0; i < pathEstrageias.Count - 1; i++)
            {
                destworkBook.Worksheets.Add(After: destworkBook.Sheets[destworkBook.Sheets.Count]);
            }

            List<String> columnasFuente = new List<string>();
            columnasFuente.Add("A");
            columnasFuente.Add("B");
            columnasFuente.Add("C");
            columnasFuente.Add("D");
            columnasFuente.Add("K");
            columnasFuente.Add("AU");
            columnasFuente.Add("U");
            columnasFuente.Add("O");
            columnasFuente.Add("W");
            columnasFuente.Add("X");
            columnasFuente.Add("AV");
            columnasFuente.Add("AW");
            columnasFuente.Add("AY");

            promotores = new List<Promotor>[pathEstrageias.Count];
            promotoresVariable = new List<Promotor>[pathEstrageias.Count];

            for (int k=1; k <= pathEstrageias.Count; k++)
            {
                promotores[k - 1] = new List<Promotor>();
                promotoresVariable[k - 1] = new List<Promotor>();

                srcworkBook = excelApplication.Workbooks.Open(pathEstrageias.ElementAt(k-1));
                srcworkSheet = (Microsoft.Office.Interop.Excel.Worksheet)srcworkBook.Sheets.get_Item(1);

                destworkSheet = (Microsoft.Office.Interop.Excel.Worksheet)destworkBook.Sheets.get_Item(k);
                //destworkSheet2 = (Microsoft.Office.Interop.Excel.Worksheet)destworkBook.Sheets.get_Item(2);

                int j = 1;

                foreach (string i in columnasFuente)
                {
                    Excel.Range sourceRange = srcworkSheet.get_Range(i + "1", i + "" + (srcworkSheet.UsedRange.Rows.Count));
                    Excel.Range destinationRange = destworkSheet.get_Range(calcularCelda(j) + "1", calcularCelda(j) + "" + (srcworkSheet.UsedRange.Rows.Count));
                    //Excel.Range destinationRange2 = destworkSheet2.get_Range(calcularCelda(j) + "1", calcularCelda(j) + "" + (srcworkSheet.UsedRange.Rows.Count - 1));
                    sourceRange.Copy(Type.Missing);

                    destinationRange.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteValues, Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);//Copiar Valores
                    destinationRange.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormats, Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);//copiar formato de columna
                    destinationRange.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteColumnWidths, Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);//copiar valores de columna

                    //destinationRange2.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteValues, Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);//Copiar Valores
                    //destinationRange2.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormats, Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);//copiar formato de columna
                    //destinationRange2.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteColumnWidths, Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);//copiar valores de columna

                    if (i.Equals("C"))
                    {
                        j++;
                        destworkSheet.Columns[j].ColumnWidth = 50.0;
                        //destworkSheet2.Columns[j].ColumnWidth = 50.0;
                        int aux = RowToInt(i);
                        for (int h = 2; h <= srcworkSheet.UsedRange.Rows.Count; h++)
                        {
                            double value = srcworkSheet.Cells[h, aux].Value;
                            promotores[k-1].Add(new Promotor((int)value));
                        }
                    }

                    if (i.Equals("K"))
                    {
                        int aux = RowToInt(i);
                        for (int h = 2; h <= srcworkSheet.UsedRange.Rows.Count; h++)
                        {
                            string value = srcworkSheet.Cells[h, aux].Value.ToString();
                            nomPaquetes.Add(value);
                        }
                        nomPaquetes = nomPaquetes.Distinct().ToList();
                    }
                    j++;
                }

                srcworkBook.Close(false);
            }
            
            //FiltroNominaVariable(destworkSheet2);
            try
            {
                //destinationRange.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteValues, Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
                //destworkBook.SaveAs(pathdes+"MercTest2.xlsx", Type.Missing, Type.Missing,
                //Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlExclusive,
                //Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                destworkBook.Save();
            }
            catch (Exception e4)
            {
                MessageBox.Show(e4.ToString());
            }  
            finally
            {
                destworkBook.Close(true);
                
                excelApplication.Quit();
            }

            /*for (int k = 0; k < promotores.Count - 1; k += 50)
            {
                Console.WriteLine(promotores.ElementAt(k).clavePromotor);
            }*/

            return pathdes + "MercTest2" + nom + ".xlsx";
            //excelApplication.Quit();
        }

        public void nominaFijaExcel(string path, List<Promotor>[] promotores)
        {
            Excel.Application excelApplication = new Excel.Application();
            excelApplication.DisplayAlerts = false;
            Excel.Workbook destworkBook;
            Excel.Worksheet destworkSheet;
            destworkBook = excelApplication.Workbooks.Open(path);
            int hojas = promotores.Length;
            int n;
            Dictionary<string, double> diccionarioPromotores;

            for (int i = 1; i <= hojas; i++)
            {
                destworkSheet = (Microsoft.Office.Interop.Excel.Worksheet)destworkBook.Sheets.get_Item(i);
                n = destworkSheet.UsedRange.Rows.Count;
                diccionarioPromotores = new Dictionary<string, double>();
                for (int k = 0; k < promotores[i - 1].Count; k++)
                {

                    if (promotores[i - 1].ElementAt(k).Ncelula > 4)
                    {
                        if (!diccionarioPromotores.ContainsKey(promotores[i - 1].ElementAt(k).clavePromotor.ToString()))
                        {
                            diccionarioPromotores.Add(promotores[i - 1].ElementAt(k).clavePromotor.ToString(), 134.62);
                        }
                    }
                    else
                    {
                        if (!diccionarioPromotores.ContainsKey(promotores[i - 1].ElementAt(k).clavePromotor.ToString()))
                        {
                            diccionarioPromotores.Add(promotores[i - 1].ElementAt(k).clavePromotor.ToString(), 115.38);
                        }
                    }


                }

                destworkSheet.Cells[1, RowToInt("N")].Value = "MONTO PAGADO";
                for (int j = 2; j <= n; j++)
                {
                    destworkSheet.Cells[j, RowToInt("N")].Value = diccionarioPromotores[destworkSheet.Cells[j, 3].Value.ToString()];
                }
            }


            destworkBook.Save();
            destworkBook.Close(true);
            excelApplication.Quit();
        }

        public void generaRespaldo1() {

            Excel.Application excelApplication = new Excel.Application();
            excelApplication.DisplayAlerts = false;
            Excel.Workbook destworkBook;
            Excel.Worksheet destworkSheet;
            destworkBook = excelApplication.Workbooks.Open(destPath);
            int m = destworkBook.Sheets.Count;
            int aux = 0;
            int n = -1;
            string path = Directory.GetCurrentDirectory() + "\\respaldos.accdb";
            String stringConexion;
            string query="";
            for (int j = 1; j <= m; j++)
            {
                destworkSheet = (Microsoft.Office.Interop.Excel.Worksheet)destworkBook.Sheets.get_Item(j);
                stringConexion = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path;
                using (OleDbConnection connection = new OleDbConnection(stringConexion))
                {
                            connection.Open();
                            OleDbDataReader reader = null;
                            n = destworkSheet.UsedRange.Rows.Count;
                            for (int i = 2; i <= n; i++)
                            {
                        reader = null;
                        OleDbCommand command1 = new OleDbCommand("select * FROM Respaldo where Folio_SIAC = '" + destworkSheet.Cells[i, 5].Value.ToString() + "'");
                        command1.Connection = connection;
                        reader = command1.ExecuteReader();
                        if (!reader.Read())
                        {
                                    query = "";
                                    for(int h =1; h <= destworkSheet.UsedRange.Columns.Count; h++)
                                    {
                                        if(destworkSheet.Cells[i, h].Value == null)
                                        {
                                            destworkSheet.Cells[i, h].Value = " ";
                                        }
                                        
                                        if(h == destworkSheet.UsedRange.Columns.Count)
                                        {
                                            query = query + "'" + destworkSheet.Cells[i, h].Value.ToString()+"'";
                                        }
                                        else
                                        {
                                        query = query + "'" + destworkSheet.Cells[i, h].Value.ToString() + "',";
                                        }
                                    }
                                    //Console.WriteLine(query);
                                    //Console.WriteLine("INSERT into Respaldo (Fecha Captura, Estrategia, Promotor, Nombre Promotor, Folio SIAC, Paquete, Otros Servicios, Campana, Telefono Asignado, Estatus PISA Multiorden, Pisa OS Fecha POSTEO Multiorden, Entrego Expediente, Tipo Entrego Expediente, Semana) VALUES " + "(" + query + ")");
                                    OleDbCommand command = new OleDbCommand("INSERT into Respaldo (Fecha_Captura, Estrategia, Promotor, Nombre_Promotor, Folio_SIAC, Paquete, Otros_Servicios, Campana, Telefono_Asignado, Estatus_PISA_Multiorden, Pisa_OS_Fecha_POSTEO_Multiorden, Entrego_Expediente, Tipo_Entrego_Expediente, Semana) VALUES " + "(" + query+  ")");
                                    command.Connection = connection;
                                    if (connection.State != ConnectionState.Open)
                                    {
                                        connection.Open();
                                    }

                                    try
                                    {
                                        command.ExecuteNonQuery();

                                    }
                                    catch (OleDbException ex)
                                    {
                                        MessageBox.Show("error al agregar en la BD");
                                        Console.WriteLine(ex.ToString());

                                    }
                                }/*
                                else
                                {
                                    aux++;
                                }*/

                            }
                            try
                            {
                                connection.Close();
                            }
                            catch (OleDbException ex)
                            {
                                Console.WriteLine(ex);
                                connection.Close();
                            }



                        }


                    }

                    destworkBook.Save();
                    destworkBook.Close(true);
                    excelApplication.Quit();

                    //Console.WriteLine("MIRA AQUI PORFA: n=" + n + " aux=" + aux);


                    if (aux == n - 1)
                    {
                       
                    }
                    else
                    {
                        
                    }
            


            }

        public void generaRespaldo()
        {

            Excel.Application excelApplication = new Excel.Application();
            excelApplication.DisplayAlerts = false;
            Excel.Workbook destworkBook;
            Excel.Worksheet destworkSheet;
            destworkBook = excelApplication.Workbooks.Open(destPath);
            string stringConexion;
            string path = Directory.GetCurrentDirectory() + "\\respaldos.accdb";
            int m = destworkBook.Sheets.Count;
            int n;
            try
            {
                for (int j = 1; j <= m; j++)
                {
                    destworkSheet = (Microsoft.Office.Interop.Excel.Worksheet)destworkBook.Sheets.get_Item(j);
                    stringConexion = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path;
                    using (OleDbConnection connection = new OleDbConnection(stringConexion))
                    {
                        connection.Open();
                        OleDbDataReader reader = null;
                        n = destworkSheet.UsedRange.Rows.Count;
                        for (int i = 2; i <= n; i++)
                        {
                            reader = null;
                            Console.WriteLine("FIJATE AQUI WEEE: " + destworkSheet.Cells[i, 5].Value.ToString());
                            OleDbCommand command1 = new OleDbCommand("select * FROM Respaldo where Folio_SIAC = '" + destworkSheet.Cells[i, 5].Value.ToString() + "'", connection);
                            //command1.Connection = connection;
                            reader = command1.ExecuteReader();
                            if (!reader.Read())
                            {
                                //"'" + nombre + "'," + "'" + destworkSheet.Cells[i, 5].Value.ToString() + "'," + "'" + destworkSheet.Cells[i, 1].Value.ToString() + "','" + destworkSheet.Cells[i, 14].Value.ToString()
                                OleDbCommand command = new OleDbCommand("INSERT into Respaldo (Fecha Captura, Estrategia, Promotor, Nombre Promotor, Folio SIAC, Paquete, Paquete, Otros Servicios, Campana, Telefono Asignado, Estatus PISA Multiorden, Pisa OS Fecha POSTEO Multiorden, Entrego Expediente, Semana) VALUES " + "(" + "'" + destworkSheet.Cells[i, 1].Value.ToString() + "'," + "'" + destworkSheet.Cells[i, 2].Value.ToString() + "'," + "'" + destworkSheet.Cells[i, 3].Value.ToString() + "'," + "'" + destworkSheet.Cells[i, 4].Value.ToString() + "'," + "'" + destworkSheet.Cells[i, 5].Value.ToString() + "'," + "'" + destworkSheet.Cells[i, 6].Value.ToString() + "'," + "'" + destworkSheet.Cells[i, 7].Value.ToString() + "'," + "'" + destworkSheet.Cells[i, 8].Value.ToString() + "'," + "'" + destworkSheet.Cells[i, 9].Value.ToString() + "'," + "'" + destworkSheet.Cells[i, 10].Value.ToString() + "'," + "'" + destworkSheet.Cells[i, 11].Value.ToString() + "'," + "'" + destworkSheet.Cells[i, 12].Value.ToString() + "'," + "'" + destworkSheet.Cells[i, 13].Value.ToString() + "'," + "'" + destworkSheet.Cells[i, 14].Value.ToString() + "')");
                                command.Connection = connection;
                                if (connection.State != ConnectionState.Open)
                                {
                                    connection.Open();
                                }

                                try
                                {
                                    command.ExecuteNonQuery();

                                }
                                catch (OleDbException ex)
                                {
                                    MessageBox.Show("error al agregar en la BD");
                                    Console.WriteLine(ex);

                                }
                            }
                        }
                        /*try
                        {
                            connection.Close();
                        }
                        catch (OleDbException ex)
                        {
                            Console.WriteLine(ex);
                            connection.Close();
                        }*/
                    }
                }
                //destworkBook.Save();               
                //excelApplication.Quit();
            }
            catch (Exception e4)
            {
                MessageBox.Show(e4.ToString());
            }
            finally
            {
                destworkBook.Close(true);
                excelApplication.Quit();
            }
        }

        public void crearRespaldo(int opc)
        {

            string DateFormat;
            string date;
            Excel.Workbook srcworkBook;
            Excel.Application excelApplication;


            

            switch (opc)
            {
                case 1:

                    DateFormat = "yyyy_MM_dd_HH_mm_ss";
                    date = DateTime.Now.ToString(DateFormat);
                    
                    excelApplication = new Excel.Application();
                    excelApplication.DisplayAlerts = false;
                    srcworkBook = excelApplication.Workbooks.Open(destPath);


                    try
                    {
                        srcworkBook.SaveAs(Directory.GetCurrentDirectory() + "\\PIPES\\I" + date + ".xlsx");

                    }
                    catch (Exception e4)
                    {
                        MessageBox.Show(e4.ToString());
                    }
                    finally
                    {
                        srcworkBook.Close(true);
                        excelApplication.Quit();

                    }

                    if (!LlenarRespaldo("I" + date + ".xlsx", Directory.GetCurrentDirectory() + "\\respaldos.accdb",1))
                    {
                        File.Delete(Directory.GetCurrentDirectory() + "\\PIPES\\I" + date + ".xlsx");
                    }

                    break;

                //respaldo posteos
                case 2:

                    DateFormat = "yyyy_MM_dd_HH_mm_ss";
                    date = DateTime.Now.ToString(DateFormat);

                    excelApplication = new Excel.Application();
                    excelApplication.DisplayAlerts = false;
                    srcworkBook = excelApplication.Workbooks.Open(destPath);


                    try
                    {
                        srcworkBook.SaveAs(Directory.GetCurrentDirectory() + "\\PIPES\\P" + date + ".xlsx");

                    }
                    catch (Exception e4)
                    {
                        MessageBox.Show(e4.ToString());
                    }
                    finally
                    {
                        srcworkBook.Close(true);
                        excelApplication.Quit();

                    }

                    if (!LlenarRespaldo("P" + date + ".xlsx", Directory.GetCurrentDirectory() + "\\respaldos.accdb",2))
                    {
                        File.Delete(Directory.GetCurrentDirectory() + "\\PIPES\\P" + date + ".xlsx");
                    }

                    break;            
            }
            
        }

        public bool LlenarRespaldo(string nombre, string path, int opc)
        {
            Excel.Application excelApplication = new Excel.Application();
            excelApplication.DisplayAlerts = false;
            Excel.Workbook destworkBook;
            Excel.Worksheet destworkSheet;
            destworkBook = excelApplication.Workbooks.Open(destPath);
            int m = destworkBook.Sheets.Count;
            int aux = 0;
            int n = -1;
            String stringConexion;

            switch (opc)
            {
                case 1:


                    for (int j = 1; j <= m; j++)
                    {
                        destworkSheet = (Microsoft.Office.Interop.Excel.Worksheet)destworkBook.Sheets.get_Item(j);
                        stringConexion = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path;
                        using (OleDbConnection connection = new OleDbConnection(stringConexion))
                        {
                            connection.Open();
                            OleDbDataReader reader = null;
                            n = destworkSheet.UsedRange.Rows.Count;
                            for (int i = 2; i <= n; i++)
                            {
                                reader = null;
                                OleDbCommand command1 = new OleDbCommand("select * FROM Respaldo_I where Folio_SIAC = '" + destworkSheet.Cells[i, 5].Value.ToString() + "'");
                                command1.Connection = connection;
                                reader = command1.ExecuteReader();
                                if (!reader.Read())
                                {
                                    OleDbCommand command = new OleDbCommand("INSERT into Respaldo_I (Nombre, Folio_SIAC, Fecha, Semana) VALUES " + "(" + "'" + nombre + "'," + "'" + destworkSheet.Cells[i, 5].Value.ToString() + "'," + "'" + destworkSheet.Cells[i, 1].Value.ToString() + "','" + destworkSheet.Cells[i, 14].Value.ToString() + "')");
                                    command.Connection = connection;
                                    if (connection.State != ConnectionState.Open)
                                    {
                                        connection.Open();
                                    }

                                    try
                                    {
                                        command.ExecuteNonQuery();

                                    }
                                    catch (OleDbException ex)
                                    {
                                        MessageBox.Show("error al agregar en la BD");
                                        Console.WriteLine(ex);

                                    }
                                }
                                else
                                {
                                    aux++;
                                }

                            }
                            try
                            {
                                connection.Close();
                            }
                            catch (OleDbException ex)
                            {
                                Console.WriteLine(ex);
                                connection.Close();
                            }



                        }


                    }

                    destworkBook.Save();
                    destworkBook.Close(true);
                    excelApplication.Quit();

                    //Console.WriteLine("MIRA AQUI PORFA: n=" + n + " aux=" + aux);


                    if (aux == n - 1)
                    {
                        return false;
                    }
                    else
                    {
                        return true;
                    }


                case 2:

                    for (int j = 1; j <= m; j++)
                    {
                        destworkSheet = (Microsoft.Office.Interop.Excel.Worksheet)destworkBook.Sheets.get_Item(j);
                        stringConexion = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path;
                        using (OleDbConnection connection = new OleDbConnection(stringConexion))
                        {
                            connection.Open();
                            OleDbDataReader reader = null;
                            n = destworkSheet.UsedRange.Rows.Count;
                            for (int i = 2; i <= n; i++)
                            {
                                reader = null;
                                OleDbCommand command1 = new OleDbCommand("select * FROM Respaldo_P where Folio_SIAC = '" + destworkSheet.Cells[i, 5].Value.ToString() + "'");
                                command1.Connection = connection;
                                reader = command1.ExecuteReader();
                                if (!reader.Read())
                                {
                                    OleDbCommand command = new OleDbCommand("INSERT into Respaldo_P (Nombre, Folio_SIAC, Fecha, Semana) VALUES " + "(" + "'" + nombre + "'," + "'" + destworkSheet.Cells[i, 5].Value.ToString() + "'," + "'" + destworkSheet.Cells[i, 1].Value.ToString() + "','" + destworkSheet.Cells[i, 14].Value.ToString() + "')");
                                    command.Connection = connection;
                                    if (connection.State != ConnectionState.Open)
                                    {
                                        connection.Open();
                                    }

                                    try
                                    {
                                        command.ExecuteNonQuery();

                                    }
                                    catch (OleDbException ex)
                                    {
                                        MessageBox.Show("error al agregar en la BD");
                                        Console.WriteLine(ex);

                                    }
                                }
                                else
                                {
                                    aux++;
                                }

                            }
                            try
                            {
                                connection.Close();
                            }
                            catch (OleDbException ex)
                            {
                                Console.WriteLine(ex);
                                connection.Close();
                            }



                        }




                        //Console.WriteLine("MIRA AQUI PORFA: n=" + n + " aux=" + aux);


                    }

                    destworkBook.Save();
                    destworkBook.Close(true);
                    excelApplication.Quit();

                    if (aux == n - 1)
                    {
                        return false;
                    }
                    else
                    {
                        return true;
                    }

                default:

                    return false;



            }



        }

        public void FiltroNominaVariable(Excel.Worksheet worksheet)
        {
            int maxRows = worksheet.UsedRange.Rows.Count;

            for (int i = 2; i <= maxRows; i++)
            {
                if (worksheet.Cells[i, 10].Value == null || !worksheet.Cells[i, 10].Value.ToString().Equals("POSTEADA"))
                {
                    worksheet.get_Range("A" + i, "A" + i).EntireRow.Delete();

                    i--;
                    maxRows--;
                }
                else
                {
                    //promotoresVariable.Add(new Promotor((int)worksheet.Cells[i, 3].Value));
                }
            }
        }

        public string calcularCelda(int numero)
        {
            string columna = "";
            while (numero != 0)
            {
                int res = numero % 26;
                res += 64;
                char c = (char)res;
                columna += "" + c;
                numero /= 26;
            }
            return Reverse(columna);
        }

        public static string Reverse(string s)
        {
            char[] charArray = s.ToCharArray();
            Array.Reverse(charArray);
            return new string(charArray);
        }

       

        public int RowToInt(string ejem)
        {
            int num = 0;
            int j = 0;
            for (int i = ejem.Length - 1; i >= 0; i--)
            {
                num += (ejem[i] - 64) * (int)Math.Pow(26, j);
                j++;
            }
            return num;
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

        public void AñadirDatos(List<Promotor> promotores)
        {
            Excel.Application excelApplication = new Excel.Application();
            excelApplication.DisplayAlerts = false;
            Excel.Workbook destworkBook;
            Excel.Worksheet destworkSheet;
            destworkBook = excelApplication.Workbooks.Open(destPath);
            
            
            for (int i=1; i <= pathEstrageias.Count; i++)
            {
                destworkSheet = (Microsoft.Office.Interop.Excel.Worksheet)destworkBook.Sheets.get_Item(i);

                int k = 1;
                destworkSheet.Cells[k, 4].Value = "Nombre Promotor";
                k++;
                
                foreach (Promotor j in this.promotores[i - 1])
                {
                    destworkSheet.Cells[k, 4].Value = j.nombrePromotor;
                    k++;
                }
            }
           
            /*Excel.Worksheet destworkSheet2;
            //Opening of first worksheet and copying
            //Console.WriteLine("weeeea" + eliminarSlash(path) + "MercTest2.xlsx");
           
            destworkSheet2 = (Microsoft.Office.Interop.Excel.Worksheet)destworkBook.Sheets.get_Item(2);

            int k = 1;
            destworkSheet.Cells[k, 4].Value = "Nombre Promotor";
            destworkSheet2.Cells[k, 4].Value = "Nombre Promotor";
            destworkSheet2.Cells[k, 14].Value = "MONTO PAGADO";
            k++;
            foreach (Promotor i in promotores)
            {
                destworkSheet.Cells[k, 4].Value = i.nombrePromotor;
                k++;
            }
            k = 2;*/
            /*foreach (Promotor i in promotoresVariable)
            {
                
                if (destworkSheet2.Cells[k, 6].Value != null)
                {
                    destworkSheet2.Cells[k, 4].Value = i.nombrePromotor;
                    destworkSheet2.Cells[k, 14].Value = packs.ObtenerComision(destworkSheet2.Cells[k, 6].Value.ToString(), destworkSheet2.Cells[k, 7].Value.ToString());
                    k++;
                }
            }*/

            destworkBook.Save();
            destworkBook.Close(true);
            excelApplication.Quit();

        }

        public void AñadirMontoFijo(string path)
        {
            Excel.Application excelApplication = new Excel.Application();
            excelApplication.DisplayAlerts = false;
            Excel.Workbook destworkBook;
            Excel.Worksheet destworkSheet;
            destworkBook = excelApplication.Workbooks.Open(destPath);
            int aux;

            for (int i = 1; i <= pathEstrageias.Count; i++)
            {
                destworkSheet = (Microsoft.Office.Interop.Excel.Worksheet)destworkBook.Sheets.get_Item(i);

                destworkSheet.Cells[1, 15].Value = "MONTO PAGADO";
                //Console.WriteLine("FURRO "+ destworkSheet.UsedRange.Rows.Count);
                for (int k = 2; k <= destworkSheet.UsedRange.Rows.Count; k++)
                {

                    string stringConexion = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path;
                    using (OleDbConnection connection = new OleDbConnection(stringConexion))
                    {
                        connection.Open();
                        OleDbDataReader reader = null;
                        //Console.WriteLine("FURROx2 " + destworkSheet.Cells[k, 3].Value.ToString());
                        OleDbCommand command = new OleDbCommand("SELECT * FROM Promotores WHERE Cve_prom = " + int.Parse(destworkSheet.Cells[k, 3].Value.ToString()), connection);
                        reader = command.ExecuteReader();
                        if (reader.Read())
                        {
                            aux = int.Parse(reader[8].ToString());
                            if(aux >= 5)
                            {
                                destworkSheet.Cells[k, 15].Value = "138.64";
                            }
                            else
                            {
                                destworkSheet.Cells[k, 15].Value = "115.38";
                            }
                            
                        }


                    }
                }
            }
          
            destworkBook.Save();
            destworkBook.Close(true);
            excelApplication.Quit();

        }


        public void SepararEstrategias(string test)
        {
            string pathdes = eliminarSlash(path);
            Excel.Workbook srcworkBook;
            Excel.Worksheet srcworkSheet;
            Excel.Workbook destworkBook;
            Excel.Worksheet destworkSheet;
            string srcPath;
            Excel.Application excelApplication = new Excel.Application();
            excelApplication.DisplayAlerts = false;
            //Opening of first worksheet and copying
            srcPath = path;
            srcworkBook = excelApplication.Workbooks.Open(srcPath);
            srcworkSheet = (Microsoft.Office.Interop.Excel.Worksheet)srcworkBook.Sheets.get_Item(1);

            Excel.Range allDataRange = srcworkSheet.get_Range("A2", "AY" + srcworkSheet.UsedRange.Rows.Count);
            allDataRange.Sort(allDataRange.Columns[2], Excel.XlSortOrder.xlAscending);

            int inicio = 2;
            int total = allDataRange.Rows.Count;
            int totalSheets = 1;

            while (inicio <= total)
            {
                destworkBook = excelApplication.Workbooks.Add();
                destworkSheet = (Microsoft.Office.Interop.Excel.Worksheet)destworkBook.Sheets.get_Item(1);
                int ini = inicio;
                string promotor;
                 string anterior = srcworkSheet.Cells[inicio, 2].Value.ToString();
                //Console.WriteLine("Estas son las estragias prro" + srcworkSheet.Cells[inicio, 2].Value.ToString());
                Estrageias.Add(srcworkSheet.Cells[inicio, 2].Value.ToString());
                while (inicio <= total && srcworkSheet.Cells[inicio, 2].Value.ToString().Equals(anterior))
                {
                    promotor = srcworkSheet.Cells[inicio, 3].Value.ToString();
                    Boolean bandera=false;
                    foreach(Promotor i in promotoresN)
                    {

                        if (i.clavePromotor== int.Parse(promotor))
                        {
                            bandera = true;
                        }
                    }
                    if (bandera == false)
                    {
                        //Console.WriteLine("El promotor con clave: " + promotor + " no existe en la BDD D:");
                        //MessageBox.Show("El promotor con clave: " + promotor + " no existe en la BDD D:");
                    }
                    inicio++;
                    
                }
                Excel.Range src;
                Excel.Range dest;

                src = srcworkSheet.get_Range("A1", "AY1");
                src.Copy(Type.Missing);
                dest = destworkSheet.get_Range("A1", "AY1");
                dest.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteValues, Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);

                src = srcworkSheet.get_Range("A" + ini, "AY" + (inicio - 1));

                //Console.WriteLine("Voy a copiar de A" + ini + ", AY" + (inicio - 1) + " a: A1, AY" + (Math.Abs(ini - inicio)));
                dest = destworkSheet.get_Range("A2", "AY" + (Math.Abs(ini - inicio) + 1));
                src.Copy(Type.Missing);

                dest.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteValues, Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);//Copiar Valores
                dest.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormats, Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);//copiar formato de columna
                dest.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteColumnWidths, Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);//copiar valores de columna

                destPath = pathdes + test + totalSheets + ".xlsx";
                pathEstrageias.Add(destPath);
                destworkBook.SaveAs(destPath);
                destworkBook.Close(true);

                totalSheets++;
            }


            srcworkBook.Close(false);

            excelApplication.Quit();
        }

        public void SepararEstrategias2(string pathF, string pathV)
        {
            string[] archivos = { pathF, pathV};

            string pathdes = eliminarSlash(path);
            Excel.Workbook srcworkBook;
            Excel.Worksheet srcworkSheet;
            Excel.Workbook destworkBook;
            Excel.Worksheet destworkSheet;
            string srcPath;
            Excel.Application excelApplication = new Excel.Application();
            excelApplication.DisplayAlerts = false;
            for (int j=0; j < 2; j++)
            {
                srcPath = archivos[j];
                srcworkBook = excelApplication.Workbooks.Open(srcPath);
                srcworkSheet = (Microsoft.Office.Interop.Excel.Worksheet)srcworkBook.Sheets.get_Item(1);

                Excel.Range allDataRange = srcworkSheet.get_Range("A2", "AY" + srcworkSheet.UsedRange.Rows.Count);
                allDataRange.Sort(allDataRange.Columns[2], Excel.XlSortOrder.xlAscending);

                int inicio = 2;
                int total = allDataRange.Rows.Count;
                int totalSheets = 1;

                while (inicio <= total)
                {
                    destworkBook = excelApplication.Workbooks.Add();
                    destworkSheet = (Microsoft.Office.Interop.Excel.Worksheet)destworkBook.Sheets.get_Item(1);
                    
                    int ini = inicio;
                    string promotor;
                    string anterior = srcworkSheet.Cells[inicio, 2].Value.ToString();
                    //Console.WriteLine("Estas son las estragias prro" + srcworkSheet.Cells[inicio, 2].Value.ToString());
                    Estrageias.Add(srcworkSheet.Cells[inicio, 2].Value.ToString());
                    while (inicio <= total && srcworkSheet.Cells[inicio, 2].Value.ToString().Equals(anterior))
                    {
                        promotor = srcworkSheet.Cells[inicio, 3].Value.ToString();
                        Boolean bandera = false;
                        foreach (Promotor i in promotoresN)
                        {

                            if (i.clavePromotor == int.Parse(promotor))
                            {
                                bandera = true;
                            }
                        }
                        if (bandera == false)
                        {
                            //Console.WriteLine("El promotor con clave: " + promotor + " no existe en la BDD D:");
                            //MessageBox.Show("El promotor con clave: " + promotor + " no existe en la BDD D:");
                        }
                        inicio++;

                    }
                    Excel.Range src;
                    Excel.Range dest;

                    src = srcworkSheet.get_Range("A1", "AY1");
                    src.Copy(Type.Missing);
                    dest = destworkSheet.get_Range("A1", "AY1");
                    dest.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteValues, Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);

                    src = srcworkSheet.get_Range("A" + ini, "AY" + (inicio - 1));

                    //Console.WriteLine("Voy a copiar de A" + ini + ", AY" + (inicio - 1) + " a: A1, AY" + (Math.Abs(ini - inicio)));
                    dest = destworkSheet.get_Range("A2", "AY" + (Math.Abs(ini - inicio) + 1));
                    src.Copy(Type.Missing);

                    dest.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteValues, Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);//Copiar Valores
                    dest.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormats, Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);//copiar formato de columna
                    dest.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteColumnWidths, Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);//copiar valores de columna

                    if (j == 0)
                    {
                        destPath = pathdes + "TestF" + totalSheets + ".xlsx";
                        pathEstrageias.Add(destPath);
                        destworkBook.SaveAs(destPath);
                        destworkBook.Close(true);
                    }
                    else
                    {
                        destPath = pathdes + "TestV" + totalSheets + ".xlsx";
                        pathEstrageias.Add(destPath);
                        destworkBook.SaveAs(destPath);
                        destworkBook.Close(true);
                    }
                    

                    totalSheets++;
                }
                srcworkBook.Close(false);
            }

            //Opening of first worksheet and copying


            

            excelApplication.Quit();
        }

        public void CalcularNominaVariable()
        {
            ExtraerPaquetesAccess();
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook book = xlApp.Workbooks.Open(destPath);
            Excel.Worksheet sheet;

            int n = book.Sheets.Count;
            for (int l = 1; l <= n; l++)
            {
                Excel.Workbook LibroAux = xlApp.Workbooks.Add();
                Excel.Worksheet hojaAux = LibroAux.Worksheets[1];

                sheet = book.Sheets.get_Item(l);
                int maxRows = sheet.UsedRange.Rows.Count;
                totalRows = maxRows - 1;
                sheet.Cells[1, 15].Value = "MONTO PAGADO";
                for (int i = 2; i <= maxRows; i++)
                {
                    if (sheet.Cells[i, 7].Value != null)
                    {
                        string tipoPack = sheet.Cells[i, 7].Value.ToString().ToUpper();
                        if (tipoPack.StartsWith("I- "))
                        {
                            tipoPack = tipoPack.Substring(3);
                        }
                        else if (tipoPack.StartsWith("I - "))
                        {
                            tipoPack = tipoPack.Substring(4);
                        }
                        string best = Mejor(tipoPack);
                        List<string> bestPacks = getFullMin();

                        string otherPack = sheet.Cells[i, 6].Value.ToString().Replace("  ", " ");

                        if (bestPacks.Count > 1)
                        {
                            string auxP = bestPacks.ElementAt(0);
                            int m = 99;
                            foreach (string str in bestPacks)
                            {
                                //auxP = str;
                                int ax = MejorPaquete2(otherPack, dic[str]);

                                Console.WriteLine("Ditancia Mejor2: " + ax);
                                if (ax < m)
                                {
                                    m = ax;
                                    auxP = str;
                                }

                            }

                            best = auxP;
                            Console.WriteLine("Habia varios, el mejor: " + auxP);
                        }

                        try
                        {
                            sheet.Cells[i, 15].Value = "" +  dic[best][otherPack];
                            aciertos++;
                        }
                        catch (System.Collections.Generic.KeyNotFoundException)
                        {
                            string bestPack;
                            if (best == "SOLO LINEA")
                            {
                                bestPack = MejorPaquete("PORTABI" + otherPack, dic[best]);
                            }
                            else
                            {
                                bestPack = MejorPaquete(otherPack, dic[best]);
                            }

                            int dis = CalcLevenshteinDistance(otherPack, bestPack);
                            int half = 0;
                            if (otherPack.Length > bestPack.Length)
                            {
                                half = otherPack.Length / 2;
                            }
                            else
                            {
                                half = bestPack.Length / 2;
                            }

                            if (dis > half)
                            {
                                Console.WriteLine("El paquete no es apto, ERROR activando plan de contingencia");

                                FullDistance(otherPack);
                                if (HayAmbiguedad())
                                {
                                    Console.WriteLine("Hubo una ambigüedad, el plan de contingencia no funciono.");
                                }
                                else
                                {

                                    dis = CalcLevenshteinDistance(otherPack, globalBestPack);
                                    half = 0;
                                    if (otherPack.Length > globalBestPack.Length)
                                    {
                                        half = otherPack.Length / 2;
                                    }
                                    else
                                    {
                                        half = globalBestPack.Length / 2;
                                    }

                                    if (dis > half)
                                    {
                                        Console.WriteLine("El plan de contingencia no funciono.");
                                    }
                                    else
                                    {
                                        Console.WriteLine("El plan de contingencia funciono, se encontro algo mejor.");
                                        Console.WriteLine("Eso mejor fue: " + globalBestPack + " de tipo: " + globalBestType);
                                        sheet.Cells[i, 15].Value = "" + dic[globalBestType][globalBestPack];

                                        aciertos++;
                                    }
                                }
                            }
                            else
                            {
                                Console.WriteLine("No encontre el paquete " + otherPack + " de tipo " + best + " y lo insertare en " + bestPack);
                                sheet.Cells[i, 15].Value = "" + dic[best][otherPack];
                                aciertos++;
                            }
                        }
                    }
                    else
                    {
                        Console.WriteLine("No se va a poder agregar paquete de la linea " + i);
                        Excel.Range colorRange = sheet.get_Range(calcularCelda(i) + 1, calcularCelda(i) + 16);
                        colorRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                    }
                }
                Console.WriteLine("Aciertos: " + aciertos + " totalRows: " + totalRows);
                double percent = (aciertos / totalRows) * 100.0;
                Console.WriteLine("Porcentaje de paquetes clasificados correctamente: " + percent);


                int x = sheet.UsedRange.Rows.Count;
                int z = sheet.UsedRange.Columns.Count;
                string lim = calcularCelda(z);
            }

            book.Save();
            book.Close();
            xlApp.Quit();
        }

        public void calcularNominaFija()
        {
            var AppExcel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook LibroExcel;
            Microsoft.Office.Interop.Excel.Worksheet HojaExcel;

            /*
             TITULOS DE LA LISTA DE MEDICIONES
             */
            //LibroExcel = AppExcel.Workbooks.Open(@"C:\Users\Juan\Documents\Merc\prueba.xlsx");
            LibroExcel = AppExcel.Workbooks.Open(destPath);
            

            int n = LibroExcel.Sheets.Count;
            for (int i = 1; i <= n; i++)
            {

                Console.WriteLine("LA I ES: " + i);
                Excel.Workbook LibroAux = AppExcel.Workbooks.Add();
                HojaExcel = LibroExcel.Worksheets[i];
                int x = HojaExcel.UsedRange.Rows.Count;
                int z = HojaExcel.UsedRange.Columns.Count;
                string lim = calcularCelda(z);
                Excel.Worksheet hojaAux = LibroAux.Worksheets[1];

                Excel.Range range = HojaExcel.get_Range("A1", "" + lim + x);
                Console.WriteLine("ALGO COSAS " + lim + x);

                //range.Value = data;
                
                ////3.
                Excel.Range oRange = range; // HojaExcel.UsedRange;
                Excel.PivotCache oPivotCache = (Excel.PivotCache)LibroExcel.PivotCaches().Add(Excel.XlPivotTableSourceType.xlDatabase, oRange);  // Set the Source data range from First sheet
                                                                                                                                        //Excel.Range oRange2 = HojaExcel.Cells[1, 1];
                Excel.PivotCaches pch = LibroExcel.PivotCaches();
                pch.Add(Excel.XlPivotTableSourceType.xlDatabase, oRange).CreatePivotTable(HojaExcel.Cells[x + 5, 4], "PivotTable"+i, Type.Missing, Type.Missing);// Create Pivot table

                Excel.PivotTable pvt = HojaExcel.PivotTables("PivotTable"+i) as Excel.PivotTable;


                pvt.ShowDrillIndicators = false;  // Used to remove the Expand/ Collapse Button from each cell

                Excel.PivotField fld = ((Excel.PivotField)pvt.PivotFields("Nombre Promotor")); // Create a Pivot Field in Pivot table

                fld.Orientation = Excel.XlPivotFieldOrientation.xlRowField;
                fld.set_Subtotals(1, true);

                fld = ((Excel.PivotField)pvt.PivotFields("Estatus PISA Multiorden"));
                fld.Orientation = Excel.XlPivotFieldOrientation.xlPageField;
                fld.EnableMultiplePageItems = true;
                string[] lista = new string[3] { "CORRECCION PROMOTOR", "SOLICITUD DUPLICADA", "SOLICITUD CANCELADA" };


                for (int j = 0; j < 3; j++)
                {
                    try
                    {
                        fld.PivotItems(lista[j]).Visible = false;
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e.ToString());
                    }


                }

                fld = ((Excel.PivotField)pvt.PivotFields("Folio SIAC"));
                fld.Orientation = Excel.XlPivotFieldOrientation.xlDataField;
                fld.Function = Excel.XlConsolidationFunction.xlCount;
                //fld.Name = "PAQ,";

                fld = ((Excel.PivotField)pvt.PivotFields("MONTO PAGADO"));
                fld.Orientation = Excel.XlPivotFieldOrientation.xlDataField;
                fld.Function = Excel.XlConsolidationFunction.xlSum;



                //fld.Name = "PAGO";

                pvt.DataPivotField.Orientation = Excel.XlPivotFieldOrientation.xlColumnField;

                //HojaExcel.UsedRange.Columns.AutoFit();  // Used to Autoset the column width according to data 
                pvt.ColumnGrand = true;  // Used to hide Grand total for columns
                pvt.RowGrand = true; // Used to hide Grand total for Rows
                Console.WriteLine("Furro 1");
                HojaExcel.Columns.AutoFit();
                Console.WriteLine("Furro 1");

                int mxCol = HojaExcel.Columns.Count;
                Excel.Range src = HojaExcel.get_Range("A1", calcularCelda(mxCol) + HojaExcel.Rows.Count);
                Excel.Range dest = hojaAux.get_Range("A1", calcularCelda(mxCol) + HojaExcel.Rows.Count);

                src.Copy(Type.Missing);
                dest.PasteSpecial(XlPasteType.xlPasteAll, XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
                
                //dest.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteValues, Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);//Copiar Valores
                //dest.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormats, Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);//copiar formato de columna
                //dest.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteColumnWidths, Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);//copiar valores de columna

                //HojaExcel.Rows.AutoFit();
                AppExcel.Application.Visible = false;

                LibroAux.SaveAs(Directory.GetCurrentDirectory() + HojaExcel.Cells[2, 2].Value.ToString() + "-Fija.xlsx");
                LibroAux.Close();

            }

            LibroExcel.Close();

        }

        public void calcularVariable()
        {
            var AppExcel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook LibroExcel;
            Microsoft.Office.Interop.Excel.Worksheet HojaExcel;

            /*
             TITULOS DE LA LISTA DE MEDICIONES
             */
            //LibroExcel = AppExcel.Workbooks.Open(@"C:\Users\Juan\Documents\Merc\prueba.xlsx");
            LibroExcel = AppExcel.Workbooks.Open(destPath);


            int n = LibroExcel.Sheets.Count;
            for (int i = 1; i <= n; i++)
            {

                Console.WriteLine("LA I ES: " + i);
                Excel.Workbook LibroAux = AppExcel.Workbooks.Add();
                HojaExcel = LibroExcel.Worksheets[i];
                int x = HojaExcel.UsedRange.Rows.Count;
                int z = HojaExcel.UsedRange.Columns.Count;
                string lim = calcularCelda(z);
                Excel.Worksheet hojaAux = LibroAux.Worksheets[1];

                Excel.Range range = HojaExcel.get_Range("A1", "" + lim + x);
                Console.WriteLine("ALGO COSAS " + lim + x);

                //range.Value = data;

                ////3.
                Excel.Range oRange = range; // HojaExcel.UsedRange;
                Excel.PivotCache oPivotCache = (Excel.PivotCache)LibroExcel.PivotCaches().Add(Excel.XlPivotTableSourceType.xlDatabase, oRange);  // Set the Source data range from First sheet
                                                                                                                                                 //Excel.Range oRange2 = HojaExcel.Cells[1, 1];
                Excel.PivotCaches pch = LibroExcel.PivotCaches();
                pch.Add(Excel.XlPivotTableSourceType.xlDatabase, oRange).CreatePivotTable(HojaExcel.Cells[x + 5, 4], "PivotTable" + i, Type.Missing, Type.Missing);// Create Pivot table

                Excel.PivotTable pvt = HojaExcel.PivotTables("PivotTable" + i) as Excel.PivotTable;


                pvt.ShowDrillIndicators = false;  // Used to remove the Expand/ Collapse Button from each cell

                Excel.PivotField fld = ((Excel.PivotField)pvt.PivotFields("Nombre Promotor")); // Create a Pivot Field in Pivot table

                fld.Orientation = Excel.XlPivotFieldOrientation.xlRowField;
                fld.set_Subtotals(1, true);

                fld = ((Excel.PivotField)pvt.PivotFields("Estatus PISA Multiorden"));
                fld.Orientation = Excel.XlPivotFieldOrientation.xlPageField;
                fld.EnableMultiplePageItems = true;
                string[] lista = new string[3] { "CORRECCION PROMOTOR", "SOLICITUD DUPLICADA", "SOLICITUD CANCELADA" };


                for (int j = 0; j < 3; j++)
                {
                    try
                    {
                        fld.PivotItems(lista[j]).Visible = false;
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e.ToString());
                    }


                }

                fld = ((Excel.PivotField)pvt.PivotFields("Folio SIAC"));
                fld.Orientation = Excel.XlPivotFieldOrientation.xlDataField;
                fld.Function = Excel.XlConsolidationFunction.xlCount;
                //fld.Name = "PAQ,";

                fld = ((Excel.PivotField)pvt.PivotFields("MONTO PAGADO"));
                fld.Orientation = Excel.XlPivotFieldOrientation.xlDataField;
                fld.Function = Excel.XlConsolidationFunction.xlSum;



                //fld.Name = "PAGO";

                pvt.DataPivotField.Orientation = Excel.XlPivotFieldOrientation.xlColumnField;

                //HojaExcel.UsedRange.Columns.AutoFit();  // Used to Autoset the column width according to data 
                pvt.ColumnGrand = true;  // Used to hide Grand total for columns
                pvt.RowGrand = true; // Used to hide Grand total for Rows
                Console.WriteLine("Furro 1");
                HojaExcel.Columns.AutoFit();
                Console.WriteLine("Furro 1");

                int mxCol = HojaExcel.Columns.Count;
                Excel.Range src = HojaExcel.get_Range("A1", calcularCelda(mxCol) + HojaExcel.Rows.Count);
                Excel.Range dest = hojaAux.get_Range("A1", calcularCelda(mxCol) + HojaExcel.Rows.Count);
                src.Copy(Type.Missing);

                src.Copy(Type.Missing);
                dest.PasteSpecial(XlPasteType.xlPasteAll, XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);

                //dest.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteValues, Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);//Copiar Valores
                //dest.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormats, Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);//copiar formato de columna
                //dest.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteColumnWidths, Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);//copiar valores de columna

                //HojaExcel.Rows.AutoFit();
                AppExcel.Application.Visible = false;

                LibroAux.SaveAs(Directory.GetCurrentDirectory() + HojaExcel.Cells[2, 2].Value.ToString() + "-Variable.xlsx");
                LibroAux.Close();

            }

            LibroExcel.Close();

        }

        public List<string> getFullMin()
        {
            List<string> bestPacks = new List<string>();
            foreach (KeyValuePair<string, int> entry in bestos)
            {
                if (entry.Value == actMin)
                {
                    bestPacks.Add(entry.Key);
                }
            }

            return bestPacks;
        }

        public string Mejor(string xlPack)
        {
            int aux = 0;
            string mejor = dic.Keys.ElementAt(0).ToUpper();
            int dis = CalcLevenshteinDistance(mejor, xlPack);
            aux = dis;
            actMin = aux;
            foreach (KeyValuePair<string, Dictionary<string, double>> entry in dic)
            {
                aux = CalcLevenshteinDistance(entry.Key.ToUpper(), xlPack);
                //Console.WriteLine("Distancia entre " + xlPack + " y " + entry.Key.ToUpper() + " es " + aux);
                bestos[entry.Key.ToUpper()] = aux;
                //Console.WriteLine("Comparando " + entry.Key + " con " + xlPack);
                if (aux < dis)
                {
                    mejor = entry.Key;
                    dis = aux;
                    actMin = aux;
                }
                // do something with entry.Value or entry.Key
            }
            return mejor;
        }

        public void FullDistance(string bpack)
        {
            dicAux = new Dictionary<string, Dictionary<string, int>>();
            foreach (KeyValuePair<string, Dictionary<string, double>> entry in dic)
            {
                dicAux.Add(entry.Key, new Dictionary<string, int>());
                foreach (KeyValuePair<string, double> entry2 in entry.Value)
                {
                    dicAux[entry.Key].Add(entry2.Key, 0);
                    dicAux[entry.Key][entry2.Key] = CalcLevenshteinDistance(bpack, entry2.Key);
                }
            }
        }

        public bool HayAmbiguedad()
        {
            bool ban = false;
            List<int> packsMinimos = new List<int>();
            int min = 99999;

            foreach (KeyValuePair<string, Dictionary<string, int>> entry in dicAux)
            {
                foreach (KeyValuePair<string, int> entry2 in entry.Value)
                {
                    if (entry2.Value < min)
                    {
                        min = entry2.Value;
                    }
                }
            }

            foreach (KeyValuePair<string, Dictionary<string, int>> entry in dicAux)
            {
                foreach (KeyValuePair<string, int> entry2 in entry.Value)
                {
                    if (entry2.Value == min)
                    {
                        globalBestType = entry.Key;
                        globalBestPack = entry2.Key;
                        packsMinimos.Add(entry2.Value);
                    }
                }
            }

            if (packsMinimos.Count > 1)
            {
                ban = true;
            }

            return ban;
        }

        public string MejorPaquete(string pack, Dictionary<string, double> dicPacks)
        {
            int percent = pack.Length / 2;
            int aux = 0;
            string mejor = dicPacks.Keys.ElementAt(0).ToUpper();
            int dis = CalcLevenshteinDistance(mejor, pack);
            aux = dis;
            bool ban = false;
            foreach (KeyValuePair<string, double> entry in dicPacks)
            {
                aux = CalcLevenshteinDistance(pack, entry.Key.ToUpper());
                //Console.WriteLine("Distancia entre " + pack + " y " + entry.Key.ToUpper() + " es " + aux);
                if (aux < dis)
                {
                    ban = true;
                    //Console.WriteLine("Encontre un mejor paquete con dist " + aux + "dist ant: " + dis);
                    mejor = entry.Key;
                    dis = aux;
                }
            }
            if (dis >= 15)
            {
                Console.WriteLine("No se encontro un buen paquete para " + pack);
            }
            //Console.WriteLine("La distacia del mejor fue: " + dis);
            return mejor;
        }

        public int MejorPaquete2(string pack, Dictionary<string, double> dicPacks)
        {
            int percent = pack.Length / 2;
            int aux = 0;
            string mejor = dicPacks.Keys.ElementAt(0).ToUpper();
            int dis = CalcLevenshteinDistance(mejor, pack);
            aux = dis;
            bool ban = false;
            foreach (KeyValuePair<string, double> entry in dicPacks)
            {
                aux = CalcLevenshteinDistance(pack, entry.Key.ToUpper());
                //Console.WriteLine("Distancia entre " + pack + " y " + entry.Key.ToUpper() + " es " + aux);
                if (aux < dis)
                {
                    ban = true;
                    //Console.WriteLine("Encontre un mejor paquete con dist " + aux + "dist ant: " + dis);
                    mejor = entry.Key;
                    dis = aux;
                }
            }
            if (dis >= 15)
            {
                Console.WriteLine("No se encontro un buen paquete para " + pack);
            }
            //Console.WriteLine("Contar 2 distancia: " + dis + " Pack: " + pack + " elegido: " + mejor);
            return dis;
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

        public void ExtraerPaquetesAccess()
        {
            
            string connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Directory.GetCurrentDirectory() +"\\comisiones1.accdb";
            string strSQL = "SELECT DISTINCT(TABLACOMISIONES) FROM Paquetes";
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
                            dic.Add(reader["TABLACOMISIONES"].ToString().ToUpper(), new Dictionary<string, double>());
                            bestos.Add(reader["TABLACOMISIONES"].ToString().ToUpper(), 99);

                        }
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.ToString());
                }
                foreach (KeyValuePair<string, Dictionary<string, double>> entry in dic)
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
                                dic[entry.Key].Add(reader["PAQUETES"].ToString().ToUpper(), double.Parse(reader["COMISIONPROMOTOR"].ToString()));
                                fullPacks.Add(reader["PAQUETES"].ToString().ToUpper());
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

    }
}
