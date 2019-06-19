using System;
using System.IO;
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
    public partial class VentanaBusqueda : Form
    {

        string path;
        string excel;
        private System.Data.DataTable dt;
        int semanaMax=-1, semanaMIN=-1,j=0;

        public VentanaBusqueda(string path)
        {
            InitializeComponent();
            agregaSemana(path);
            radioButton1.Checked = true;
            tipoNomina.Visible = false;
            comboBox1.Visible = false;
            this.path = path;
            
        }     
        
        public void agregaSemana(string path)
        {
            string semanAux;
            int aux;
            string[] cadenas;
            string stringConexion = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path;
            using (OleDbConnection connection = new OleDbConnection(stringConexion))
            {
                connection.Open();
                OleDbDataReader reader = null;
                OleDbCommand command = new OleDbCommand("SELECT * FROM Respaldo;", connection);
                reader = command.ExecuteReader();
                while (reader.Read())
                {
                    j++;
                    semanAux = reader[14].ToString();
                    cadenas = semanAux.Split();
                    aux = int.Parse(cadenas[1]);
                    if (semanaMax == -1)
                    {
                        semanaMax = aux;
                    }
                    else
                    {
                        if (semanaMax < aux)
                        {
                            semanaMax = aux;
                        }
                    }

                    if (semanaMIN == -1)
                    {
                        semanaMIN = aux;
                    }
                    else
                    {
                        if (semanaMIN > aux)
                        {
                            semanaMIN = aux;
                        }
                    }


                }

            }
         
            for(int i = semanaMIN; i <= semanaMax; i++)
            {
                tipoNomina.Items.Add("Semana "+ i);
                comboBox1.Items.Add("Semana " + i);
            }

        }

        private void button1_Click(object sender, EventArgs e)
        {


            dt = new System.Data.DataTable();
            dt.Columns.Add("Fecha Captura");
            dt.Columns.Add("Estrategia");
            dt.Columns.Add("Promotor");
            dt.Columns.Add("Nombre Promotor");
            dt.Columns.Add("Folio SIAC");
            dt.Columns.Add("Paquete");
            dt.Columns.Add("Otros Servicios");
            dt.Columns.Add("Campana");
            dt.Columns.Add("Telefono Asignado");
            dt.Columns.Add("Estatus PISA Multiorden");
            dt.Columns.Add("Pisa OS Fecha POSTEO Multiorden");
            dt.Columns.Add("Entrego Expediente");
            dt.Columns.Add("Tipo Entrego Expediente");
            dt.Columns.Add("Semana");
            dataGridView1.DataSource = dt;


            int inicio, final;
            string[] aux;
            if (radioButton1.Checked){

                string stringConexion = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path;
                using (OleDbConnection connection = new OleDbConnection(stringConexion))
                {
                    connection.Open();
                    OleDbDataReader reader = null;
                    OleDbCommand command = new OleDbCommand("select * FROM Respaldo WHERE Folio_SIAC = '" + textBox1.Text + "'", connection);
                    reader = command.ExecuteReader();
                    if (reader.Read())
                    {
                        DataRow row = dt.NewRow();
                        row["Fecha Captura"] = reader[1].ToString();
                        row["Estrategia"] = reader[2].ToString();
                        row["Promotor"] = reader[3].ToString();
                        row["Nombre Promotor"] = reader[4].ToString();
                        row["Folio SIAC"] = reader[5].ToString();
                        row["Paquete"] = reader[6].ToString();
                        row["Otros Servicios"] = reader[7].ToString();
                        row["Campana"] = reader[8].ToString();
                        row["Telefono Asignado"] = reader[9].ToString();
                        row["Estatus PISA Multiorden"] = reader[10].ToString();
                        row["Pisa OS Fecha POSTEO Multiorden"] = reader[11].ToString();
                        row["Entrego Expediente"] = reader[12].ToString();
                        row["Tipo Entrego Expediente"] = reader[13].ToString();
                        row["Semana"] = reader[14].ToString();
                        dt.Rows.Add(row);
                    }
                    else
                    {
                        MessageBox.Show("Folio SIAC no encontrado");
                    }

                }

            }

            if (radioButton2.Checked)
            {
                aux = tipoNomina.Text.Split();
                inicio = int.Parse(aux[1]);
                aux = comboBox1.Text.Split();
                final = int.Parse(aux[1]);
                string stringConexion = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path;
                using (OleDbConnection connection = new OleDbConnection(stringConexion))
                {
                    connection.Open();
                    OleDbDataReader reader = null;
                    for(int i = inicio; i<= final; i++)
                    {
                        OleDbCommand command = new OleDbCommand("select * FROM Respaldo WHERE Semana = '" + "Semana " +i+ "'", connection);
                        reader = command.ExecuteReader();
                        while (reader.Read())
                        {
                            DataRow row = dt.NewRow();
                            row["Fecha Captura"] = reader[1].ToString();
                            row["Estrategia"] = reader[2].ToString();
                            row["Promotor"] = reader[3].ToString();
                            row["Nombre Promotor"] = reader[4].ToString();
                            row["Folio SIAC"] = reader[5].ToString();
                            row["Paquete"] = reader[6].ToString();
                            row["Otros Servicios"] = reader[7].ToString();
                            row["Campana"] = reader[8].ToString();
                            row["Telefono Asignado"] = reader[9].ToString();
                            row["Estatus PISA Multiorden"] = reader[10].ToString();
                            row["Pisa OS Fecha POSTEO Multiorden"] = reader[11].ToString();
                            row["Entrego Expediente"] = reader[12].ToString();
                            row["Tipo Entrego Expediente"] = reader[13].ToString();
                            row["Semana"] = reader[14].ToString();
                            dt.Rows.Add(row);
                        }
                    }
                    

                }
                
            }





            /*
            Boolean bandera = false;
            if (tipoNomina.SelectedItem.ToString() == "Ingresos")
            {
                string stringConexion = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path;
                using (OleDbConnection connection = new OleDbConnection(stringConexion))
                {
                    connection.Open();
                    OleDbDataReader reader = null;
                    OleDbCommand command = new OleDbCommand("select * FROM Respaldo_I WHERE Folio_SIAC = '" + textBox1.Text + "'", connection);
                    reader = command.ExecuteReader();
                    if (reader.Read())
                    {
                        excel = reader[1].ToString();
                    }

                }

                if (excel != null)
                {
                    Excel.Application excelApplication = new Excel.Application();
                    Excel.Workbook destworkBook;
                    Excel.Worksheet destworkSheet;
                    destworkBook = excelApplication.Workbooks.Open(Directory.GetCurrentDirectory() + "\\PIPES\\" + excel);
                    destworkSheet = (Microsoft.Office.Interop.Excel.Worksheet)destworkBook.Sheets.get_Item(1);

                    Console.WriteLine("MIRAAAAA porfis : "+ destworkSheet.UsedRange.Rows.Count);

                    int i = 2;
                    Console.WriteLine("MIRAAAAA porfis : " + destworkSheet.Cells[i, 5].Value.ToString());
                    while (destworkSheet.Cells[i, 5].Value.ToString() != textBox1.Text && i <= destworkSheet.UsedRange.Rows.Count)
                    {
                        Console.WriteLine("MIRAAAAA porfis : " + i);
                        i++;
                    }

                    if (i == destworkSheet.UsedRange.Columns.Count + 1)
                    {
                        MessageBox.Show("Folio SIAC no encontrado");
                    }
                    else
                    {

                        for (int j = 1; j <= destworkSheet.UsedRange.Columns.Count; j++)
                        {
                            if (destworkSheet.Cells[i, j].Value == null)
                            {
                                destworkSheet.Cells[i, j].Value = " ";
                            }
                        }
                        //Fecha Captura, Estrategia, Promotor, Nombre Promotor, Folio SIAC, Paquete, Paquete, Otros Servicios, Campana, Telefono Asignado, Estatus PISA Multiorden, Pisa OS Fecha POSTEO Multiorden, Entrego Expediente, Semana
                        DataRow row = dt.NewRow();
                        row["Fecha Captura"] = destworkSheet.Cells[i, 1].Value.ToString();
                        row["Estrategia"] = destworkSheet.Cells[i, 2].Value.ToString();
                        row["Promotor"] = destworkSheet.Cells[i, 3].Value.ToString();
                        row["Nombre Promotor"] = destworkSheet.Cells[i, 4].Value.ToString();
                        row["Folio SIAC"] = destworkSheet.Cells[i, 5].Value.ToString();
                        row["Paquete"] = destworkSheet.Cells[i, 6].Value.ToString();
                        row["Otros Servicios"] = destworkSheet.Cells[i, 7].Value.ToString();
                        row["Campana"] = destworkSheet.Cells[i, 8].Value.ToString();
                        row["Telefono Asignado"] = destworkSheet.Cells[i, 9].Value.ToString();
                        row["Estatus PISA Multiorden"] = destworkSheet.Cells[i, 10].Value.ToString();
                        row["Pisa OS Fecha POSTEO Multiorden"] = destworkSheet.Cells[i, 11].Value.ToString();
                        row["Entrego Expediente"] = destworkSheet.Cells[i, 12].Value.ToString();
                        row["Tipo Entrego Expediente"] = destworkSheet.Cells[i, 13].Value.ToString();
                        row["Semana"] = destworkSheet.Cells[i, 14].Value.ToString();
                        dt.Rows.Add(row);

                        destworkBook.Close(false);
                        excelApplication.Quit();

                    }


                }
                else
                {
                    MessageBox.Show("Folio SIAC no encontrado");
                }
            }
            else
            {
                string stringConexion = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path;
                using (OleDbConnection connection = new OleDbConnection(stringConexion))
                {
                    connection.Open();
                    OleDbDataReader reader = null;
                    OleDbCommand command = new OleDbCommand("select * FROM Respaldo_P WHERE Folio_SIAC = '" + textBox1.Text + "'", connection);
                    reader = command.ExecuteReader();
                    if (reader.Read())
                    {
                        excel = reader[1].ToString();
                    }

                }

                if (excel != null)
                {
                    Excel.Application excelApplication = new Excel.Application();
                    Excel.Workbook destworkBook;
                    Excel.Worksheet destworkSheet;
                    destworkBook = excelApplication.Workbooks.Open(Directory.GetCurrentDirectory() + "\\PIPES\\" + excel);
                    destworkSheet = (Microsoft.Office.Interop.Excel.Worksheet)destworkBook.Sheets.get_Item(1);

                    int i = 2;
                    while (destworkSheet.Cells[i, 5].Value.ToString() != textBox1.Text && i <= destworkSheet.UsedRange.Columns.Count)
                    {
                        i++;
                    }

                    if (i == destworkSheet.UsedRange.Columns.Count + 1)
                    {
                        MessageBox.Show("Folio SIAC no encontrado");
                    }
                    else
                    {

                        for (int j = 1; j <= destworkSheet.UsedRange.Rows.Count; j++)
                        {
                            if (destworkSheet.Cells[i, j].Value == null)
                            {
                                destworkSheet.Cells[i, j].Value = " ";
                            }
                        }

                        DataRow row = dt.NewRow();
                        row["Fecha Captura"] = destworkSheet.Cells[i, 1].Value.ToString();
                        row["Estrategia"] = destworkSheet.Cells[i, 2].Value.ToString();
                        row["Promotor"] = destworkSheet.Cells[i, 3].Value.ToString();
                        row["Nombre Promotor"] = destworkSheet.Cells[i, 4].Value.ToString();
                        row["Folio SIAC"] = destworkSheet.Cells[i, 5].Value.ToString();
                        row["Paquete"] = destworkSheet.Cells[i, 6].Value.ToString();
                        row["Otros Servicios"] = destworkSheet.Cells[i, 7].Value.ToString();
                        row["Campana"] = destworkSheet.Cells[i, 8].Value.ToString();
                        row["Telefono Asignado"] = destworkSheet.Cells[i, 9].Value.ToString();
                        row["Estatus PISA Multiorden"] = destworkSheet.Cells[i, 10].Value.ToString();
                        row["Pisa OS Fecha POSTEO Multiorden"] = destworkSheet.Cells[i, 11].Value.ToString();
                        row["Entrego Expediente"] = destworkSheet.Cells[i, 12].Value.ToString();
                        row["Tipo Entrego Expediente"] = destworkSheet.Cells[i, 13].Value.ToString();
                        dt.Rows.Add(row);

                        destworkBook.Close(false);
                        excelApplication.Quit();

                    }


                }
                else
                {
                    MessageBox.Show("Folio SIAC no encontrado");
                }
            }
           */
        }

        private void VentanaBusqueda_Load(object sender, EventArgs e)
        {
            dt = new System.Data.DataTable();
            dt.Columns.Add("Fecha Captura");
            dt.Columns.Add("Estrategia");
            dt.Columns.Add("Promotor");
            dt.Columns.Add("Nombre Promotor");
            dt.Columns.Add("Folio SIAC");
            dt.Columns.Add("Paquete");
            dt.Columns.Add("Otros Servicios");
            dt.Columns.Add("Campana");
            dt.Columns.Add("Telefono Asignado");
            dt.Columns.Add("Estatus PISA Multiorden");
            dt.Columns.Add("Pisa OS Fecha POSTEO Multiorden");
            dt.Columns.Add("Entrego Expediente");
            dt.Columns.Add("Tipo Entrego Expediente");
            dt.Columns.Add("Semana");
            dataGridView1.DataSource = dt;
        }

        private void tipoNomina_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            tipoNomina.Visible = false;
            comboBox1.Visible = false;
            label1.Visible = true;
            textBox1.Visible = true;
        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {
            //radioButton2.Checked = false;
            
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            //radioButton1.Checked = false;
            tipoNomina.Visible = true;
            comboBox1.Visible = true;
            label1.Visible = false;
            textBox1.Visible = false;
        }

        private void dataGridView1_Scroll(object sender, ScrollEventArgs e)
        {

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}
