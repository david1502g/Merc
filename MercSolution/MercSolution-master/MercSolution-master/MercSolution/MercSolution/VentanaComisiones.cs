using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;

namespace MercSolution
{
    public partial class VentanaComisiones : Form
    {
        string path;
        List<Comisiones> Comision;
        public VentanaComisiones(List<Comisiones> aux, string path)
        {
            this.Comision = aux;
            InitializeComponent();
            this.path = Directory.GetCurrentDirectory() + "\\comisiones1.accdb";

            foreach (Comisiones i in Comision)
            {
                comboBox1.Items.Add(i.paquetes);
            }
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            OleDbConnection con;//Representa una conexión abierta a un origen de datos
            String stringConexion = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path;
            using (OleDbConnection connection = new OleDbConnection(stringConexion))
            {
                connection.Open();


                OleDbCommand command = new OleDbCommand("INSERT into Paquetes (TABLACOMISIONES, COMISIONTOTAL, COMISIONPROMOTOR, COMISIONGERENTE, PAQUETES) VALUES " + "('" + textBox1.Text + "'," + textBox2.Text + "," + textBox3.Text + "," + textBox4.Text + "," + "'" + comboBox1.Text + "')");
                command.Connection = connection;
                if (connection.State != ConnectionState.Open)
                {
                    connection.Open();
                }

                try
                {
                    command.ExecuteNonQuery();
                    MessageBox.Show("Data Added");
                    Comisiones nuevo = new Comisiones();

                    nuevo.comisionGerente = float.Parse(textBox3.Text);
                    nuevo.comisionPromotor = float.Parse(textBox2.Text);
                    nuevo.tablasComisiones = textBox4.Text;
                    nuevo.comisionTotal = float.Parse(textBox1.Text);
                    nuevo.paquetes = comboBox1.Text;



                    Comision.Add(nuevo);
                    connection.Close();
                }
                catch (OleDbException ex)
                {
                    MessageBox.Show("error al agregar en la BD");
                    Console.WriteLine(ex);
                    connection.Close();
                }

            }


        }

        private void button3_Click(object sender, EventArgs e)
        {

            OleDbConnection con;//Representa una conexión abierta a un origen de datos
            String stringConexion = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path;
            using (OleDbConnection connection = new OleDbConnection(stringConexion))
            {
                connection.Open();

                OleDbCommand command = new OleDbCommand("DELETE FROM Paquetes Where PAQUETES = '" + comboBox1.Text + "'");
                command.Connection = connection;
                if (connection.State != ConnectionState.Open)
                {
                    connection.Open();
                }

                try
                {
                    command.ExecuteNonQuery();
                    MessageBox.Show("Data Delete");
                    int aux = 0;
                    foreach (Comisiones i in Comision)
                    {
                        if (i.paquetes == comboBox1.Text)
                        {
                            aux = Comision.IndexOf(i);
                        }
                    }
                    Comision.RemoveAt(aux);
                    connection.Close();
                }
                catch (OleDbException ex)
                {
                    MessageBox.Show("error al eliminar en la BD");
                    Console.WriteLine(ex);
                    connection.Close();
                }

            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            OleDbConnection con;//Representa una conexión abierta a un origen de datos
            String stringConexion = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path;
            using (OleDbConnection connection = new OleDbConnection(stringConexion))
            {
                connection.Open();


                OleDbCommand command = new OleDbCommand("UPDATE Paquetes SET TABLACOMISIONES = '" + textBox4.Text + "', COMISIONTOTAL = " + textBox1.Text + ", COMISIONPROMOTOR = '" + textBox2.Text + "', COMISIONGERENTE = " + textBox3.Text + ", PAQUETES = '" + comboBox1.Text + "' WHERE PAQUETES = " + "'" + comboBox1.Text + "'");
                command.Connection = connection;
                if (connection.State != ConnectionState.Open)
                {
                    connection.Open();
                }

                try
                {
                    command.ExecuteNonQuery();
                    MessageBox.Show("Data Update");
                    Comisiones nuevo = new Comisiones();
                    nuevo.comisionGerente = float.Parse(textBox3.Text);
                    nuevo.comisionPromotor = float.Parse(textBox2.Text);
                    nuevo.tablasComisiones = textBox4.Text;
                    nuevo.comisionTotal = float.Parse(textBox1.Text);
                    nuevo.paquetes = comboBox1.Text;
                    int aux = 0;
                    foreach (Comisiones i in Comision)
                    {
                        if (i.paquetes == comboBox1.Text)
                        {
                            Console.WriteLine("i.paquetes:::::" + i.paquetes);
                            Console.WriteLine("combox1" + comboBox1.Text);
                            aux = Comision.IndexOf(i);
                        }
                    }
                    Comision.RemoveAt(aux);
                    Comision.Add(nuevo);
                    connection.Close();
                }
                catch (OleDbException ex)
                {
                    MessageBox.Show("error al editar en la BD");
                    Console.WriteLine(ex);
                    connection.Close();
                }

            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            


        }

        private void comboBox1_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            string nombrePquete = comboBox1.SelectedItem.ToString();
            foreach (Comisiones i in Comision)
            {
                if (i.paquetes == nombrePquete)
                {
                    textBox1.Text = "" + i.comisionTotal;
                    textBox2.Text = "" + i.comisionPromotor;
                    textBox3.Text = "" + i.comisionGerente;
                    textBox4.Text = i.tablasComisiones;


                }
            }
        }

        private void label1_Click_1(object sender, EventArgs e)
        {

        }
    }
}
