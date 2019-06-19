using System;
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
    public partial class VentanaPromotores : Form
    {
        string path;
        List<Promotor> promotores;
        public VentanaPromotores(List<Promotor> aux, string path)
        {

            InitializeComponent();
            this.path = path;
            promotores = aux;
            foreach (Promotor i in promotores)
            {
                comboBox1.Items.Add(i.clavePromotor);
            }

        }

        private void label6_Click(object sender, EventArgs e)
        {

        }

        private void VentanaPromotores_Load(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            int idPromotor = int.Parse(comboBox1.SelectedItem.ToString());
            foreach (Promotor i in promotores)
            {
                if (i.clavePromotor == idPromotor)
                {
                    textBox1.Text = i.nombrePromotor;
                    textBox2.Text = i.celula;
                    textBox3.Text = "" + i.estrategia;
                    textBox4.Text = "" + i.nominaFija;
                    textBox5.Text = i.claveInter;
                    textBox6.Text = "" + i.nominaVariable;
                    textBox7.Text = i.banco;
                    textBox8.Text = ""+i.Ncelula;
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {

            OleDbConnection con;//Representa una conexión abierta a un origen de datos
            String stringConexion = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path;
            using (OleDbConnection connection = new OleDbConnection(stringConexion))
            {
                connection.Open();


                OleDbCommand command = new OleDbCommand("INSERT into Promotores (nombre_promotor, Cve_prom, Estrategia, Celula, CLAVE_INTERBANCARIA, Banco, Ncelula) VALUES " + "('" + textBox1.Text + "'," + comboBox1.Text + "," + textBox3.Text + "," + "'" + textBox2.Text + "'," + "'" + textBox5.Text + "'," + "'" + textBox7.Text + "'," + textBox8 + ")");
                command.Connection = connection;
                if (connection.State != ConnectionState.Open)
                {
                    connection.Open();
                }

                try
                {
                    command.ExecuteNonQuery();
                    MessageBox.Show("Data Added");
                    Promotor nuevo = new Promotor();
                    nuevo.clavePromotor = int.Parse(comboBox1.Text);
                    nuevo.nombrePromotor = textBox1.Text;
                    nuevo.celula = textBox2.Text;
                    nuevo.estrategia = int.Parse(textBox3.Text);
                    nuevo.claveInter = textBox5.Text;
                    nuevo.banco = textBox7.Text;
                    nuevo.Ncelula = int.Parse(textBox8.Text);
                    promotores.Add(nuevo);
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

        private void button2_Click(object sender, EventArgs e)
        {
            OleDbConnection con;//Representa una conexión abierta a un origen de datos
            String stringConexion = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path;
            using (OleDbConnection connection = new OleDbConnection(stringConexion))
            {
                connection.Open();


                OleDbCommand command = new OleDbCommand("UPDATE Promotores SET nombre_promotor = '" + textBox1.Text + "', Estrategia = "+ textBox3.Text+ ", Celula = '" + textBox2.Text + "', CLAVE_INTERBANCARIA = '"+ textBox5.Text+ "', Banco = '"+ textBox7.Text +"', Ncelula = "+ textBox8.Text+ " Where Cve_prom = "+comboBox1.Text);
                command.Connection = connection;
                if (connection.State != ConnectionState.Open)
                {
                    connection.Open();
                }

                try
                {
                    command.ExecuteNonQuery();
                    MessageBox.Show("Data Update");
                    Promotor nuevo = new Promotor();
                    nuevo.clavePromotor = int.Parse(comboBox1.Text);
                    nuevo.nombrePromotor = textBox1.Text;
                    nuevo.celula = textBox2.Text;
                    nuevo.estrategia = int.Parse(textBox3.Text);
                    nuevo.claveInter = textBox5.Text;
                    nuevo.banco = textBox7.Text;
                    nuevo.Ncelula = int.Parse(textBox8.Text);
                    int aux=0;
                    foreach (Promotor i in promotores)
                    {
                        if(i.clavePromotor == int.Parse(comboBox1.Text))
                        {
                            aux=promotores.IndexOf(i);                           
                        }
                    }
                    promotores.RemoveAt(aux);
                    promotores.Add(nuevo);
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


                OleDbCommand command = new OleDbCommand("DELETE FROM Promotores Where Cve_prom = " + comboBox1.Text);
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
                    foreach (Promotor i in promotores)
                    {
                        if (i.clavePromotor == int.Parse(comboBox1.Text))
                        {
                            aux = promotores.IndexOf(i);
                        }
                    }
                    promotores.RemoveAt(aux);
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

        private void button1_Click_1(object sender, EventArgs e)
        {
            OleDbConnection con;//Representa una conexión abierta a un origen de datos
            String stringConexion = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path;
            using (OleDbConnection connection = new OleDbConnection(stringConexion))
            {
                connection.Open();


                OleDbCommand command = new OleDbCommand("INSERT into Promotores (nombre_promotor, Cve_prom, Estrategia, Celula, CLAVE_INTERBANCARIA, Banco, Ncelula) VALUES " + "('" + textBox1.Text + "'," + comboBox1.Text + "," + textBox3.Text + "," + "'" + textBox2.Text + "'," + "'" + textBox5.Text + "'," + "'" + textBox7.Text + "'," + textBox8 + ")");
                command.Connection = connection;
                if (connection.State != ConnectionState.Open)
                {
                    connection.Open();
                }

                try
                {
                    command.ExecuteNonQuery();
                    MessageBox.Show("Data Added");
                    Promotor nuevo = new Promotor();
                    nuevo.clavePromotor = int.Parse(comboBox1.Text);
                    nuevo.nombrePromotor = textBox1.Text;
                    nuevo.celula = textBox2.Text;
                    nuevo.estrategia = int.Parse(textBox3.Text);
                    nuevo.claveInter = textBox5.Text;
                    nuevo.banco = textBox7.Text;
                    nuevo.Ncelula = int.Parse(textBox8.Text);
                    promotores.Add(nuevo);
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
    }

}
