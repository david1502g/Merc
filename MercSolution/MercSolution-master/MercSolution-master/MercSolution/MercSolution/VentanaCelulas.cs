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
    public partial class VentanaCelulas : Form
    {

        string path;
        string pathA;
        List<Celula> celulas;
        List<Promotor> promotores;       

        public VentanaCelulas(List<Celula> aux, string path, string pathA, List<Promotor> auxP)
        {
            InitializeComponent();
            this.path = path;
            celulas = aux;
            promotores = auxP;
            this.pathA = pathA;
            foreach (Celula i in celulas)
            {
                comboBox1.Items.Add(i.claveCelula);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {

            OleDbConnection connection;

            
            List<Promotor> auxProm = new List<Promotor>();

            String stringConexion = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path;
            using (connection = new OleDbConnection(stringConexion))
            {
                connection.Open();


                OleDbCommand command = new OleDbCommand("UPDATE CELULAS SET nomCelula = '" + textBox1.Text + "', numCelula = " + textBox2.Text + " Where cveCelula = " + comboBox1.Text);
                command.Connection = connection;
                if (connection.State != ConnectionState.Open)
                {
                    connection.Open();
                }

                try
                {
                    command.ExecuteNonQuery();
                    MessageBox.Show("Data Update");
                    Celula nuevo = new Celula();
                    nuevo.claveCelula = int.Parse(comboBox1.Text);
                    nuevo.nombreCelula = textBox1.Text;
                    nuevo.Ncelula = int.Parse(textBox2.Text);

                    int aux = 0;
                    foreach (Celula i in celulas)
                    {
                        if (i.claveCelula == int.Parse(comboBox1.Text))
                        {
                            aux = celulas.IndexOf(i);                      
                        }
                    }
                    celulas.RemoveAt(aux);
                    celulas.Add(nuevo);
                    connection.Close();
                }
                catch (OleDbException ex)
                {
                    MessageBox.Show("error al agregar en la BD");
                    Console.WriteLine(ex);
                    connection.Close();
                }
            
            }
            
            foreach (Promotor i in promotores)
            {
                if (i.estrategia == int.Parse(comboBox1.Text))
                {
                    stringConexion = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + pathA;
                    using (connection = new OleDbConnection(stringConexion))
                    {
                        connection.Open();


                        OleDbCommand command = new OleDbCommand("UPDATE Promotores SET Celula = '" + textBox1.Text + "', Ncelula = " + textBox2.Text + " Where Cve_prom = " + i.clavePromotor);
                        command.Connection = connection;
                        if (connection.State != ConnectionState.Open)
                        {
                            connection.Open();
                        }

                        try
                        {
                            command.ExecuteNonQuery();
                            ///MessageBox.Show("Data Update");
                            Promotor nuevo = new Promotor();
                            nuevo = i;
                            nuevo.celula = textBox1.Text;
                            nuevo.Ncelula = int.Parse(textBox2.Text);
                            auxProm.Add(nuevo);
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
                else
                {
                    auxProm.Add(i);
                }
            }

            promotores.Clear();
            promotores = auxProm;
          
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            int cveCelula = int.Parse(comboBox1.SelectedItem.ToString());
            foreach (Celula i in celulas)
            {
                if (i.claveCelula == cveCelula)
                {
                    textBox1.Text = i.nombreCelula;
                    textBox2.Text = i.Ncelula+"";                   
                }
            }
        }

        private void agregar_Click(object sender, EventArgs e)
        {
            String stringConexion = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path;
            using (OleDbConnection connection = new OleDbConnection(stringConexion))
            {
                connection.Open();


                OleDbCommand command = new OleDbCommand("INSERT into CELULAS (cveCelula, nomCelula, numCelula) VALUES " + "(" + comboBox1.Text + ", '" + textBox1.Text + "', " + textBox2.Text + ")");
                command.Connection = connection;
                if (connection.State != ConnectionState.Open)
                {
                    connection.Open();
                }

                try
                {
                    command.ExecuteNonQuery();
                    MessageBox.Show("Data Added");
                    Celula nuevo = new Celula();
                    nuevo.claveCelula = int.Parse(comboBox1.Text);
                    nuevo.nombreCelula = textBox1.Text;
                    nuevo.Ncelula = int.Parse(textBox2.Text);
                    celulas.Add(nuevo);
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

        private void eliminar_Click(object sender, EventArgs e)
        {

            String stringConexion = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path;
            using (OleDbConnection connection = new OleDbConnection(stringConexion))
            {
                connection.Open();


                OleDbCommand command = new OleDbCommand("DELETE FROM CELULAS Where cveCelula = " + comboBox1.Text);
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
                    foreach (Celula i in celulas)
                    {
                        if (i.claveCelula == int.Parse(comboBox1.Text))
                        {
                            aux = celulas.IndexOf(i);
                        }
                    }
                    celulas.RemoveAt(aux);
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
