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

namespace MercSolution
{
    public partial class Login : Form
    {
        Boolean acceso = false;
        Boolean aprovado = false;

        
        string pathE=null;
        string pathEI = null;
        string pathA=null;
        string usuario = null;
        string contrasena = null;
        public Login()
        {
            InitializeComponent();


        }

        private void login_Load(object sender, EventArgs e)
        {

        }

        private void acces_Click(object sender, EventArgs e)
        {
            OpenFileDialog file = new OpenFileDialog();
            file.Filter = "Acces Files|*.accdb;*.laccdb;";
            if (file.ShowDialog() == DialogResult.OK)
            {
                pathA = file.FileName;
                Console.WriteLine(pathA);             
            }
        }

        private void excel_Click(object sender, EventArgs e)
        {
            OpenFileDialog file = new OpenFileDialog();
            file.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            if (file.ShowDialog() == DialogResult.OK)
            {
                pathE = file.FileName;
                Console.WriteLine(pathE);
                checkBox2.Checked = true;
            }
        }

        private void aceptar_Click(object sender, EventArgs e)
        {
            if (!acceso)
            {
                string[] lines = null;
                try
                {

                    lines = System.IO.File.ReadAllLines(Directory.GetCurrentDirectory() + "\\usuario.txt");
                    usuario = lines[0];
                    contrasena = lines[1];
                    if (usuario != null && contrasena != null)
                    {
                        if (textBox1.Text == usuario && textBox2.Text == contrasena)
                        {
                            acceso = true;
                            
                            checkBox2.Visible = false;
                            checkBox3.Visible = false;


                        }
                        else
                        {
                            MessageBox.Show("Datos incorrectos");
                        }
                    }
                    else
                    {
                        MessageBox.Show("No se cargaron bien los campos");
                    }
                }
                catch (Exception e1)
                {
                    Console.WriteLine("Exception: " + e1.Message);
                }
                finally
                {
                    Console.WriteLine("Executing finally block.");
                }
            }
            if (acceso)
            {
                textBox1.Visible = false;
                textBox2.Visible = false;
                excel.Visible = true;
                excelIngresos.Visible = true;
                checkBox2.Visible = true;
                checkBox3.Visible = true;
                //acces.Visible = true;
                pictureBox1.Visible = false;
                pictureBox2.Visible = false;
                checkBox1.Visible = false;

            }
            

                
                if (checkBox2.Checked && !checkBox3.Checked)
                {
                    if (pathE != null)
                    {
                        this.Hide();
                        Form1 form1 = new Form1(pathE);
                        
                        form1.Show();
                        
                    }
                    else
                    {
                        MessageBox.Show("Debe seleccionar Excel");
                    }
                    
                }

                if (!checkBox2.Checked && checkBox3.Checked)
                {
                    if (pathEI != null)
                    {
                        this.Hide();
                        Form1 form1 = new Form1(pathEI, 1);
                        form1.Show();
                        
                        
                    }
                    else
                    {
                        MessageBox.Show("Debe seleccionar Excel");
                    }
                }

                if (checkBox2.Checked && checkBox3.Checked)
                {
                    if (pathE != null && pathEI != null)
                    {
                        this.Hide();
                        Form1 form1 = new Form1(pathE, pathEI);
                        form1.Show();
                        
                        
                    }
                    else
                    {
                        MessageBox.Show("Debe seleccionar Excel");
                    }
                }


            


        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            textBox2.UseSystemPasswordChar = true;
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            
        }

        private void pictureBox1_Click_1(object sender, EventArgs e)
        {

        }

        private void textBox1_Click(object sender, EventArgs e)
        {
            textBox1.Text = "";
            
        }

        private void textBox2_Click(object sender, EventArgs e)
        {
            textBox2.Text = "";
            
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                textBox2.UseSystemPasswordChar = false;
            }
            else
            {
                textBox2.UseSystemPasswordChar = true;
            }
        }

        private void excelIngresos_Click(object sender, EventArgs e)
        {
            OpenFileDialog file = new OpenFileDialog();
            file.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            if (file.ShowDialog() == DialogResult.OK)
            {
                pathEI = file.FileName;
                Console.WriteLine(pathE);
                checkBox3.Checked = true;

            }
        }

        private void progressBar1_Click(object sender, EventArgs e)
        {

        }
    }
}
