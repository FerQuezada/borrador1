using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;



using System.IO;
using System.Reflection;
using Word = Microsoft.Office.Interop.Word;

using Microsoft.Office.Interop.Word;
using Microsoft.Office.Core;
using System.Diagnostics;
using System.Drawing.Drawing2D;




namespace SmartNotaría
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            textBox1.MaxLength = 19;
            textBox2.MaxLength = 19;
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                SendKeys.Send("{TAB}");
            }
        }
        private void textBox2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                SendKeys.Send("{TAB}");
            }
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Form2 f2 = new Form2();// la Form2 se convierte en f2
            //Form3 f3 = new Form3();// la Form2 se convierte en f2
            //this.Hide();//esconde la primera form
            //f2.ShowDialog();
            //f2.ShowDialog();


            if (textBox1.Text == "FERNANDO QUEZADA")
            {
                if(textBox2.Text == "NOTARIA35")
                {
                    //MessageBox.Show("Username and password is correct");
                    Form2 f2 = new Form2();// la Form2 se convierte en f2
                    this.Hide();//esconde la primera form
                    

                    f2.ShowDialog();
                }
                else
                {
                    MessageBox.Show("Username or Password is not correct .. ");
                }
                
            }
            else
            {
                MessageBox.Show("Username or Password is not correct .. ");
            }
        }
    }
}
