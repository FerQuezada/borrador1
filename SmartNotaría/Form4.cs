using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SmartNotaría
{
    public partial class Form4 : Form
    {
        public Form4()
        {
            InitializeComponent();
        }

        private void Form4_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            Form3 f3 = new Form3();// la Form2 se convierte en f2
            //this.Hide();//esconde la primera form
            f3.ShowDialog();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            
            Form5 f5 = new Form5();// la Form2 se convierte en f2
            //this.Hide();//esconde la primera form
            f5.ShowDialog();
            
        }

        private void button3_Click(object sender, EventArgs e)
        {
            
            Form6 f6 = new Form6();// la Form2 se convierte en f2
            //this.Hide();//esconde la primera form
            f6.ShowDialog();
            
        }

        private void button4_Click(object sender, EventArgs e)
        {
            
            Form7 f7 = new Form7();// la Form2 se convierte en f2
            //this.Hide();//esconde la primera form
            f7.ShowDialog();
            
        }

        private void button5_Click(object sender, EventArgs e)
        {
            
            Form8 f8 = new Form8();// la Form2 se convierte en f2
            //this.Hide();//esconde la primera form
            f8.ShowDialog();
            
        }
    }
}
