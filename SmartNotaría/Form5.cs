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
    public partial class Form5 : Form
    {
        int con1 = 1;
        int con2 = 1;
        int con3 = 1;
        int con4 = 1;
        int con5 = 1;
        public Form5()
        {
            InitializeComponent();
        }
        private void FindAndReplace(Word.Application wordApp, object ToFindText, object replaceWithText)
        {
            object matchCase = true;
            object matchWholeWord = true;
            object matchWildCards = false;
            object matchSoundLike = false;
            object nmatchAllforms = false;
            object forward = true;
            object format = false;
            object matchKashida = false;
            object matchDiactitics = false;
            object matchAlefHamza = false;
            object matchControl = false;
            object read_only = false;
            object visible = true;
            object replace = 2;
            object wrap = 1;

            wordApp.Selection.Find.Execute(ref ToFindText,
                ref matchCase, ref matchWholeWord,
                ref matchWildCards, ref matchSoundLike,
                ref nmatchAllforms, ref forward,
                ref wrap, ref format, ref replaceWithText,
                ref replace, ref matchKashida,
                ref matchDiactitics, ref matchAlefHamza,
                ref matchControl);
        }



        private void CreateWordDocument(object filename, object SaveAs)//file name is the temple file y saveAs es el que generamos
        {
            // List<int> processesbeforegen = getRunningProcesses();


            Word.Application wordApp = new Word.Application();// se crea una word aplication
            object missing = Missing.Value;//la función tiene muchos paramtros y no queremos utilizarlos todos
            Word.Document myWordDoc = null;//myWordDoc es el documento que vamos a generar y guardar en saveAs file

            if (File.Exists((string)filename))
            {
                object readOnly = false;//tenemos que escribir en ese documento
                object isVisible = false;
                wordApp.Visible = false;

                myWordDoc = wordApp.Documents.Open(ref filename, ref missing, ref readOnly,//abrimos el machote, ponemos el nombre del maquite que quereos abrir
                                                  ref missing, ref missing, ref missing,
                                                  ref missing, ref missing, ref missing,
                                                  ref missing, ref missing, ref missing,
                                                  ref missing, ref missing, ref missing, ref missing);
                myWordDoc.Activate();

                //find and replace
                if (textBox1.Text == "")
                {

                }
                else
                {
                    this.FindAndReplace(wordApp, "<albacea>", textBox1.Text);
                }
                if (textBox2.Text == "")
                {

                }
                else
                {
                    this.FindAndReplace(wordApp, "<de cujus>", textBox2.Text);
                }
                if (textBox3.Text == "")
                {

                }
                else
                {
                    this.FindAndReplace(wordApp, "<escrituranumeroradicacion>", textBox3.Text);
                }
                if (textBox4.Text == "")
                {

                }
                else
                {
                    this.FindAndReplace(wordApp, "<volumenradicacion>", textBox4.Text);
                }


                this.FindAndReplace(wordApp, "<fecha>", dateTimePicker1.Value.ToLongDateString());
                /*
                this.FindAndReplace(wordApp, "<fechatestamento>", dateTimePicker2.Value.ToLongDateString());
                this.FindAndReplace(wordApp, "<fechanacimientoalbacea>", dateTimePicker3.Value.ToLongDateString());
                this.FindAndReplace(wordApp, "<fechafallecimiento>", dateTimePicker4.Value.ToLongDateString());
                this.FindAndReplace(wordApp, "<fechacertificadotestamento>", dateTimePicker5.Value.ToLongDateString());
                */
                this.FindAndReplace(wordApp, "<fecha1>", DateTime.Now.ToLongDateString());
            }
            else
            {
                MessageBox.Show("File not Found!");
            }
            //SaveAs2 es el lugar en el que se va a guardar, SaveAs es el camino para generar el documento
            myWordDoc.SaveAs2(ref SaveAs, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing,//
                ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing);

            myWordDoc.Close(); //esto sí estaba
            wordApp.Quit();// esto sí estaba
            //MessageBox.Show("File Created!");

            //List<int> processesaftergen = getRunningProcesses();
            //killProcesses(processesbeforegen, processesaftergen);



        }

        private void Form5_Load(object sender, EventArgs e)
        {

        }

        private void abrirToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                label2.Text = openFileDialog1.FileName;//label2 es la etiqueta al lado del boton para cargar un documento anterior
                //textBox1.Text = openFileDialog1.FileName;
                //tEnabled(true);
            }
        }

        private void guardarToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                if (label2.Text == "")//no hay nada en label2 porque no se tomó ningun documento anterior
                {
                    //C: \Users\fer_q\OneDrive\Documentos\PROYECTONOTARIA\SmartNotaría\SmartNotaría
                    // C:\Users\fer_q\OneDrive\Documentos\PROYECTONOTARIA\SmartNotaría\SmartNotaría\RADICACION


                    //CreateWordDocument(@"C:\Users\fer_q\OneDrive\Documentos\PROYECTONOTARIA\SmartNotaría\SmartNotaría\RADICACION.docx", saveFileDialog1.FileName + @"Radicación.docx");
                    //MessageBox.Show("Se generó la Radicación!");
                    CreateWordDocument(@"C:\Users\fer_q\OneDrive\Documentos\PROYECTONOTARIA\SmartNotaría\SmartNotaría\AVISO.docx", saveFileDialog1.FileName + @"Edictos.docx");
                    MessageBox.Show("Se generaron los Edictos!");
                    //CreateWordDocument(@"C:\Users\fer_q\OneDrive\Documentos\PROYECTONOTARIA\SmartNotaría\SmartNotaría\ADJUDICACION.docx", saveFileDialog1.FileName + @"Adjudicación.docx");
                    //MessageBox.Show("Se generó la Adjudicación!");
                    label4.Text = saveFileDialog1.FileName;//label4 es la que está al lado de "se guardó en:"
                }
                else
                {

                    //CreateWordDocument(label2.Text, saveFileDialog1.FileName + @"Radicación.docx");//(DE DONDE LO TOMO, EN DONDE LO GUARDO)
                    //MessageBox.Show("Se modificó la Radicación!");
                    
                    
                    CreateWordDocument(label2.Text, saveFileDialog1.FileName + @"Edictos.docx");
                    MessageBox.Show("Se modificaron los Edictos!");
                    label4.Text = saveFileDialog1.FileName;

                    string sdato = label2.Text, sdato2;//estoy cambiando el nombre del documento que abrí de "Radicación" a "Edictos"
                    //string sdato3 = label2.Text, sdato4;
                    sdato2 = sdato.Replace("Edictos", "Adjudicación");

                    //sdato4 = sdato3.Replace("Radicación", "Adjudicación");
                    CreateWordDocument(sdato2, saveFileDialog1.FileName + @"Adjudicación.docx");
                    MessageBox.Show("Se modificó la Adjudicación!");
                    //label41.Text = sdato2;//lo imprimo en esta etiqueta para ver si funciona




                }
                
            }
        }

        private void label2_Click(object sender, EventArgs e)
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

        private void textBox3_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                SendKeys.Send("{TAB}");
            }
        }

        private void textBox4_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                SendKeys.Send("{TAB}");
            }
        }

        private void dateTimePicker1_KeyDown(object sender, KeyEventArgs e)
        {
            progressBar1.Increment(20);
            label11.Text = progressBar1.Value.ToString() + "%";
            if (e.KeyCode == Keys.Enter)
            {
                SendKeys.Send("{TAB}");
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            con1 = con1 + 1;
            if (con1 == 5)
            {
                progressBar1.Increment(20);
                label11.Text = progressBar1.Value.ToString() + "%";
            }
            if (textBox1.Text == "")
            {
                progressBar1.Increment(-20);
                label11.Text = progressBar1.Value.ToString() + "%";
                con1 = 1;
            }
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            con2 = con2 + 1;
            if (con2 == 5)
            {
                progressBar1.Increment(20);
                label11.Text = progressBar1.Value.ToString() + "%";
            }
            if (textBox2.Text == "")
            {
                progressBar1.Increment(-20);
                label11.Text = progressBar1.Value.ToString() + "%";
                con2 = 1;
            }
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            con3 = con3 + 1;
            if (con3 == 5)
            {
                progressBar1.Increment(20);
                label11.Text = progressBar1.Value.ToString() + "%";
            }
            if (textBox3.Text == "")
            {
                progressBar1.Increment(-20);
                label11.Text = progressBar1.Value.ToString() + "%";
                con3 = 1;
            }
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            con4 = con4 + 1;
            if (con4 == 5)
            {
                progressBar1.Increment(20);
                label11.Text = progressBar1.Value.ToString() + "%";
            }
            if (textBox4.Text == "")
            {
                progressBar1.Increment(-20);
                label11.Text = progressBar1.Value.ToString() + "%";
                con4 = 1;
            }
        }

        private void dateTimePicker1_KeyPress(object sender, KeyPressEventArgs e)
        {
            
        }

        private void documentoToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }
    }
}
