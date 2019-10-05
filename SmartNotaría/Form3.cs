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
    public partial class Form3 : Form
    {
        int con1 = 1;
        
        int con2 = 1;
        int con3 = 1;
        int con4 = 1;
        int con5 = 1;
        int con6 = 1;
        int con7 = 1;
        int con8 = 1;
        int con9 = 1;
        int con10 = 1;
        int con11 = 1;
        int con12 = 1;
        int con13 = 1;
        int con14 = 1;
        int con15 = 1;
        int con16 = 1;
        int con17 = 1;
        int con18 = 1;
        int con19 = 1;
        int con20 = 1;
        //88888
        


        public Form3()
        {
            InitializeComponent();
            
            //Form2 f2 = new Form2();// la Form2 se convierte en f2
            //this.Hide();//esconde la primera form
            //f2.ShowDialog();
            
        }
        /*
        private void Form3_Load(object sender, EventArgs e)
        {
            Panel my_panel = new Panel();
            VScrollBar vScroller = new VScrollBar();
            vScroller.Dock = DockStyle.Right;
            vScroller.Width = 30;
            vScroller.Height = 4000;
            vScroller.Name = "VScrollBar1";
            my_panel.Controls.Add(vScroller);
        }
        */
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
                    this.FindAndReplace(wordApp, "<escrituranumeroradicacion>", textBox1.Text);
                }
                if (textBox2.Text == "")
                {

                }
                else
                {
                    this.FindAndReplace(wordApp, "<volumenradicacion>", textBox2.Text);
                }
                if (textBox3.Text == "")
                {

                }
                else
                {
                    this.FindAndReplace(wordApp, "<de cujus>", textBox3.Text);
                }
                if (textBox4.Text == "")
                {

                }
                else
                {
                    this.FindAndReplace(wordApp, "<albacea>", textBox4.Text);
                }

                if (textBox5.Text == "")
                {

                }
                else
                {
                    this.FindAndReplace(wordApp, "<testamentonumero>", textBox5.Text);
                }

                if (textBox6.Text == "")
                {

                }
                else
                {
                    this.FindAndReplace(wordApp, "<testamentovolumen>", textBox6.Text);
                }

                if (textBox7.Text == "")
                {

                }
                else
                {
                    this.FindAndReplace(wordApp, "<notariotestamento>", textBox7.Text);
                }

                if (textBox8.Text == "")
                {

                }
                else
                {
                    this.FindAndReplace(wordApp, "<numeronotariotestamento>", textBox8.Text);
                }

                if (textBox9.Text == "")
                {

                }
                else
                {
                    string contenido = textBox9.Text, contenido1="", contenido2 = "", contenido3 = "", contenido4 = "", contenido5 = "";
                    string contenido6 = "", contenido7 = "", contenido8 = "", contenido9 = "", contenido10 = "";
                    int conte = contenido.Length;
                    int conte1;
                    
                    conte1 = conte / 10;
                    contenido1 = contenido.Substring(0, conte1);//(DE DÓNDE, CUÁNTOS)
                    
                    contenido2 = contenido.Substring(conte1, conte1);
                    
                    contenido3 = contenido.Substring((conte1 + conte1), (conte1));
                    
                    contenido4 = contenido.Substring((3*conte1), conte1);
                    contenido5 = contenido.Substring((4*conte1), conte1);
                    contenido6 = contenido.Substring((5*conte1), conte1);
                    contenido7 = contenido.Substring((6*conte1), conte1);
                    contenido8 = contenido.Substring((7*conte1), conte1);
                    contenido9 = contenido.Substring((8*conte1), conte1);
                    contenido10 = contenido.Substring((9*conte1), conte1);
                    


                    //this.FindAndReplace(wordApp, "<clausulastestamento>", textBox9.Text);//dynamicTextBox.Text;textBox9.Text
                    this.FindAndReplace(wordApp, "<clausulastestamento1>", contenido1);
                    
                    this.FindAndReplace(wordApp, "<clausulastestamento2>", contenido2);
                    
                    this.FindAndReplace(wordApp, "<clausulastestamento3>", contenido3);
                    
                    this.FindAndReplace(wordApp, "<clausulastestamento4>", contenido4);
                    this.FindAndReplace(wordApp, "<clausulastestamento5>", contenido5);
                    this.FindAndReplace(wordApp, "<clausulastestamento6>", contenido6);
                    this.FindAndReplace(wordApp, "<clausulastestamento7>", contenido7);
                    this.FindAndReplace(wordApp, "<clausulastestamento8>", contenido8);
                    this.FindAndReplace(wordApp, "<clausulastestamento9>", contenido9);
                    this.FindAndReplace(wordApp, "<clausulastestamento10>", contenido10);
                    



                }

                if (textBox10.Text == "")
                {

                }
                else
                {
                    this.FindAndReplace(wordApp, "<nacionalidad>", textBox10.Text);
                }

                if (textBox11.Text == "")
                {

                }
                else
                {
                    this.FindAndReplace(wordApp, "<ciudadfallecimiento>", textBox11.Text);
                }

                if (textBox12.Text == "")
                {

                }
                else
                {
                    this.FindAndReplace(wordApp, "<estadocivilalbacea>", textBox12.Text);
                }

                if (textBox13.Text == "")
                {

                }
                else
                {
                    this.FindAndReplace(wordApp, "<ocupacionalbacea>", textBox13.Text);
                }

                if (textBox14.Text == "")
                {

                }
                else
                {
                    this.FindAndReplace(wordApp, "<originarioalbacea>", textBox14.Text);
                }

                if (textBox15.Text == "")
                {

                }
                else
                {
                    this.FindAndReplace(wordApp, "<domicilioalbacea>", textBox15.Text);
                }

                if (textBox16.Text == "")
                {

                }
                else
                {
                    this.FindAndReplace(wordApp, "<codigopostalalbacea>", textBox16.Text);
                }

                if (textBox17.Text == "")
                {

                }
                else
                {
                    this.FindAndReplace(wordApp, "<curpalbacea>", textBox17.Text);
                }

                if (textBox18.Text == "")
                {

                }
                else
                {
                    this.FindAndReplace(wordApp, "<rfcalbacea>", textBox18.Text);
                }

                if (textBox19.Text == "")
                {

                }
                else
                {
                    this.FindAndReplace(wordApp, "<inealbacea>", textBox19.Text);
                }


                this.FindAndReplace(wordApp, "<fecha>", dateTimePicker1.Value.ToLongDateString());
                this.FindAndReplace(wordApp, "<fechatestamento>", dateTimePicker2.Value.ToLongDateString());
                this.FindAndReplace(wordApp, "<fechanacimientoalbacea>", dateTimePicker3.Value.ToLongDateString());
                this.FindAndReplace(wordApp, "<fechafallecimiento>", dateTimePicker4.Value.ToLongDateString());
                this.FindAndReplace(wordApp, "<fechacertificadotestamento>", dateTimePicker5.Value.ToLongDateString());

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




        private void Form3_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                label4.Text = openFileDialog1.FileName;//label4 es la etiqueta al lado del boton para cargar un documento anterior
                //textBox1.Text = openFileDialog1.FileName;
                //tEnabled(true);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                if(label4.Text == "")//no hay nada en label4 porque no se tomó ningun documento anterior
                {
                //C: \Users\fer_q\OneDrive\Documentos\PROYECTONOTARIA\SmartNotaría\SmartNotaría
                       // C:\Users\fer_q\OneDrive\Documentos\PROYECTONOTARIA\SmartNotaría\SmartNotaría\RADICACION


                    CreateWordDocument(@"C:\Users\fer_q\OneDrive\Documentos\PROYECTONOTARIA\SmartNotaría\SmartNotaría\RADICACION.docx", saveFileDialog1.FileName + @"Radicación.docx");
                    MessageBox.Show("Se generó la Radicación!");
                    CreateWordDocument(@"C:\Users\fer_q\OneDrive\Documentos\PROYECTONOTARIA\SmartNotaría\SmartNotaría\AVISO.docx", saveFileDialog1.FileName + @"Edictos.docx");
                    MessageBox.Show("Se generaron los Edictos!");
                    CreateWordDocument(@"C:\Users\fer_q\OneDrive\Documentos\PROYECTONOTARIA\SmartNotaría\SmartNotaría\ADJUDICACION.docx", saveFileDialog1.FileName + @"Adjudicación.docx");
                    MessageBox.Show("Se generó la Adjudicación!");
                    label3.Text = saveFileDialog1.FileName;
                }
                else
                {
                    
                    CreateWordDocument(label4.Text, saveFileDialog1.FileName + @"Radicación.docx");//(DE DONDE LO TOMO, EN DONDE LO GUARDO)
                    MessageBox.Show("Se modificó la Radicación!");
                    label3.Text = saveFileDialog1.FileName;
                    string sdato = label4.Text, sdato2;//estoy cambiando el nombre del documento que abrí de "Radicación" a "Edictos"
                    string sdato3 = label4.Text, sdato4;
                    sdato2 = sdato.Replace("Radicación", "Edictos");
                    CreateWordDocument(sdato2, saveFileDialog1.FileName + @"Edictos.docx");
                    MessageBox.Show("Se modificaron los Edictos!");
                    sdato4 = sdato3.Replace("Radicación", "Adjudicación");
                    CreateWordDocument(sdato4, saveFileDialog1.FileName + @"Adjudicación.docx");
                    MessageBox.Show("Se modificó la Adjudicación!");
                    //label41.Text = sdato2;//lo imprimo en esta etiqueta para ver si funciona




                }
                /*
                if (textBox1.Text == "")
                {
                    label3.Text = saveFileDialog1.FileName;
                    CreateWordDocument(@"C:\Users\fer_q\OneDrive\Documentos\PROYECTONOTARIA\SmartNotaría\SmartNotaría\AVISO.docx", saveFileDialog1.FileName + @".docx");//no reconoce los / por lo que hay que poner @
                }
                else
                {
                    CreateWordDocument(textBox1.Text, saveFileDialog1.FileName + @".docx");
                }
                */
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            
            con1=con1+1;
            if (con1 == 5)
            {
                progressBar1.Increment(5);
                label43.Text = progressBar1.Value.ToString() + "%";
            }
            if (textBox1.Text == "")
            {
                progressBar1.Increment(-5);
                label43.Text = progressBar1.Value.ToString() + "%";
                con1 = 1;
            }
            
        }

        private void textBox10_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                SendKeys.Send("{TAB}");
                
            }
        }

        private void dateTimePicker3_ValueChanged(object sender, EventArgs e)
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

        private void textBox16_TextChanged(object sender, EventArgs e)
        {
            
            con10 = con10 + 1;
            if (con10 == 5)
            {
                progressBar1.Increment(5);
                label43.Text = progressBar1.Value.ToString() + "%";
            }
            if (textBox16.Text == "")
            {
                progressBar1.Increment(-5);
                label43.Text = progressBar1.Value.ToString() + "%";
                con10 = 1;
            }
            

        }

        private void textBox10_KeyDown_1(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                SendKeys.Send("{TAB}");
            }

        }

        private void textBox12_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                SendKeys.Send("{TAB}");
            }

        }

        private void textBox13_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                SendKeys.Send("{TAB}");
            }

        }

        private void textBox14_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                SendKeys.Send("{TAB}");
            }

        }

        private void textBox15_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                SendKeys.Send("{TAB}");
            }

        }

        private void textBox16_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                SendKeys.Send("{TAB}");
            }

        }

        private void textBox17_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                SendKeys.Send("{TAB}");
            }

        }

        private void textBox18_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                SendKeys.Send("{TAB}");
            }

        }

        private void textBox19_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                SendKeys.Send("{TAB}");
            }

        }

        private void textBox5_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                SendKeys.Send("{TAB}");
            }

        }

        private void textBox6_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                SendKeys.Send("{TAB}");
            }

        }

        private void textBox7_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                SendKeys.Send("{TAB}");
            }

        }

        private void textBox9_KeyDown(object sender, KeyEventArgs e)
        {
            /*
            if (e.KeyCode == Keys.Enter)
            {
                SendKeys.Send("{TAB}");
            }
            */
        }

        private void button2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                SendKeys.Send("{TAB}");
            }

        }

        private void button1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                SendKeys.Send("{TAB}");
            }

        }

        private void dateTimePicker1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                SendKeys.Send("{TAB}");
            }
        }

        private void dateTimePicker3_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                SendKeys.Send("{TAB}");
            }
        }

        private void dateTimePicker2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                SendKeys.Send("{TAB}");
            }
        }

        private void textBox11_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                SendKeys.Send("{TAB}");
            }
        }

        private void dateTimePicker4_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                SendKeys.Send("{TAB}");
            }
        }

        private void dateTimePicker5_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                SendKeys.Send("{TAB}");
            }
        }

        
        private void button5_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                label41.Text = openFileDialog1.FileName;//label4 es la etiqueta al lado del boton para cargar un documento anterior
                //textBox1.Text = openFileDialog1.FileName;
                //tEnabled(true);
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                label42.Text = openFileDialog1.FileName;//label4 es la etiqueta al lado del boton para cargar un documento anterior
                //textBox1.Text = openFileDialog1.FileName;
                //tEnabled(true);
            }
        }
        
        private void button3_Click(object sender, EventArgs e)
        {
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                if (label41.Text == "")//no hay nada en label41 porque no se tomó ningun documento anterior
                {
                    //C: \Users\fer_q\OneDrive\Documentos\PROYECTONOTARIA\SmartNotaría\SmartNotaría
                    // C:\Users\fer_q\OneDrive\Documentos\PROYECTONOTARIA\SmartNotaría\SmartNotaría\RADICACION


                    CreateWordDocument(@"C:\Users\fer_q\OneDrive\Documentos\PROYECTONOTARIA\SmartNotaría\SmartNotaría\AVISO.docx", saveFileDialog1.FileName + @".docx");
                    label36.Text = saveFileDialog1.FileName;
                }
                else
                {
                    CreateWordDocument(label41.Text, saveFileDialog1.FileName + @".docx");
                    label36.Text = saveFileDialog1.FileName;
                }
                
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                if (label42.Text == "")//no hay nada en label42 porque no se tomó ningun documento anterior
                {


                    CreateWordDocument(@"C:\Users\fer_q\OneDrive\Documentos\PROYECTONOTARIA\SmartNotaría\SmartNotaría\ADJUDICACION.docx", saveFileDialog1.FileName + @".docx");
                    label37.Text = saveFileDialog1.FileName;
                }
                else
                {
                    CreateWordDocument(label42.Text, saveFileDialog1.FileName + @".docx");
                    label37.Text = saveFileDialog1.FileName;
                }

            }
        }

        private void textBox9_TextChanged(object sender, EventArgs e)
        {
            
            con19 = con19 + 1;
            if (con19 == 5)
            {
                progressBar1.Increment(10);
                label43.Text = progressBar1.Value.ToString() + "%";
            }
            if (textBox9.Text == "")
            {
                progressBar1.Increment(-10);
                label43.Text = progressBar1.Value.ToString() + "%";
                con19 = 1;
            }
            
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            
            con3 = con3 + 1;
            if (con3 == 5)
            {
                progressBar1.Increment(5);
                label43.Text = progressBar1.Value.ToString() + "%";
            }
            if (textBox3.Text == "")
            {
                progressBar1.Increment(-5);
                label43.Text = progressBar1.Value.ToString() + "%";
                con3 = 1;
            }
            
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            
            con2 = con2 + 1;
            if (con2 == 5)
            {
                progressBar1.Increment(5);
                label43.Text = progressBar1.Value.ToString() + "%";
            }
            if (textBox2.Text == "")
            {
                progressBar1.Increment(-5);
                label43.Text = progressBar1.Value.ToString() + "%";
                con2 = 1;
            }
            
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            
            con4 = con4 + 1;
            if (con4 == 5)
            {
                progressBar1.Increment(5);
                label43.Text = progressBar1.Value.ToString() + "%";
            }
            if (textBox4.Text == "")
            {
                progressBar1.Increment(-5);
                label43.Text = progressBar1.Value.ToString() + "%";
                con4 = 1;
            }
            
        }

        private void textBox10_TextChanged(object sender, EventArgs e)
        {
            
            con5 = con5 + 1;
            if (con5 == 5)
            {
                progressBar1.Increment(5);
                label43.Text = progressBar1.Value.ToString() + "%";
            }
            if (textBox10.Text == "")
            {
                progressBar1.Increment(-5);
                label43.Text = progressBar1.Value.ToString() + "%";
                con5 = 1;
            }
            
        }

        private void textBox12_TextChanged(object sender, EventArgs e)
        {
            
            con6 = con6 + 1;
            if (con6 == 5)
            {
                progressBar1.Increment(5);
                label43.Text = progressBar1.Value.ToString() + "%";
            }
            if (textBox12.Text == "")
            {
                progressBar1.Increment(-5);
                label43.Text = progressBar1.Value.ToString() + "%";
                con6 = 1;
            }
            
        }

        private void textBox13_TextChanged(object sender, EventArgs e)
        {
            
            con7 = con7 + 1;
            if (con7 == 5)
            {
                progressBar1.Increment(5);
                label43.Text = progressBar1.Value.ToString() + "%";
            }
            if (textBox13.Text == "")
            {
                progressBar1.Increment(-5);
                label43.Text = progressBar1.Value.ToString() + "%";
                con7 = 1;
            }
            
        }

        private void textBox14_TextChanged(object sender, EventArgs e)
        {
            
            con8 = con8 + 1;
            if (con8 == 5)
            {
                progressBar1.Increment(5);
                label43.Text = progressBar1.Value.ToString() + "%";
            }
            if (textBox14.Text == "")
            {
                progressBar1.Increment(-5);
                label43.Text = progressBar1.Value.ToString() + "%";
                con8 = 1;
            }
            
        }

        private void textBox15_TextChanged(object sender, EventArgs e)
        {
            
            con9 = con9 + 1;
            if (con9 == 5)
            {
                progressBar1.Increment(5);
                label43.Text = progressBar1.Value.ToString() + "%";
            }
            if (textBox15.Text == "")
            {
                progressBar1.Increment(-5);
                label43.Text = progressBar1.Value.ToString() + "%";
                con9 = 1;
            }
            
        }

        private void textBox17_TextChanged(object sender, EventArgs e)
        {
            
            con11 = con11 + 1;
            if (con11 == 5)
            {
                progressBar1.Increment(5);
                label43.Text = progressBar1.Value.ToString() + "%";
            }
            if (textBox17.Text == "")
            {
                progressBar1.Increment(-5);
                label43.Text = progressBar1.Value.ToString() + "%";
                con11 = 1;
            }
            
        }

        private void textBox18_TextChanged(object sender, EventArgs e)
        {
            
            con12 = con12 + 1;
            if (con12 == 5)
            {
                progressBar1.Increment(5);
                label43.Text = progressBar1.Value.ToString() + "%";
            }
            if (textBox18.Text == "")
            {
                progressBar1.Increment(-5);
                label43.Text = progressBar1.Value.ToString() + "%";
                con12 = 1;
            }
            
        }

        private void textBox19_TextChanged(object sender, EventArgs e)
        {
            
            con13 = con13 + 1;
            if (con13 == 5)
            {
                progressBar1.Increment(5);
                label43.Text = progressBar1.Value.ToString() + "%";
            }
            if (textBox19.Text == "")
            {
                progressBar1.Increment(-5);
                label43.Text = progressBar1.Value.ToString() + "%";
                con13 = 1;
            }
            
        }

        private void textBox11_TextChanged(object sender, EventArgs e)
        {
            
            con14 = con14 + 1;
            if (con14 == 5)
            {
                progressBar1.Increment(5);
                label43.Text = progressBar1.Value.ToString() + "%";
            }
            if (textBox11.Text == "")
            {
                progressBar1.Increment(-5);
                label43.Text = progressBar1.Value.ToString() + "%";
                con14 = 1;
            }
            
        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {
            
            con15 = con15 + 1;
            if (con15 == 5)
            {
                progressBar1.Increment(5);
                label43.Text = progressBar1.Value.ToString() + "%";
            }
            if (textBox5.Text == "")
            {
                progressBar1.Increment(-5);
                label43.Text = progressBar1.Value.ToString() + "%";
                con15 = 1;
            }
            
        }

        private void textBox6_TextChanged(object sender, EventArgs e)
        {
            
            con16 = con16 + 1;
            if (con16 == 5)
            {
                progressBar1.Increment(5);
                label43.Text = progressBar1.Value.ToString() + "%";
            }
            if (textBox6.Text == "")
            {
                progressBar1.Increment(-5);
                label43.Text = progressBar1.Value.ToString() + "%";
                con16 = 1;
            }
            
        }

        private void textBox7_TextChanged(object sender, EventArgs e)
        {
            
            con17 = con17 + 1;
            if (con17 == 5)
            {
                progressBar1.Increment(5);
                label43.Text = progressBar1.Value.ToString() + "%";
            }
            if (textBox7.Text == "")
            {
                progressBar1.Increment(-5);
                label43.Text = progressBar1.Value.ToString() + "%";
                con17 = 1;
            }
            
        }

        private void textBox8_TextChanged(object sender, EventArgs e)
        {
            
            con18 = con18 + 1;
            if (con18 == 5)
            {
                progressBar1.Increment(5);
                label43.Text = progressBar1.Value.ToString() + "%";
            }
            if (textBox8.Text == "")
            {
                progressBar1.Increment(-5);
                label43.Text = progressBar1.Value.ToString() + "%";
                con18 = 1;
            }
            
        }
    }
}
