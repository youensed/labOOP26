using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using Word = Microsoft.Office.Interop.Word;

namespace OOP_26
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            comboBox1.Items.Add(@"C:\Users\bogda\source\repos\OOP_26\OOP_26\DEF.dotx");
        }
        Word.Application word = new Word.Application();
        Word.Document doc;

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                Object missingObj = System.Reflection.Missing.Value;
                Object templatePathObj = comboBox1.SelectedItem.ToString();

                doc = word.Documents.Add(ref templatePathObj, ref missingObj, ref missingObj, ref missingObj);
                doc.Activate();

                foreach (Word.FormField f in doc.FormFields)
                {
                    switch (f.Name)
                    {
                        case "Сonfirmed":
                            f.Range.Text = textBox1.Text;
                            break;
                        case "DateDay":
                            f.Range.Text = textBox6.Text;
                            break;
                        case "DateMonth":
                            f.Range.Text = textBox8.Text;
                            break;
                        case "Items":
                            f.Range.Text = textBox5.Text;
                            break;
                        case "AdressStreet":
                            f.Range.Text = textBox7.Text;
                            break;
                        case "AdressCity":
                            f.Range.Text = textBox2.Text;
                            break;
                        case "Issue":
                            f.Range.Text = textBox13.Text;
                            break;
                        case "WrittenBy":
                            f.Range.Text = textBox10.Text;
                            break;
                        case "CheckedBy":
                            f.Range.Text = textBox14.Text;
                            break;
                        case "Taker":
                            f.Range.Text = textBox15.Text;
                            break;
                        case "Maker":
                            f.Range.Text = textBox9.Text;
                            break;
                    }
                }
                //Збереження по визначеному шляху
                Object savePath = @"C:\Users\bogda\OneDrive\Документы\Збережений файл1.doc";
                doc.SaveAs2(ref savePath);
                //Пошук 
                string findText = textBox11.Text;
                string replaceWith = textBox12.Text;
                bool found = false;

                foreach (Word.Range range in doc.StoryRanges)
                {
                    
                    Word.Find find = range.Find;
                    find.Text = findText;
                    find.Replacement.Text = replaceWith;//заміна тектсу

                    if (find.Execute(Replace: WdReplace.wdReplaceAll))
                    {
                        found = true;
                    }
                }
                
                if (found)
                {
                    MessageBox.Show($"Текст '{findText}' було знайдено та змінено на '{replaceWith}'", "Успіх", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show($"Текст '{findText}' не було знайдено", "Результат", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

                word.Visible = true;


            }
            catch(Exception ex) 
            {
                if (doc != null)
                {
                    doc.Close(WdSaveOptions.wdDoNotSaveChanges);
                    doc = null;
                }

                if (word != null)
                {
                    word.Quit();
                    word = null;
                }

                MessageBox.Show("Виникла помилка: " + ex.Message, "Помилка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
           

        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (doc != null)
            {
                doc.Close(WdSaveOptions.wdDoNotSaveChanges);
                doc = null;
            }

            if (word != null)
            {
                word.Quit();
                word = null;
            }

        }


    }
}
