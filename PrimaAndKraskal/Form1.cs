using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using System.Diagnostics;
using System.Reflection;
using System.IO;
//using System.Reflection;

namespace PrimaAndKraskal
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            alfavit = new string[] {"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", };
                
            LoadExample(5, 1, 6);
            Calculate();
        }

        Word._Application oWord;
        Word._Document oDoc;
        object oMissing = System.Reflection.Missing.Value;
        object oEndOfDoc = "\\endofdoc";

        private string[] alfavit;

        //Для Прима
        List<Edge> E1 = new List<Edge>();  //Подсчитываем рёбра
        List<Edge> MST1 = new List<Edge>();
        //Для Краскала
        List<Edge> E2 = new List<Edge>();  //Подсчитываем рёбра
        List<Edge> MST2 = new List<Edge>();

        PrimsAlgoritm pa = new PrimsAlgoritm();
        KruskalsAlgoritm ka = new KruskalsAlgoritm();

        int CurrentGraf = 1;

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                if (dataGridView1.RowCount == 0 || dataGridView1.ColumnCount == 0)
                    throw new Exception("Таблица не создана!");
                for (int i = 0; i < dataGridView1.RowCount; i++)
                    if (dataGridView1.Rows[i].Cells[i].Value == null || dataGridView1.Rows[i].Cells[i].Value.ToString() != "0")
                        throw new Exception("Матрица не симметрична!");
                
                for (int i = 0; i < dataGridView1.RowCount; i++)
                    for (int j = 0; j < dataGridView1.ColumnCount; j++)
                    if (((dataGridView1.Rows[i].Cells[j].Value == null) || (dataGridView1.Rows[j].Cells[i].Value == null))  || (dataGridView1.Rows[i].Cells[j].Value.ToString() != dataGridView1.Rows[j].Cells[i].Value.ToString()))
                            throw new Exception("Матрица не симметрична!");
                Calculate();
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public void Calculate()
        {
            LoadEdgeToList(E1, MST1);
            LoadEdgeToList(E2, MST2);

            pa.Calculate(dataGridView1.RowCount, E1, MST1);
            E2.Sort();
            ka.Calculate(dataGridView1.RowCount, E2, MST2);

            PrintMSTtoListBox(listBox1, MST1);
            PrintMSTtoListBox(listBox2, MST2);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                int countVersh = Convert.ToInt32(textBox1.Text);
                if (countVersh > 26)
                    throw new Exception("Количество вершин не должно превашать 26!");
                int minDigit = Convert.ToInt32(textBox2.Text);
                int maxDigit = Convert.ToInt32(textBox3.Text);
                if (minDigit >= maxDigit)
                    throw new Exception("Значение min не должно равняться или превышать max");
                LoadExample(countVersh, minDigit, maxDigit);
                Calculate();
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
                for (int i = 0; i < dataGridView1.RowCount; i++)
                    if (dataGridView1.Rows[i].Cells[i].Value == null || dataGridView1.Rows[i].Cells[i].Value.ToString() != "0")
                        throw new Exception("Матрица не симметрична");

                for (int i = 0; i < dataGridView1.RowCount; i++)
                    for (int j = 0; j < dataGridView1.ColumnCount; j++)
                        if (((dataGridView1.Rows[i].Cells[j].Value == null) || (dataGridView1.Rows[j].Cells[i].Value == null)) || (dataGridView1.Rows[i].Cells[j].Value.ToString() != dataGridView1.Rows[j].Cells[i].Value.ToString()))
                            throw new Exception("Матрица не симметрична");
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
                return;
            }

            object oRng;
            Word.Paragraph oPara4;
            oRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            oPara4 = oDoc.Content.Paragraphs.Add(ref oRng);
            oPara4.Range.InsertParagraphBefore();
            oPara4.Range.Text = "Граф "+ CurrentGraf.ToString();
            oPara4.Format.SpaceAfter = 5;
            oPara4.Range.InsertParagraphAfter();

            Word.Table oTable;
            Word.Range wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            oTable = oDoc.Tables.Add(wrdRng, dataGridView1.RowCount + 1, dataGridView1.ColumnCount + 1, ref oMissing, ref oMissing);
            oTable.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
            oTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
            oTable.Range.ParagraphFormat.SpaceAfter = 1;

            for(int i = 2; i <= dataGridView1.ColumnCount + 1; i++)
            {
                oTable.Cell(1, i).Range.Text = alfavit[i - 2];
                oTable.Cell(1, i).Range.Font.Bold = 1;
                oTable.Cell(1, i).Range.Font.Italic = 1;
            }

            for (int i = 2; i <= dataGridView1.RowCount + 1; i++)
            {
                oTable.Cell(i, 1).Range.Text = alfavit[i - 2];
                oTable.Columns[1].Cells[i].Range.Font.Bold = 1;
                oTable.Columns[1].Cells[i].Range.Font.Italic = 1;
            }

            for (int i = 1; i <= dataGridView1.ColumnCount + 1; i++)
            {
                oTable.Columns[i].Width = 30;
            }

            for (int i = 0; i < dataGridView1.RowCount ; i++)
                for (int j = 0; j < dataGridView1.ColumnCount; j++)
                {
                    oTable.Cell(i + 2, j + 2).Range.Text = dataGridView1.Rows[i].Cells[j].Value.ToString();
                }


            oRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            oPara4 = oDoc.Content.Paragraphs.Add(ref oRng);
            oPara4.Range.InsertParagraphBefore();
            oPara4.Range.Text = "Решение (Граф " + CurrentGraf.ToString() + ")";
            oPara4.Format.SpaceAfter = 1;//24;
            oPara4.Range.InsertParagraphAfter();


            wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            oTable = oDoc.Tables.Add(wrdRng, 2, 2, ref oMissing, ref oMissing);
            oTable.Columns[1].Cells[1].Range.Text = "Прима";
            oTable.Columns[1].Width = 200;
            oTable.Columns[2].Cells[1].Range.Text = "Краскал";
            oTable.Columns[2].Width = 500;
            oTable.Rows[1].Height = 5;

            for (int i = 1; i <= 2; i++)
            {
                oTable.Columns[i].Width = 100;
            }

            List<string> MSTPRIMA = SpecialForWord(MST1);
            List<string> MSTKRUSKAL = SpecialForWord(MST2);

            oTable.Columns[1].Cells[2].Range.Text = "G(0)";
            for (int i = 0; i < MSTPRIMA.Count; i++)
            {
                oTable.Columns[1].Cells[2].Range.Text += MSTPRIMA[i];
            }

            oTable.Columns[2].Cells[2].Range.Text = "G(0)";
            for (int i = 0; i < MSTKRUSKAL.Count; i++)
            {
                oTable.Columns[2].Cells[2].Range.Text += MSTKRUSKAL[i];
            }

            oTable.Columns[1].Cells[2].Range.ParagraphFormat.TextboxTightWrap = Word.WdTextboxTightWrap.wdTightNone;
            oTable.Columns[1].Cells[2].Range.ParagraphFormat.SpaceAfter = 0.01f;
            oTable.Columns[2].Cells[2].Range.ParagraphFormat.SpaceAfter = 0.01f;


            for (int i = 1; i <= 2; i++)
            {
                oTable.Columns[i].Width = 150;
            }

            
            oTable.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
            oTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;

            oRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            oPara4 = oDoc.Content.Paragraphs.Add(ref oRng);
            oPara4.Range.InsertParagraphBefore();
            oPara4.Format.SpaceAfter = 1;
            oPara4.Range.InsertParagraphAfter();

            ++CurrentGraf;
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            listBox1.Items.Clear();
            listBox2.Items.Clear();

            try
            {
                dataGridView1.RowCount = 0;
                dataGridView1.ColumnCount = 0;
                int razmer = Convert.ToInt32(textBox4.Text);
                if (razmer > 26)
                    throw new Exception("Количество вершин не должно превашать 26!");
                dataGridView1.RowCount = razmer;
                dataGridView1.ColumnCount = razmer;
                for (int i = 0; i < dataGridView1.ColumnCount; i++)
                    dataGridView1.Columns[i].Width = 40;
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void LoadEdgeToList(List<Edge> E, List<Edge> MST)
        {
            E.Clear();
            for (int i = 0; i < dataGridView1.ColumnCount; i++)
                for (int j = 0; j < dataGridView1.RowCount; j++)
                {
                    if (dataGridView1.Rows[i].Cells[j].Value.ToString() != "0")
                        E.Add(new Edge(i, j, Convert.ToInt32(dataGridView1.Rows[i].Cells[j].Value)));
                }
            MST.Clear();
        }
        private void PrintMSTtoListBox(ListBox lb, List<Edge> MST)
        {
            lb.Items.Clear();
            lb.Items.Add("G(0)");
            int sum = 0;
            for (int i = 0; i < MST.Count; i++)
            {
                lb.Items.Add(string.Format("G({0}) = G({1}) + ({2},{3}): {4};", i + 1, i, MST[i].v1 + 1, MST[i].v2 + 1, MST[i].weight));
                sum += MST[i].weight;
            }
            lb.Items.Add("______________________");
            lb.Items.Add(string.Format("SUMM = {0}", sum));
        }
        public List<string> SpecialForWord(List<Edge> MST)
        {
            List<string> outMST = new List<string>(MST.Count+1);

            int sum = 0;
            for (int i = 0; i < MST.Count; i++)
            {
                outMST.Add(string.Format("G({0}) = G({1}) + ({2},{3}): {4};", i + 1, i, alfavit[MST[i].v1], alfavit[MST[i].v2], MST[i].weight));
                sum += MST[i].weight;
            }
            outMST.Add("______________________");
            outMST.Add(string.Format("SUMM = {0}", sum));

            return outMST;
        }
        private void LoadExample(int countVershin, int minDig, int maxDig)
        {
            Random rand = new Random();

            int size = countVershin; // textBox1.Text == "" ?rand.Next(4,10) : razmer;
            dataGridView1.RowCount = size;
            dataGridView1.ColumnCount = size;
            for (int i = 0; i < dataGridView1.ColumnCount; i++)
                dataGridView1.Columns[i].Width = 40;


            for (int i = 0; i < dataGridView1.RowCount; i++)
                for (int j = 0; j < i + 1; j++)
                    dataGridView1.Rows[i].Cells[j].Value = j == i ? 0 : rand.Next(minDig, maxDig);
            
            for (int j = 0; j < dataGridView1.ColumnCount; j++)
                for (int i = 0; i < dataGridView1.RowCount; i++)
                {
                    dataGridView1.Rows[j].Cells[i].Value = dataGridView1.Rows[i].Cells[j].Value;
                }
        }
        private void button8_Click(object sender, EventArgs e)
        {
            int razmer = dataGridView1.RowCount;
            dataGridView1.RowCount = 0;
            dataGridView1.ColumnCount = 0;
            dataGridView1.RowCount = razmer;
            dataGridView1.ColumnCount = razmer;
            for (int i = 0; i < dataGridView1.ColumnCount; i++)
                dataGridView1.Columns[i].Width = 40;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            button5.Enabled = false;
            button4.Enabled = true;
            button6.Enabled = true;
            textBox5.Enabled = true;
            checkBox1.Enabled = true;
            oWord = new Word.Application();
            oDoc = oWord.Documents.Add(ref oMissing, ref oMissing,
                ref oMissing, ref oMissing);
        }

        private void button6_Click(object sender, EventArgs e)
        {
            string name = string.Empty;
            try
            {
                if (textBox5.Text == "")
                    throw new Exception("Введите название долкумента!");
                if (textBox5.Text != "")
                {
                    if (File.Exists(Environment.CurrentDirectory + string.Format(@"\{0}.doc", textBox5.Text)))
                        throw new Exception("Файл c таким именем существует!");
                    name = textBox5.Text;
                }
                name = Environment.CurrentDirectory + string.Format(@"\{0}.doc", name);

                Object fileName = name;
                Object fileFormat = Word.WdSaveFormat.wdFormatDocument;
                Object lockComments = false;
                Object password = "";
                Object addToRecentFiles = false;
                Object writePassword = "";
                Object readOnlyRecommended = false;
                Object embedTrueTypeFonts = false;
                Object saveNativePictureFormat = false;
                Object saveFormsData = false;
                Object saveAsAOCELetter = Type.Missing;
                Object encoding = Type.Missing;
                Object insertLineBreaks = Type.Missing;
                Object allowSubstitutions = Type.Missing;
                Object lineEnding = Type.Missing;
                Object addBiDiMarks = Type.Missing;

                oDoc.SaveAs(ref fileName,
                ref fileFormat, ref lockComments,
                ref password, ref addToRecentFiles, ref writePassword,
                ref readOnlyRecommended, ref embedTrueTypeFonts,
                ref saveNativePictureFormat, ref saveFormsData,
                ref saveAsAOCELetter, ref encoding, ref insertLineBreaks,
                ref allowSubstitutions, ref lineEnding, ref addBiDiMarks);
                oWord.Quit();
                button5.Enabled = true;
                button4.Enabled = false;
                button6.Enabled = false;
                textBox5.Enabled = false;
                checkBox1.Checked = false;
                checkBox1.Enabled = false;
                oWord = null;
                textBox5.Text = "";

                CurrentGraf = 1;
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            oWord.Visible = checkBox1.Checked ? true : false;
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {

            if (oWord != null) 
			{
                Object saveChanges = Word.WdSaveOptions.wdDoNotSaveChanges;
                Object originalFormat = Word.WdOriginalFormat.wdWordDocument;
                Object routeDocument = Type.Missing;

                oWord.Quit(ref saveChanges,
                             ref originalFormat, ref routeDocument);
			    oWord = null;
			}
        }
    }
}
