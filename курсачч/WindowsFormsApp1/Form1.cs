using ExcelDataReader;
using System;
using System.IO;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using SD = System.Data;


namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        string fullPath = Application.StartupPath.ToString() + "\\1.xlsx";
        private SD.DataTableCollection TableCollection = null;

        public Form1()
        {
            InitializeComponent();
            form2 = new Form2();
        }
        Form2 form2;
        private void Form1_Load(object sender, EventArgs e)
        {
            label1.Parent = pictureBox1;
            label2.Parent = pictureBox1;
            label3.Parent = pictureBox1;
            label4.Parent = pictureBox1;
            label5.Parent = pictureBox1;
            label6.Parent = pictureBox1;
            OpenExcelFile(fullPath);
        }

        private void OpenExcelFile(string path)
        {
            FileStream stream = File.Open(fullPath, FileMode.Open, FileAccess.Read);
            IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream);
            SD.DataSet db = reader.AsDataSet(new ExcelDataSetConfiguration()
            {
                ConfigureDataTable = (x) => new ExcelDataTableConfiguration()
                {
                    UseHeaderRow = true
                }
            });

            TableCollection = db.Tables;
            toolStripComboBox1.Items.Clear();
            foreach (SD.DataTable table in TableCollection)
            {
                toolStripComboBox1.Items.Add(table.TableName);
            }

            toolStripComboBox1.SelectedIndex = 0;
            stream.Close();
            reader.Close();
        }

        private void toolStripComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            SD.DataTable table = TableCollection[Convert.ToString(toolStripComboBox1.SelectedItem)];
            dataGridView1.DataSource = table;
        }

        string val1, val2, val3;

        private void button1_Click(object sender, EventArgs e)
        {
            string fileName = fullPath; //имя Excel файла  
            Excel.Application xlApp = new Excel.Application();

            xlApp.ScreenUpdating = false;// !!!ускорение кода!!!

            Excel.Workbook xlWb = xlApp.Workbooks.Open(fileName); //открываем Excel файл
            Excel.Worksheet xlSht = xlWb.Sheets[1]; //первый лист в файле
            int iLastRow = xlSht.Cells[xlSht.Rows.Count, "A"].End[Excel.XlDirection.xlUp].Row;  //последняя заполненная строка в столбце А

            //---------------------------------------
            string num = Convert.ToString(textBox1.Text);
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                if (Convert.ToString(dataGridView1.Rows[i].Cells[0].Value) == num)
                {
                    val1 = dataGridView1.Rows[i].Cells[0].Value.ToString();
                    val2 = dataGridView1.Rows[i].Cells[1].Value.ToString();
                    val3 = dataGridView1.Rows[i].Cells[2].Value.ToString();
                    xlSht.Rows[i + 2].Delete();

                    MessageBox.Show(val2.ToString() + " уехал");
                }
            }
            xlSht = xlWb.Sheets[2]; //первый лист в файле
            iLastRow = xlSht.Cells[xlSht.Rows.Count, "A"].End[Excel.XlDirection.xlUp].Row;  //последняя заполненная строка
            iLastRow++;
            xlSht.Cells[iLastRow, "A"].Value = val1; //запись в ячейку
            xlSht.Cells[iLastRow, "B"].Value = val2;
            xlSht.Cells[iLastRow, "C"].Value = val3;

            xlApp.ScreenUpdating = true;// !!!ускорение кода!!!

            //xlApp.Visible = true;
            xlWb.Close(true); //закрыть и сохранить книгу
            xlApp.Quit();
            //MessageBox.Show("Файл успешно сохранён!");
            OpenExcelFile(fullPath);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string fileName = fullPath; //имя Excel файла  
            Excel.Application xlApp = new Excel.Application();

            xlApp.ScreenUpdating = false;// !!!ускорение кода!!!

            Excel.Workbook xlWb = xlApp.Workbooks.Open(fileName); //открываем Excel файл
            Excel.Worksheet xlSht = xlWb.Sheets[2]; //первый лист в файле
            int iLastRow = xlSht.Cells[xlSht.Rows.Count, "A"].End[Excel.XlDirection.xlUp].Row;  //последняя заполненная строка в столбце А

            //---------------------------------------
            string num = Convert.ToString(textBox1.Text);
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                if (Convert.ToString(dataGridView1.Rows[i].Cells[0].Value) == num)
                {
                    val1 = dataGridView1.Rows[i].Cells[0].Value.ToString();
                    val2 = dataGridView1.Rows[i].Cells[1].Value.ToString();
                    val3 = dataGridView1.Rows[i].Cells[2].Value.ToString();
                    xlSht.Rows[i + 2].Delete();

                    MessageBox.Show(val2.ToString() + " приехал");


                }


            }
            xlSht = xlWb.Sheets[1]; //первый лист в файле
            iLastRow = xlSht.Cells[xlSht.Rows.Count, "A"].End[Excel.XlDirection.xlUp].Row;  //последняя заполненная строка
            iLastRow++;
            xlSht.Cells[iLastRow, "A"].Value = val1; //запись в ячейку
            xlSht.Cells[iLastRow, "B"].Value = val2;
            xlSht.Cells[iLastRow, "C"].Value = val3;

            xlApp.ScreenUpdating = true;// !!!ускорение кода!!!

            //xlApp.Visible = true;
            xlWb.Close(true); //закрыть и сохранить книгу
            xlApp.Quit();
            //MessageBox.Show("Файл успешно сохранён!");
            OpenExcelFile(fullPath);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string fileName = fullPath; //имя Excel файла  
            Excel.Application xlApp = new Excel.Application();

            xlApp.ScreenUpdating = false;// !!!ускорение кода!!!

            Excel.Workbook xlWb = xlApp.Workbooks.Open(fileName); //открываем Excel файл
            Excel.Worksheet xlSht = xlWb.Sheets[1]; //первый лист в файле
            int iLastRow = xlSht.Cells[xlSht.Rows.Count, "A"].End[Excel.XlDirection.xlUp].Row;  //последняя заполненная строка в столбце А

            iLastRow++;
            xlSht.Cells[iLastRow, "A"].Value = textBox2.Text.ToString(); //запись в ячейку
            xlSht.Cells[iLastRow, "B"].Value = textBox3.Text.ToString();
            xlSht.Cells[iLastRow, "C"].Value = textBox4.Text.ToString();

            xlApp.ScreenUpdating = true;// !!!ускорение кода!!!

            //xlApp.Visible = true;
            xlWb.Close(true); //закрыть и сохранить книгу
            xlApp.Quit();
            //MessageBox.Show("Файл успешно сохранён!");
            OpenExcelFile(fullPath);
        }

        private void toolStripMenuItem3_Click(object sender, EventArgs e)
        {
            form2.Show();
        }

    }
}
