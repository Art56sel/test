using Microsoft.Office.Interop.Excel;
using MySql.Data.MySqlClient;
using MySqlX.XDevAPI.Relational;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using Excel = Microsoft.Office.Interop.Excel;

namespace TESTVER0._1
{
    public partial class Form1 : Form
    {
        System.Data.DataTable originalDataCitylink = new System.Data.DataTable();
        System.Data.DataTable originalDataOzon = new System.Data.DataTable();
        System.Data.DataTable originalDataYnMarket = new System.Data.DataTable();
        public Form1()
        {
            InitializeComponent();
            #region Настройка передвижения формы
            this.MouseDown += new MouseEventHandler(Form_MouseDown);
            this.MouseMove += new MouseEventHandler(Form_MouseMove);
            this.MouseUp += new MouseEventHandler(Form_MouseUp);
            menuStrip1.MouseDown += new MouseEventHandler(Form_MouseDown);
            menuStrip1.MouseMove += new MouseEventHandler(Form_MouseMove);
            menuStrip1.MouseUp += new MouseEventHandler(Form_MouseUp);
            panel1.MouseDown += new MouseEventHandler(Form_MouseDown);
            panel1.MouseMove += new MouseEventHandler(Form_MouseMove);
            panel1.MouseUp += new MouseEventHandler(Form_MouseUp);

        }
        private bool isDragging = false;
        private System.Drawing.Point lastCursor;
        private System.Drawing.Point lastForm;

        private void Form_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                isDragging = true;
                lastCursor = Cursor.Position;
                lastForm = this.Location;
            }
        }


        private void Form_MouseMove(object sender, MouseEventArgs e)
        {
            if (isDragging)
            {
                int xDiff = Cursor.Position.X - lastCursor.X;
                int yDiff = Cursor.Position.Y - lastCursor.Y;
                this.Location = new System.Drawing.Point(lastForm.X + xDiff, lastForm.Y + yDiff);
            }
        }


        private void Form_MouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                isDragging = false;
            }
        }
        #endregion
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {


        }


        private void Form1_Load(object sender, EventArgs e)
        {
            this.TopMost = false;
            this.WindowState = FormWindowState.Normal;
            #region Код для определения файлов

            string basePath = Path.Combine(Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName, "Resources");
            string[] fileNames = { "Ситилинк.xlsx", "Озон.xlsx", "ЯндексМаркет.xlsx" };

            bool allFilesExist = true;
            foreach (string fileName in fileNames)
            {
                string filePath = Path.Combine(basePath, fileName);
                if (!File.Exists(filePath))
                {
                    allFilesExist = false;
                }
            }

            if (!allFilesExist)
            {
                DialogResult result = MessageBox.Show("Один или несколько файлов отсутствуют. Нажмите 'ОК', чтобы продолжить.", "Внимание", MessageBoxButtons.OK);
                if (result == DialogResult.OK)
                {
                    ExcelHandler eh1 = new ExcelHandler();
                    eh1.CreateAndFillExcel();
                } else
                {
                    System.Windows.Forms.Application.Exit();
                }
            }
            
            dataGridView1.RowHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.EnableHeadersVisualStyles = false;
            dataGridView1.ColumnHeadersDefaultCellStyle.Font = new System.Drawing.Font(dataGridView1.ColumnHeadersDefaultCellStyle.Font.FontFamily, 12f, FontStyle.Bold);
            CitylinkFilling();
           
            dataGridView2.RowHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView2.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView2.EnableHeadersVisualStyles = false;
            dataGridView2.ColumnHeadersDefaultCellStyle.Font = new System.Drawing.Font(dataGridView2.ColumnHeadersDefaultCellStyle.Font.FontFamily, 12f, FontStyle.Bold);
            OzonFilling();
            
            dataGridView3.RowHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView3.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView3.EnableHeadersVisualStyles = false;
            dataGridView3.ColumnHeadersDefaultCellStyle.Font = new System.Drawing.Font(dataGridView3.ColumnHeadersDefaultCellStyle.Font.FontFamily, 12f, FontStyle.Bold);
            YnMarketFilling();
        }
        public void CitylinkFilling() {
            // Ситилинк       
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(Path.Combine(Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName, @"Resources\Ситилинк.xlsx"));
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;
            // Проверяем наличие столбцов в DataGridView
            if (dataGridView1.Columns.Count == 0)
            {
                // Устанавливаем заголовки столбцов из первой строки Excel
                for (int j = 1; j <= colCount; j++)
                {
                    dataGridView1.Columns.Add("Сolumn" + j, xlRange.Cells[1, j].Value2.ToString());
                }
                dataGridView1.Columns[0].Width = 150;
                dataGridView1.Columns[1].Width = 80;
                dataGridView1.Columns[2].Width = 90;
                dataGridView1.Columns[3].Width = 100;
                dataGridView1.Columns[4].Width = 100;
                dataGridView1.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dataGridView1.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dataGridView1.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }
            // Заполняем DataGridView данными из Excel, начиная со второй строки
            for (int i = 2; i <= rowCount; i++)
            {
                dataGridView1.Rows.Add();
                for (int j = 1; j <= colCount; j++)
                {
                    if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                    {
                        dataGridView1.Rows[i - 2].Cells[j - 1].Value = xlRange.Cells[i, j].Value2.ToString();
                    }
                }
            }
           
            xlWorkbook.Close();
            xlApp.Quit();
        }
        public void OzonFilling()
        {
            //Озон
            Excel.Application xlApp2 = new Excel.Application();
            Excel.Workbook xlWorkbook2 = xlApp2.Workbooks.Open(Path.Combine(Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName, @"Resources\Озон.xlsx"));
            Excel._Worksheet xlWorksheet2 = xlWorkbook2.Sheets[1];
            Excel.Range xlRange2 = xlWorksheet2.UsedRange;

            int rowCount2 = xlRange2.Rows.Count;
            int colCount2 = xlRange2.Columns.Count;
            // Проверяем наличие столбцов в DataGridView
            if (dataGridView2.Columns.Count == 0)
            {
                // Устанавливаем заголовки столбцов из первой строки Excel
                for (int j = 1; j <= colCount2; j++)
                {
                    dataGridView2.Columns.Add("Column" + j, xlRange2.Cells[1, j].Value2.ToString());
                }
                dataGridView2.Columns[0].Width = 150;
                dataGridView2.Columns[1].Width = 80;
                dataGridView2.Columns[2].Width = 90;
                dataGridView2.Columns[3].Width = 100;
                dataGridView2.Columns[4].Width = 100;
                dataGridView2.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dataGridView2.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dataGridView2.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }
            // Заполняем DataGridView данными из Excel, начиная со второй строки
            for (int i = 2; i <= rowCount2; i++)
            {
                dataGridView2.Rows.Add();
                for (int j = 1; j <= colCount2; j++)
                {
                    if (xlRange2.Cells[i, j] != null && xlRange2.Cells[i, j].Value2 != null)
                    {
                        dataGridView2.Rows[i - 2].Cells[j - 1].Value = xlRange2.Cells[i, j].Value2.ToString();
                    }
                }
            }

            xlWorkbook2.Close();
            xlApp2.Quit();
            
        }

        public void YnMarketFilling()
        {
            //Яндекс маркет
            Excel.Application xlApp3 = new Excel.Application();
            Excel.Workbook xlWorkbook3 = xlApp3.Workbooks.Open(Path.Combine(Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName, @"Resources\ЯндексМаркет.xlsx"));
            Excel._Worksheet xlWorksheet3 = xlWorkbook3.Sheets[1];
            Excel.Range xlRange3 = xlWorksheet3.UsedRange;

            int rowCount3 = xlRange3.Rows.Count;
            int colCount3 = xlRange3.Columns.Count;
            // Проверяем наличие столбцов в DataGridView
            if (dataGridView3.Columns.Count == 0)
            {
                for (int j = 1; j <= colCount3; j++)
                {
                    dataGridView3.Columns.Add("Column" + j, xlRange3.Cells[1, j].Value2.ToString());
                }
            }
            // Заполняем DataGridView данными из Excel, начиная со второй строки
            for (int i = 2; i <= rowCount3; i++)
            {
                dataGridView3.Rows.Add();
                for (int j = 1; j <= colCount3; j++)
                {
                    if (xlRange3.Cells[i, j] != null && xlRange3.Cells[i, j].Value2 != null)
                    {
                        dataGridView3.Rows[i - 2].Cells[j - 1].Value = xlRange3.Cells[i, j].Value2.ToString();
                    }
                }
                dataGridView3.Columns[0].Width = 150;
                dataGridView3.Columns[1].Width = 80;
                dataGridView3.Columns[2].Width = 90;
                dataGridView3.Columns[3].Width = 100;
                dataGridView3.Columns[4].Width = 100;
                dataGridView3.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dataGridView3.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dataGridView3.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }
            xlWorkbook3.Close();
            xlApp3.Quit();
            #endregion

        }
    
        private void label1_Click(object sender, EventArgs e)
        {

        }
        public float a, b, c, sred;
        private void button1_Click(object sender, EventArgs e)
        {
            
            string searchItem = textBox1.Text;// Получаем текст из поисковой строки
            try
            {


                // Проходим по каждой таблице и ищем схожий товар
                ContrastEx ex = new ContrastEx();
                    ex.FilterAndDisplaySimilarProducts(searchItem, dataGridView1,dataGridView2,dataGridView3);
               button3.Enabled = true;

                


            } 
            catch (Exception ex) {
                MessageBox.Show($"{ex.Message}", "Внимание", MessageBoxButtons.OK);
            }
        }
        private void pictureBox1_Click(object sender, EventArgs e)
        {
            if (DialogResult.Yes == MessageBox.Show("Вы уверены, что хотите выйти из приложения?", "Выход", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation))
            {
                System.Windows.Forms.Application.Exit();
            }
            else
            {
                return;
            }
        }


        private void спавкаToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void toolStripMenuItem2_Click(object sender, EventArgs e)
        {

        }

        private void ситилинкToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CitylinkFilling();
        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
            dataGridView2.Rows.Clear();
            dataGridView3.Rows.Clear();
            CitylinkFilling();
            OzonFilling();
            YnMarketFilling();
            textBox1.Clear();
            button3.Enabled = false;
        }

        private void озонToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OzonFilling();
        }

        private void яндексМаркетToolStripMenuItem_Click(object sender, EventArgs e)
        {
            YnMarketFilling();
        }

       

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                // Получаем индекс строки, на которую был сделан клик
                int rowIndex = e.RowIndex;

                // Получаем значение ячейки, содержащей цену, из выбранной строки
                string price = dataGridView1.Rows[rowIndex].Cells[4].Value.ToString().Trim();

                // Записываем цену в текстовое поле
                textBox2.Text = price;
            }
        }

        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                // Получаем индекс строки, на которую был сделан клик
                int rowIndex = e.RowIndex;

                // Получаем значение ячейки, содержащей цену, из выбранной строки
                string price = dataGridView2.Rows[rowIndex].Cells[4].Value.ToString().Trim().Replace(" ", "");

                // Записываем цену в текстовое поле
                textBox3.Text = price;
               
            }
        }

        private void dataGridView3_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (e.RowIndex >= 0)
                {
                    // Получаем индекс строки, на которую был сделан клик
                    int rowIndex = e.RowIndex;

                    // Получаем значение ячейки, содержащей цену, из выбранной строки
                    string price = dataGridView3.Rows[rowIndex].Cells[4].Value.ToString().Trim().Replace(" ", "");

                    // Записываем цену в текстовое поле
                    textBox4.Text = price;
                }
            }catch (Exception ex)
            {

            }
            
        }

        private void button4_Click(object sender, EventArgs e)
        {
            float a, c, sred;

            if (!float.TryParse(textBox2.Text, out a))
            {
                MessageBox.Show("Ошибка ввода для значения a. Пожалуйста, введите корректное число.");
                return;
            }

            float b;
            string inputB = new string(textBox3.Text.Where(c => char.IsDigit(c) || c == '.').ToArray());

            if (string.IsNullOrWhiteSpace(inputB) || !float.TryParse(inputB, out b))
            {
                MessageBox.Show("Ошибка ввода для значения b. Пожалуйста, введите корректное число.");
                return;
            }

            string inputC = new string(textBox3.Text.Where(c => char.IsDigit(c) || c == '.').ToArray());

            if (string.IsNullOrWhiteSpace(inputB) || !float.TryParse(inputB, out c))
            {
            
                MessageBox.Show("Ошибка ввода для значения c. Пожалуйста, введите корректное число.");
                return;
            }
            
                sred = (a + b + c) / 3;
                textBox5.Text = sred.ToString();
            button5.Enabled = true;
            
        }

        private void button5_Click(object sender, EventArgs e)
        {
            textBox2.Clear();
            textBox3.Clear();
            textBox4.Clear();
            textBox5.Clear();

            button5.Enabled = false;
        }

        private void оПрограммеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AboutBox1 AB = new AboutBox1();
            AB.Show();
        }

        private void обАвтореToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        private void Exit_Click(object sender, EventArgs e)
        {
            if (DialogResult.Yes == MessageBox.Show("Вы уверены, что хотите выйти из приложения?", "Выход", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation))
            {
                System.Windows.Forms.Application.Exit();
            }
            else
            {
                return;
            }
        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }
    }
}
