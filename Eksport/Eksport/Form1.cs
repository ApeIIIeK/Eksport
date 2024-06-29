using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms.DataVisualization.Charting;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using _Application = Microsoft.Office.Interop.Word._Application;

namespace Eksport
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            string[] lines = File.ReadAllLines("C:\\Users\\Leonid Fon Bom\\Desktop\\Долги\\Eksport\\Eksport\\bin\\Debug\\1.txt");
            foreach (string line in lines)
            {
                dataGridView1.Rows.Add(line); // Добавление ФИО студента в таблицу
            }
            // Добавление строки для среднего балла
            int index = dataGridView1.Rows.Add();
            dataGridView1.Rows[index].Cells[0].Value = "Средний балл";

            Random rand = new Random();
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (row.IsNewRow || row.Cells[0].Value.ToString() == "Средний балл") continue;
                for (int i = 1; i < dataGridView1.Columns.Count; i++)
                {
                    int mark = rand.Next(2, 6); // Оценки от 2 до 5
                    row.Cells[i].Value = mark;
                    row.Cells[i].Style.ForeColor = mark == 2 ? Color.Red : Color.Blue;
                }
            }
            
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string subjectName = comboBox1.SelectedItem.ToString();
            if (!dataGridView1.Columns.Contains(subjectName))
            {
                dataGridView1.Columns.Add(subjectName, subjectName);
            }
        }
        private void CalculateAverage()
        {
            for (int i = 1; i < dataGridView1.Columns.Count; i++)
            {
                double sum = 0;
                int count = 0;
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    if (row.IsNewRow || row.Cells[0].Value.ToString() == "Средний балл") continue;
                    if (row.Cells[i].Value != null)
                    {
                        sum += Convert.ToDouble(row.Cells[i].Value);
                        count++;
                    }
                }
                double average = count == 0 ? 0 : sum / count;
                dataGridView1.Rows[dataGridView1.Rows.Count - 1].Cells[i].Value = average;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            CalculateAverage();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            chart1.Series.Clear(); // Очистите текущие серии
            chart1.ChartAreas.Clear(); // Очистите текущие области диаграммы

            ChartArea chartArea1 = new ChartArea();
            chart1.ChartAreas.Add(chartArea1);

            Series series1 = new Series
            {
                Name = "Series1",
                IsVisibleInLegend = true,
                ChartType = SeriesChartType.Pie
            };
            chart1.Series.Add(series1);

            // Добавление данных в диаграмму
            for (int i = 1; i < dataGridView1.Columns.Count; i++)
            {
                double average = Convert.ToDouble(dataGridView1.Rows[dataGridView1.Rows.Count - 1].Cells[i].Value);
                string subjectName = dataGridView1.Columns[i].HeaderText;
                series1.Points.AddXY(subjectName, average);
            }

            chart1.Invalidate(); // Обновите диаграмму

        }

        private void button4_Click(object sender, EventArgs e)
        {
            chart1.Series.Clear(); // Очистите текущие серии
            chart1.ChartAreas.Clear(); // Очистите текущие области диаграммы

            ChartArea chartArea2 = new ChartArea();
            chart1.ChartAreas.Add(chartArea2);

            Series series2 = new Series
            {
                Name = "Series2",
                IsVisibleInLegend = true,
                ChartType = SeriesChartType.Column
            };
            chart1.Series.Add(series2);

            // Добавление данных в диаграмму
            for (int i = 1; i < dataGridView1.Columns.Count; i++)
            {
                double average = Convert.ToDouble(dataGridView1.Rows[dataGridView1.Rows.Count - 1].Cells[i].Value);
                string subjectName = dataGridView1.Columns[i].HeaderText;
                series2.Points.AddXY(subjectName, average);
            }

            chart1.Invalidate(); // Обновите диаграмму

        }

        private void button6_Click(object sender, EventArgs e)
        {

            // Создаем новое приложение Word
            Word.Application wordApp = new Word.Application();
            // Создаем новый документ
            Word.Document document = wordApp.Documents.Add();

            // Получаем количество строк и столбцов в dataGridView1
            int rowsCount = dataGridView1.Rows.Count;
            int columnsCount = dataGridView1.Columns.Count;

            // Создаем таблицу в документе Word с соответствующим количеством строк и столбцов
            Word.Table table = document.Tables.Add(document.Range(), rowsCount + 1, columnsCount);

            // Заполняем заголовки таблицы
            for (int i = 0; i < columnsCount; i++)
            {
                table.Cell(1, i + 1).Range.Text = dataGridView1.Columns[i].HeaderText;
            }

            // Заполняем ячейки таблицы данными из dataGridView1
            for (int i = 0; i < rowsCount; i++)
            {
                for (int j = 0; j < columnsCount; j++)
                {
                    table.Cell(i + 2, j + 1).Range.Text = dataGridView1.Rows[i].Cells[j].Value?.ToString() ?? "";
                }
            }

            // Делаем приложение Word видимым
            wordApp.Visible = true;
        }

        private void button7_Click(object sender, EventArgs e)
        {
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook workbook = excelApp.Workbooks.Add(Type.Missing);
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Sheets["Лист1"];

            for (int i = 1; i <= dataGridView1.Columns.Count; i++)
            {
                worksheet.Cells[1, i] = dataGridView1.Columns[i - 1].HeaderText;
            }

            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                for (int j = 0; j < dataGridView1.Columns.Count; j++)
                {
                    worksheet.Cells[i + 2, j + 1] = dataGridView1.Rows[i].Cells[j].Value?.ToString() ?? "";
                }
            }

            excelApp.Visible = true;
        }
    }
}
