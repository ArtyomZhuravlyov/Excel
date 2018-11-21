using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace excel
{
    public partial class Form1 : Form
    {
        private Excel.Application excelapp;
        private Excel.Workbooks excelappworkbooks;
        private Excel.Workbook excelappworkbook;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            int i = Convert.ToInt32(((Button)(sender)).Tag);
            switch (i)
            {
                case 1:
                    // Создаём экземпляр нашего приложения
                    Excel.Application excelApp = new Excel.Application();
                    // Создаём экземпляр рабочей книги Excel
                    Excel.Workbook workBook;
                    // Создаём экземпляр листа Excel
                    Excel.Worksheet workSheet;

                    workBook = excelApp.Workbooks.Add();
                    workSheet = (Excel.Worksheet)workBook.Worksheets.get_Item(1); //выбор листа, в котором будем совершать действия

                    // Заполняем первую строку числами от 1 до 10
                    for (int j = 1; j <= 10; j++)
                    {
                        workSheet.Cells[1, j] = j;
                    }
                    // Вычисляем сумму этих чисел
                    Excel.Range rng = workSheet.Range["A2"];
                    rng.Formula = "=SUM(A1:L1)";


                    // rng.FormulaHidden = false; 
                    rng = workSheet.Range["A3"];
                    rng.Value = 5.3;
                    //rng.Value2 = 6;
                    rng = workSheet.Range["B3"];
                    rng.Value2 = 7.1;
                    workSheet.Cells[3, 3] = 4.5;
                    rng.FormulaHidden = false; //непонятно, работает только когда лист защищён
                    
                    // Выделяем границы у этой ячейки
                    Excel.Borders border = rng.Borders;
                    border.LineStyle = Excel.XlLineStyle.xlContinuous;
                    //border.LineStyle = Excel.XlLineStyle.xlDash;
                    // Строим круговую диаграмму
                    Excel.ChartObjects chartObjs = (Excel.ChartObjects)workSheet.ChartObjects();
                    Excel.ChartObject chartObj = chartObjs.Add(5, 50, 300, 300); //expression. Add( _Left_ , _Top_ , _Width_ , _Height_ )
                    Excel.Chart xlChart = chartObj.Chart;
                    Excel.Range rng2 = workSheet.Range["A1:L1"]; //expression. Range( _Cell1_ , _Cell2_ )
                    // Устанавливаем тип диаграммы
                    xlChart.ChartType = Excel.XlChartType.xlPie;
                    // Устанавливаем источник данных (значения от 1 до 10)
                    xlChart.SetSourceData(rng2);
                    
                    rng = workSheet.Range["F3"];
                    rng.FormulaHidden = false; //непонятно, работает только когда лист защищён

                    /****************вторая диаграмма*******************/

                    // Выделяем границы у этой ячейки
                    Excel.Borders border2 = rng.Borders;
                    border.LineStyle = Excel.XlLineStyle.xlContinuous;
                    //border.LineStyle = Excel.XlLineStyle.xlDash;
                    // Строим круговую диаграмму
                    Excel.ChartObjects chartObjs2 = (Excel.ChartObjects)workSheet.ChartObjects();
                    Excel.ChartObject chartObj2 = chartObjs2.Add(500, 50, 300, 300); //expression. Add( _Left_ , _Top_ , _Width_ , _Height_ )
                    Excel.Chart xlChart2 = chartObj2.Chart;
                    rng2 = workSheet.Range["A1:L1"]; //expression. Range( _Cell1_ , _Cell2_ )
                    // Устанавливаем тип диаграммы
                    xlChart2.ChartType = Excel.XlChartType.xlPie;
                    // Устанавливаем источник данных (значения от 1 до 10)
                    xlChart2.SetSourceData(rng2);

                    // Открываем созданный excel-файл
                    excelApp.Visible = true;
                    excelApp.UserControl = true;
                    //excelApp.UserControl = true;
                    
                    break;
                case 2:

                    try
                    {
                        excelapp.Quit();
                    }
                    catch
                    {

                    }
                    break;
                default:
                    Close();
                    break;
            }
        }

        private void Form1_KeyDown(object sender, KeyEventArgs e)
        {
            
            if (e.KeyData == Keys.A && this.FormBorderStyle == FormBorderStyle.FixedSingle)
            {
                label1.Text = "555";
                this.FormBorderStyle = FormBorderStyle.None;

            }
            else if(e.KeyData == Keys.A && this.FormBorderStyle == FormBorderStyle.None)
            { 
                this.FormBorderStyle = FormBorderStyle.FixedSingle;
                label1.Text = "111";
            }
        }
    }
}
