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
        private Excel.Sheets excelsheets;
        private Excel.Worksheet excelworksheet;
        private Excel.Range excelcells;
        private Excel.Range excelcells2;

        public struct SDatas
        {
            public static double[] time = new double[500];
            public static double[] T_0 = new double[500];
            public static double[] T_3 = new double[500];
            public static int Quantityrow = 0;
        }

        public struct SItemsListBox
        {
            public static string[] Items = new string[6];
            public static string[] PathFile = new string[6];
            public static int QuantityFiles = 0;
        }
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            int m, n;
            int i = Convert.ToInt32(((Button)(sender)).Tag);
            SItemsListBox.QuantityFiles = 1; //!! временно
            if (SItemsListBox.QuantityFiles == 0 && i == 1 ) return; //i=1 первая кнопка 
                                                                     //разобраться почему не сработало String.Empty
            SItemsListBox.PathFile[0] = @"C:\Users\Zhuravlev\Desktop\для проги\ИПТ4.xlsx"; //исправить в будущем
            label1.Text = SItemsListBox.PathFile[0];
            switch (i)
            {
                case 1:
                    excelapp = new Excel.Application();
                    excelapp.Visible = true;
                    //Получаем набор ссылок на объекты Workbook
                    excelappworkbooks = excelapp.Workbooks;
                    //Открываем книгу и получаем на нее ссылку
                    excelappworkbook = excelapp.Workbooks.Open(SItemsListBox.PathFile[0],
                                       Type.Missing, Type.Missing, Type.Missing,
                     "WWWWW", "WWWWW", Type.Missing, Type.Missing, Type.Missing,
                      Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                      Type.Missing, Type.Missing);
                    //Если бы мы открыли несколько книг, то получили ссылку так
                    //excelappworkbook=excelappworkbooks[1];
                    //Получаем массив ссылок на листы выбранной книги
                    excelsheets = excelappworkbook.Worksheets;
                    //Получаем ссылку на лист 1
                    excelworksheet = (Excel.Worksheet)excelsheets.get_Item(1);
                             //Выбираем ряд для времени!!!!!!!!!!!
                    excelcells = excelworksheet.get_Range("A1", "D1");
                    excelcells = excelworksheet.get_Range("H11", "H11");
                    label1.Text = Convert.ToString(excelcells.Value2);  //получаем число из экселя
                    //Выводим число
                    excelcells = excelworksheet.get_Range("E11", "H11");
                    excelcells.Value2 = 10.5;
                    //выбираю нужный формат
                    excelcells = excelworksheet.get_Range("C11", "C11");
                    //excelcells.NumberFormat = "ДД.ММ.ГГГГ ч:мм";
                    label1.Text = Convert.ToString(excelcells.NumberFormat);
                    //excelcells.Clear();
                    // excelcells.NumberFormat = "ДД.ММ.ГГГГ ч:мм:cc" ; // "Д ММММ, ГГГГ"
                    //excelcells.Value2 = "19.11.2018  16:54:00";
                    // копируем из одного ряда в другой
                    excelcells2 = excelworksheet.get_Range("O"+"11", "O11");
                    excelcells2.Clear();
                    //excelcells2.NumberFormat = "ДД.ММ.ГГГГ ч:мм:cc"; // "Д ММММ, ГГГГ"
                    excelcells2.NumberFormat = excelcells.NumberFormat; //забираем формат
                    excelcells2.Value2 = excelcells.Value2; // забираем число
                    
                    /* потом вернуть
                    //создание 
                    excelapp.SheetsInNewWorkbook = 1;
                    excelapp.Workbooks.Add(Type.Missing);
                    */
                    ////Выбираем лист 2
                    //excelworksheet = (Excel.Worksheet)excelsheets.get_Item(2);
                    ////При выборе одной ячейки можно не указывать вторую границу 
                    //excelcells = excelworksheet.get_Range("A1", Type.Missing);
                    ////Выводим значение текстовую строку
                    //excelcells.Value2 = "Лист 2";
                    //excelcells.Font.Size = 20;
                    //excelcells.Font.Italic = true;
                    //excelcells.Font.Bold = true;
                    ////Выбираем лист 3
                    //excelworksheet = (Excel.Worksheet)excelsheets.get_Item(3);
                    ////Делаем третий лист активным
                    //excelworksheet.Activate();
                    ////Вывод в ячейки используя номер строки и столбца Cells[строка, столбец]
                    //for (m = 1; m < 20; m++)
                    //{
                    //    for (n = 1; n < 15; n++)
                    //    {
                    //        excelcells = (Excel.Range)excelworksheet.Cells[m, n];
                    //        //Выводим координаты ячеек
                    //        excelcells.Value2 = m.ToString() + " " + n.ToString();
                    //    }
                    //}
                    break;
                case 2:
                   // excelapp.Quit();
                    break;
                default:
                    try
                    {
                        excelapp.Quit();
                        Close();
                    }
                    catch
                    {
                        Close();
                    }
                    
                    break;
            }
                
        }

        private void listBox1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop, false) == true)
            {
                e.Effect = DragDropEffects.All;
            }

        }

        private void listBox1_DragDrop(object sender, DragEventArgs e)
        {
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);

            foreach (string file in files)
            {

                //label1.Text = System.IO.Path.GetExtension(file);
                if (System.IO.Path.GetExtension(file) == ".xls" || System.IO.Path.GetExtension(file) == ".xlsx" || System.IO.Path.GetExtension(file) == ".csv")
                {
                    SItemsListBox.PathFile[SItemsListBox.QuantityFiles] = file;
                    SItemsListBox.Items[SItemsListBox.QuantityFiles] = System.IO.Path.GetFileName(file); //Только имя файла (с расширением):
                                                                                                       // dbf_File = System.IO.Path.GetFileName(dbf_File);
                    label1.Text = SItemsListBox.PathFile[SItemsListBox.QuantityFiles];
                    listBox1.Items.Add(SItemsListBox.Items[SItemsListBox.QuantityFiles]);
                    //Только содержащий каталог:
                    //string dbf_Path = System.IO.Path.GetDirectoryName(dbf_File);
                    SItemsListBox.QuantityFiles++;
                }
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog
            {
                Multiselect = true,
                Title = "Выберите файлы",
                InitialDirectory = @"C:\"
            };
            dlg.ShowDialog();
            // пользователь вышел из диалога ничего не выбрав
            if (dlg.FileName == String.Empty)
                return;
            foreach (string file in dlg.FileNames)
            {
                SItemsListBox.PathFile[SItemsListBox.QuantityFiles] = file;
                SItemsListBox.Items[SItemsListBox.QuantityFiles] = System.IO.Path.GetFileName(file); //Только имя файла (с расширением):
                                                                                                   // dbf_File = System.IO.Path.GetFileName(dbf_File);
                label1.Text = SItemsListBox.PathFile[SItemsListBox.QuantityFiles];
                listBox1.Items.Add(SItemsListBox.Items[SItemsListBox.QuantityFiles]);
                SItemsListBox.QuantityFiles++;
            }
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            try
            { 
            excelapp.Quit();
            }
            catch
            {

            }
        }
    }
}
