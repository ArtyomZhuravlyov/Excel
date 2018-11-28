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
        public interface Prototipe
        { 
            void writtenIPT(int offset); //int offset в прототипе можно не указывать
            void writtenBU(int offset);           
        }
        

        //общее //приложение и книги
        private Excel.Application excelapp;
        private Excel.Workbooks excelappworkbooks;

        //книга , листы и ячейки
        private Excel.Workbook excelappworkbook;
        private Excel.Sheets excelsheets;
        private Excel.Worksheet excelworksheet;
        private Excel.Range excelcells;
        private Excel.Range excelcells2;

        //книга , листы и ячейки
        private Excel.Workbook excelappworkbooknew;
        private Excel.Sheets excelsheetsnew;
        private Excel.Worksheet excelworksheetnew;
        private Excel.Range excelcellsnew;

        //определение диапазона
        private string InceptionRange; //чтение с текстбокс
        private string EndRange;
        private int CountCells; // количество ячеек в этом диапазоне 

        //текущее количество пройденных ячеек
        private int CountCurentCell = 1; //для тройного цикла (переключения ячеек) //1 потому что excelcells2[1, 1] начинается с 1
        private int CountCurentCellback;

        public struct SDatas
        {
            public static double[] time = new double[500];
            public static double[] T_0 = new double[500];
            public static double[] T_3 = new double[500];
            public static int Quantityrow = 0;
        }

        public struct SItemsListBox //если использую static не могу создать объекты IPT;BU;
        {
            public  string[] Items; //= { "", "", "" };
            public  string[] PathFile;//= { "", "", "" };
            public  int QuantityFiles ; //=0;

            public SItemsListBox(string[] Items, string[] PathFile, int QuantityFiles)
            {
                this.Items = Items; // хранит имя файла
                this.PathFile = PathFile; //хранит путь к файлу
                this.QuantityFiles = QuantityFiles; // кол-во файлов
            }

            public void clear()
            {
                Items = new string[3];
                Items[0] = "";
                Items[1] = "";
                Items[2] = "";

                PathFile = new string[3];
                PathFile[0] = "";
                PathFile[1] = "";
                PathFile[2] = "";

                QuantityFiles = 0;
            }
        }

        SItemsListBox BU; //т.к. 2 бокса на файлы
        SItemsListBox IPT;

        //такая запись работает если не создаются объекты этой структуры
        public struct Stime
        {
            public static string[] seconds_05 = {"00", "05", "10" , "15" , "20", "25", "30","35", "40","45", "50", "55" };
            public static string[] minutes = { "0", "01", "02", "03", "04", "05", "06", "07", "08", "09", "10",
                                               "11", "12", "13", "14", "15", "16", "17", "18", "19", "20",
                                               "21", "22", "23", "24", "25", "26", "27", "28", "29", "30",
                                               "31", "32", "33", "34", "35", "36", "37", "38", "39", "40",
                                               "41", "42", "43", "44", "45", "46", "47", "48", "49", "50",
                                               "51", "52", "53", "54", "55", "56", "57", "58", "59"};
            public static string[] hours = { "00", "01", "02" };
        }

        public Form1()
        {
            InitializeComponent();
            BU.clear();
            IPT.clear();
        }

        

        private void button1_Click(object sender, EventArgs e)
        {
             
            int i = Convert.ToInt32(((Button)(sender)).Tag);
            IPT.QuantityFiles = 1; //!! временно
            if (IPT.QuantityFiles == 0 && i == 1 ) return; //i=1 первая кнопка 
                                                           //разобраться почему не сработало String.Empty
            IPT.PathFile[0] = @"C:\Users\Zhuravlev\Desktop\для проги\ИПТ4.xlsx"; //исправить в будущем
            //SItemsListBoxIPT.PathFile[0] = @"C:\Users\Артём\Desktop\для тестов";
            label1.Text = IPT.PathFile[0];
            
            switch (i)
            {
                case 1:
                    excelapp = new Excel.Application();
                    excelapp.Visible = true;
                    //Получаем набор ссылок на объекты Workbook
                    excelappworkbooks = excelapp.Workbooks;
                    //Открываем книгу и получаем на нее ссылку
                    excelappworkbook = excelapp.Workbooks.Open(IPT.PathFile[0],
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

                        
                    //выбираю нужный формат
                    excelcells = excelworksheet.get_Range("C11", "C11");
                    //excelcells.NumberFormat = "ДД.ММ.ГГГГ ч:мм";
                    // label1.Text = Convert.ToString(excelcells.NumberFormat);
                    label1.Text = (excelcells.NumberFormat).ToString();
                    //excelcells.Clear();
                    // excelcells.NumberFormat = "ДД.ММ.ГГГГ ч:мм:cc" ; // "Д ММММ, ГГГГ"
                    //excelcells.Value2 = "19.11.2018  16:54:00";
                    // копируем из одного ряда в другой
                    excelcells2 = excelworksheet.get_Range("O"+"11", "O11"); // вы
                    excelcells2.Clear();
                    //excelcells2.NumberFormat = "ДД.ММ.ГГГГ ч:мм:cc"; // "Д ММММ, ГГГГ"
                    excelcells2.NumberFormat = excelcells.NumberFormat; //забираем формат
                    excelcells2.Value2 = excelcells.Value2; // забираем число

                    /***************************чтение диапазона ячеек и cчёт количества ячеек***************************/
                    InceptionRange =  textBox1.Text; //начало диапазона //большое число//пример 20
                    EndRange = textBox2.Text; //конец диапазона
                    CountCells = Convert.ToInt32(textBox1.Text) - Convert.ToInt32(textBox2.Text) + 1; //количчетсво ячеек диапазона
                    CountCurentCellback = CountCells; // необходимо для переворота ибо значения идёт снизу вверх //обратный счётчик
                    label1.Text = CountCells.ToString();
                    
                    excelcells2 = excelworksheet.get_Range("B" + EndRange, "B" + EndRange);

                    //excelcells2[1, 1].EntireColumn.NumberFormat = "мм:сс";
                    //!!!!! проверить !! CountHour = CountCells / 12 / 60; //12 т.к. в 1 мин 5сек по 12 раз //60 минут в часе

                    //int CountSeconds = CountCells % 12;
                    // excelcells2[1, 1].EntireColumn.NumberFormat = "мм:сс";
                    //      sheet.Cells[1, "A"].Value2 = "Id"; 

                   /*****************************Создание новой пустой книги***************************************************/
                   excelapp.SheetsInNewWorkbook = 1; //1 количество листов в новой книге
                   excelapp.Workbooks.Add(Type.Missing);
                   excelappworkbooknew = excelappworkbooks[2]; ////Получаем ссылку на книгу 1 - нумерация от 1
                   excelsheetsnew = excelappworkbooknew.Worksheets;   //Получаем массив ссылок на листы выбранной книги
                   excelworksheetnew = (Excel.Worksheet)excelsheetsnew.get_Item(1); //получаем ссылку на первый лист
                   excelcellsnew = excelworksheetnew.get_Range("B" + EndRange, "B" + EndRange); //куда будет записываться

                    //что нужно реализовать
                    //1 подсчёт кол-ва файлов ИПТ и БУ отдельно
                    //запись значений вынести как отдельную функцию
                    //с помощью свитча отправлять определённое кол-во раз в эту функцию
                    // менять куда будет записываться в виде аргументов в зависимости от свитча

                    /*****************************Запись времени и значений для ИПТ***************************************************/
                    if (IPT.PathFile[0] !=  String.Empty) writtenIPT(0); //вписать аргументы
                    if (IPT.PathFile[1] != String.Empty) label102.Text = "chf,";
                    
                 
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
        /********************************* выбор файлов для ИПТ **********************************************************/
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
                    IPT.PathFile[IPT.QuantityFiles] = file;
                    IPT.Items[IPT.QuantityFiles] = System.IO.Path.GetFileName(file); //Только имя файла (с расширением):
                                                                                                       // dbf_File = System.IO.Path.GetFileName(dbf_File);
                    label1.Text = IPT.PathFile[IPT.QuantityFiles];
                    listBox1.Items.Add(IPT.Items[IPT.QuantityFiles]);
                    //Только содержащий каталог:
                    //string dbf_Path = System.IO.Path.GetDirectoryName(dbf_File);
                    IPT.QuantityFiles++;
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
                IPT.PathFile[IPT.QuantityFiles] = file;
                IPT.Items[IPT.QuantityFiles] = System.IO.Path.GetFileName(file); //Только имя файла (с расширением):
                                                                                                   // dbf_File = System.IO.Path.GetFileName(dbf_File);
                label1.Text = IPT.PathFile[IPT.QuantityFiles];
                listBox1.Items.Add(IPT.Items[IPT.QuantityFiles]);
                IPT.QuantityFiles++;
            }
        }
        /**************************/
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
        /*****************************Запись времени и значений для ИПТ***************************************************/
        void writtenIPT(int offset) // в аргументе смещение потому что будет записываться столбцы от каждого ИПТ и надо их свдвигать вправо
        {
            for (int ih = 0; ih <= 1; ih++)
            {
                if ((CountCurentCell - 1) == CountCells) break;
                for (int im = 0; im < (60); im++)
                {
                    if ((CountCurentCell - 1) == CountCells) break;
                    for (int isec = 0; isec < (12); isec++)
                    {
                        if ((CountCurentCell - 1) == CountCells) break;
                        excelcellsnew[CountCurentCell, 6 + offset].Value2 = excelcells2[CountCurentCellback, 6].Value2;// Т_3
                        excelcellsnew[CountCurentCell, 5 + offset].Value2 = excelcells2[CountCurentCellback, 5].Value2; //D:D //Т_0

                        excelcellsnew[CountCurentCell, 1 + offset].EntireColumn.NumberFormat = "[$-ru-RU,1] ДД.ММ.ГГГГ ч:мм:сс";
                         excelcellsnew[CountCurentCell, 1 + offset].Value2 = excelcells2[CountCurentCellback, 2].Value2;//Bnew:C //время датчика
                        excelcellsnew[CountCurentCell, 2 + offset].Value2 = Stime.hours[ih] + ":" + Stime.minutes[im] + ":" + Stime.seconds_05[isec];
                        CountCurentCell++;
                        CountCurentCellback--;
                    }
                }

            }
        }






    }



        
    
}



//sheet = wb.Sheets.Add(); 
//      sheet.Name = "TestSheet1"; 
//      sheet.Cells[1, "A"].Value2 = "Id"; 
//      sheet.Cells[1, "B"].Value2 = "Name"; 



////Если бы мы открыли несколько книг, то получили ссылку так
////excelappworkbook=excelappworkbooks[1];
////Получаем массив ссылок на листы выбранной книги
//excelsheets = excelappworkbook.Worksheets;



//                    //Получаем ссылку на лист 1
//                    excelworksheet = (Excel.Worksheet) excelsheets.get_Item(1);
////Выбираем ряд для времени!!!!!!!!!!!
//excelcells = excelworksheet.get_Range("A1", "D1");
//                    excelcells = excelworksheet.get_Range("H11", "H11");
//                    label1.Text = Convert.ToString(excelcells.Value2);  //получаем число из экселя



//                    //Выводим число
//                    excelcells = excelworksheet.get_Range("E11", "H11");
//                    excelcells.Value2 = 10.5;
//                    //выбираю нужный формат
//                    excelcells = excelworksheet.get_Range("C11", "C11");
//                    //excelcells.NumberFormat = "ДД.ММ.ГГГГ ч:мм";
//                    // label1.Text = Convert.ToString(excelcells.NumberFormat);
//                    label1.Text = (excelcells.NumberFormat).ToString();
////excelcells.Clear();
//// excelcells.NumberFormat = "ДД.ММ.ГГГГ ч:мм:cc" ; // "Д ММММ, ГГГГ"
////excelcells.Value2 = "19.11.2018  16:54:00";
//// копируем из одного ряда в другой
//excelcells2 = excelworksheet.get_Range("O"+"11", "O11");
//                    excelcells2.Clear();
//                    //excelcells2.NumberFormat = "ДД.ММ.ГГГГ ч:мм:cc"; // "Д ММММ, ГГГГ"
//                    excelcells2.NumberFormat = excelcells.NumberFormat; //забираем формат
//                    excelcells2.Value2 = excelcells.Value2; // забираем число
                    
//                    /* потом вернуть
//                    //создание 
//                    excelapp.SheetsInNewWorkbook = 1;
//                    excelapp.Workbooks.Add(Type.Missing);
//                    */



//                    ////Выбираем лист 2
//                    //excelworksheet = (Excel.Worksheet)excelsheets.get_Item(2);
//                    ////При выборе одной ячейки можно не указывать вторую границу 
//                    //excelcells = excelworksheet.get_Range("A1", Type.Missing);
//                    ////Выводим значение текстовую строку
//                    //excelcells.Value2 = "Лист 2";
//                    //excelcells.Font.Size = 20;
//                    //excelcells.Font.Italic = true;
//                    //excelcells.Font.Bold = true;
//                    ////Выбираем лист 3
//                    //excelworksheet = (Excel.Worksheet)excelsheets.get_Item(3);
//                    ////Делаем третий лист активным
//                    //excelworksheet.Activate();
//                    ////Вывод в ячейки используя номер строки и столбца Cells[строка, столбец]
//                    //for (m = 1; m < 20; m++)
//                    //{
//                    //    for (n = 1; n < 15; n++)
//                    //    {
//                    //        excelcells = (Excel.Range)excelworksheet.Cells[m, n];
//                    //        //Выводим координаты ячеек
//                    //        excelcells.Value2 = m.ToString() + " " + n.ToString();
//                    //    }
//                    //}