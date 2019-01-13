using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection; //необходимо было для использования Missing.Value
 


namespace excel
{
    public partial class Form1 : Form
    {
        //public interface Prototipe
        //{ 
        //    void writtenIPT(int offset); //int offset в прототипе можно не указывать
        //    void writtenBU(int offset);           
        //}
        
        
        enum Unit  // для появления нужных label
        {
            BU,
            IPT
        }
        
        //общее //приложение и книги
        private Excel.Application excelapp;
        private Excel.Workbooks excelappworkbooks;

        
        Excel.Chart oChart;
        Excel.Chart oChart2;

        //книга , листы и ячейки
        private Excel.Workbook excelappworkbook;
        private Excel.Sheets excelsheets;
        private Excel.Worksheet excelworksheet;
        private Excel.Range excelcells;
        private Excel.Range excelcells2;

        //книга , листы и ячейки
        private Excel.Workbook excelappworkbooknew;
        Excel.Series oSeries;
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
        int CountMax = 0;

        public struct SDatas
        {
            public static double[] time = new double[500];
            public static double[] T_0 = new double[500];
            public static double[] T_3 = new double[500];
            public static int Quantityrow = 0;
        }

         struct SItemsListBox //если использую static не могу создать объекты IPT;BU;
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
         struct Stime
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

        Stack<int> numbers = new Stack<int>(); //добавялем в стек количество ячеек (чтобы выделить диапазон для графиков)
        public int[] CountCellGeneral = new int[6]; //ИСПОЛЬЗУЕТСЯ ДЛЯ ОПРЕДЕЛЕНИЯ дипазаона ячеек всех файлов

        private void button1_Click(object sender, EventArgs e)
        {
             
            int i = Convert.ToInt32(((Button)(sender)).Tag);
           // IPT.QuantityFiles = 1; //!! временно
            if (IPT.QuantityFiles == 0 && i == 1 && BU.QuantityFiles == 0) return; //i=1 первая кнопка 
                                                           //разобраться почему не сработало String.Empty
            //IPT.PathFile[0] = @"C:\Users\Zhuravlev\Desktop\для проги\ИПТ4.xlsx"; //исправить в будущем
          // IPT.PathFile[0] = @"C:\Users\Артём\Desktop\для тестов";
           
            int offset = 0; //смещение для записи столбцов между ИПТ и БУ // 
            switch (i)
            {
                case 1:
                    label1.Text = "Загружается";
                    excelapp = new Excel.Application();
                    
                    //Получаем набор ссылок на объекты Workbook
                    excelappworkbooks = excelapp.Workbooks;

                    /*****************************Создание новой пустой книги, в которую осуществляется запись***************************************************/
                    excelapp.SheetsInNewWorkbook = 1; //1 количество листов в новой книге
                    excelapp.Workbooks.Add(Type.Missing);
                    excelappworkbooknew = excelappworkbooks[1]; ////Получаем ссылку на книгу 1 - нумерация от 1
                    excelsheetsnew = excelappworkbooknew.Worksheets;   //Получаем массив ссылок на листы выбранной книги
                    excelworksheetnew = (Excel.Worksheet)excelsheetsnew.get_Item(1); //получаем ссылку на первый лист                    
                    excelcellsnew = excelworksheetnew.get_Range("B" + "12", "B" + "12"); //куда будет записываться

                    /*****************************Открытие файлов, чтение и запись значений для ИПТ***************************************************/
                    
                    for (int j=0; j<IPT.QuantityFiles; j++)
                    {
                        OpenFiles(IPT.PathFile[j], j , Unit.IPT); //открытие файлов и чтение значений
                        writtenIPT(offset,IPT.Items[j]);
                        offset += 5;
                        excelappworkbook.Close(); //закрытие книги с которой уже считали данные, чтобы открыть новую использовать один метод
                       if(j==1) label1.Text = "Осталось совсем немного";
                       
                    }
                    /*****************************Открытие файлов, чтение и запись значений для БУ ***************************************************/
                    for (int j = 0; j < BU.QuantityFiles; j++)
                    {
                        OpenFiles(BU.PathFile[j], j , Unit.BU); //открытие файлов и чтение значений
                        writtenBU(offset, BU.Items[j]);
                        offset += 5;
                        excelappworkbook.Close(); //закрытие книги с которой уже считали данные, чтобы открыть новую и использовать один метод
                        if (j == 1) label1.Text = "Осталось ещё чуть-чуть";
                    }
                    writtenBigTime(); //запиисываем самый ьольшой ряд времени для ОСи х
                    BuildingCharts(); //строим графики

                    excelapp.Visible = true;
                    label1.Text = "";

                    break;
                case 2:
                    //применение одного диапазона
                    textBox1.Text = textBox13.Text;
                    textBox3.Text = textBox13.Text;
                    textBox5.Text = textBox13.Text;
                    textBox7.Text = textBox13.Text;
                    textBox9.Text = textBox13.Text;
                    textBox11.Text = textBox13.Text;

                    textBox2.Text = textBox14.Text;
                    textBox4.Text = textBox14.Text;
                    textBox6.Text = textBox14.Text;
                    textBox8.Text = textBox14.Text;
                    textBox10.Text = textBox14.Text;
                    textBox12.Text = textBox14.Text;

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

        //открытие файлов и чтение значений
        void OpenFiles(string pathfile,int Qfile , Unit unit1)
        {
           
        //Открываем книгу и получаем на нее ссылку
        excelappworkbook = excelapp.Workbooks.Open(pathfile,
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
            //excelcells.Clear();
            ////excelcells2.NumberFormat = "ДД.ММ.ГГГГ ч:мм:cc"; // "Д ММММ, ГГГГ"
            //excelcells2.NumberFormat = excelcells.NumberFormat; //забираем формат

            /***************************чтение диапазона ячеек и cчёт количества ячеек (с текстбокса)***************************/
            if (unit1 == Unit.IPT)
            {
                switch (Qfile)
                {
                    case 0: InceptionRange = textBox1.Text; EndRange = textBox2.Text ; break;
                    case 1: InceptionRange = textBox3.Text; EndRange = textBox4.Text; break; 
                    case 2: InceptionRange = textBox5.Text; EndRange = textBox6.Text; break;
                    default: return;
                }
            }
            else
            {
                switch (Qfile)
                {
                    case 0: InceptionRange = textBox7.Text; EndRange = textBox8.Text; break;
                    case 1: InceptionRange = textBox9.Text; EndRange = textBox10.Text; break;
                    case 2: InceptionRange = textBox11.Text; EndRange = textBox12.Text; break;
                    default: return;
                }
            }
            

            //InceptionRange = textBox1.Text; //начало диапазона //большое число//пример 20
            //EndRange = textBox2.Text; //конец диапазона
            CountCells = Convert.ToInt32(InceptionRange) - Convert.ToInt32(EndRange) + 1; //количчетсво ячеек диапазона
            if (CountMax < CountCells) CountMax = CountCells; //для построения самого большого ряда времени
            CountCurentCellback = CountCells; // необходимо для переворота ибо значения идёт снизу вверх //обратный счётчик           
            CountCurentCell = 1; //для тройного цикла (переключения ячеек) //1 потому что excelcells2[1, 1] начинается с 1
            excelcells2 = excelworksheet.get_Range("B" + EndRange, "B" + EndRange); //откуда смотрим ячейки

            if (unit1 == Unit.BU) Qfile += 3;    

            CountCellGeneral[Qfile] = CountCells;
            numbers.Push(CountCells); // в стеке 

        }




        /********************************* выбор файлов для ИПТ и БУ **********************************************************/
        private void listBox1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop, false) == true)
            {
                e.Effect = DragDropEffects.All;
            }

        }

        private void listBox1_DragDrop(object sender, DragEventArgs e)
        {
            Unit unit;
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
              int i = Convert.ToInt32(((ListBox)(sender)).Tag); //для того чтобы не писать 2 одинаковые функции для лист бокс 1 и листбокс2
              
            foreach (string file in files)
            {

                //label1.Text = System.IO.Path.GetExtension(file);
                if (System.IO.Path.GetExtension(file) == ".xls" || System.IO.Path.GetExtension(file) == ".xlsx" || System.IO.Path.GetExtension(file) == ".csv")
                {
                    if (i == 1)
                    {    if (IPT.QuantityFiles == 3) return; // защита от перетаскивания 4 файлов
                        IPT.PathFile[IPT.QuantityFiles] = file;
                        IPT.Items[IPT.QuantityFiles] = System.IO.Path.GetFileName(file); //Только имя файла (с расширением):
                                                                                         // dbf_File = System.IO.Path.GetFileName(dbf_File);
                                                                                         //label1.Text = IPT.PathFile[IPT.QuantityFiles];
                        listBox1.Items.Add(IPT.Items[IPT.QuantityFiles]);
                        //Только содержащий каталог:
                        //string dbf_Path = System.IO.Path.GetDirectoryName(dbf_File);
                        unit = Unit.IPT;
                        Appear_label_drag(IPT.QuantityFiles, file, unit); //чтобы появлялись лейблы 
                        IPT.QuantityFiles++;       
                    }
                    else
                    {
                        if (BU.QuantityFiles == 3) return; // защита от перетаскивания 4 файлов
                        BU.PathFile[BU.QuantityFiles] = file;
                        BU.Items[BU.QuantityFiles] = System.IO.Path.GetFileName(file); //Только имя файла (с расширением):
                                                                                         // dbf_File = System.IO.Path.GetFileName(dbf_File);
                                                                                         //label1.Text = IPT.PathFile[IPT.QuantityFiles];
                        listBox2.Items.Add(BU.Items[BU.QuantityFiles]);
                        //Только содержащий каталог:
                        //string dbf_Path = System.IO.Path.GetDirectoryName(dbf_File);
                        unit = Unit.BU;
                        Appear_label_drag(BU.QuantityFiles, file, unit); //чтобы появлялись лейблы 
                        BU.QuantityFiles++;
                       
                    }
                    //появление кнопки "применение диапазона" и подписей
                    if (BU.QuantityFiles + IPT.QuantityFiles >= 2)
                    {
                        button2.Visible = true;
                        label20.Visible = true;
                        label18.Visible = true;
                        label19.Visible = true;
                        textBox13.Visible = true;
                        textBox14.Visible = true;

                    }

                }
            }
        }

        private void button4_Click(object sender, EventArgs e) //выбор файлов через OpenFileDialog
        {
            OpenFileDialog dlg = new OpenFileDialog
            {
                Multiselect = true,
                Title = "Выберите файлы",
                InitialDirectory = @"C:\"
            };
            dlg.ShowDialog();
            // пользователь вышел из диалога ничего не выбрав
            if (dlg.FileName == String.Empty) return;

            int i = Convert.ToInt32(((Button)(sender)).Tag);
            if (i==1)
            {         
                 foreach (string file in dlg.FileNames)
                 {
                IPT.PathFile[IPT.QuantityFiles] = file;
                IPT.Items[IPT.QuantityFiles] = System.IO.Path.GetFileName(file); //Только имя файла (с расширением):
                                                                                                   // dbf_File = System.IO.Path.GetFileName(dbf_File);
                //label1.Text = IPT.PathFile[IPT.QuantityFiles];
                listBox1.Items.Add(IPT.Items[IPT.QuantityFiles]);
                IPT.QuantityFiles++;
                 }
            }
            else
            {
                foreach (string file in dlg.FileNames)
                {
                    BU.PathFile[BU.QuantityFiles] = file;
                    BU.Items[BU.QuantityFiles] = System.IO.Path.GetFileName(file); //Только имя файла (с расширением):
                                                                                     // dbf_File = System.IO.Path.GetFileName(dbf_File);
                                                                                     //label1.Text = IPT.PathFile[IPT.QuantityFiles];
                    listBox2.Items.Add(BU.Items[BU.QuantityFiles]);
                    BU.QuantityFiles++;
                }
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
        void writtenIPT(int offset, string name) // в аргументе смещение потому что будет записываться столбцы от каждого ИПТ и надо их свдвигать вправо
        {
            excelcellsnew[CountCurentCell, 3 + offset].Value2 = name; //две строчки под название файла
            CountCurentCell++;
            for (int ih = 0; ih <= 1; ih++)
            {
                
                for (int im = 0; im < (60); im++)
                {
                   
                    for (int isec = 0; isec < (12); isec++)
                    {
                        if ((CountCurentCell-2) == CountCells) break;
                        excelcellsnew[CountCurentCell, 4 + offset].Value2 = excelcells2[CountCurentCellback, 6].Value2;// Т_3
                        excelcellsnew[CountCurentCell, 3 + offset].Value2 = excelcells2[CountCurentCellback, 5].Value2; //D:D //Т_0
                        
                        excelcellsnew[CountCurentCell, 1 + offset].EntireColumn.NumberFormat = "[$-ru-RU,1] ДД.ММ.ГГГГ ч:мм:сс";
                        excelcellsnew[CountCurentCell, 1 + offset].Value2 = excelcells2[CountCurentCellback, 2].Value2;//Bnew:C //время датчика

                        excelcellsnew[CountCurentCell, 2 + offset].Value2 = Stime.hours[ih] + ":" + Stime.minutes[im] + ":" + Stime.seconds_05[isec];
                        CountCurentCell++;
                        CountCurentCellback--;
                    }
                }

            }
        }
        /*****************************Запись времени и значений для БУ***************************************************/
        void writtenBU(int offset , string name) // в аргументе смещение потому что будет записываться столбцы от каждого прибора и надо их свдвигать вправо
        {
            excelcellsnew[CountCurentCell, 3 + offset].Value2 = name; //две строчки под название файла
            CountCurentCell++;
            for (int ih = 0; ih <= 1; ih++)
            {
               
                for (int im = 0; im < (60); im++)
                {
                    
                    for (int isec = 0; isec < (12); isec++)
                    {
                        if ((CountCurentCell-2) == CountCells) break;
                        
                        excelcellsnew[CountCurentCell, 3 + offset].Value2 = excelcells2[CountCurentCellback, 5].Value2; //D:D //Т
                        excelcellsnew[CountCurentCell, 1 + offset].EntireColumn.NumberFormat = "[$-ru-RU,1] ДД.ММ.ГГГГ ч:мм:сс";
                        excelcellsnew[CountCurentCell, 1 + offset].Value2 = excelcells2[CountCurentCellback, 2].Value2;//Bnew:C //время датчика

                       excelcellsnew[CountCurentCell, 2 + offset].Value2 = Stime.hours[ih] + ":" + Stime.minutes[im] + ":" + Stime.seconds_05[isec];
                        CountCurentCell++;
                        CountCurentCellback--;
                    }
                }

            }
        }
        
        void Appear_label_drag(int QuantityFile, string fileName, Unit Ustr) // появление label в зависимости от момента перетаскивания файлов
        {
            if (Ustr == Unit.IPT)
            {
                switch (QuantityFile)
                {
                    case 0: label100.Text = System.IO.Path.GetFileName(fileName); label100.Visible = true;
                        label2.Visible = true; label3.Visible = true; label4.Visible = true;
                        textBox1.Visible = true; textBox2.Visible = true; break; 

                    case 1: label101.Text = System.IO.Path.GetFileName(fileName); label101.Visible = true;
                        label7.Visible = true; label8.Visible = true; 
                        textBox3.Visible = true; textBox4.Visible = true; break;
                        
                    case 2: label102.Text = System.IO.Path.GetFileName(fileName); label102.Visible = true;
                        label11.Visible = true; label13.Visible = true;
                        textBox5.Visible = true; textBox6.Visible = true; break;

                    default: break;
                }
                
                
            }
            else
            {
                switch (QuantityFile)
                {
                    case 0: label103.Text = System.IO.Path.GetFileName(fileName); label103.Visible = true;
                         label12.Visible = true; label14.Visible = true; label15.Visible = true;
                        textBox7.Visible = true; textBox8.Visible = true; break;

                    case 1: label104.Text = System.IO.Path.GetFileName(fileName); label104.Visible = true; 
                        label9.Visible = true; label10.Visible = true;
                        textBox9.Visible = true; textBox10.Visible = true; break;

                    case 2: label105.Text = System.IO.Path.GetFileName(fileName); label105.Visible = true;
                        label16.Visible = true; label17.Visible = true;
                        textBox11.Visible = true; textBox12.Visible = true; break;

                    default: break;
                }
            }
        }

        //строим графики         
        void BuildingCharts()
        {
            Excel.Range[] rngIPT = new Excel.Range[IPT.QuantityFiles];
            Excel.Range[] rngBU = new Excel.Range[BU.QuantityFiles];
            char nameColumn = 'D';
            // label6.Text = ((char)((int)nameColumn + 1)).ToString(); 

            //Add a Chart for the selected data.
            for (int i = 0; i < IPT.QuantityFiles; i++)
            {
                rngIPT[i] = excelworksheetnew.get_Range(nameColumn.ToString() + "13:" + ((char)((int)nameColumn + 1)).ToString() + (CountCellGeneral[i] + 12).ToString(), Type.Missing); // D and C column
                nameColumn = (char)((int)nameColumn + 5);
            }

            for (int i = 0; i < BU.QuantityFiles; i++)
            {
                rngBU[i] = excelworksheetnew.get_Range(nameColumn.ToString() + "13:" + nameColumn.ToString() + (CountCellGeneral[i + 3] + 12).ToString(), Type.Missing); // D and C column
                nameColumn = (char)((int)nameColumn + 5);
            }
            //объединение рядов
            switch (IPT.QuantityFiles)
            {
                case 0:
                    switch (BU.QuantityFiles)
                    {

                        case 1:
                            excelcellsnew = rngBU[0]; break;
                        case 2:
                            excelcellsnew = excelapp.Union(rngBU[0], rngBU[1],
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing); break;
                        case 3:
                            excelcellsnew = excelapp.Union(rngBU[0], rngBU[1],
                            rngBU[2], Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing); break;
                        default: break;
                    }
                    break;
                case 1:
                    switch (BU.QuantityFiles)
                    {
                        case 0:
                            excelcellsnew = rngIPT[0]; break;
                        case 1:
                            excelcellsnew = excelapp.Union(rngIPT[0], rngBU[0],
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing); break;
                        case 2:
                            excelcellsnew = excelapp.Union(rngIPT[0], rngBU[0],
                            rngBU[1], Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing); break;
                        case 3:
                            excelcellsnew = excelapp.Union(rngIPT[0], rngBU[0],
                            rngBU[1], rngBU[2], Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing); break;
                        default: break;
                    }
                    break;
                case 2:
                    switch (BU.QuantityFiles)
                    {
                        case 0:
                            excelcellsnew = excelapp.Union(rngIPT[0], rngIPT[1],
                           Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                           Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                           Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                           Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                           Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                           Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                           Type.Missing, Type.Missing, Type.Missing, Type.Missing); break;
                        case 1:
                            excelcellsnew = excelapp.Union(rngIPT[0], rngBU[0],
                            rngIPT[1], Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing); break;
                        case 2:
                            excelcellsnew = excelapp.Union(rngIPT[0], rngBU[0],
                            rngBU[1], rngIPT[1], Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing); break;
                        case 3:
                            excelcellsnew = excelapp.Union(rngIPT[0], rngBU[0],
                            rngBU[1], rngBU[2], rngIPT[1], Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing); break;
                        default: break;
                    } break;

                case 3:
                    switch (BU.QuantityFiles)
                    {
                        case 0:
                            excelcellsnew = excelapp.Union(rngIPT[0], rngIPT[1],
                           rngIPT[2], Type.Missing, Type.Missing, Type.Missing,
                           Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                           Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                           Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                           Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                           Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                           Type.Missing, Type.Missing, Type.Missing, Type.Missing); break;
                        case 1:
                            excelcellsnew = excelapp.Union(rngIPT[0], rngBU[0],
                            rngIPT[1], rngIPT[2], Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing); break;
                        case 2:
                            excelcellsnew = excelapp.Union(rngIPT[0], rngBU[0],
                            rngBU[1], rngIPT[1], rngIPT[2], Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing); break;
                        case 3:
                            excelcellsnew = excelapp.Union(rngIPT[0], rngBU[0],
                            rngBU[1], rngBU[2], rngIPT[1], rngIPT[2],
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing); break;
                        default: break;
                    }
                    break;

                default: break;
            }

            // 3 строчки ниже чтобы задать размеры
            Excel.ChartObjects chartObjs = (Excel.ChartObjects)excelworksheetnew.ChartObjects();
            Excel.ChartObject chartObj = chartObjs.Add(5, 50, 1000, 500);
            oChart = chartObj.Chart;


            //oChart = (Excel.Chart)excelappworkbooknew.Charts.Add(Missing.Value, Missing.Value,
            //Missing.Value, Missing.Value);



            //Excel.ChartObject myChart = (Excel.ChartObject)xlCharts.Add(10, 80, 300, 250);
            //Use the ChartWizard to create a new chart from the selected data.                                                     
            // rng1 = excelapp.get_Range("B137:Y137, B139:Y139, B141:Y141", Missing.Value); // эта запись не сработала, сработало только Union для объеждинения нескольких рядов
            //!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

            oChart.ChartWizard(excelcellsnew, Excel.XlChartType.xlLine, Missing.Value,
            Excel.XlRowCol.xlColumns, Missing.Value, Missing.Value, Missing.Value,
            Missing.Value, Missing.Value, Missing.Value, Missing.Value);



            //ChartWizard метод для создания трехмерной гистограммы, отображающий ряд данных в диапазоне ячеек 

            ////Перемещаем диаграмму в нужное место
            //excelworksheetnew.Shapes.Item(1).IncrementLeft(-201);
            //excelworksheetnew.Shapes.Item(1).IncrementTop((float)20.5);

            ////Задаем размеры диаграммы
            //excelworksheetnew.Shapes.Item(1).Height = 1000;
            //excelworksheetnew.Shapes.Item(1).Width = 100;


            oSeries = (Excel.Series)oChart.SeriesCollection(1); //подпись рядов справа

            string arg2 = "A" + (CountMax + 12).ToString();
            oSeries.XValues = excelworksheetnew.get_Range("A13", arg2); //подпись оси Х
            int iRet = 0;
            int j = 0;
            //char[] xlsx = new char[5] {'.','x','l','s','x' };
            string xlsx = ".xlsx";
            for (iRet = 0,  j = 0; iRet < IPT.QuantityFiles; iRet += 2, j++)
            {
                oSeries = (Excel.Series)oChart.SeriesCollection(iRet+1);
                String seriesName;
                seriesName = IPT.Items[j] + " T0";
                if (seriesName.IndexOf(xlsx) > -1) seriesName = seriesName.Replace(xlsx, ""); 
                //seriesName = String.Concat(seriesName, iRet);  // склеивание стринг файлов
                //seriesName = String.Concat(seriesName, "\"");
                oSeries.Name = seriesName;

                oSeries = (Excel.Series)oChart.SeriesCollection(iRet + 2);
                String seriesName2;
                seriesName2 = IPT.Items[j] + " T3";
                if (seriesName2.IndexOf(xlsx) > -1) seriesName2 = seriesName2.Replace(xlsx, " ");
                oSeries.Name = seriesName2;
            }
            j = 0;
            for (int iRet2 = iRet ; j < BU.QuantityFiles; iRet++,j++)
            {
                oSeries = (Excel.Series)oChart.SeriesCollection(iRet2 + 1);
                String seriesName3;
                seriesName3 = BU.Items[j];
                if (seriesName3.IndexOf(xlsx) > -1) seriesName3 = seriesName3.Replace(xlsx, "");
                //  лучше использовать  - if (files[i].EndsWith(".exe")) для определения заканчитвается ли строка на .exe
                oSeries.Name = seriesName3;
            }

            oChart.Location(Excel.XlChartLocation.xlLocationAsObject, excelworksheetnew.Name);

            




 

            ////Excel.Range chartRange;
            //Excel.Range chartRange;
            //Excel.Range misValue;
            //Excel.Range misValue2;

            //Excel.ChartObjects xlCharts = (Excel.ChartObjects)excelworksheetnew.ChartObjects(Type.Missing);
            //Excel.ChartObject myChart = (Excel.ChartObject)xlCharts.Add(10, 80, 300, 250);
            //Excel.Chart chartPage = myChart.Chart;
            //Excel.Chart chartPage2 = myChart.Chart;

            //// chartRange = excelworksheetnew.get_Range("C13", "D33");
            //misValue = excelworksheetnew.get_Range("C13", "E33");
            ////chartPage.SetSourceData(chartRange, Type.Missing);
            //chartPage.SetSourceData(misValue, Type.Missing);
            //chartPage.ChartType = Excel.XlChartType.xlLine;
            ////Excel.Series ser = (Excel.Series)chartPage.SeriesCollection(1);

            //misValue2 = excelworksheetnew.get_Range("F13", "I33");
            ////chartPage.SetSourceData(chartRange, Type.Missing);
            //chartPage2.SetSourceData(misValue2, Type.Missing);
            //chartPage2.ChartType = Excel.XlChartType.xlLine;
            ////ser.XValues = ws.Range[ws.cells[row, col], ws.cells[row, col]];
            ////chartPage.HasLegend = true;

            ////chartRange = excelworksheetnew.get_Range("C13", "D13");
            ////chartPage.SetSourceData(chartRange, misValue);
            ////chartPage.ChartType = Excel.XlChartType.xlColumnClustered;
            
        }



        void writtenBigTime()
        {
            
            CountCurentCell = 1; //для тройного цикла (переключения ячеек) //1 потому что excelcells2[1, 1] начинается с 1
            excelcellsnew = excelworksheetnew.get_Range("A13" , "A13" ); //откуда смотрим ячейки

            for (int ih = 0; ih <= 1; ih++)
            {
                for (int im = 0; im < (60); im++)
                {      
                   if ((CountCurentCell - 1) >= CountMax) break;

                   excelcellsnew[CountCurentCell, 1].Value2 = Stime.hours[ih] + ":" + Stime.minutes[im];                      
                   CountCurentCell+= 12;
                    
                }

            }
        }
        

    }





}


//// Строим круговую диаграмму
//Excel.ChartObjects chartObjs = (Excel.ChartObjects)workSheet.ChartObjects();
//Excel.ChartObject chartObj = chartObjs.Add(5, 50, 300, 300);
//Excel.Chart xlChart = chartObj.Chart;
//Excel.Range rng2 = workSheet.Range["A1:L1"];
//// Устанавливаем тип диаграммы
//xlChart.ChartType = Excel.XlChartType.xlPie;
//      // Устанавливаем источник данных (значения от 1 до 10)
//      xlChart.SetSourceData(rng2);


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



//void BuildingCharts ()
//{
//    ////Выделяем ячейки с данными  в таблице
//    //excelappworkbooknew = excelappworkbooks[1]; ////Получаем ссылку на книгу 1 - нумерация от 1
//    //excelsheetsnew = excelappworkbooknew.Worksheets;   //Получаем массив ссылок на листы выбранной книги
//    //excelworksheetnew = (Excel.Worksheet)excelsheetsnew.get_Item(1); //получаем ссылку на первый лист 
//    excelcellsnew = excelworksheetnew.get_Range("C13", "D13");
//    //И выбираем их
//   excelcellsnew.Select();
//    //Создаем объект Excel.Chart диаграмму по умолчанию
//    Excel.Chart excelchart = (Excel.Chart)excelapp.Charts.Add(Type.Missing,
//     Type.Missing, Type.Missing, Type.Missing);
//    //Выбираем диграмму - отображаем лист с диаграммой
//    excelchart.Activate();
//    excelchart.Select(Type.Missing);
//    //Изменяем тип диаграммы
//    excelapp.ActiveChart.ChartType = Excel.XlChartType.xlConeCol;
//    //Создаем надпись - Заглавие диаграммы
//    excelapp.ActiveChart.HasTitle = true;
//    excelapp.ActiveChart.ChartTitle.Text
//       = "Продажи фирмы Рога и Копыта за неделю";
//    //Меняем шрифт, можно поменять и другие параметры шрифта
//    excelapp.ActiveChart.ChartTitle.Font.Size = 14;
//    excelapp.ActiveChart.ChartTitle.Font.Color = 255;
//    //Обрамление для надписи c тенями
//    excelapp.ActiveChart.ChartTitle.Shadow = true;
//    excelapp.ActiveChart.ChartTitle.Border.LineStyle
//         = Excel.Constants.xlSolid;
//    //Даем названия осей
//    ((Excel.Axis)excelapp.ActiveChart.Axes(Excel.XlAxisType.xlCategory,
//        Excel.XlAxisGroup.xlPrimary)).HasTitle = true;
//    ((Excel.Axis)excelapp.ActiveChart.Axes(Excel.XlAxisType.xlCategory,
//        Excel.XlAxisGroup.xlPrimary)).AxisTitle.Text = "День недели";
//    ((Excel.Axis)excelapp.ActiveChart.Axes(Excel.XlAxisType.xlSeriesAxis,
//        Excel.XlAxisGroup.xlPrimary)).HasTitle = false;
//    ((Excel.Axis)excelapp.ActiveChart.Axes(Excel.XlAxisType.xlValue,
//        Excel.XlAxisGroup.xlPrimary)).HasTitle = true;
//    ((Excel.Axis)excelapp.ActiveChart.Axes(Excel.XlAxisType.xlValue,
//        Excel.XlAxisGroup.xlPrimary)).AxisTitle.Text = "Рогов/Копыт";
//    //Координатная сетка - оставляем только крупную сетку
//    ((Excel.Axis)excelapp.ActiveChart.Axes(Excel.XlAxisType.xlCategory,
//       Excel.XlAxisGroup.xlPrimary)).HasMajorGridlines = true;
//    ((Excel.Axis)excelapp.ActiveChart.Axes(Excel.XlAxisType.xlCategory,
//      Excel.XlAxisGroup.xlPrimary)).HasMinorGridlines = false;
//    ((Excel.Axis)excelapp.ActiveChart.Axes(Excel.XlAxisType.xlSeriesAxis,
//      Excel.XlAxisGroup.xlPrimary)).HasMajorGridlines = true;
//    ((Excel.Axis)excelapp.ActiveChart.Axes(Excel.XlAxisType.xlSeriesAxis,
//      Excel.XlAxisGroup.xlPrimary)).HasMinorGridlines = false;
//    ((Excel.Axis)excelapp.ActiveChart.Axes(Excel.XlAxisType.xlValue,
//      Excel.XlAxisGroup.xlPrimary)).HasMinorGridlines = false;
//    ((Excel.Axis)excelapp.ActiveChart.Axes(Excel.XlAxisType.xlValue,
//      Excel.XlAxisGroup.xlPrimary)).HasMajorGridlines = true;
//    //Будем отображать легенду и уберем строки, 
//    //которые отображают пустые строки таблицы
//    excelapp.ActiveChart.HasLegend = true;
//    //Расположение легенды
//    excelapp.ActiveChart.Legend.Position
//       = Excel.XlLegendPosition.xlLegendPositionLeft;
//    //Можно изменить шрифт легенды и другие параметры 
//   // потом разобраться с исключением
//        ((Excel.LegendEntry)excelapp.ActiveChart.Legend.LegendEntries(1)).Font.Size = 12;
//   // ((Excel.LegendEntry)excelapp.ActiveChart.Legend.LegendEntries(2)).Font.Size = 12;
//    //Легенда тесно связана с подписями на осях - изменяем надписи
//    // - меняем легенду, удаляем чтото на оси - изменяется легенда
//    Excel.SeriesCollection seriesCollection =
//     (Excel.SeriesCollection)excelapp.ActiveChart.SeriesCollection(Type.Missing);
//    Excel.Series series = seriesCollection.Item(1);
//    series.Name = "Рога";
//    //Помним, что у нас объединенные ячейки, значит каждая второя строка - пустая
//    //Удаляем их из диаграммы и из легенды
//    //!!!!! тоже проблеммы
//    //series = seriesCollection.Item(2);
//    //series.Delete();
//    ////После удаления второго (пустого набора значений) третий занял его место
//    //series = seriesCollection.Item(2);
//    // series = seriesCollection.Item(2);
//    //  series.Name = "Копыта";
//    //series = seriesCollection.Item(3);
//    //series.Delete();
//    //series = seriesCollection.Item(1);
//    //Переименуем ось X
//    //series.XValues = "Понедельник;Вторник;Среда;Четверг;Пятница;Суббота;Воскресенье;Итог";
//    //Если закончить код на этом месте то у нас Диаграммы на отдельном листе - Рис.9.
//    //Строку легенды можно удалить здесь, но строка на оси не изменится
//    //Поэтому мы удаляли в Excel.Series
//    //((Excel.LegendEntry)excelapp.ActiveChart.Legend.LegendEntries(2)).Delete();
//    //Перемещаем диаграмму на лист 1


//    //excelapp.SheetsInNewWorkbook = 1; //1 количество листов в новой книге



//    //excelapp.ActiveChart.Location(Excel.XlChartLocation.xlLocationAsObject, "Лист1");
//    ////Получаем ссылку на лист 1
//    //excelsheetsnew = excelappworkbook.Worksheets;
//    excelworksheetnew = (Excel.Worksheet)excelsheetsnew.get_Item(1);
//    //Перемещаем диаграмму в нужное место
//    //excelworksheetnew.Shapes.Item(1).IncrementLeft(1);
//    //excelworksheetnew.Shapes.Item(1).IncrementTop((float)20.5);
//    //Задаем размеры диаграммы
//    //excelworksheetnew.Shapes.Item(1).Height = 550;
//    //excelworksheetnew.Shapes.Item(1).Width = 500;
//    //Конец кода - диаграммы на листе там где и таблица



//http://www.ishodniki.ru/art/474.html есть вообще всё

 //    Пример от Майкрософт!!!!!!!!!!!!!!

//    using System;
//using System.Collections.Generic;
//using System.ComponentModel;
//using System.Data;
//using System.Drawing;
//using System.Linq;
//using System.Text;
//using System.Threading.Tasks;
//using System.Windows.Forms;
//using Excel = Microsoft.Office.Interop.Excel;
//using System.Reflection;

//namespace WindowsFormsApp1
//{
//    public partial class Form1 : Form
//    {
//        public Form1()
//        {
//            InitializeComponent();
//        }

//        private void button1_Click(object sender, EventArgs e)
//        {
//            Excel.Application oXL;
//            Excel._Workbook oWB;
//            Excel._Worksheet oSheet;
//            Excel.Range oRng;

//            try
//            {
//                //Start Excel and get Application object.
//                oXL = new Excel.Application();
//                oXL.Visible = true;

//                //Get a new workbook.
//                oWB = (Excel._Workbook)(oXL.Workbooks.Add(Missing.Value));
//                oSheet = (Excel._Worksheet)oWB.ActiveSheet;

//                //Add table headers going cell by cell.
//                oSheet.Cells[1, 1] = "First Name";
//                oSheet.Cells[1, 2] = "Last Name";
//                oSheet.Cells[1, 3] = "Full Name";
//                oSheet.Cells[1, 4] = "Salary";

//                //Format A1:D1 as bold, vertical alignment = center.
//                oSheet.get_Range("A1", "D1").Font.Bold = true;
//                oSheet.get_Range("A1", "D1").VerticalAlignment =
//                Excel.XlVAlign.xlVAlignCenter;

//                // Create an array to multiple values at once.
//                string[,] saNames = new string[5, 2];

//                saNames[0, 0] = "John";
//                saNames[0, 1] = "Smith";
//                saNames[1, 0] = "Tom";
//                saNames[1, 1] = "Brown";
//                saNames[2, 0] = "Sue";
//                saNames[2, 1] = "Thomas";
//                saNames[3, 0] = "Jane";
//                saNames[3, 1] = "Jones";
//                saNames[4, 0] = "Adam";
//                saNames[4, 1] = "Johnson";

//                //Fill A2:B6 with an array of values (First and Last Names).
//                oSheet.get_Range("A2", "B6").Value2 = saNames;

//                //Fill C2:C6 with a relative formula (=A2 & " " & B2).
//                oRng = oSheet.get_Range("C2", "C6");
//                oRng.Formula = "=A2 & \" \" & B2";

//                //Fill D2:D6 with a formula(=RAND()*100000) and apply format.
//                oRng = oSheet.get_Range("D2", "D6");
//                oRng.Formula = "=RAND()*100000";
//                oRng.NumberFormat = "$0.00";

//                //AutoFit columns A:D.
//                oRng = oSheet.get_Range("A1", "D1");
//                oRng.EntireColumn.AutoFit(); //автоподбор высоты и ширины строк

//                //Manipulate a variable number of columns for Quarterly Sales Data.
//                DisplayQuarterlySales(oSheet);

//                //Make sure Excel is visible and give the user control
//                //of Microsoft Excel's lifetime.
//                oXL.Visible = true;
//                oXL.UserControl = true;
//            }
//            catch (Exception theException)
//            {
//                String errorMessage;
//                errorMessage = "Error: ";
//                errorMessage = String.Concat(errorMessage, theException.Message);
//                errorMessage = String.Concat(errorMessage, " Line: ");
//                errorMessage = String.Concat(errorMessage, theException.Source);

//                MessageBox.Show(errorMessage, "Error");
//            }
//        }

//        private void DisplayQuarterlySales(Excel._Worksheet oWS)
//        {
//            Excel._Workbook oWB;
//            Excel.Series oSeries;
//            Excel.Range oResizeRange;
//            Excel._Chart oChart;
//            String sMsg;
//            int iNumQtrs;

//            //Determine how many quarters to display data for.
//            for (iNumQtrs = 4; iNumQtrs >= 2; iNumQtrs--)
//            {
//                sMsg = "Enter sales data for ";
//                sMsg = String.Concat(sMsg, iNumQtrs);
//                sMsg = String.Concat(sMsg, " quarter(s)?");

//                DialogResult iRet = MessageBox.Show(sMsg, "Quarterly Sales?",
//                MessageBoxButtons.YesNo);
//                if (iRet == DialogResult.Yes)
//                    break;
//            }

//            sMsg = "Displaying data for ";
//            sMsg = String.Concat(sMsg, iNumQtrs); //скрепляет string переменные
//            sMsg = String.Concat(sMsg, " quarter(s).");

//            MessageBox.Show(sMsg, "Quarterly Sales");

//            //Starting at E1, fill headers for the number of columns selected.
//            oResizeRange = oWS.get_Range("E1", "E1").get_Resize(Missing.Value, iNumQtrs);
//            //  oResizeRange.Formula = "=\"Q\" & COLUMN()-4 & CHAR(10) & \"Sales\""; //char(10) это перенос строки, но у меня почему-то не работает

//            //Change the Orientation and WrapText properties for the headers.
//            oResizeRange.Orientation = 38;
//            oResizeRange.WrapText = true;

//            //Fill the interior color of the headers.
//            oResizeRange.Interior.ColorIndex = 36;

//            //Fill the columns with a formula and apply a number format.
//            oResizeRange = oWS.get_Range("E2", "E6").get_Resize(Missing.Value, iNumQtrs); //get resize позволяет применить тот же эффект для ячеек справа ( iNumQtrs их количество)
//            oResizeRange.Formula = "=RAND()*100";
//            oResizeRange.NumberFormat = "$0.00";

//            //Apply borders to the Sales data and headers.
//            oResizeRange = oWS.get_Range("E1", "E6").get_Resize(Missing.Value, iNumQtrs); //
//            oResizeRange.Borders.Weight = Excel.XlBorderWeight.xlThin;

//            //Add a Totals formula for the sales data and apply a border.
//            oResizeRange = oWS.get_Range("E8", "E8").get_Resize(Missing.Value, iNumQtrs);
//            oResizeRange.Formula = "=SUM(E2:E6)";
//            oResizeRange.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).LineStyle
//            = Excel.XlLineStyle.xlDouble;
//            oResizeRange.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).Weight
//            = Excel.XlBorderWeight.xlThick;

//            //Add a Chart for the selected data.
//            oWB = (Excel._Workbook)oWS.Parent; //Возвращает родительский объект для указанного объекта
//            oChart = (Excel._Chart)oWB.Charts.Add(Missing.Value, Missing.Value,
//            Missing.Value, Missing.Value);

//            //Use the ChartWizard to create a new chart from the selected data.
//            oResizeRange = oWS.get_Range("E2:E6", Missing.Value).get_Resize(
//            Missing.Value, iNumQtrs);

//            //  oResizeRange = oWS.get_Range("B137:Y137", "B139:Y139", Missing.Value);

//            oChart.ChartWizard(oResizeRange, Excel.XlChartType.xl3DColumn, Missing.Value,
//            Excel.XlRowCol.xlColumns, Missing.Value, Missing.Value, Missing.Value,
//            Missing.Value, Missing.Value, Missing.Value, Missing.Value);
//            oSeries = (Excel.Series)oChart.SeriesCollection(1);
//            oSeries.XValues = oWS.get_Range("A2", "A6");
//            for (int iRet = 1; iRet <= iNumQtrs; iRet++)
//            {
//                oSeries = (Excel.Series)oChart.SeriesCollection(iRet);
//                String seriesName;
//                seriesName = "=\"Q";
//                seriesName = String.Concat(seriesName, iRet);
//                seriesName = String.Concat(seriesName, "\"");
//                oSeries.Name = seriesName;
//            }

//            oChart.Location(Excel.XlChartLocation.xlLocationAsObject, oWS.Name);

//            //Move the chart so as not to cover your data.
//            oResizeRange = (Excel.Range)oWS.Rows.get_Item(10, Missing.Value);
//            oWS.Shapes.Item("Chart 1").Top = (float)(double)oResizeRange.Top;
//            oResizeRange = (Excel.Range)oWS.Columns.get_Item(2, Missing.Value);
//            oWS.Shapes.Item("Chart 1").Left = (float)(double)oResizeRange.Left;
//        }
//    }
//}