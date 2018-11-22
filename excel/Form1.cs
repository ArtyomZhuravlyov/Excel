﻿ using System;
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

        private OpenFileDialog dlg = new OpenFileDialog();

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
                    excelapp = new Excel.Application();
                    excelapp.Visible = true; 
                    excelapp.SheetsInNewWorkbook = 3; // set количество листов
                    excelapp.Workbooks.Add(Type.Missing);
                    excelapp.SheetsInNewWorkbook = 5;
                    excelapp.Workbooks.Add(Type.Missing);
                    //Запрашивать сохранение
                    excelapp.DisplayAlerts = true;
                    //Получаем набор ссылок на объекты Workbook (на созданные книги)
                    excelappworkbooks = excelapp.Workbooks;
                    //Получаем ссылку на книгу 1 - нумерация от 1
                    excelappworkbook = excelappworkbooks[1];
                    //Ссылку можно получить и так, но тогда надо знать имена книг,
                    //причем, после сохранения - знать расширение файла
                    //excelappworkbook=excelappworkbooks["Книга 1"];
                    //Запроса на сохранение для книги не должно быть
                    excelappworkbook.Saved = true;
                    //Используем свойство Count, число Workbook в Workbooks 
                    if (excelappworkbooks.Count > 1)
                    {
                        excelappworkbook = excelappworkbooks[2];
                        //Запрос на сохранение  книги 2  должен быть
                        excelappworkbook.Saved = false;
                    }
                    break;
                case 2:
                    excelapp.Quit();
                    break;
                default:
                    Close();
                    break;
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {



            dlg.Multiselect = true;
            dlg.Title = "Выберите файлы";
            dlg.InitialDirectory = @"C:\Users\";
            //один из вариантов хорошей записи
            //OpenFileDialog dlg = new OpenFileDialog
            //{
            //    Multiselect = true,
            //    Title = "Выберите файлы",
            //    InitialDirectory = @"C:\"
            //};
            //dlg.ShowDialog();
            dlg.ShowDialog();
            // пользователь вышел из диалога ничего не выбрав
            if (dlg.FileName == String.Empty)
                return;

            textBox1.Visible = true;
            int kk = 0;
            foreach (string file in dlg.FileNames)
            {
                kk++;
                textBox1.Text += file +"\r\n"+ "\r\n";
                //MessageBox.Show(file);
            }

        }
    }
}
