using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelImportExport
{
    public partial class Form1 : Form
    {
        List<List<string>> list_ = new List<List<string>>();
        List<List<string>> estimatesId = new List<List<string>>();
        List<string> allEstimates = new List<string>();
        List<string> remarks = new List<string>();
        List<string> missing = new List<string>();
        int count = 0;

        public Form1()
        {
            InitializeComponent();
        }

        private void ImportBtn_Click(object sender, EventArgs e)
        {
            int[] n = ExportExcel();

            ResultsList.Items.Clear();
            
            string s;

            
            foreach (var item in list_)
            {
                allEstimates.Add(item[0].ToString());

                if (!string.IsNullOrEmpty(item[5]))
                {
                    remarks.Add(item[5]);
                }                           
                                
            }
            remarksCount.Text = remarks.Count.ToString();

            foreach(var item in remarks)
            {
                ResultsList.Items.Add(item);
            }

            foreach (var item in list_)
            {
                if (remarks.Contains(item[0]))
                {
                    estimatesId.Add(item);
                }
            }

            foreach (var item in estimatesId)
            {
                s = "";
                s = item[0] + "|" + item[1];
               
                coincidencesBox.Items.Add(s);
            }
            coincidences.Text = estimatesId.Count.ToString();

            foreach (var item in list_)
            {
                if (remarks.Contains(item[0]))
                {
                    estimatesId.Add(item);
                }
            }
            foreach (var item in remarks)
            {
                if (!allEstimates.Contains(item))
                {
                    missingBox.Items.Add(item);
                    count++;
                }
            }
            missingCountLb.Text = count.ToString();
        }
        // Импорт данных из Excel-файла (не более 5 столбцов и любое количество строк <= 50.
        private int[] ExportExcel()
        {
            
            int[] result = new int[]{0, 0};
            // Выбрать путь и имя файла в диалоговом окне
            OpenFileDialog ofd = new OpenFileDialog();

            // Задаем расширение имени файла по умолчанию (открывается папка с программой)
            ofd.DefaultExt = "*.xls;*.xlsx";

            // Задаем строку фильтра имен файлов, которая определяет варианты
            ofd.Filter = "файл Excel (Spisok.xlsx)|*.xlsx";

            // Задаем заголовок диалогового окна
            ofd.Title = "Выберите файл базы данных";

            if (!(ofd.ShowDialog() == DialogResult.OK)) // если файл БД не выбран -> Выход
                return result;

            Excel.Application ObjWorkExcel = new Excel.Application();

            Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(ofd.FileName);

            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1]; //получить 1-й лист

            var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);//последнюю ячейку
                                                                                                // размеры базы
            int lastColumn = (int)lastCell.Column;

            int lastRow = (int)lastCell.Row;

            result[0] = lastRow;
            result[1] = lastColumn;


            for (int j = 0; j < lastRow; j++) //по всем колонкам
            {
                List<string> temp = new List<string>();

                for (int i = 0; i < lastColumn; i++) // по всем строкам
                {
                    
                    temp.Add(ObjWorkSheet.Cells[j + 1, i + 1].Text.ToString());
                    
                }
                list_.Add(temp);
                
            }

            ObjWorkBook.Close(false, Type.Missing, Type.Missing); //закрыть не сохраняя
            
            ObjWorkExcel.Quit(); // выйти из Excel
            
            GC.Collect(); // убрать за собой
            
            return result;
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}
