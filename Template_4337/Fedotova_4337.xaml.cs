using GSF.IO;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;

namespace Template_4337
{
    /// <summary>
    /// Логика взаимодействия для Fedotova_4337.xaml
    /// </summary>
    public partial class Fedotova_4337 : Window
    {
        public Fedotova_4337()
        {
            InitializeComponent();
        }

        private void BnImport_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                OpenFileDialog ofd = new OpenFileDialog()
                {
                    DefaultExt = "*.xls;*.xlsx",
                    Filter = "файл Excel (Spisok.xlsx)|*.xlsx",
                    Title = "Выберите файл базы данных"
                };
                if (!(ofd.ShowDialog() == true))
                {
                    return;
                }

                string[,] list; 
                Excel.Application ObjWorkExcel = new Excel.Application();
                Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(ofd.FileName);
                Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];
                var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
                int _columns = (int)lastCell.Column;
                int _rows = (int)lastCell.Row;
                list = new string[_rows, _columns];
                for (int j = 0; j < _columns; j++)
                {
                    for (int i = 0; i < _rows; i++)
                    {
                        list[i, j] = ObjWorkSheet.Cells[i + 1, j + 1].Text;
                    }
                }
                ObjWorkBook.Close(false, Type.Missing, Type.Missing);
                ObjWorkExcel.Quit();
                GC.Collect();

                using (isrpoEntities entities = new isrpoEntities())
                {
                    for (int i = 0; i < _rows; i++)
                    {
                        entities.import2.Add(new import2() { id = Int32.Parse(list[i, 0]), name = list[i, 1], vid_uslugi = list[i, 2], id_uslugi = list[i, 3], price = Int32.Parse(list[i, 4]) });
                    }
                    MessageBox.Show("Успешно!");
                    entities.SaveChanges();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка сохранения данных в базе данных: " + ex.Message);
            }
        }

        private void BnExport_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                List<import2> list;
                using (isrpoEntities isrpo = new isrpoEntities())
                {
                    list = isrpo.import2.ToList().OrderBy(s => s.id).ToList();
                }

                Excel.Application app = new Excel.Application();
                app.SheetsInNewWorkbook = 1;  
                Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);
                Excel.Worksheet worksheet = app.Worksheets.Item[1];
                worksheet.Name = "Группировка по категориям";

                int startRowIndex = 1;
                worksheet.Cells[1][startRowIndex] = "ID";
                worksheet.Cells[2][startRowIndex] = "Наименование услуги";
                worksheet.Cells[3][startRowIndex] = "Вид услуги";
                worksheet.Cells[4][startRowIndex] = "Стоимость, руб. за час";
                startRowIndex++;

                List<import2> category1 = new List<import2>();
                List<import2> category2 = new List<import2>();
                List<import2> category3 = new List<import2>();

                foreach (var item in list)
                {
                    if (item.price >= 0 && item.price < 350)
                    {
                        category1.Add(item);
                    }
                    else if (item.price >= 250 && item.price < 800)
                    {
                        category2.Add(item);
                    }
                    else if (item.price >= 800)
                    {
                        category3.Add(item);
                    }
                }

                WriteDataToWorksheet(worksheet, category1, ref startRowIndex);
                WriteDataToWorksheet(worksheet, category2, ref startRowIndex);
                WriteDataToWorksheet(worksheet, category3, ref startRowIndex);

                CreateWorksheetForCategory(workbook, category1, "Категория 1");
                CreateWorksheetForCategory(workbook, category2, "Категория 2");
                CreateWorksheetForCategory(workbook, category3, "Категория 3");

                SaveFileDialog sfd = new SaveFileDialog()
                {
                    DefaultExt = "*.xlsx",
                    Filter = "Файл Excel (*.xlsx)|*.xlsx",
                    Title = "Сохранить файл Excel"
                };
                if (sfd.ShowDialog() == true)
                {
                    workbook.SaveAs(sfd.FileName);
                    workbook.Close();
                    app.Quit();
                    GC.Collect();

                    MessageBox.Show("Экспорт завершен.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка сохранения данных в базе данных: " + ex.Message);
            }

        }

        private void WriteDataToWorksheet(Excel.Worksheet worksheet, List<import2> data, ref int startRowIndex)
        {
            foreach (var item in data)
            {
                worksheet.Cells[1][startRowIndex] = item.id;
                worksheet.Cells[2][startRowIndex] = item.name;
                worksheet.Cells[3][startRowIndex] = item.vid_uslugi;
                worksheet.Cells[4][startRowIndex] = item.price;
                startRowIndex++;
            }
        }

        private void CreateWorksheetForCategory(Excel.Workbook workbook, List<import2> data, string categoryName)
        {
            Excel.Worksheet worksheet = workbook.Worksheets.Add(Type.Missing, workbook.Sheets[workbook.Sheets.Count], 1, Type.Missing) as Excel.Worksheet;
            worksheet.Name = categoryName;

            int startRowIndex = 1;
            worksheet.Cells[1][startRowIndex] = "ID";
            worksheet.Cells[2][startRowIndex] = "Наименование услуги";
            worksheet.Cells[3][startRowIndex] = "Вид услуги";
            worksheet.Cells[4][startRowIndex] = "Стоимость, руб. за час";
            startRowIndex++;

            foreach (var item in data)
            {
                worksheet.Cells[1][startRowIndex] = item.id;
                worksheet.Cells[2][startRowIndex] = item.name;
                worksheet.Cells[3][startRowIndex] = item.vid_uslugi;
                worksheet.Cells[4][startRowIndex] = item.price;
                startRowIndex++;
            }
        }

    }
}
