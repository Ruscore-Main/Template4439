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

namespace Template4439
{
    /// <summary>
    /// Логика взаимодействия для _4439_Bikbaev.xaml
    /// </summary>
    public partial class _4439_Bikbaev : Window
    {
        public _4439_Bikbaev()
        {
            InitializeComponent();
        }

        public static int GetAge(DateTime birthDate)
        {
            var now = DateTime.Today;
            return now.Year - birthDate.Year - 1 +
                ((now.Month > birthDate.Month || now.Month == birthDate.Month && now.Day >= birthDate.Day) ? 1 : 0);
        }

        private void ExportBtn_Click(object sender, RoutedEventArgs e)
        {

            List<Client> clients;

            using (ISRPO_LR2_EXCELEntities db = new ISRPO_LR2_EXCELEntities())
            {
                clients = db.Clients.ToList();
            }

            int countAgeCategories = 3;
            List<int[]> AgeCategories = new List<int[]>() {
                new int[]{ 20, 29 },
                new int[]{ 30, 39 },
                new int[]{ 40, int.MaxValue },
            };
            var app = new Microsoft.Office.Interop.Excel.Application();
            app.SheetsInNewWorkbook = countAgeCategories;
            Microsoft.Office.Interop.Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);

            for (int i = 0; i < countAgeCategories; i++)
            {
                int startRowIndex = 1;
                Microsoft.Office.Interop.Excel.Worksheet worksheet = app.Worksheets.Item[i + 1];
                worksheet.Name = $"Категория - {i + 1}";

                Microsoft.Office.Interop.Excel.Range headerRange = worksheet.Range[worksheet.Cells[1][1], worksheet.Cells[3][1]];
                headerRange.Merge();
                headerRange.Value = $"Категория - {i + 1}";
                headerRange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                headerRange.Font.Italic = true;
                startRowIndex++;


                worksheet.Cells[1][startRowIndex] = "Код клиента";
                worksheet.Cells[2][startRowIndex] = "ФИО";
                worksheet.Cells[3][startRowIndex] = "E-mail";

                startRowIndex++;



                foreach (Client client in clients)
                {
                    if (client.Age >= AgeCategories[i][0] && client.Age <= AgeCategories[i][1])
                    {
                        worksheet.Cells[1][startRowIndex] = client.ClientCode;
                        worksheet.Cells[2][startRowIndex] = client.FIO;
                        worksheet.Cells[3][startRowIndex] = client.Email;
                        startRowIndex++;
                    }
                }

                Microsoft.Office.Interop.Excel.Range rangeBorders = worksheet.Range[worksheet.Cells[1][1], worksheet.Cells[3][startRowIndex - 1]];
                rangeBorders.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle =
                rangeBorders.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].LineStyle =
                rangeBorders.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].LineStyle =
                rangeBorders.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].LineStyle =
                rangeBorders.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal].LineStyle =
                rangeBorders.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                worksheet.Columns.AutoFit();
            }

            MessageBox.Show("Export success");
            app.Visible = true;
        }
        // Импорт данных из Excel-таблицы
        private void ImportBtn_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.xls;*.xlsx",
                Filter = "файл Excel (Spisok.xlsx)|*.xlsx",
                Title = "Выберите файл базы данных"
            };
            if (!(ofd.ShowDialog() == true))
                return;

            string[,] list;
            Microsoft.Office.Interop.Excel.Application ObjWorkExcel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(ofd.FileName);
            Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[1];

            var lastCell = ObjWorkSheet.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell);
            int _columns = (int)lastCell.Column;
            int _rows = (int)lastCell.Row;
            list = new string[_rows, _columns];

            for (int j = 0; j < _columns; j++)
                for (int i = 0; i < _rows; i++)
                    list[i, j] = ObjWorkSheet.Cells[i + 1, j + 1].Text;

            ObjWorkBook.Close(false, Type.Missing, Type.Missing);
            ObjWorkExcel.Quit();
            GC.Collect();

            using (ISRPO_LR2_EXCELEntities db = new ISRPO_LR2_EXCELEntities())
            {
                for (int i = 1; i < _rows; i++)
                {
                    try
                    {
                        db.Clients.Add(new Client()
                        {
                            FIO = list[i, 0],
                            ClientCode = Convert.ToInt32(list[i, 1]),
                            Birthday = Convert.ToDateTime(list[i, 2]),
                            ClientIndex = list[i, 3],
                            City = list[i, 4],
                            Street = list[i, 5],
                            HouseNumber = Convert.ToInt32(list[i, 6]),
                            FlatNumber = Convert.ToInt32(list[i, 7]),
                            Email = list[i, 8],
                            Age = GetAge(Convert.ToDateTime(list[i, 2]))
                        });
                    }
                    catch { }
                }
                db.SaveChanges();
                MessageBox.Show("Import success");
            }

        }
    }
}
