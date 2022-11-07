using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
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
        List<int[]> AgeCategories = new List<int[]>() {
                new int[]{ 20, 29 },
                new int[]{ 30, 39 },
                new int[]{ 40, int.MaxValue },
            };
        int countAgeCategories = 3;

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
        // 3LR
        class ClientJson
        {
            public int Id { get; set; }
            public string FullName { get; set; }
            public string CodeClient { get; set; }
            public string BirthDate { get; set; }
            public string Index { get; set; }
            public string City { get; set; }
            public string Street { get; set; }
            public int Home { get; set; }
            public int Kvartira { get; set; }
            public string E_mail { get; set; }


        }



        private void ImportJsonBtn_Click(object sender, RoutedEventArgs e)
        {
            // чтение данных

            string json = File.ReadAllText(@"C:\4course\3.json");
            var clients = JsonSerializer.Deserialize<List<ClientJson>>(json);

            using (ISRPO_LR2_EXCELEntities db = new ISRPO_LR2_EXCELEntities())
            {
                foreach (ClientJson client in clients)
                {
                    try
                    {
                        db.Clients.Add(new Client()
                        {
                            FIO = client.FullName,
                            ClientCode = Convert.ToInt32(client.CodeClient),
                            Birthday = Convert.ToDateTime(client.BirthDate),
                            ClientIndex = client.Index,
                            City = client.City,
                            Street = client.Street,
                            HouseNumber = client.Home,
                            FlatNumber = client.Kvartira,
                            Email = client.E_mail,
                            Age = GetAge(Convert.ToDateTime(client.BirthDate))
                        });
                    }
                    catch { }
                }
                db.SaveChanges();
                MessageBox.Show("Import success");
            }
        }

        private void ExportDocxBtn_Click(object sender, RoutedEventArgs e)
        {
            List<Client> clients;

            using (ISRPO_LR2_EXCELEntities db = new ISRPO_LR2_EXCELEntities())
            {
                clients = db.Clients.ToList();
            }


            var app = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Document document = app.Documents.Add();

            for (int i = 0; i < countAgeCategories; i++)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph = document.Paragraphs.Add();
                Microsoft.Office.Interop.Word.Range range = paragraph.Range;
                range.Text = Convert.ToString($"Категория - {i + 1}");
                paragraph.set_Style("Заголовок 1");
                range.InsertParagraphAfter();

                List<Client> currentClients = clients.Where(c => c.Age >= AgeCategories[i][0] && c.Age <= AgeCategories[i][1]).ToList();
                int countClientInCategory = currentClients.Count();


                Microsoft.Office.Interop.Word.Paragraph tableParagraph = document.Paragraphs.Add();
                Microsoft.Office.Interop.Word.Range tableRange = tableParagraph.Range;
                Microsoft.Office.Interop.Word.Table clientsTable = document.Tables.Add(tableRange, countClientInCategory + 1, 3);
                clientsTable.Borders.InsideLineStyle =
                clientsTable.Borders.OutsideLineStyle =
                Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;
                clientsTable.Range.Cells.VerticalAlignment =
                Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

                Microsoft.Office.Interop.Word.Range cellRange = clientsTable.Cell(1, 1).Range;
                cellRange.Text = "Код клиента";
                cellRange = clientsTable.Cell(1, 2).Range;
                cellRange.Text = "ФИО";
                cellRange = clientsTable.Cell(1, 3).Range;
                cellRange.Text = "E-mail";
                clientsTable.Rows[1].Range.Bold = 1;
                clientsTable.Rows[1].Range.ParagraphFormat.Alignment =
                Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;

                int j = 1;
                foreach (var currentClient in currentClients)
                {
                    cellRange = clientsTable.Cell(j + 1, 1).Range;
                    cellRange.Text = $"{currentClient.ClientCode}";
                    cellRange.ParagraphFormat.Alignment =
                    Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    cellRange = clientsTable.Cell(j + 1, 2).Range;
                    cellRange.Text = currentClient.FIO;
                    cellRange.ParagraphFormat.Alignment =
                    Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    cellRange = clientsTable.Cell(j + 1, 3).Range;
                    cellRange.Text = currentClient.Email;
                    cellRange.ParagraphFormat.Alignment =
                    Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;

                    j++;
                }
            }

        document.Words.Last.InsertBreak(Microsoft.Office.Interop.Word.WdBreakType.wdPageBreak);
        app.Visible = true;
        /*document.SaveAs2(@"C:\outputFileWord.docx");
        document.SaveAs2(@"C:\outputFilePdf.pdf",
        Microsoft.Office.Interop.Word.WdExportFormat.wdExportFormatPDF);*/
        }

        private void ChangeLR2Btn_Click(object sender, RoutedEventArgs e)
        {
            Grid_LR2.Visibility = Visibility.Visible;
            Grid_LR3.Visibility = Visibility.Hidden;
            ChangeLR2Btn.Background = Brushes.Bisque;
            ChangeLR3Btn.Background = Brushes.LightGray;
        }

        private void ChangeLR3Btn_Click(object sender, RoutedEventArgs e)
        {
            Grid_LR3.Visibility = Visibility.Visible;
            Grid_LR2.Visibility = Visibility.Hidden;
            ChangeLR2Btn.Background = Brushes.LightGray;
            ChangeLR3Btn.Background = Brushes.Bisque;
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            ChangeLR2Btn.Background = Brushes.Bisque;
            ChangeLR3Btn.Background = Brushes.LightGray;
            /*
            ImportJsonBtn.Click += async (o, el) =>
            {
                using (FileStream fs = new FileStream(@"C:\4course\3.json", FileMode.OpenOrCreate))
                {
                    List<ClientJson> person = await JsonSerializer.DeserializeAsync<List<ClientJson>>(fs);
                    MessageBox.Show(person[0].FullName);
                }

            };*/
        }
    }
}
