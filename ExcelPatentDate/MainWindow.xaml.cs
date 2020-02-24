using Microsoft.Win32;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Xml;

namespace ExcelPatentDate
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void BtnImportExcel_Click(object sender, RoutedEventArgs e)
        {

        }

        private void btnOpenFile_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            if (openFileDialog.ShowDialog() == true)
            {   //txtEditor.Text = File.ReadAllText(openFileDialog.FileName);
                // path to your excel file
                //string path = "C:/****/sample_data.xlsx";
                FileInfo fileInfo = new FileInfo(openFileDialog.FileName);

                ExcelPackage package = new ExcelPackage(fileInfo);
                ExcelWorksheet worksheet = package.Workbook.Worksheets[2];

                // get number of rows and columns in the sheet
                int rows = worksheet.Dimension.Rows; // 20
                int columns = worksheet.Dimension.Columns; // 7

                // loop through the worksheet rows and columns
                //for (int i = 1; i <= rows; i++)
                //{
                //    for (int j = 1; j <= columns; j++)
                //    {
                //        try
                //        {
                //            txtEditor.Text = worksheet.Cells[i, j].Value == null ? string.Empty : worksheet.Cells[i, j].Value.ToString() ;

                //        }
                //        catch (Exception ex)
                //        {

                //            throw;
                //        }
                //        /* Do something ...*/
                //    }
                //}
                bool hasHeader = true;
                DataTable tbl = new DataTable();
                foreach (var firstRowCell in worksheet.Cells[1, 1, 1, worksheet.Dimension.End.Column])
                {
                    tbl.Columns.Add(hasHeader ? firstRowCell.Text : string.Format("Column {0}", firstRowCell.Start.Column));
                }
                tbl.Columns.Add("Patent Link for Google");
                var startRow = hasHeader ? 2 : 1;
                for (int rowNum = startRow; rowNum <= worksheet.Dimension.End.Row; rowNum++)
                {
                    var wsRow = worksheet.Cells[rowNum, 1, rowNum, worksheet.Dimension.End.Column];
                    int i = wsRow.Hyperlink.AbsoluteUri.IndexOf("/patent");
                    if (i >= 0)
                    {
                        string a = wsRow.Hyperlink.AbsoluteUri.Substring(i + 1);
                        tbl.Columns[tbl.Columns.Count - 1].DefaultValue = "https://patents.google.com/" + a.Replace("_", "");
                        HttpClient hc = new HttpClient();
                        //HttpClient hc = new HttpClient();
                        var response = await hc.GetByteArrayAsync(b);
                        string source = Encoding.GetEncoding("utf-8").GetString(response, 0, response.Length - 1);
                        source = WebUtility.HtmlDecode(source);
                        XmlDocument test = new XmlDocument();
                        test.LoadXml(source);
                    }
                    DataRow row = tbl.Rows.Add();
                    foreach (var cell in wsRow)
                    {
                        row[cell.Start.Column - 1] = cell.Text;
                    }
                }

                excelGrid.ItemsSource = tbl.DefaultView;
            }
        }


    }
}
