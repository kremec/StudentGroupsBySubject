using Microsoft.Win32;
using System;
using System.Windows;
using System.IO;
using OfficeOpenXml;
using System.Collections.Generic;
using System.Windows.Controls;
using System.Linq;
using System.Windows.Media;
using System.Threading.Tasks;

namespace WPF_Test
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

        private async void btnOpenFile_Click(object sender, RoutedEventArgs e)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            string file = "";
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Title = "Odpri Excel datoteko";
            openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            if (openFileDialog.ShowDialog() == true)
            {
                file = openFileDialog.FileName;
            }

            openFilePanel.Visibility = Visibility.Hidden;

            var ep = new ExcelPackage(new FileInfo(file));
            var ws = ep.Workbook.Worksheets["Sheet1"];

            MakeTable(ws.Dimension.Columns, ws.Dimension.Rows);

            var groups = new List<List<string>>();

            for (int col = 1; col <= ws.Dimension.End.Column; col++)
            {
                var column = new List<string>();
                for (int rw = 1; rw <= ws.Dimension.End.Row; rw++)
                {
                    if (ws.Cells[rw, col].Value != null)
                        column.Add(ws.Cells[rw, col].Value.ToString());
                }
                groups.Add(column);
            }
            groups.ToArray();

            int smallestGroupSize = 100;
            foreach (var group in groups)
            {
                if (group.Count < smallestGroupSize) smallestGroupSize = group.Count;
            }

            Random rnd = new Random();
            int tempr = 0;
            int r = 0;
            int c = 0;
            foreach (var group in groups)
            {
                string[] names = group.ToArray();
                names = names.OrderBy(x => rnd.Next()).ToArray();

                for (int i = 0; i <= ws.Dimension.Rows; i++)
                {
                    await Task.Delay(500);
                    if (i >= names.Length)
                    {
                        FillTableCell(r, c, "");
                    }
                    else
                    {
                        if (r == smallestGroupSize)
                        {
                            FillTableCell(r, c, "---");
                            r++;
                            await Task.Delay(500);
                        }
                        FillTableCell(r, c, names[i]);
                    }
                    r++;
                }
                tempr = r;
                r = 0;
                c++;
            }
            await Task.Delay(500);
            FillTableCell(tempr, c, "");
        }

        public void MakeTable(int columns, int rows)
        {
            for (int x = 0; x < columns; x++)
            {
                MyGrid.ColumnDefinitions.Add(new ColumnDefinition());
            }
            for (int y = 0; y < rows + 2; y++)
            {
                RowDefinition r = new RowDefinition();
                r.Height = GridLength.Auto;
                MyGrid.RowDefinitions.Add(r);
            }
        }

        public async void FillTableCell(int row, int column, string text)
        {
            TextBox tb = new TextBox();
            //tb.Visibility = Visibility.Hidden;
            //tb.Visibility = Visibility.Visible;
            tb.Text = text;
            tb.FontSize = 20;
            if (row % 2 == 0)
            {
                tb.Background = Brushes.FloralWhite;
            }
            Grid.SetColumn(tb, column);
            Grid.SetRow(tb, row);
            MyGrid.Children.Add(tb);
        }
    }
}
