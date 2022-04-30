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
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private async void btnOpenFile_Click(object sender, RoutedEventArgs e)
        {
            openFilePanel.Visibility = Visibility.Hidden;

            #region EXCEL
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            string file = "";
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Title = "Odpri Excel datoteko";
            openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            if (openFileDialog.ShowDialog() == true)
            {
                file = openFileDialog.FileName;
            }
            #endregion

            var ep = new ExcelPackage(new FileInfo(file));
            var ws = ep.Workbook.Worksheets[0];

            MakeTable(ws.Dimension.Columns, ws.Dimension.Rows);

            #region MAKING LIST
            // Student groups are sublists of the "groupsDefault" list

            var groupsDefault = new List<List<string>>();
            for (int col = 1; col <= ws.Dimension.End.Column; col++)
            {
                var column = new List<string>();
                for (int rw = 1; rw <= ws.Dimension.End.Row; rw++)
                {
                    if (ws.Cells[rw, col].Value != null)
                        column.Add(ws.Cells[rw, col].Value.ToString());
                }
                groupsDefault.Add(column);
            }
            groupsDefault.ToArray();

            int smallestGroupSize = 100;
            foreach (var group in groupsDefault)
            {
                if (group.Count < smallestGroupSize) smallestGroupSize = group.Count;
            }
            #endregion

            Random rnd = new Random();

            List<string[]> groupsRandom = new List<string[]>();
            foreach (var group in groupsDefault)
            {
                string[] names = group.ToArray();
                names = names.OrderBy(x => rnd.Next()).ToArray();
                groupsRandom.Add(names);
            }

            int c = 0;
            for (int row = 0; row < MyGrid.RowDefinitions.Count; row++)
            {
                c = 0;
                foreach (var group in groupsRandom)
                {
                    await Task.Delay(500);
                    if (row == smallestGroupSize)
                    {
                        FillTableCell(row, c, "---");
                    }
                    else if ((row < smallestGroupSize && row >= group.Length) || (row > smallestGroupSize && row - 1 >= group.Length))
                    {
                        FillTableCell(row, c, "");
                    }
                    else
                    {
                        if (row < smallestGroupSize) FillTableCell(row, c, group[row]);
                        else FillTableCell(row, c, group[row-1]);
                    }
                    c++;
                }
            }
        }

        public void MakeTable(int columns, int rows)
        {
            for (int x = 0; x < columns; x++)
            {
                ColumnDefinition column = new ColumnDefinition();
                column.Width = new GridLength(1, GridUnitType.Star);
                MyGrid.ColumnDefinitions.Add(column);
            }
            for (int y = 0; y < rows + 1; y++)
            {
                RowDefinition r = new RowDefinition();
                r.Height = GridLength.Auto;
                MyGrid.RowDefinitions.Add(r);
            }
        }

        public void FillTableCell(int row, int column, string text)
        {
            TextBox tb = new TextBox();
            tb.HorizontalContentAlignment = HorizontalAlignment.Center;
            tb.VerticalContentAlignment = VerticalAlignment.Center;
            tb.Text = text;
            tb.FontSize = 20;
            if (row % 2 == 0)
            {
                tb.Background = Brushes.Orange;
            }
            Grid.SetColumn(tb, column);
            Grid.SetRow(tb, row);
            MyGrid.Children.Add(tb);
        }
    }
}