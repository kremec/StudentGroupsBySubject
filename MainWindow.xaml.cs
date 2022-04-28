using Microsoft.Win32;
using System;
using System.Windows;
using System.IO;
using OfficeOpenXml;
using System.Collections.Generic;
using System.Windows.Controls;
using System.Linq;
using System.Windows.Media;

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

        private void btnOpenFile_Click(object sender, RoutedEventArgs e)
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
            int r = 0;
            int c = 0;
            foreach (var group in groups)
            {
                string[] names = group.ToArray();
                names = names.OrderBy(x => rnd.Next()).ToArray();

                foreach (string name in names)
                {
                    if (c == smallestGroupSize)
                    {
                        if (c % 2 == 0)
                        {
                            Border br = new Border();
                            br.Background = Brushes.FloralWhite;
                            Grid.SetColumn(br, r);
                            Grid.SetRow(br, c);
                            MyGrid.Children.Add(br);
                        }
                        TextBlock txb = new TextBlock();
                        txb.Text = "---";
                        Grid.SetColumn(txb, r);
                        Grid.SetRow(txb, c);
                        MyGrid.Children.Add(txb);
                        c++;
                    }

                    if (c % 2 == 0)
                    {
                        Border br = new Border();
                        br.Background = Brushes.FloralWhite;
                        Grid.SetColumn(br, r);
                        Grid.SetRow(br, c);
                        MyGrid.Children.Add(br);
                    }
                    TextBlock tb = new TextBlock();
                    tb.Text = name;
                    Grid.SetColumn(tb, r);
                    Grid.SetRow(tb, c);
                    MyGrid.Children.Add(tb);

                    c++;
                }
                c = 0;
                r++;
            }

            //for (int x = 0; x < ws.Dimension.Columns; x++)
            //{
            //    for (int y = 0; y < ws.Dimension.Rows; y++)
            //    {
            //        TextBox tb = new TextBox();
            //        tb.Text = "my text for " + x + " " + y;
            //        Grid.SetColumn(tb, x);
            //        Grid.SetRow(tb, y);
            //        MyGrid.Children.Add(tb);
            //    }
            //}
        }

        public void MakeTable(int columns, int rows)
        {
            for (int x = 0; x < columns; x++)
                MyGrid.ColumnDefinitions.Add(new ColumnDefinition());
            for (int y = 0; y < rows + 1; y++)
            {
                RowDefinition r = new RowDefinition();
                r.Height = GridLength.Auto;
                MyGrid.RowDefinitions.Add(r);
            }


        }
    }
}
