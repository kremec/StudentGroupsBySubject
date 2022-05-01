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
            // Disabling the button to open files
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

            var excelFile = new ExcelPackage(new FileInfo(file));
            var sheet0 = excelFile.Workbook.Worksheets[0];
            var sheet1 = excelFile.Workbook.Worksheets[1];
            #endregion

            MakeTable(sheet0.Dimension.Columns, sheet0.Dimension.Rows);

            #region MAKING EXCEPTION GROUPS LIST
            var xGroups = new List<List<string>>();
            for (int col = 1; col <= sheet1.Dimension.End.Column; col++)
            {
                var column = new List<string>();
                for (int rw = 1; rw <= sheet1.Dimension.End.Row; rw++)
                {
                    if (sheet1.Cells[rw, col].Value != null)
                        column.Add(sheet1.Cells[rw, col].Value.ToString());
                }
                xGroups.Add(column);
            }
            xGroups.ToArray();

            /*
            foreach (var group in xGroups)
            {
                string output = "";
                foreach (var name in group)
                {
                    output += name + ", ";
                }
                MessageBox.Show(output);
            }
            */
            #endregion

            #region MAKING LIST
            // Student groups are sublists of the "groupsDefault" list

            var groupsDefault = new List<List<string>>();
            for (int col = 1; col <= sheet0.Dimension.End.Column; col++)
            {
                var column = new List<string>();
                for (int rw = 1; rw <= sheet0.Dimension.End.Row; rw++)
                {
                    if (sheet0.Cells[rw, col].Value != null)
                        column.Add(sheet0.Cells[rw, col].Value.ToString());
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

            #region SHUFFLING GROUPS OF NAMES
            List<string[]> groupsRandom = new List<string[]>();

            ShuffleAgain:
            groupsRandom.Clear();
            Random rnd = new Random();
            foreach (var group in groupsDefault)
            {
                string[] names = group.ToArray();
                names = names.OrderBy(x => rnd.Next()).ToArray();
                groupsRandom.Add(names);
            }
            #endregion

            #region HANDLING EXCEPTIONS

            //Za vsak exception group
            foreach (var xGroup in xGroups)
            {
                string[] tempXGroup = xGroup.ToArray();

                // Za vsako ime v exceptions grupi
                for (int indexOfExceptionNameInException = 0; indexOfExceptionNameInException < tempXGroup.Length; indexOfExceptionNameInException++)
                {
                    //Preverimo, če je exception ime v zmešani grupi
                    int indexOfExceptionNameInShuffled = -1;
                    foreach (var group in groupsRandom)
                    {
                        if (group.Contains(tempXGroup[indexOfExceptionNameInException]))
                        {
                            indexOfExceptionNameInShuffled = Array.IndexOf(group, tempXGroup[indexOfExceptionNameInException]);

                            // Preverimo znotraj ostalih grup s tem indeksom, če je kak exception znotraj iste grupe
                            foreach (var g in groupsRandom)
                            {
                                if (g.Length > indexOfExceptionNameInShuffled && tempXGroup.Contains(g[indexOfExceptionNameInShuffled]) && g[indexOfExceptionNameInShuffled] != tempXGroup[indexOfExceptionNameInException])
                                {
                                    //MessageBox.Show("Exception group found in group with " + g[indexOfExceptionNameInShuffled] + " and " + tempXGroup[indexOfExceptionNameInException]);

                                    // Resolvamo exception
                                    goto ShuffleAgain;
                                }
                            }
                            break;
                        }
                    }
                }
            }
            #endregion

            #region FILLING OUT THE GRID
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
            #endregion
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