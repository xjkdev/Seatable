using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Microsoft.Office.Interop.Excel;


namespace Seatable
{
    using PeopleList = Dictionary<String, Gender>;
    using People = KeyValuePair<String, Gender>;
    using Desk = KeyValuePair<KeyValuePair<String, Gender>, KeyValuePair<String, Gender>>;
    using System.Data;
    

    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    enum Gender
    {
        Male,
        Femele
    }
    public partial class MainWindow : System.Windows.Window
    {

        Microsoft.Office.Interop.Excel.Application xlApp;
        const string Exceptionfilename = "exception.txt";
        const string Namelistfilename = "list.txt";
        const string DeskHistoryfilename = "Desklist.txt";
        //bool autosaving;
        DataTable a;
        List<Desk> desklist;
        string[] leaders;
        
        public MainWindow()
        {
            InitializeComponent();
            
        }

        private async void abutton_Click(object sender, RoutedEventArgs e)
        {
            await Readlist();
            dataGrid.ItemsSource = this.a.AsDataView();
        }

        

        private async void exportButton_Click(object sender, RoutedEventArgs e)
        {
            progressbar1.Visibility = Visibility.Visible;
            exportButton.IsEnabled = false;
            exportButton.Content = "正在导出";
            await Task.Run(() =>
            {
                XLAinit();
            });
            await Task.Run(() => CreateExcelDocument(this.a));
            xlApp.Visible = true;
            progressbar1.Visibility = Visibility.Hidden;
            exportButton.IsEnabled = false;
            exportButton.Content = "导出";
        }

        private void setExceptionButton_Click(object sender, RoutedEventArgs e)
        {
            System.Diagnostics.Process p = new System.Diagnostics.Process();
            p.StartInfo.FileName = "notepad";
            p.StartInfo.Arguments = Exceptionfilename;
            p.StartInfo.CreateNoWindow = false;
            p.Start();
        }

       

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {

            //if (!xlApp?.Visible ?? false)
            //{
            //    xlApp.Quit();
            //}
            //else if (!isquit)
            //{
            //    MessageBox.Show("请先关闭Excel");
            //    e.Cancel = true;
            //}
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            progressbar1.Visibility = Visibility.Hidden;
            exportButton.IsEnabled = true;
            exportButton.Content = "导出";
        }

        private void XlApp_WorkbookBeforeSave(Workbook Wb, bool SaveAsUI, ref bool Cancel)
        {
            Worksheet ws = Wb.Worksheets[1];
            foreach (DataRow i in this.a.Rows)
            {
                for (int j = 0; j < i.ItemArray.Length; j++)
                {
                    string range = this.a.Columns[j].Caption + (int)(this.a.Rows.IndexOf(i) + 1);
                    Range a = ws.get_Range(range);
                    if (a != null)
                    {
                        i[j] = a.Value;
                    }
                }
            }
            List<string> desklines = new List<string>();
            foreach (DataRow i in this.a.Rows)
            {
                for (int j = 0; j < i.ItemArray.Length; j += 2)
                {
                    string a = string.Empty, b = string.Empty;
                    if (i[j] != System.DBNull.Value)
                        a = (string)i[j];
                    if (i[j + 1] != System.DBNull.Value)
                        b = (string)i[j + 1];
                    if (!string.IsNullOrEmpty(a + b))
                        desklines.Add($"{a}\t{b}");
                }
            }
            
            WriteFile(DeskHistoryfilename, desklines.ToArray());
            Writeintowb(Wb);
        }

        private void XlApp_WorkbookBeforeClose(Workbook Wb, ref bool Cancel)
        {
            //isquit = true;
            //exportButton.IsEnabled = true;
            exportButton.Dispatcher.Invoke(() => { exportButton.IsEnabled = true; });
        }

        private void Window_SizeChanged(object sender, SizeChangedEventArgs e)
        {

        }
    }
}
