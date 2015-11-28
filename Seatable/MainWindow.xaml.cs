using System;
using System.Collections.Generic;
using System.IO;
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
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.Office.Interop.Excel;

namespace Seatable
{
    using PeopleList = Dictionary<String, Gender>;
    using People = KeyValuePair<String, Gender>;
    using Desk = KeyValuePair<KeyValuePair<String, Gender>, KeyValuePair<String, Gender>>;
    using System.Data;
    using System.Reflection;

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

        DataTable a;
        
        public MainWindow()
        {
            InitializeComponent();
        }

        private async void abutton_Click(object sender, RoutedEventArgs e)
        {
            Task<String> waitfortext = ReadFileAsync("list.txt");
            String text = await waitfortext;
            PeopleList peoplelist = TextToPeopleList(text);
            List<Desk> desklist = MakeDesk(peoplelist);
            var rows = MakeRow(desklist);
            rows.Reverse();

            this.a = new DataTable();

            for (int i = 0; i < rows.Count; i++)
            {
                List<string> tmp = new List<string>();
                foreach(var j in rows[i])
                {
                    tmp.Add(j.Key.Key);
                    tmp.Add(j.Value.Key);
                }
                while (this.a.Columns.Count < tmp.Count)
                {
                    this.a.Columns.Add(new DataColumn(((char)((char)'A'+this.a.Columns.Count)).ToString(),typeof(string)));
                    //a.Columns.Add(new DataColumn())
                }
                if(tmp.Count >= 0 )
                    this.a.Rows.Add(tmp.ToArray());
            }


            dataGrid.ItemsSource = this.a.AsDataView();

        }

        private List<List<Desk>> MakeRow(List<Desk> desklist)
        {
            List<List<Desk>> result = new List<List<Desk>>();
            while(desklist.Count >=5 )
            {                
                result.Add(desklist.GetRange(0, 5));
                desklist.RemoveRange(0, 5);
            }
            Desk emptydesk = new Desk();
            int left = 0, right = 0;
            switch (desklist.Count)
            {
                case 1:
                    left = 2;
                    right = 2;
                    break;
                case 2:
                    left = 2;
                    right = 1;
                    break;
                case 3:
                    left = 1;
                    right = 1;
                    break;
                case 4:
                    left = 0;
                    right = 1;
                    break;
                case 0:
                    return result;
                default:
                    return result;
            }
            List<Desk> tmp = MakeRepeatElementList(emptydesk, left);
            tmp.AddRange(desklist);
            tmp.AddRange(MakeRepeatElementList(emptydesk, right));
            result.Add(tmp);
            return result;
        }

        private List<List<T>> Rotate<T> (List<List<T>> list)
        {
            List<List<T>> result = new List<List<T>>();
            for(int i=0; i < list.Count; i++)
            {
                
                for(int j = 0; j < list[i].Count; j++)
                {
                    if (result.Count < list[i].Count)
                        result.Add(new List<T>());
                    result[j].Add(list[i][j]);
                }
            }
            return result;
        }

        private List<T> MakeRepeatElementList<T> (T element,int count)
        {
            List<T> result = new List<T>();
            for(int i=0;i< count; i++)
            {
                result.Add(element);
            }
            return result;
        }

        private List<Desk> MakeDesk(PeopleList peoplelist)//男女混坐同桌
        {
            PeopleList malelist = ToPeopleList(peoplelist.Where(x => x.Value == Gender.Male));
            PeopleList femalelist = ToPeopleList(peoplelist.Where(x => x.Value == Gender.Femele));
            Shuffle(ref malelist);
            Shuffle(ref femalelist);
            var desklist = new List<Desk>();
            while(malelist.Count>0 || femalelist.Count > 0)
            {

                People m = new People();
                People f = new People();
                if (femalelist.Count > 0)
                {
                    f = femalelist.First();
                    femalelist.Remove(f.Key);
                }
                else if (malelist.Count > 0)
                {
                    f = malelist.First();
                    malelist.Remove(f.Key);
                }

                if (malelist.Count > 0)
                {
                    m = malelist.First();
                    malelist.Remove(m.Key);
                }
                else if(femalelist.Count > 0)
                {
                    m = femalelist.First();
                    femalelist.Remove(m.Key);
                }
                if (desklist.Count % 5 == 0)
                {
                    var tmp = malelist;
                    malelist = femalelist;
                    femalelist = tmp;
                }
                desklist.Add(new Desk(f,m));

            }
            
            return desklist;
        }

        

        private async Task<string> ReadFileAsync(string path)
        {
            try
            {
                using (StreamReader sr = new StreamReader(path, encoding: Encoding.UTF8))
                {
                    String line = await sr.ReadToEndAsync();
                    return line;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Could not read the file\n" +
                    "Error:" + ex.GetType());
                return String.Empty;
            }
        }        

        private String printPeoplelist(PeopleList a)
        {
            String result = "";
            foreach (var i in a)
            {
                result += i.Key + '\t' + i.Value + '\n';
            }
            return result;
        }

        private String printDeskList(List<Desk> a)
        {
            String result = "";
            foreach(var i in a)
            {
                result += i.Key.Key + '\t' + i.Key.Value + "\t\t";
                result += i.Value.Key + '\t' + i.Value.Value + '\n';       
            }
            return result;
        }

        private Dictionary<TKey,TValue> Shuffle<TKey, TValue>(ref Dictionary<TKey, TValue> l)
        {
            Random random = new Random();
            l = l.OrderBy(x => random.Next()).ToDictionary(x=>x.Key,x=>x.Value);
            return l;
        }

        private PeopleList TextToPeopleList(String text)
        {
            Dictionary<String, Gender> result = new PeopleList();
            if (String.IsNullOrEmpty(text))
            {
                return result;
            }
            var lines = text.Split('\n');
            foreach (var line in lines)
            {
                String[] tmp = line.Split('\t');
                result.Add(
                    tmp?[0] ?? "Null",
                    (tmp?[1].Equals("男") ?? true) ? Gender.Male : Gender.Femele);
            }
            return result;
        }

        private PeopleList ToPeopleList(IEnumerable<People> ie)
        {
            return ie.ToDictionary(t => t.Key, t => t.Value);
        }

        private async Task<DataTable> ChangeGroup(DataTable table)
        {
            var result = new DataTable();
            await Task.Run(() =>
            {
                while (result.Columns.Count < table.Columns.Count)
                {
                    result.Columns.Add(new DataColumn(((char)((char)'A' + result.Columns.Count)).ToString(), typeof(string)));                    
                }
                for (int i = 3; i < table.Rows.Count; i++)
                {
                    result.Rows.Add(table.Rows[i].ItemArray);
                }
                for (int i = 0; i < 3; i++)
                {
                    result.Rows.Add(table.Rows[i].ItemArray);
                }

            });
            return result;
        }

        private void ChangeColumns(ref DataTable obj,int a,int b)
        {
            foreach(DataRow row in obj.Rows)
            {
                var tmp = row.ItemArray[a];
                row.ItemArray[a] = row.ItemArray[b];
                row.ItemArray[b] = tmp;
            }
        }

        private async void exportButton_Click(object sender, RoutedEventArgs e)
        {
            progressbar1.Visibility = Visibility.Visible;
            exportButton.IsEnabled = false;
            exportButton.Content = "exporting";
            await CreateExcelDocument(this.a);
            progressbar1.Visibility = Visibility.Hidden;
            exportButton.IsEnabled = true;
            exportButton.Content = "export";
        }


        private async Task<bool> CreateExcelDocument (DataTable table)
        {
            if (table == null)
                return false;
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            if (xlApp == null)
            {
                MessageBox.Show("EXCEL could not be started. Check that your office installation and project references are correct.");
                return false;
            }

            xlApp.Visible = false;
            Workbook wb = xlApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            wb.Worksheets.Add();
            xlApp.Caption = $" {DateTime.Now.Month}月{DateTime.Now.Day}日";
            Worksheet ws = (Worksheet)wb.Worksheets[1];
            await makesheet(table, ws, xlApp, $"{DateTime.Now.Month}月{DateTime.Now.Day}日");
            DataTable newtable;
            newtable = await ChangeGroup(table);            
            ws = (Worksheet)wb.Worksheets[2];
            await makesheet(newtable, ws, xlApp, $"{DateTime.Now.AddDays(14).Month}月{DateTime.Now.AddDays(14).Day}日");
            xlApp.Visible = true;
            return true;

        }
        private async Task<bool> makesheet(DataTable table, Worksheet ws, Microsoft.Office.Interop.Excel.Application xlApp,String Name)
        {
            return await Task.Run(()=>
            {            
            ws.Name = Name;
            if (ws == null)
            {
                MessageBox.Show("Worksheet could not be created. Check that your office installation and project references are correct.");
                return false;
            }

            foreach (DataRow i in table.Rows)
            {
                for (int j = 0; j < i.ItemArray.Length; j++)
                {
                    string range = table.Columns[j].Caption + (int)(table.Rows.IndexOf(i) + 1);
                    Range a = ws.get_Range(range);
                    if (a != null)
                    {
                        a.Value = i[j].ToString();
                        a.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                        a.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                        a.RowHeight = 50;
                        a.ColumnWidth = 13.3;
                        a.Font.Size = 20;
                    }
                }
            }
            Range ab = ws.get_Range("A8", "B8");
            if (ab != null)
            {
                ab.Merge();
                ab.Value = "第一组";
                ab.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                ab.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                ab.RowHeight = 50;
                ab.Font.Size = 20;
            }

            ab = ws.get_Range("C8", "D8");
            if (ab != null)
            {
                ab.Merge();
                ab.Value = "第二组";
                ab.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                ab.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                ab.RowHeight = 50;
                ab.Font.Size = 20;
            }

            ab = ws.get_Range("E8", "F8");
            if (ab != null)
            {
                ab.Merge();
                ab.Value = "第三组";
                ab.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                ab.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                ab.RowHeight = 50;
                ab.Font.Size = 20;
            }

            ab = ws.get_Range("G8", "H8");
            if (ab != null)
            {
                ab.Merge();
                ab.Value = "第四组";
                ab.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                ab.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                ab.RowHeight = 50;
                ab.Font.Size = 20;
            }

            ab = ws.get_Range("I8", "J8");
            if (ab != null)
            {
                ab.Merge();
                ab.Value = "第五组";
                ab.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                ab.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                ab.RowHeight = 50;
                ab.Font.Size = 20;
            }

            ab = ws.get_Range("A9", "J9");
            if (ab != null)
            {
                ab.Merge();
                ab.Value = "讲台";
                ab.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                ab.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                ab.RowHeight = 50;
                ab.Font.Size = 20;
            }
            
            return true;
        });
        }
    }
}
