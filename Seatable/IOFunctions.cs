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

    public partial class MainWindow : System.Windows.Window
    {
        Workbook WB;
        //TextFiles
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

        private string ReadFile(string path)
        {
            try
            {
                using (StreamReader sr = new StreamReader(path, encoding: Encoding.UTF8))
                {
                    String line = sr.ReadToEnd();
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

        private async Task<bool> WriteFileAsync(string path, string[] lines)
        {
            using (StreamWriter file = new StreamWriter(path, false, Encoding.UTF8))
            {
                foreach (string line in lines)
                {
                    await file.WriteLineAsync(line);
                }
            }
            return true;
        }

        private bool WriteFile(string path, string[] lines)
        {
            using (StreamWriter file = new StreamWriter(path, false, Encoding.UTF8))
            {
                foreach (string line in lines)
                {
                    file.WriteLine(line);
                }
            }
            return true;
        }

        //DebugOutput
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
            foreach (var i in a)
            {
                result += i.Key.Key + '\t' + i.Key.Value + "\t\t";
                result += i.Value.Key + '\t' + i.Value.Value + '\n';
            }
            return result;
        }

        //ExcelIO
        private void XLAinit()
        {
            xlApp = new Microsoft.Office.Interop.Excel.Application();
            //xlApp.WorkbookAfterSave += XlApp_WorkbookAfterSave;
            xlApp.WorkbookBeforeSave += XlApp_WorkbookBeforeSave;
            xlApp.WorkbookBeforeClose += XlApp_WorkbookBeforeClose;
            //isquit = false;            
            WB = xlApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            WB.Worksheets.Add();
            Worksheet ws = (Worksheet)WB.Worksheets[1];
            if (ws == null)
            {
                MessageBox.Show("1:Worksheet could not be created. Check that your office installation and project references are correct.");
                return;
            }
            Firstset(ws);
            ws = (Worksheet)WB.Worksheets[2];
            Firstset(ws);
        }



        private bool CreateExcelDocument(DataTable table)
        {
            if (table == null)
                return false;

            var exceptionfile = ReadFile(Exceptionfilename);
            List<string> exception = new List<string>();
            exception.AddRange(exceptionfile.Split('\n'));
            exception.Add("\u8C22\u4FCA\u7428");


            if (xlApp == null)
            {
                MessageBox.Show("EXCEL could not be started. Check that your office installation and project references are correct.");
                return false;
            }
            xlApp.Visible = false;
            Worksheet ws = (Worksheet)WB.Worksheets[1];
            makesheet(table, ws, $"{DateTime.Now.Month}月{DateTime.Now.Day}日");
            ws = (Worksheet)WB.Worksheets[2];
            makesheet(ChangeGroup(table), ws, $"{DateTime.Now.AddDays(14).Month}月{DateTime.Now.AddDays(14).Day}日");
            xlApp.Visible = true;
            WB.Activate();
            return true;

        }

        private void Writeintowb(Workbook wb)
        {         
            leaders = Pickleaders(this.a);
            Worksheet ws = (Worksheet)wb.Worksheets[1];
            Firstset(ws, false);
            makesheet(this.a, ws, $"{DateTime.Now.Month}月{DateTime.Now.Day}日");
            DataTable newtable;
            newtable = ChangeGroup(this.a);
            ws = (Worksheet)wb.Worksheets[2];
            Firstset(ws, false);
            makesheet(newtable, ws, $"{DateTime.Now.AddDays(14).Month}月{DateTime.Now.AddDays(14).Day}日");
        }

        private void Firstset(Worksheet ws,bool completely = true)
        {
            Range ab = ws.get_Range("A1", "J9");
            if (ab != null)
            {
                if (completely)
                {
                    ab.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    ab.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    ab.RowHeight = 50;
                    ab.ColumnWidth = 13.3;
                    ab.Font.Size = 20;
                }
                ab.Font.Bold = 0;
            }

            if (completely)
            {
                ab = ws.get_Range("A8", "B8");
                if (ab != null)
                {
                    ab.Merge();
                    ab.Value = "第一组";
                }

                ab = ws.get_Range("C8", "D8");
                if (ab != null)
                {
                    ab.Merge();
                    ab.Value = "第二组";
                }

                ab = ws.get_Range("E8", "F8");
                if (ab != null)
                {
                    ab.Merge();
                    ab.Value = "第三组";
                }

                ab = ws.get_Range("G8", "H8");
                if (ab != null)
                {
                    ab.Merge();
                    ab.Value = "第四组";
                }

                ab = ws.get_Range("I8", "J8");
                if (ab != null)
                {
                    ab.Merge();
                    ab.Value = "第五组";
                }

                ab = ws.get_Range("A9", "J9");
                if (ab != null)
                {
                    ab.Merge();
                    ab.Value = "讲台";
                }

            }
        }

        private bool makesheet(DataTable table, Worksheet ws, string Name)
        {
            Range a;
            ws.Name = Name;
            foreach (DataRow i in table.Rows)
            {
                for (int j = 0; j < i.ItemArray.Length; j++)
                {
                    string range = table.Columns[j].Caption + (int)(table.Rows.IndexOf(i) + 1);
                    a = ws.get_Range(range);
                    if (a != null)
                    {
                        a.Value = i[j].ToString();
                        if (leaders.Contains(i[j].ToString()))
                            a.Font.Bold = 1;
                    }
                }
            }
            return true;
        }
    }
}