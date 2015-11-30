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
        private List<List<Desk>> MakeRow(List<Desk> desklist)
        {
            List<List<Desk>> result = new List<List<Desk>>();
            while (desklist.Count >= 5)
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

        private List<List<T>> Rotate<T>(List<List<T>> list)
        {
            List<List<T>> result = new List<List<T>>();
            for (int i = 0; i < list.Count; i++)
            {

                for (int j = 0; j < list[i].Count; j++)
                {
                    if (result.Count < list[i].Count)
                        result.Add(new List<T>());
                    result[j].Add(list[i][j]);
                }
            }
            return result;
        }

        private List<T> MakeRepeatElementList<T>(T element, int count)
        {
            List<T> result = new List<T>();
            for (int i = 0; i < count; i++)
            {
                result.Add(element);
            }
            return result;
        }

        private List<Desk> MakeDesk(PeopleList peoplelist)//男女混坐同桌
        {
            string[] lasttime = ReadFile(DeskHistoryfilename).Split('\n');

            makerowstart:

            PeopleList malelist = ToPeopleList(peoplelist.Where(x => x.Value == Gender.Male));
            PeopleList femalelist = ToPeopleList(peoplelist.Where(x => x.Value == Gender.Femele));
            Shuffle(ref malelist);
            Shuffle(ref femalelist);
            var desklist = new List<Desk>();
            while (malelist.Count > 0 || femalelist.Count > 0)
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
                else if (femalelist.Count > 0)
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
                if (isSameaslasttime(lasttime, f.Key, m.Key))
                    goto makerowstart;
                desklist.Add(new Desk(f, m));
            }

            return desklist;
        }

        private bool isSameaslasttime(string[] lasttime, string a, string b)
        {
            foreach (var i in lasttime)
            {
                string[] lt = i.Split('\t');

                if (lt.Count() == 2)
                {
                    if (lt[0].Equals(a) && lt[1].Equals(b) || lt[0].Equals(b) && lt[1].Equals(a))
                        return true;
                }
            }
            return false;
        }

        private string[] Pickleaders(DataTable table, string[] eception)
        {
            List<string> res = new List<string>();
            Random random = new Random();
            for (int i = 0; i < table.Columns.Count; i += 2)
            {
                List<string> tmp = new List<string>();
                foreach (DataRow row in table.Rows)
                {
                    if (row[i] != System.DBNull.Value)
                        tmp.Add((string)row[i]);
                    if (row[i + 1] != System.DBNull.Value)
                        tmp.Add((string)row[i + 1]);
                }
                string leader = tmp[random.Next(tmp.Count)];
                if (eception.Contains(leader))
                    i -= 2;
                else
                    res.Add(leader);
            }
            return res.ToArray();
        } 

        private Dictionary<TKey, TValue> Shuffle<TKey, TValue>(ref Dictionary<TKey, TValue> l)
        {
            Random random = new Random();
            l = l.OrderBy(x => random.Next()).ToDictionary(x => x.Key, x => x.Value);
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

        private DataTable ChangeGroup(DataTable table)
        {
            var result = new DataTable();

            while (result.Columns.Count < table.Columns.Count)
            {
                result.Columns.Add(new DataColumn(((char)((char)'A' + result.Columns.Count)).ToString(), typeof(string)));
            }
            for (int i = 4; i < table.Rows.Count; i++)
            {
                result.Rows.Add(table.Rows[i].ItemArray);
            }
            for (int i = 0; i < 4; i++)
            {
                result.Rows.Add(table.Rows[i].ItemArray);
            }
            for (int i = result.Rows.Count - 1; i > 0; i--)
            {
                for (int j = 0; j < result.Columns.Count; j++)
                {
                    if (result.Rows[i][j] == System.DBNull.Value)
                    {
                        result.Rows[i][j] = result.Rows[i - 1][j];
                        result.Rows[i - 1][j] = System.DBNull.Value;
                    }
                }
            }
            for (int i = 1; i < result.Rows.Count; i++)
            {
                ExchangeColumns(result.Rows[i], 0, 2);
                ExchangeColumns(result.Rows[i], 4, 6);
                ExchangeColumns(result.Rows[i], 4, 8);
            }

            return result;
        }

        private void ExchangeColumns(DataRow obj, int a, int b)
        {
            var tmp = obj[a];
            var tmp1 = obj[a + 1];
            obj[a] = obj[b];
            obj[a + 1] = obj[b + 1];
            obj[b] = tmp;
            obj[b + 1] = tmp1;
        }

        private async Task Readlist()
        {
            Task<String> waitfortext = ReadFileAsync(Namelistfilename);
            String text = await waitfortext;
            PeopleList peoplelist = TextToPeopleList(text);
            desklist = MakeDesk(peoplelist);        
            var rows = MakeRow(desklist);

            rows.Reverse();

            this.a = new DataTable();

            for (int i = 0; i < rows.Count; i++)
            {
                List<string> tmp = new List<string>();
                foreach (var j in rows[i])
                {
                    tmp.Add(j.Key.Key);
                    tmp.Add(j.Value.Key);
                }
                while (this.a.Columns.Count < tmp.Count)
                {
                    this.a.Columns.Add(new DataColumn(((char)((char)'A' + this.a.Columns.Count)).ToString(), typeof(string)));
                    //a.Columns.Add(new DataColumn())
                }
                if (tmp.Count >= 0)
                    this.a.Rows.Add(tmp.ToArray());
            }
        }
    }
}