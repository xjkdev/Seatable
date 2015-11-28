﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Threading.Tasks;

using System.Reflection;
using Microsoft.Office.Interop.Excel;

namespace ExcelExperiment
{
    class Program
    {
        static void Main(string[] args)
        {
            //Random random = new Random();
            //int[] values = { 4, 6, 18, 2, 1, 76, 0, 3, 11 };
            //for(int i = 0; i < 100; i++)
            //{
            //    Console.WriteLine(random.Next(2));
            //}
            CreateWorkbook();
            Console.ReadKey();
        }

        static void CreateWorkbook()
        {
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            if (xlApp == null)
            {
                Console.WriteLine("EXCEL could not be started. Check that your office installation and project references are correct.");
                return;
            }
            xlApp.Visible = true;

            Workbook wb = xlApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            Worksheet ws = (Worksheet)wb.Worksheets[1];

            if (ws == null)
            {
                Console.WriteLine("Worksheet could not be created. Check that your office installation and project references are correct.");
            }

            // Select the Excel cells, in the range c1 to c7 in the worksheet.
            Range aRange = ws.get_Range("C1", "C7");

            if (aRange == null)
            {
                Console.WriteLine("Could not get a range. Check to be sure you have the correct versions of the office DLLs.");
            }

            // Fill the cells in the C1 to C7 range of the worksheet with the number 6.
            Object[] args = new Object[1];
            args[0] = 6;
            aRange.GetType().InvokeMember("Value", BindingFlags.SetProperty, null, aRange, args);

            // Change the cells in the C1 to C7 range of the worksheet to the number 8.
            aRange.Value2 = 8;
            Console.WriteLine(ws.get_Range("C2").Value);
            //Range b = ws.get_Range("A1");
            //b.ColumnWidth = 20;
        }
    }
}
