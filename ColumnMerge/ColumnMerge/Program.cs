using ColumnMerge;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using X15 = DocumentFormat.OpenXml.Office2013.Excel;

namespace ColumnMergess
{
    public class Program
    {
        public static void Main(string[] args)
        {
            for (int i = 100001; i < 100010; i++)
            {
                GenerateExcel.GenerateOrderDetailsExcel(i.ToString(), 
                                                        @"C:\temp\orders.xlsx",
                                                        "OrderNumber,Part #",
                                                        0,
                                                        @"C:\temp\result"+i.ToString() +".xlsx"
                                                        );

            }
            Console.WriteLine();
        }
    }
}