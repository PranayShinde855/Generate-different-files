using System;
using Microsoft.Office.Interop;

namespace Genearting_excel_file_in_console_application
{
    static class Program
    {
        public static void GenerateExcel()
        {
            string fileName, Id, Name;
            Console.WriteLine("Enter file name :");
            fileName = Console.ReadLine();

            Console.WriteLine("Enter ID");
            Id = Console.ReadLine();

            Console.WriteLine("Enter Name");
            Name = Console.ReadLine();

            Microsoft.Office.Interop.Excel.Application xlFile = new Microsoft.Office.Interop.Excel.Application();
            if (xlFile == null)
            {
                Console.WriteLine("Excel is not installed.");
                Console.ReadKey();
                return;
            }

            Microsoft.Office.Interop.Excel.Workbook xlWorkbook;
            Microsoft.Office.Interop.Excel.Worksheet xlSheet;
            object misValue = System.Reflection.Missing.Value;

            xlWorkbook = xlFile.Workbooks.Add(misValue);
            xlSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkbook.Worksheets.get_Item(1);
            xlSheet.Cells[1, 1] = "Id";
            xlSheet.Cells[1, 2] = "Name";
            xlSheet.Cells[2, 1] = Id;
            xlSheet.Cells[2, 2] = Name;


            string location = @"D:\" + fileName + ".xls";
            xlWorkbook.SaveAs(location, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal,
                misValue, misValue, misValue, misValue, 
                Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue,
                misValue, misValue, misValue);
            xlWorkbook.Close(true, misValue, misValue);
            xlFile.Quit();

            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlFile);
                xlFile = null;
            }
            catch(Exception ex)
            {
                xlFile = null;
                Console.Write("Error" + ex.ToString()); ;
            }
            finally
            {
                GC.Collect();
            }
        }
        static void Main(string[] args)
        {
            GenerateExcel();
        }
    }
}
