using System;
using Microsoft.Office.Interop.Excel;
using _excel = Microsoft.Office.Interop.Excel;
namespace ExportToExcel
{
    class Program
    {
        string path = "";
        _Application excel = new _excel.Application();
        public Workbook wb;
        public Worksheet ws;


        public _Workbook Workbook;
       public Program(string path, string sheet)
        {
            this.path = path;
            wb = excel.Workbooks.Open(path);
            ws = wb.Worksheets[sheet];
        }

        public string ReadCell(int i, int j)
        {
            i++;
            j++;
            if (ws.Cells[i, j].ToString()==null)
            {
                return ws.Cells[i, j].ToString();
            }

        }

        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");
        }
    }
}
