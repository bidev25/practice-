using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace CSharpandHTML
{
    class Program
    {
        static void Main()
        {
            Excel.Application SourceApplication;
            Excel.Workbook Sourceworkbook;
            Excel.Worksheet Sourceworksheet;
            Excel.Range Sourcerange;
            String SourcePath;
            SourcePath = FILEPATH;
            SourceApplication = new Excel.Application();
            Sourceworkbook = SourceApplication.Workbooks.Open(SourcePath);
            Sourceworksheet = Sourceworkbook.Worksheets.get_Item(1);
            Sourcerange = Sourceworksheet.UsedRange;

        int rowcount= Sourcerange.Rows.Count;
        int columncount =Sourcerange.Columns.Count;
        int i;
        int j;

            for (i = 1; i<=rowcount; i++ )
            {
                for (j=1; j< columncount;j++)
                {
                    if (j == 1)
                        Console.Write("\r\n");

                    if (Sourcerange.Cells[i, j] != null && Sourcerange.Cells[i, j].Value2 != null)
                        Console.Write(Sourcerange.Cells[i, j].Value2.ToString() + "\t");
                }
            }
            Sourceworkbook.Close(true, null, null);
            SourceApplication.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(Sourceworksheet);
            Marshal.ReleaseComObject(Sourceworkbook);
            Marshal.ReleaseComObject(SourceApplication);
        }
}
}
