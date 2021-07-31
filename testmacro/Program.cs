using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace testmacro
{
    class Program
    {
        static void Main(string[] args)
        {
            runMacro();
        }

        public static void runMacro()
        {
            Excel.Application xlApp = new Excel.Application();

            Excel.Workbook xlWorkBook;

            //~~> Start Excel and open the workbook.
            xlWorkBook = xlApp.Workbooks.Open("file path");

            //~~> Run the macros by supplying the necessary arguments
            xlApp.Run("macro name");

            //~~> Clean-up: Close the workbook
            xlWorkBook.Close(false);

            //~~> Quit the Excel Application
            xlApp.Quit();
            Console.WriteLine("Macro was executed succesfully");
            //~~> Clean Up
            releaseObject(xlApp);
            releaseObject(xlWorkBook);
        }

        public static void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}
