using System;
using System.IO;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace PingAuswertung
{
    class Program
    {
        static Excel.Application xlApp = new Excel.Application();
        static Excel.Workbook xlWorkBook;
        static Excel.Worksheet xlWorkSheet;
        static string path;

        static void Main(string[] args)
        {
            if (args.Length == 0)
            {
                Console.WriteLine("Ordnerpfad eingeben: ");
                path = Console.ReadLine();

                goDo(path);
            }
            else
            {
                path = args[0];
                if (!path.Contains("\\"))
                    Console.WriteLine("Parameter falsch angegeben!");
                else
                    goDo(path);
            }

            xlApp.Quit();
            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);
        }

        static void goDo(string folderPath)
        {
            xlWorkBook = xlApp.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Sheets.Item[1];

            int cursorHor = 2;
            foreach (string file in Directory.GetFiles(folderPath))
            {
                FileInfo fi = new FileInfo(file);
                if (fi.Name.Contains("Computers") && fi.Name.Contains("Output"))
                    continue;

                int cursorVer = 2;

                Console.WriteLine(cursorHor + " " + cursorVer);

                if (fi.Name.Contains("PingResult"))
                {
                    xlWorkSheet.Cells[1, cursorHor] = fi.Name.Substring(11, 10);
                    foreach (string comp in File.ReadAllText(file).Split('\n'))
                    {
                        if (!comp.Contains("["))
                            continue;

                        Console.WriteLine(comp);
                        xlWorkSheet.Cells[cursorVer, 1] = comp.Split('[')[1].Substring(0, comp.Split('[')[1].Length - 2);
                        xlWorkSheet.Cells[cursorVer, cursorHor] = comp.Split(' ')[0];
                        cursorVer++;
                    }
                    cursorHor++;
                }
            }

            xlWorkBook.SaveAs(folderPath + "\\Output.xls", Excel.XlFileFormat.xlWorkbookNormal, null, null, null, null, Excel.XlSaveAsAccessMode.xlNoChange, null, null, null, null, null);
            xlWorkBook.Close();
        }
    }
}
