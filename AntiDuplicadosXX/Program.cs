using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel;

namespace AntiDuplicadosXX
{
    class Program
    {
        static void Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var pathExcel = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) +
                            @"\Downloads\Compras Indirectas.xlsx";
            var file = new FileInfo(pathExcel);

            using var package = new ExcelPackage(file);

            var worksheet = package.Workbook.Worksheets[0];
            int lastIndex = worksheet.Dimension.End.Row;
            List<int> rowsToKeep = new List<int>();
            List<int> rowsToDelete = new List<int>();
            var cellsInSheet = worksheet.Cells;

            //while (true)
            //{
            //    var temp = cellsInSheet["C" + lastIndex];
            //    if (temp.Value != null)
            //    {
            //        lastIndex++;
            //    }
            //    else
            //    {
            //        lastIndex--;
            //        break;
            //    }
            //} //get index of last row with text

            for (int mainIndex = 2; mainIndex <= lastIndex; mainIndex++)
            {
                if (cellsInSheet["C" + mainIndex].Value.Equals(cellsInSheet["C" + (mainIndex + 1)].Value))
                {
                    var counter = 1;
                    var auxIndex = mainIndex;
                    var constIndex = mainIndex;
                    if (!rowsToKeep.Contains(mainIndex))
                    {
                        Console.WriteLine("Good row: " + mainIndex);
                        rowsToKeep.Add(mainIndex);
                    }

                    if (cellsInSheet["G" + auxIndex].Value.Equals(cellsInSheet["G" + (auxIndex + 1)].Value))
                    {
                        while (cellsInSheet["G" + constIndex].Value.Equals(cellsInSheet["G" + (auxIndex + 1)].Value))
                        {
                            counter++;
                            auxIndex++;
                        }

                        auxIndex = constIndex;
                        while (cellsInSheet["G" + constIndex].Value.Equals(cellsInSheet["G" + (auxIndex - 1)].Value))
                        {
                            counter++;
                            auxIndex--;
                        }
                        mainIndex += counter;
                    }
                    
                }
                else
                {
                    if (!rowsToKeep.Contains(mainIndex))
                    {
                        Console.WriteLine("Good row: " + mainIndex);
                        rowsToKeep.Add(mainIndex);
                    }
                }
                        
            } //Get good rows

            for (int rowNumber = lastIndex; rowNumber >= 2; rowNumber--)
            {
                if (!(rowsToKeep.Contains(rowNumber)))
                {
                    worksheet.DeleteRow(rowNumber);
                }
            }

            package.Save();
            Console.WriteLine("Press anything to end it all.");
            Console.ReadKey();

        }
    }
}
