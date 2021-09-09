using System;
using System.Collections.Generic;
using System.IO;
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
            int lastIndex = 2;

            List<int> rowsToDelete = new List<int>();
            var cellsInSheet = worksheet.Cells;
            while (true)
            {
                var temp = cellsInSheet["C" + lastIndex];
                if (temp.Value != null)
                {
                    lastIndex++;
                }
                else
                {
                    lastIndex--;
                    break;
                }
            } //get index of last row with text

            for (int mainIndex = 2; mainIndex <= lastIndex; mainIndex++)
                if (cellsInSheet["C" + mainIndex].Value == cellsInSheet["C" + mainIndex + 1])
                {
                    if (cellsInSheet["G" + mainIndex].Value == cellsInSheet["G" + mainIndex + 1].Value)
                    {

                    }
                    var counter = 1;
                    var auxIndex = mainIndex;
                    while (cellsInSheet["C" + auxIndex].Value.Equals(cellsInSheet["C" + auxIndex + 1].Value))
                    {
                        counter++;
                        auxIndex++;
                    }
                }
            }
        }
    }
}
