using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;

namespace TimesTable2Excel
{
    class Program
    {
        static void Main(string[] args)
        {
            string outputFilePath = "timesTable.xlsx";

            int seriesSize = 12;

            int total = 40;

            var generator = new ExcelGenerator(outputFilePath, seriesSize, total);
            generator.Generate();
        }
    }

    internal static class Randomizer
    {
        static Random _random = new Random();

        internal static int GetRandomNumber() => _random.Next();
    }

    internal class Operation
    {
        internal int Multiplicand { get; set; }
        internal int Multiplier { get; set; }
    }


    internal class ExcelGenerator
    {
        private string outputFilePath;
        private int seriesSize;
        private int total;
        private int multiplicand = 3;

        public ExcelGenerator(string outputFilePath, int seriesSize, int total)
        {
            this.outputFilePath = outputFilePath;
            this.seriesSize = seriesSize;
            this.total = total;
        }

        internal void Debug(List<Operation> operations)
        {
            for (int iter = 0; iter < operations.Count; iter++)
            {
                if (iter > 0 && iter % seriesSize == 0)
                {
                    Console.WriteLine("------");
                }

                Console.WriteLine($"{operations[iter].Multiplicand}x{operations[iter].Multiplier}");
            }
        }

        internal void Generate()
        {
            //divide runs
            List<Operation> operations = GenerateRuns();
            Debug(operations);            

            var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("Times table");

            for (int column = 0; column <= 2; column++)
            {
                for (int iter = 0; iter < operations.Count; iter++)
                {
                    var cell = worksheet.Cell(1 + (iter), 1 + (column * 3));

                    cell.Style.Font.FontName = "Courier";
                    cell.Style.Font.FontSize = 14;
                    cell.WorksheetRow().Height = 18;
                    cell.Value
                        = string.Format("{0} x {1} =",
                            operations[iter].Multiplicand.ToString().PadRight(2),
                            operations[iter].Multiplier.ToString().PadRight(2));
                }
            }

            workbook.SaveAs(outputFilePath);
        }

        private List<Operation> GenerateRuns()
        {
            List<Operation> operations = new List<Operation>();

            int numberOfFullSeries = Math.DivRem(total, seriesSize, out int lastSeriesSize);

            for (int iter = 0; iter < numberOfFullSeries; iter++)
            {
                operations.AddRange(GenerateUniqueRun(seriesSize));
            }

            if (lastSeriesSize > 0)
            {
                operations.AddRange(GenerateUniqueRun(lastSeriesSize));
            }

            return operations;
        }

        private List<Operation> GenerateUniqueRun(int runSize)
        {
            return Enumerable.Range(1, seriesSize)
                .OrderBy(_ => Randomizer.GetRandomNumber())
                .Take(runSize)
                .Select(p => new Operation
                {
                    Multiplicand = multiplicand,
                    Multiplier = p,
                })
                .ToList();
        }
    }
}
