namespace Incrementor.Logic;

using ClosedXML.Excel;

public class Incrementor
{
    public (bool, string) ProcessData(string inputFilePath)
    {
        try
        {
            const string outputFilePath = "Output.xlsx";
            
            var workbook = new XLWorkbook(inputFilePath);
            var worksheet = workbook.Worksheet(1);

            var columnNumber = worksheet.LastColumnUsed()!.ColumnNumber();

            for (var i = 1; i < worksheet.LastRowUsed()!.RowNumber() + 1; i++)
            {
                for (var j = 1; j < columnNumber + 2; j++)
                {
                    // worksheet.Cell(i, j).Value = i + ":" + j;
                    if (j > 1 && worksheet.Cell(i, j).IsEmpty()
                              && !worksheet.Cell(i, j - 1).IsEmpty())
                    {
                        worksheet.Cell(i, j).Value = decimal.Parse(
                            worksheet.Cell(i, j - 1).Value.ToString()) + 1;
                    }
                }
            }

            workbook.SaveAs(outputFilePath);

            return (true, outputFilePath);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Unhandled exception: ${0}", ex);
        }

        return (false, "");
    }
}