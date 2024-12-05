namespace Incrementor.Logic;
using Interop = Microsoft.Office.Interop.Excel; 

public class Incrementor
{
    public bool CalculateData(string inputFilePath)
    {
        var result = false;
        
        // load
        var excelApplication = new Interop.Application();
        var excelWorkbook = excelApplication.Workbooks
            .Open(
                inputFilePath, 
                0,
                true, 
                5, 
                "", 
                "", 
                true, 
                Interop.XlPlatform.xlWindows, 
                "\t", 
                false, 
                false, 
                0, 
                true, 
                1, 
                0);
        var excelWorksheet = (Interop.Worksheet)excelWorkbook.Worksheets.Item[1];

        var usedRange = excelWorksheet.UsedRange;
        var rowsCount = usedRange.Rows.Count;
        var columnsCount = usedRange.Columns.Count;
        
        for (var i = 1; i <= rowsCount; i++)
        {
            for (var j = 1; j <= columnsCount; j++)
            {
                Console.WriteLine((string)(usedRange.Cells[i, j] as Interop.Range)?.Value2!);
            }
        }

        excelWorkbook.Close(true, null, null);
        excelApplication.Quit();
        
        // calculate
        
        // save
        
        return result;
    }
}