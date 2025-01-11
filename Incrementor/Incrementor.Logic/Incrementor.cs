using ClosedXML.Excel;

namespace Incrementor.Logic;

public enum IncrementorParsingResultErrorType
{
    None = -1,
    FileNotFound,
    Io,
    Argument,
    Unhandled
}

public class IncrementorParsingResult
{
    public bool ParsingResult { get; set; }
    public string OutputFilePath { get; set; } = string.Empty;
    public IncrementorParsingResultErrorType ErrorType { get; set; } = IncrementorParsingResultErrorType.None;
    public string ErrorMessage { get; set; } = string.Empty;
}

public static class Incrementor
{
    public static IncrementorParsingResult ProcessData(string inputFilePath)
    {
        var result = new IncrementorParsingResult();
        try
        {
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

            const string outputFilePath = "Output.xlsx";

            workbook.SaveAs(outputFilePath);

            result.ParsingResult = true;
            result.OutputFilePath = outputFilePath;
            result.ErrorMessage = string.Empty;
        }
        catch (FileNotFoundException)
        {
            // TODO: Log.
            result.ParsingResult = false;
            result.OutputFilePath = string.Empty;
            result.ErrorType = IncrementorParsingResultErrorType.FileNotFound;
            result.ErrorMessage = "Error: File not found.";
        }
        catch (IOException)
        {
            // TODO: Log.
            result.ParsingResult = false;
            result.OutputFilePath = string.Empty;
            result.ErrorType = IncrementorParsingResultErrorType.Io;
            result.ErrorMessage = "Error: File path not found or file used by other programs.";
        }
        catch (ArgumentException)
        {
            // TODO: Log.
            result.ParsingResult = false;
            result.OutputFilePath = string.Empty;
            result.ErrorType = IncrementorParsingResultErrorType.Argument;
            result.ErrorMessage = "Error: File path or file extension invalid.";
        }
        catch (Exception ex)
        {
            // TODO: Log.
            result.ParsingResult = false;
            result.OutputFilePath = string.Empty;
            result.ErrorType = IncrementorParsingResultErrorType.Unhandled;
            result.ErrorMessage = ex.ToString();
        }

        return result;
    }
}