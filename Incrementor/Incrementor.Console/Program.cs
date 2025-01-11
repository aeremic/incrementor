var inputFilePath = @"C:\Users\Andrija\Downloads\Book1.xlsx"; // Console.ReadLine() ?? string.Empty; //  

var incrementorParsingResult = Incrementor.Logic.Incrementor.ProcessData(inputFilePath);

Console.WriteLine(incrementorParsingResult.ParsingResult
    ? $"New file saved as {incrementorParsingResult.OutputFilePath}"
    : $"New file not saved. Code: {(int)incrementorParsingResult.ErrorType}. {incrementorParsingResult.ErrorMessage}");