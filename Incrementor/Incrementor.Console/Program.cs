Console.WriteLine("Enter input file path: ");
var inputFilePath = Console.ReadLine() ?? string.Empty; // @"C:\Users\Andrija\Downloads\Book1.xlsx"; 
Console.WriteLine("Enter output file path: ");
var outputFilePath = Console.ReadLine() ?? string.Empty; // @"C:\Users\Andrija\Downloads\Output.xlsx";  

var incrementorParsingResult = Incrementor.Logic.Incrementor.ProcessData(inputFilePath, outputFilePath);

Console.WriteLine(incrementorParsingResult.ParsingResult
    ? $"New file saved at {incrementorParsingResult.OutputFilePath}"
    : $"New file not saved. Code: {(int)incrementorParsingResult.ErrorType}. {incrementorParsingResult.ErrorMessage}");

Console.WriteLine("Press any key to continue...");
Console.ReadKey();