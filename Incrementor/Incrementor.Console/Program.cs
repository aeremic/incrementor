var incrementor = new Incrementor.Logic.Incrementor();

var inputFilePath = @"C:\Users\Andrija\Downloads\Book1.xlsx"; // Console.ReadLine() ?? string.Empty; //  

var (parsingResult, outputFilePath) = incrementor.ProcessData(inputFilePath);

Console.WriteLine(parsingResult ? $"New file saved as ${outputFilePath}" : "New file not saved.");