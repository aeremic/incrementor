var incrementor = new Incrementor.Logic.Incrementor();

var inputFilePath = Console.ReadLine() ?? string.Empty;

var parsingResult = false;
try
{ 
    parsingResult = incrementor.CalculateData(inputFilePath);
}
catch (Exception ex)
{
    Console.WriteLine("Unhandled exception: ${0}", ex);
}

Console.WriteLine(parsingResult ? $"New file saved at ${inputFilePath}" : "New file not saved.");
