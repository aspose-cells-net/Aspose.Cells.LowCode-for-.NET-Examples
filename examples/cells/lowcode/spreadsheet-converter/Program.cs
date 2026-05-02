using System;
using System.IO;
using Aspose.Cells.LowCode;

string inputPath = Path.Combine(AppContext.BaseDirectory, "input.xlsx");
if (!File.Exists(inputPath))
    throw new FileNotFoundException("Input fixture not found", inputPath);

string outputPath = Path.Combine(AppContext.BaseDirectory, "output.xlsx");

// Remove any previous output to ensure deterministic behavior
if (File.Exists(outputPath))
    File.Delete(outputPath);

// Call the simplest overload of SpreadsheetConverter.Process
SpreadsheetConverter.Process(inputPath, outputPath);

// Verify that the output file was created
if (File.Exists(outputPath))
{
    var info = new FileInfo(outputPath);
    Console.WriteLine($"Done. Output: {outputPath} ({info.Length} bytes)");
}
else
{
    throw new InvalidOperationException("Output file was not created");
}