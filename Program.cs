using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using Newtonsoft.Json;

class Program
{
    static int Main(string[] args)
    {
        if (args.Length != 2)
        {
            Console.WriteLine("Usage: program <input.xlsx> <output.json>");
            Console.WriteLine("Extracts rich text from an Excel file and saves it as a structured JSON.");
            return 1;
        }

        string inputFile = args[0];
        string outputFile = args[1];

        if (!File.Exists(inputFile))
        {
            Console.WriteLine($"Error: File '{inputFile}' not found.");
            return 1;
        }

        if (File.Exists(outputFile))
        {
            Console.WriteLine($"Error: Output file '{outputFile}' already exists.");
            return 1;
        }

        try
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var package = new ExcelPackage(new FileInfo(inputFile)))
            {
                // sheet name -> list[cell info]
                var result = new Dictionary<string, List<CellInfo>>();

                // each worksheet (tab)
                foreach (var worksheet in package.Workbook.Worksheets)
                {
                    var cellList = new List<CellInfo>();

                    for (int row = 1; row <= worksheet.Dimension.Rows; row++)
                    {
                        for (int col = 1; col <= worksheet.Dimension.Columns; col++)
                        {
                            var cell = worksheet.Cells[row, col];

                            var cellInfo = new CellInfo
                            {
                                Row = row,
                                Column = col,
                                FullText = cell.Text,
                                RichTextChildren = new List<RichTextPart>()
                            };

                            // if there are multiple parts, like parts of words
                            // underlined or bolded, but the rest not
                            if (cell.RichText.Count > 0)
                            {
                                foreach (var richText in cell.RichText)
                                {
                                    cellInfo.RichTextChildren.Add(new RichTextPart
                                    {
                                        Text = richText.Text,
                                        IsBold = richText.Bold,
                                        IsItalic = richText.Italic,
                                        IsUnderline = richText.UnderLine
                                    });
                                }
                            }
                            else
                            {
                                // If no rich text, store as a single rich text child
                                cellInfo.RichTextChildren.Add(new RichTextPart
                                {
                                    Text = cell.Text,
                                    IsBold = false,
                                    IsItalic = false,
                                    IsUnderline = false
                                });
                            }

                            cellList.Add(cellInfo);
                        }
                    }

                    result[worksheet.Name] = cellList;
                }

                string json = JsonConvert.SerializeObject(result, Formatting.Indented);

                File.WriteAllText(outputFile, json);
                Console.WriteLine($"JSON export completed successfully: {outputFile}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
            return 1;
        }

        return 0;
    }
}

// Cell structure
public class CellInfo
{
    public int Row { get; set; }
    public int Column { get; set; }
    public string? FullText { get; set; } // Store full cell text for reference
    public List<RichTextPart> RichTextChildren { get; set; } = new List<RichTextPart>();
}

// Rich text fragment
public class RichTextPart
{
    public string? Text { get; set; }
    public bool IsBold { get; set; }
    public bool IsItalic { get; set; }
    public bool IsUnderline { get; set; }
}

