using System.Globalization;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace ConsoleApp1;

class SearchAndReplaceExcelParams
{
    public static void SearchAndReplace(string filePath, Dictionary<string, string> replacements)
    {
        IWorkbook workbook;
        using (FileStream fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
        {
            workbook = new XSSFWorkbook(fileStream);
        }

        try
        {
            ISheet sheet = workbook.GetSheetAt(0);
            bool isModified = false;

            for (int rowIndex = 0; rowIndex <= sheet.LastRowNum; rowIndex++)
            {
                IRow row = sheet.GetRow(rowIndex);
                if (row == null) continue;

                for (int colIndex = 0; colIndex < row.LastCellNum; colIndex++)
                {
                    ICell? cell = row.GetCell(colIndex);
                    if (cell == null) continue;

                    string cellValue = GetCellValueAsString(cell).Trim();

                    foreach (var entry in replacements)
                    {
                        if (string.Equals(cellValue, entry.Key, StringComparison.OrdinalIgnoreCase))
                        {
                            Console.WriteLine(
                                $"Replacing '{cellValue}' with '{entry.Value}' at [{rowIndex}, {colIndex}]");
                            cell.SetCellValue(entry.Value);
                            isModified = true;
                        }
                    }
                }
            }

            if (isModified)
            {
                using (FileStream outputStream = new FileStream(filePath, FileMode.Create, FileAccess.Write))
                {
                    workbook.Write(outputStream);
                }

                Console.WriteLine("✅ Excel file updated successfully!");
            }
            else
            {
                Console.WriteLine("❌ No matches found. Nothing was replaced.");
            }
        }
        finally
        {
            workbook.Close();
        }
    }

    private static string GetCellValueAsString(ICell? cell)
    {
        if (cell == null) return string.Empty;

        string value = cell.CellType switch
        {
            CellType.String => cell.StringCellValue ?? string.Empty,
            CellType.Numeric => cell.NumericCellValue.ToString(CultureInfo.InvariantCulture),
            CellType.Boolean => cell.BooleanCellValue.ToString(),
            CellType.Formula => cell.CellFormula ?? string.Empty,
            _ => string.Empty
        };

        return value.Replace("\u00A0", "").Trim();
    }

    static void Main()
    {
        string filePath = "/Users/mani/Documents/Projects/noter/ConsoleApp1/ConsoleApp1/data.xlsx";
        var replacements = new Dictionary<string, string>
        {
            { "ja Baby", "Ja Baby" },
            { "OldText2", "NewText2" },
            { "FindMe", "ReplacedValue" }
        };

        SearchAndReplace(filePath, replacements);
    }
}