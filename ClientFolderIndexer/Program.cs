using ClosedXML.Excel;
using System;
using System.IO;
using System.Linq;

namespace FolderIndexer
{
    public class Program
    {
        public static void Main(string[] args)
        {
            //The folder containing the directories we want to index
            string folderPath = @"Copy Folder Address Here";

            if (!Directory.Exists(folderPath))
            {
                Console.WriteLine("Error The following directory does not exist: " + folderPath);
                return;
            }

            //Paste the full path AND the desired file name
            string outputFile = @"\Path\To\Your\Folder\MyIndexName" + ".xlsx";

            //Ensures the output directory exists
            string? outputDir = Path.GetDirectoryName(outputFile);
            if (!string.IsNullOrEmpty(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            string[] folders;

            try
            {
                //Grabs directories
                folders = Directory.GetDirectories(folderPath);
            }
            catch (UnauthorizedAccessException)
            {
                Console.WriteLine("Access denied to the folder.");
                return;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to read folders: {ex.Message}");
                return;
            }

            //Sort folders alphabetically by name
            var orderedFolders = folders.OrderBy(f => f, StringComparer.OrdinalIgnoreCase);

            using (var workbook = new XLWorkbook())
            {
                //Name the sheet "Indexed Folders"
                var ws = workbook.Worksheets.Add("Indexed Folders");

                //Headers
                ws.Cell(1, 1).Value = "Index";
                ws.Cell(1, 2).Value = "Folder Name";
                ws.ShowGridLines = false;

                //Styles the Header Row a dark blue color
                var headerRange = ws.Range(1, 1, 1, 2);
                headerRange.Style.Font.Bold = true;
                headerRange.Style.Fill.BackgroundColor = XLColor.LightBlue;

                int row = 2;
                int count = 1;

                foreach (string dirPath in orderedFolders)
                {
                    ws.Cell(row, 1).Value = count++;

                    //Extracts just the folder name from the path
                    ws.Cell(row, 2).Value = new DirectoryInfo(dirPath).Name;

                    row++;
                }

                //Styles the Data Rows a light blue color
                var dataRange = ws.Range(2, 1, row - 1, 2);
                dataRange.Style.Fill.BackgroundColor = XLColor.AliceBlue;

                //Draws a black box around the table
                var fullTableRange = ws.Range(1, 1, row - 1, 2);
                fullTableRange.Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
                fullTableRange.Style.Border.OutsideBorderColor = XLColor.Black;

                //This creates a dark border around the Excel Columns and Row borders
                dataRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                dataRange.Style.Border.InsideBorderColor = XLColor.LightGray;
                dataRange.Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
                dataRange.Style.Border.OutsideBorderColor = XLColor.Black;

                //Auto adjusts the columns to match folder name length
                ws.Columns().AdjustToContents();
                ws.Column(2).Width = 40;

                //Align the index numbers to the left to create a gap before the title
                ws.Column(1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                ws.Column(2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                //Saves file, will error if file is open
                workbook.SaveAs(outputFile);
            }

            Console.WriteLine("Done. Excel file created: " + outputFile);
        }
    }
}