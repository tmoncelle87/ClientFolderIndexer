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
            //Set variables
            bool renameHeader = false;
            string headerText;


            //This code below asks the user if they want custom text displayed as the column header to indicate what the list is of/for.
            Console.WriteLine("Would you like a custom header displaying what the list contains? If so please type and enter a 'y' for yes or a 'n' for no");
            int numberOfTries = 0;
            while (numberOfTries < 10)
            {
                string userInput = (Console.ReadLine() ?? string.Empty).ToLower();

                if (userInput == "n")
                {
                    renameHeader = false;
                    break;
                }
                else if (userInput == "y")
                {
                    renameHeader = true;
                    break;
                }
                else
                {
                    Console.WriteLine("Would you like a custom header displaying what the list contains? If so please type and enter a 'y' for yes or a 'n' for no");
                    numberOfTries++;
                }
            }
            if (renameHeader)
            {
                Console.WriteLine("Please enter the custom header name: ");
                headerText = (Console.ReadLine() ?? string.Empty);
            }
            else
            {
                headerText = "Indexed Folders";
            }

            //The code below asks the user for the file path to the directory they want indexed

            int numberOfAttempts = 0;
            bool folderExists = false;
            string filePath = "";
            while (numberOfAttempts < 10)
            {
                Console.WriteLine("Please enter the file path to the directory you want indexed: ");
                filePath = (Console.ReadLine() ?? string.Empty).Trim(' ', '"', '\'');

                if (!Directory.Exists(filePath))
                {
                    Console.WriteLine("Error The following directory does not exist: " + filePath);
                    numberOfAttempts++;
                }
                else
                {
                    folderExists = true;
                    break;
                }
            }
            if (!folderExists)
            {
                Console.WriteLine("Maximum attempts reached or invalid directory. Exiting.");
                return;
            }

            //The code below asks the user for the file path in which the spreadsheet should be created
            numberOfAttempts = 0;
            folderExists = false;
            string outputDirectory = "";
            while (numberOfAttempts < 10)
            {
                Console.WriteLine("Please enter the file path to the directory you want the excel file saved to: ");
                outputDirectory = (Console.ReadLine() ?? string.Empty).Trim(' ', '"', '\'');

                if (!Directory.Exists(outputDirectory))
                {
                    Console.WriteLine("Error The following directory does not exist: " + outputDirectory);
                    numberOfAttempts++;
                }
                else
                {
                    folderExists = true;
                    break;
                }
            }
            if (!folderExists)
            {
                Console.WriteLine("Maximum attempts reached or invalid directory. Exiting.");
                return;
            }







            //Asks the user for a custom file name
            Console.WriteLine("Please enter a name for the Excel file (without extension): ");
            string fileNameInput = (Console.ReadLine() ?? string.Empty).Trim();
            string fileName = string.IsNullOrWhiteSpace(fileNameInput) ? "IndexedFolders" : fileNameInput;

            string outputFile = Path.Combine(outputDirectory, fileName + ".xlsx");


            //Ensures the output directory exists
            string? outputDir = Path.GetDirectoryName(outputFile);
            if (!string.IsNullOrEmpty(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            string[] folders;

            try
            {
                folders = Directory.GetDirectories(filePath);
            }
            catch (UnauthorizedAccessException)
            {
                Console.WriteLine("Access denied to the folder.");
                return;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to read folder: {ex.Message}");
                return;
            }

            //Sort folders alphabetically by name
            var orderedFolders = folders
                //Handles and hides hidden folders or "." named folders 
                .Where(f => !Path.GetFileName(f).StartsWith("."))
                .OrderBy(f => f, StringComparer.OrdinalIgnoreCase);

            using (var workbook = new XLWorkbook())
            {
                //Name the sheet "Indexed Folders"
                var ws = workbook.Worksheets.Add("Indexed Folders");

                //Headers styling
                ws.Cell(1, 1).Value = "Index";
                ws.Cell(1, 2).Value = headerText;
                ws.ShowGridLines = false;

                //Styles the Header Row a dark blue color
                var headerRange = ws.Range(1, 1, 1, 2);
                headerRange.Style.Font.Bold = true;
                headerRange.Style.Fill.BackgroundColor = XLColor.LightBlue;

                int row = 2;
                int count = 1;

                foreach (string folderPathName in orderedFolders)
                {
                    ws.Cell(row, 1).Value = count++;
                    ws.Cell(row, 2).Value = Path.GetFileName(folderPathName);
                    row++;
                }

                //Styles excel data rows but only if folders are found
                if (row > 2)
                {
                    //Styles the Data Rows a light blue color
                    var dataRange = ws.Range(2, 1, row - 1, 2);
                    dataRange.Style.Fill.BackgroundColor = XLColor.LightBlue;

                    //Draws a black box around the table
                    var fullTableRange = ws.Range(1, 1, row - 1, 2);
                    fullTableRange.Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
                    fullTableRange.Style.Border.OutsideBorderColor = XLColor.Black;

                    //This creates a dark border around the Excel Columns and Row borders
                    dataRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                    dataRange.Style.Border.InsideBorderColor = XLColor.LightGray;
                }
                else
                {
                    //Fallback if the folder is empty: just puts a border around the header row only
                    var headerOnlyRange = ws.Range(1, 1, 1, 2);
                    headerOnlyRange.Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
                    headerOnlyRange.Style.Border.OutsideBorderColor = XLColor.Black;
                }


                //Auto adjusts the columns to match folder name length
                ws.Columns().AdjustToContents();

                //Aligns the index numbers to the left to create a gap before the title
                ws.Column(1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                ws.Column(2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                //Saves file, handles error if file is already open
                try
                {
                    workbook.SaveAs(outputFile);
                    Console.WriteLine("Done. Excel file created: " + outputFile);
                }
                catch (IOException)
                {
                    Console.WriteLine("Error: Cannot save the file '" + fileName + ".xlsx' because it is currently open in another program (like Excel). Please close it and try again.");
                }
                catch (Exception exceptionMessage)
                {
                    Console.WriteLine($"An unexpected error occurred while saving: {exceptionMessage.Message}");
                }
            }
        }
    }
}