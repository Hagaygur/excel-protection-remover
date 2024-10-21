using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Security.Principal;
using System.Text;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;

namespace RemoveSheetProtection
{
    class Program
    {
        static void Main(string[] args)
        {
            // Ensure the console uses UTF-8 encoding
            Console.OutputEncoding = Encoding.UTF8;

            // Check if running as administrator
            if (!IsAdministrator())
            {
                // Restart the application with administrator privileges
                ElevateToAdministrator();
                return;
            }

            Console.WriteLine("Searching for .xlsx files in the current directory...\n");

            string currentDirectory = Directory.GetCurrentDirectory();
            string[] xlsxFiles = Directory.GetFiles(currentDirectory, "*.xlsx");

            if (xlsxFiles.Length == 0)
            {
                Console.WriteLine("No .xlsx files found in the current directory.");
                Console.WriteLine("Press any key to exit.");
                Console.ReadKey();
                return;
            }

            // Display files with numbers
            Dictionary<int, string> fileDict = new Dictionary<int, string>();
            for (int i = 0; i < xlsxFiles.Length; i++)
            {
                int index = i + 1;
                string fileName = Path.GetFileName(xlsxFiles[i]);
                fileDict.Add(index, xlsxFiles[i]);
                Console.WriteLine($"{index}: {fileName}");
            }

            Console.Write("\nEnter the number of the file you want to process: ");
            string input = Console.ReadLine();

            if (!int.TryParse(input, out int choice) || !fileDict.ContainsKey(choice))
            {
                Console.WriteLine("Invalid selection.");
                Console.WriteLine("Press any key to exit.");
                Console.ReadKey();
                return;
            }

            string selectedFile = fileDict[choice];
            string selectedFileName = Path.GetFileName(selectedFile);
            Console.WriteLine($"\nYou selected: {selectedFileName}");

            try
            {
                // Make a backup of the file
                string backupFile = selectedFile + ".bak";
                Console.WriteLine("\nMaking a backup of the file...");
                File.Copy(selectedFile, backupFile, overwrite: true);

                // Rename .xlsx file to .zip
                string zipFile = Path.ChangeExtension(selectedFile, ".zip");
                Console.WriteLine("Renaming .xlsx file to .zip...");
                if (File.Exists(zipFile))
                {
                    File.Delete(zipFile);
                }
                File.Move(selectedFile, zipFile);

                // Create temporary extraction folder
                string tempFolder = Path.Combine(currentDirectory, "temp_extract");
                Console.WriteLine("Creating temporary extraction folder...");
                if (Directory.Exists(tempFolder))
                {
                    Directory.Delete(tempFolder, true);
                }
                Directory.CreateDirectory(tempFolder);

                // Extract zip file
                Console.WriteLine("Extracting zip file...");
                ZipFile.ExtractToDirectory(zipFile, tempFolder);

                // Process XML files
                Console.WriteLine("Processing XML files...");
                string worksheetsPath = Path.Combine(tempFolder, "xl", "worksheets");
                string[] xmlFiles = Directory.GetFiles(worksheetsPath, "*.xml", SearchOption.AllDirectories);

                foreach (string xmlFile in xmlFiles)
                {
                    Console.WriteLine($"Processing {xmlFile}");
                    RemoveSheetProtection(xmlFile);
                }

                // Repackaging files into zip
                Console.WriteLine("Repackaging files into zip...");
                if (File.Exists(zipFile))
                {
                    File.Delete(zipFile);
                }
                ZipFile.CreateFromDirectory(tempFolder, zipFile);

                // Rename zip file back to .xlsx
                Console.WriteLine("Renaming zip file back to .xlsx...");
                File.Move(zipFile, selectedFile);

                // Cleaning up temporary files
                Console.WriteLine("Cleaning up temporary files...");
                Directory.Delete(tempFolder, true);

                // Compare the original and modified files
                Console.WriteLine("\nVerifying the modified file against the backup...");
                var discrepancies = CompareExcelFiles(backupFile, selectedFile);

                if (discrepancies.Any())
                {
                    Console.WriteLine("\nDiscrepancies found:");
                    foreach (var discrepancy in discrepancies)
                    {
                        Console.WriteLine(discrepancy);
                    }

                    Console.Write("\nDo you want to proceed with the changes? (Y/N): ");
                    string userInput = Console.ReadLine();
                    if (userInput.Equals("Y", StringComparison.OrdinalIgnoreCase))
                    {
                        Console.WriteLine("Changes accepted.");
                    }
                    else
                    {
                        // Revert to backup
                        Console.WriteLine("Reverting to backup...");
                        File.Copy(backupFile, selectedFile, overwrite: true);
                    }
                }
                else
                {
                    Console.WriteLine("Verification successful. No discrepancies found.");
                }

                Console.WriteLine("\nDone.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"\nAn error occurred: {ex.Message}");
            }

            Console.WriteLine("\nPress any key to exit.");
            Console.ReadKey();
        }

        static void RemoveSheetProtection(string xmlFile)
        {
            // Load the XML document
            var xdoc = System.Xml.Linq.XDocument.Load(xmlFile);

            // Define the namespace
            var ns = xdoc.Root.GetDefaultNamespace();

            // Find the sheetProtection element
            var sheetProtection = xdoc.Root.Element(ns + "sheetProtection");
            if (sheetProtection != null)
            {
                sheetProtection.Remove();
                xdoc.Save(xmlFile);
            }
        }

        static List<string> CompareExcelFiles(string file1, string file2)
        {
            List<string> discrepancies = new List<string>();

            // Configure EPPlus to allow non-commercial use
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package1 = new ExcelPackage(new FileInfo(file1)))
            using (var package2 = new ExcelPackage(new FileInfo(file2)))
            {
                var workbook1 = package1.Workbook;
                var workbook2 = package2.Workbook;

                // Compare the number of worksheets
                if (workbook1.Worksheets.Count != workbook2.Worksheets.Count)
                {
                    discrepancies.Add("Number of worksheets differ between files.");
                    return discrepancies;
                }

                for (int i = 0; i < workbook1.Worksheets.Count; i++)
                {
                    var ws1 = workbook1.Worksheets[i];
                    var ws2 = workbook2.Worksheets[i];

                    if (ws1.Dimension == null && ws2.Dimension == null)
                    {
                        continue; // Both sheets are empty
                    }

                    if (ws1.Dimension == null || ws2.Dimension == null)
                    {
                        discrepancies.Add($"Worksheet '{ws1.Name}' differs: one sheet is empty and the other is not.");
                        continue;
                    }

                    var startRow = Math.Min(ws1.Dimension.Start.Row, ws2.Dimension.Start.Row);
                    var endRow = Math.Max(ws1.Dimension.End.Row, ws2.Dimension.End.Row);
                    var startCol = Math.Min(ws1.Dimension.Start.Column, ws2.Dimension.Start.Column);
                    var endCol = Math.Max(ws1.Dimension.End.Column, ws2.Dimension.End.Column);

                    for (int row = startRow; row <= endRow; row++)
                    {
                        for (int col = startCol; col <= endCol; col++)
                        {
                            var cell1 = ws1.Cells[row, col];
                            var cell2 = ws2.Cells[row, col];

                            object val1 = GetCellValue(cell1);
                            object val2 = GetCellValue(cell2);

                            if (!object.Equals(val1, val2))
                            {
                                discrepancies.Add($"Worksheet '{ws1.Name}', Cell [{row},{col}] differs. Original: '{val1}' | Modified: '{val2}'");
                            }
                        }
                    }
                }
            }

            return discrepancies;
        }

        static object GetCellValue(ExcelRange cell)
        {
            if (cell == null)
                return null;

            if (cell.Formula != null)
            {
                // Evaluate the formula
                return cell.Value;
            }
            else
            {
                return cell.Value;
            }
        }

        static bool IsAdministrator()
        {
            WindowsIdentity identity = WindowsIdentity.GetCurrent();
            WindowsPrincipal principal = new WindowsPrincipal(identity);
            return principal.IsInRole(WindowsBuiltInRole.Administrator);
        }

        static void ElevateToAdministrator()
        {
            ProcessStartInfo psi = new ProcessStartInfo
            {
                FileName = Process.GetCurrentProcess().MainModule.FileName,
                UseShellExecute = true,
                Verb = "runas"
            };

            try
            {
                Process.Start(psi);
            }
            catch (Exception)
            {
                Console.WriteLine("This operation requires administrator privileges.");
            }
        }
    }
}
