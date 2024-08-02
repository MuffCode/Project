using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
// Other using statements...
using Insphere.Connectivity.Common;
using Insphere.Connectivity.Common.ToolModel;
using Insphere.Connectivity.Application.MessageServices;
using Insphere.Connectivity.Application.SecsToOpc;
using System.IO;
using OfficeOpenXml; // open excel 
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using OfficeOpenXml.Table; //Open Excel table 
using System.Text.RegularExpressions;


namespace MyCustomHandler
{
    public class RecipeDownload : IGEMOPCService
    {
        private GEMOPCController gemOPCController;

        public void Initialize(GEMOPCController controller)
        {
            this.gemOPCController = controller;
            gemOPCController.LogInfo("Inside RecipeDownload");
            gemOPCController.AddGEMSubscription(this, 7, 3);
        }

        public void OPCNotification(string opcItem, object value, string logicalName)
        {
            throw new NotImplementedException();
        }

        public async void SECsMessageNotification(string streamFunction, SECsTransaction transaction)
        {
            try
            {
                gemOPCController.LogInfo("start");
                // Handle the S7F3 event asynchronously
                await ProcessS7F3Async(transaction);
                SECsMessage s7f4 = this.gemOPCController.GEMService.Services.CustomMessage.CreateMessage(7, 4, false);
                s7f4.DataItem.Add("ACK", 0, SECsFormat.Binary);
                gemOPCController.GEMService.SendReply(s7f4, transaction.Id);
            }
            catch (Exception ex)
            {
                // Handle or log the exception
                gemOPCController.LogInfo($"Error: {ex.Message}");
            }
        }

        private async Task ProcessS7F3Async(SECsTransaction transaction)
        {
            try
            {
                gemOPCController.LogInfo("start2");

                // Extract recipeId, headerDataValues, and dataItems from the transaction
                string recipeId = transaction.Primary.DataItem[0][0]?.ToString()?.Trim();
                gemOPCController.LogInfo($"Recipe ID: {recipeId}");

                string headerDataValue = transaction.Primary.DataItem[0][1]?.ToString().Trim();
                string[] lines = headerDataValue.Split(new string[] { "\r\n" }, StringSplitOptions.None);

                // Iterate through each line and add a space after the colon, then remove patterns like ": 199" and ": 1000"
                List<string> modifiedLinesList = new List<string>();
                foreach (var line in lines)
                {
                    // Remove patterns like ": 199" and ": 1000"
                    string modifiedLine = Regex.Replace(line, @":\s*\d+", string.Empty).Trim();
                    modifiedLinesList.Add(modifiedLine); 
                }

                gemOPCController.LogInfo($"HeaderDataValues: {string.Join(", ", modifiedLinesList)}");

                string dataItemsValue = transaction.Primary.DataItem[0][1]?.ToString();
                string[] dataItems = dataItemsValue.Split(new string[] { "\r\n" }, StringSplitOptions.None);

                List<string> modifiedLinesList1 = new List<string>();
                foreach (var line in dataItems)
                {
                    string[] parts = line.Split(':');   // Split the line using ':' and extract the second part
                    if (parts.Length > 1)
                    {
                        modifiedLinesList1.Add(parts[1].Trim());
                    }
                    else

                    {
                        modifiedLinesList1.Add(line.Trim()); //If no second part, trim the line and add it to the list
                    }
                }

                gemOPCController.LogInfo($"Dataitems: {string.Join(", ", modifiedLinesList1)}");

                using (var package = new ExcelPackage(new FileInfo("C:\\Users\\intern1\\Documents\\Programming\\RECIPE.xlsx")))
                {
                    var worksheet = package.Workbook.Worksheets.FirstOrDefault();

                    if (worksheet != null)
                    {
                        // Iterate through modified header data values and find corresponding columns
                        for (int i = 0; i < modifiedLinesList.Count && i < modifiedLinesList1.Count; i++) 
                        {
                            int specificColumnIndex = -1;

                            // Find the specific column index for the current header data value
                            for (int col = 2; col <= worksheet.Dimension.End.Column; col++)
                            {
                                var headerCell = worksheet.Cells[1, col]?.Value?.ToString()?.Trim();

                                if (string.Equals(modifiedLinesList[i], headerCell, StringComparison.OrdinalIgnoreCase))
                                {
                                    specificColumnIndex = col;
                                    break;
                                }
                            }

                            if (specificColumnIndex != -1)
                            {
                                int specificRowIndex = -1;

                                // Find the specific row for the given recipeId
                                for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                                {
                                    var recipeCell = worksheet.Cells[row, 1]?.Value?.ToString()?.Trim();

                                    if (string.Equals(recipeCell, recipeId, StringComparison.OrdinalIgnoreCase))
                                    {
                                        specificRowIndex = row;  
                                        break;
                                    }
                                }

                                if (specificRowIndex != -1)
                                {
                                    // Update data item for the specific recipeId, headerDataValues
                                    // and corresponding column
                                    worksheet.Cells[specificRowIndex, specificColumnIndex].Value = modifiedLinesList1[i];
                                    gemOPCController.LogInfo($"Updated Recipe ID '{recipeId}' at Row {specificRowIndex} in Column {specificColumnIndex} with data '{modifiedLinesList1[i]}' for header '{modifiedLinesList[i]}'.");
                                    // Save the changes 
                                    await package.SaveAsync();
                                }
                                else
                                {
                                    int newRow = worksheet.Dimension?.End.Row ?? 1;

                                    while (worksheet.Cells[newRow, 1]?.Value != null)
                                    {
                                        newRow++;
                                    }

                                    // Update the new row with the data
                                    worksheet.Cells[newRow, 1].Value = recipeId;
                                    worksheet.Cells[newRow, specificColumnIndex].Value = modifiedLinesList1[i];
                                    gemOPCController.LogInfo($"Added Recipe ID '{recipeId}' at Row {newRow} in Column {specificColumnIndex} with data '{modifiedLinesList1[i]}' for header '{modifiedLinesList[i]}'.");
                                }

                                // Save the changes
                                await package.SaveAsync();
                                gemOPCController.LogInfo("Row data updated successfully.");
                            }
                            else
                            {
                                gemOPCController.LogInfo($"HeaderDataValues '{modifiedLinesList[i]}' not found in the first row from column 2 to the end.");
                            }
                        }

                        gemOPCController.LogInfo("Row data updated successfully.");
                    }
                    else
                    {
                        gemOPCController.LogInfo("Worksheet is null.");
                    }
                }

                SECsMessage s7f4 = gemOPCController.GEMService.Services.CustomMessage.CreateMessage(7, 4, false);
                s7f4.DataItem.Add("ACK", 0, SECsFormat.Binary);
                gemOPCController.GEMService.SendReply(s7f4, transaction.Id);
            }
            catch (Exception ex)
            {
                gemOPCController.LogInfo($"Error: {ex.Message}");
                // Handle the exception or log the error as needed
            }
        }
        public string ServiceName
        {
            get { return "CustomRecipeDownloadService"; }
        }
    }
}

