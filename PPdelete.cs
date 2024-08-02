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
using static System.Net.Mime.MediaTypeNames;
using System.Diagnostics;
using System.Windows.Media.Media3D;

namespace MyCustomHandler
{
    public class PPdelete : IGEMOPCService
    {
        //private string ServiceName = "delete";
        private GEMOPCController gemOPCController;  
        //private SECsDataItem PPID_delete;

        public void Initialize(GEMOPCController controller)
        {
            this.gemOPCController = controller;
            gemOPCController.LogInfo("Inside PPdelete");
            gemOPCController.AddGEMSubscription(this, 7, 17);
        }

        public void OPCNotification(string opcItem, object value, string logicalName)
        {
            throw new NotImplementedException();
        }

        public async void SECsMessageNotification(string streamFunction, SECsTransaction transaction)
        {
            {
                // Handle the S2F41 event asynchronously
                await ProcessS7F17Async(transaction);
            }
        }

        private async Task ProcessS7F17Async(SECsTransaction transaction)
        {
            
            {
                gemOPCController.LogInfo("Start1");
                gemOPCController.LogInfo("S7F17 processed");

                var recipeIdsToDelete = new List<string>();

                // Extract recipe IDs from the transaction if available
                if (transaction.Primary.DataItem.Count > 0 && transaction.Primary.DataItem[0].Count > 0)
                {
                    for (int i = 0; i < transaction.Primary.DataItem[0].Count; i++)
                    {
                        recipeIdsToDelete.Add(transaction.Primary.DataItem[0][i]?.ToString()?.Trim());
                    }
                }
                else
                {
                    // If no recipe IDs are provided, delete all rows in the Excel file
                    gemOPCController.LogInfo("No recipe IDs provided. Deleting all rows in Excel.");
                }

                using (var package = new ExcelPackage(new FileInfo("C:\\Users\\intern1\\Documents\\Programming\\RECIPE.xlsx")))
                {
                    var worksheet = package.Workbook.Worksheets.FirstOrDefault();

                    if (worksheet != null)
                    {
                        // If recipe IDs are provided, delete specific rows; otherwise, delete all rows
                        if (recipeIdsToDelete.Any())
                        {
                            // Delete rows for specific recipe IDs
                            foreach (var recipeIdToDelete in recipeIdsToDelete)
                            {
                                bool recipeIdMatched = false;
                                for (int row = 2; row <= worksheet.Dimension.End.Row; row++) 
                                {
                                    var recipeCell = worksheet.Cells[row, 1]?.Value?.ToString()?.Trim();
                                    if (string.Equals(recipeCell, recipeIdToDelete, StringComparison.OrdinalIgnoreCase))
                                    {
                                        // Clear the row and set '0' in all columns except the first column of the cleared row
                                        worksheet.Cells[row, 1, row, worksheet.Dimension.End.Column].Clear();
                                        for (int column = 2; column <= worksheet.Dimension.End.Column; column++)
                                        {
                                            worksheet.Cells[row, column].Value = 0;
                                        }
                                        gemOPCController.LogInfo($"Cleared in row {row}.");

                                        recipeIdMatched = true;
                                        Sends7f18Reply(transaction, recipeIdMatched);
                                    }
                                }

                                if (!recipeIdMatched)
                                {
                                    recipeIdMatched = false;
                                    gemOPCController.LogInfo($"Recipe ID '{recipeIdToDelete}' does not match. Stopping further processing.");
                                    Sends7f18Reply(transaction, recipeIdMatched);

                                    //  break; // Stop processing if the recipe ID does not match
                                }

                            }

                        }
                        else
                        {
                            // Delete all rows except the first row
                            for (int row = worksheet.Dimension.End.Row; row > 1; row--)
                            {
                                // Clear the row and set '0' in all columns except the first column of the cleared row
                                worksheet.Cells[row, 1, row, worksheet.Dimension.End.Column].Clear();
                                for (int column = 2; column <= worksheet.Dimension.End.Column; column++)
                                {
                                    worksheet.Cells[row, column].Value = 0;
                                }

                                gemOPCController.LogInfo($"Cleared and set '0' in row {row}.");


                                SECsMessage s7f18 = gemOPCController.GEMService.Services.CustomMessage.CreateMessage(7, 4, false);
                                s7f18.DataItem.Add("ACK", 0, SECsFormat.Binary);
                                gemOPCController.GEMService.SendReply(s7f18, transaction.Id);
                            }
                        }

                        // Save the changes 
                        await package.SaveAsync();
                        gemOPCController.LogInfo("Rows deleted successfully.");
                    }
                    else
                    {
                        gemOPCController.LogInfo("Worksheet is null.");
                    }
                }

            }
        }

        void Sends7f18Reply(SECsTransaction transaction, bool recipeIdMatched)
        {

            SECsMessage s7f18 = this.gemOPCController.GEMService.Services.CustomMessage.CreateMessage(7, 18, false);

            if (recipeIdMatched) // Assuming matchFound is a boolean variable indicating whether a match is found or not
            {
                object ack = 0;
                gemOPCController.LogInfo("S7F4");
                s7f18.DataItem.Add("ACK", ack, SECsFormat.Binary);
                gemOPCController.GEMService.SendReply(s7f18, transaction.Id);
            }
            else
            {
                object ack = 2;
                s7f18.DataItem.Add("ACK", ack, SECsFormat.Binary);
                gemOPCController.LogInfo("ACK" + ack.ToString());
                gemOPCController.GEMService.SendReply(s7f18, transaction.Id);
            }

        }

        public string ServiceName
        {
            get { return "CustomPPdeleteService"; }
        }
    }
}