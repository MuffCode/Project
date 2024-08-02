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
using System.Data.SqlClient;

namespace MyCustomHandler
{
    public class EnRemoteCommand : IGEMOPCService
    {
        private GEMOPCController gemOPCController;


        public void Initialize(GEMOPCController controller)
        {
            this.gemOPCController = controller;
            gemOPCController.LogInfo("Inside EnRemoteCommand");
            gemOPCController.AddGEMSubscription(this, 2, 41);
        }

        public void OPCNotification(string opcItem, object value, string logicalName)
        {
            throw new NotImplementedException();
        }

        public async void SECsMessageNotification(string streamFunction, SECsTransaction transaction)
        {
            // Handle the S2F41 event asynchronously
            await ProcessS2F41Async(transaction);
        }

        private async Task ProcessS2F41Async(SECsTransaction transaction)
        {


            //SECsMessage s2f42 = this.gemOPCController.GEMService.Services.CustomMessage.CreateMessage(2, 42, false);
            //s2f42.DataItem.AddList(); // Sending S2f4

            // Read data from excel
            using (var package = new ExcelPackage(new FileInfo("C:\\Users\\intern1\\Documents\\Programming\\RECIPE.xlsx")))
            {
                var worksheet = package.Workbook.Worksheets[0]; // accessing the number of sheets, for this case will be the first sheet only
                var recipeIdsToCheck = new List<string>();  // Original list of recipe IDs

                // Extract recipe IDs from the transaction if available
                for (int i = 0; i < transaction.Primary.DataItem.Count; i++)
                {
                    for (int j = 0; j < transaction.Primary.DataItem[i].Count; j++)
                    {
                        var recipeId = transaction.Primary.DataItem[i][j]?.ToString()?.Trim();
                        if (!string.IsNullOrEmpty(recipeId))
                        {
                            recipeIdsToCheck.Add(recipeId);
                        }
                    }
                }

                gemOPCController.LogInfo($"recipeIdsToCheck: {string.Join(", ", recipeIdsToCheck)}");

                if (recipeIdsToCheck.Contains("PP-SELECT"))
                {
                    gemOPCController.LogInfo("Start2"); 
                    var newRecipeIdsList = new List<List<string>>();  // List to store new recipe IDs

                    bool matchFound = false;  // Flag to track if a match is found

                    bool secondAsciiRecipeIdLogged = false;

                    // Iterate through each row in the Excel file and check the second ASCII recipe ID
                    for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                    {
                        var excelRecipeId = worksheet.Cells[row, 1]?.Value?.ToString()?.Trim(); 
                        if (!string.IsNullOrEmpty(excelRecipeId))
                        {
                            var secondAsciiRecipeIdObject = transaction.Primary.DataItem[0][1][0][1]?.ToString()?.Trim();
                            var secondAsciiRecipeId = (secondAsciiRecipeIdObject != null) ? secondAsciiRecipeIdObject.ToString().Trim() : null;

                            if (!secondAsciiRecipeIdLogged && !string.IsNullOrEmpty(secondAsciiRecipeId))
                            {
                                gemOPCController.LogInfo($"secondAsciiRecipeId: {secondAsciiRecipeId}");
                                secondAsciiRecipeIdLogged = true;   
                            }

                            if (!string.IsNullOrEmpty(secondAsciiRecipeId) && excelRecipeId.Equals(secondAsciiRecipeId, StringComparison.OrdinalIgnoreCase))
                            {
                                int matchedRow = row;

                                var newRecipeIdPair = new List<string>
                                {
                                    "PPID",
                                    secondAsciiRecipeId,
                                    matchedRow.ToString()
                                };
                                newRecipeIdsList.Add(newRecipeIdPair); 
                                gemOPCController.LogInfo($"Found match: {secondAsciiRecipeId} at row {matchedRow}");

                                string[] array = new string[1];
                                int startIndex = 0;
                                int count = 1;

                                ArraySegment<string> segment = new ArraySegment<string>(array, startIndex, count);

                                string joinedString = string.Join("\r\n", segment);
                                joinedString += $"\r\n  {matchedRow}";

                                //s2f42.DataItem[0].Add($"", joinedString, SECsFormat.Ascii); 

                                int arkNumber;

                                if (matchedRow >= 2)
                                {
                                    arkNumber = matchedRow;
                                    gemOPCController.LogInfo($"At row {matchedRow}");
                                    
                                    await Task.Run(() => gemOPCController.WriteToOPC("ns=2;i=0", arkNumber));
                                    await Task.Run(() => gemOPCController.WriteToOPC("ns=2;i=1", 1));
                                    await Task.Delay(2000); 
                                    await Task.Run(() => gemOPCController.WriteToOPC("ns=2;i=1", 0));
                                    matchFound = true;
                                    await SendS2F42ReplyAsync(transaction,matchFound);
 
                                }
                                else
                                {
                                    gemOPCController.LogInfo("No match found in Excel data.");
                                    //  string[] array = new string[1];
                                    //  int startIndex = 0;
                                    //  int count = 1;

                                    //  ArraySegment<string> segment = new ArraySegment<string>(array, startIndex, count);

                                    //  string joinedString = string.Join("\r\n", segment);
                                    matchFound = false;
                                    await SendS2F42ReplyAsync(transaction, matchFound);

                                }

                            }

                           // matchFound = true;
                            //break;
                        }
                    }

                    if (!matchFound)
                    {
                        gemOPCController.LogInfo("No match found in Excel data.");
                        string[] array = new string[1];
                        int startIndex = 0;
                        int count = 1;

                        ArraySegment<string> segment = new ArraySegment<string>(array, startIndex, count);

                        string joinedString = string.Join("\r\n", segment);
                        joinedString += $"\r\n No match found in Excel data";
                        matchFound = false;
                        await SendS2F42ReplyAsync(transaction, matchFound);
 

                        //s2f42.DataItem[0].Add($"", joinedString, SECsFormat.Ascii);
                    }

                    gemOPCController.LogInfo("s2f42");
                }
                else
                {
                    gemOPCController.LogInfo("'PP-SELECT' is not in the list. Continuing with other logic.");
                }
                
                gemOPCController.LogInfo("BeforeSendReply");
            }
        }

      //  readonly bool matchFound = true;  // Flag to track if a match is found
        async Task SendS2F42ReplyAsync(SECsTransaction transaction, bool matchFound)
        {

            SECsMessage s2f42 = this.gemOPCController.GEMService.Services.CustomMessage.CreateMessage(2, 42, false);

            if (matchFound) // Assuming matchFound is a boolean variable indicating whether a match is found or not
            {
                object ack = await Task.Run(() => gemOPCController.ReadOPC("ns=2;i=2"));
                s2f42.DataItem.Add("ACK", ack, SECsFormat.Binary);
                gemOPCController.LogInfo("ACK" + ack.ToString());
                gemOPCController.GEMService.SendReply(s2f42, transaction.Id);
            }
            else
            {
                object ack = 2;
                s2f42.DataItem.Add("ACK", ack, SECsFormat.Binary);
                gemOPCController.LogInfo("ACK" + ack.ToString());
                gemOPCController.GEMService.SendReply(s2f42, transaction.Id);
            }

        }

        public string ServiceName
        {
            get { return "EnRemoteCommand"; }
        }
    }
}