using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Insphere.Connectivity.Common;
using Insphere.Connectivity.Common.ToolModel;
using Insphere.Connectivity.Application.MessageServices;
using Insphere.Connectivity.Application.SecsToOpc;
using System.Threading;
using OfficeOpenXml; // open excel 
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using OfficeOpenXml.Table; //Open Excel table 
using System.IO;
using OfficeOpenXml.Drawing.Style.ThreeD;
using System.Text.RegularExpressions;
using System.Net.NetworkInformation;

namespace MyCustomHandler
{
    public class PPRUpload : IGEMOPCService
    {
        private GEMOPCController gemOPCController;

        public void Initialize(GEMOPCController controller)
        {
            this.gemOPCController = controller;
            gemOPCController.LogInfo("Inside PPUpload");
            gemOPCController.AddGEMSubscription(this, 7, 5); // Add GEM event subscription for Carrier Action Request

        }

        public void OPCNotification(string opcItem, object value, string logicalName)
        {
            throw new NotImplementedException();
        }


        public void SECsMessageNotification(string streamFunction, SECsTransaction transaction)
        {
            // Log the received stream function
            gemOPCController.LogInfo("Received: " + streamFunction);

            // Check if the stream function is equal to "S7F5"
            if (streamFunction == "S7F5")
            {
                // Log that "S7F5" has been processed
                gemOPCController.LogInfo("S7F5 processed");

                // Extract the recipeId from the transaction and trim it
                string recipeId = transaction.Primary.DataItem[0].ToString().Trim();

                // Open an Excel file
                using (var package = new ExcelPackage(new FileInfo("C:\\Users\\intern1\\Documents\\Programming\\RECIPE.xlsx")))
                {
                    // first worksheet from the Excel file
                    var worksheet = package.Workbook.Worksheets[0];

                    // Define variables for iterating through rows
                    int startRow = 2;
                    int endRow = worksheet.Dimension.End.Row;
                    int rowIndex = -1;

                    // Find the row index that matches the recipeId
                    for (int row = startRow; row <= endRow; row++)
                    {
                        // Get the value from the first column of the current row and trim it
                        var cellValue = worksheet.Cells[row, 1].Text.Trim();

                        // Define a regular expression pattern to match the trimmed recipeId
                        string pattern = @"^\s*" + Regex.Escape(recipeId) + @"\s*$";

                        // Check if the cell value matches the pattern (case insensitive)
                        if (Regex.IsMatch(cellValue, pattern, RegexOptions.IgnoreCase))
                        {
                            // Set the rowIndex to the current row and exit the loop
                            rowIndex = row;
                            break;
                        }
                    }
                    SECsMessage s7f6 = this.gemOPCController.GEMService.Services.CustomMessage.CreateMessage(7, 6, false);
                    // Check if a matching row was found
                    if (rowIndex != -1)
                    {
                        // Initialize lists to store header and data values
                        List<string> rowData = new List<string>();   
                        List<string> headerData = new List<string>();

                        // Loop through the columns of the worksheet (skipping the first column)
                        for (int col = 2; col <= worksheet.Dimension.End.Column; col++)
                        { 
                            // Get the header value from the first row of the current column and trim it
                            var headerValue = worksheet.Cells[1, col].Text.Trim();

                            // Get the cell value from the matching row and current column
                            var cellValue = worksheet.Cells[rowIndex, col].Text;

                            // Skip columns with empty, "0," or "nil" headers
                            if (!string.IsNullOrWhiteSpace(headerValue))  //&& headerValue != "0" && headerValue != "nil"
                            {
                                if (!string.IsNullOrWhiteSpace(cellValue))// && cellValue != "0" && cellValue != "nil"
                                {
                                    headerData.Add(headerValue);
                                    rowData.Add(cellValue);     
                                }
                            }
                        }

                        // Check if a matching row was found
                        if (rowIndex != -1)
                        {
                            // Create an SECsMessage object with message type 7, subtype 6, and not an acknowledge message


                            // Log a message indicating that strings have been joined
                            gemOPCController.LogInfo("JoinedString");

                            // Add a list to the message's data item
                            s7f6.DataItem.AddList(); 

                            // Add "PPID" as the header and the constructed header string as data
                            s7f6.DataItem[0].Add("PPID", recipeId, SECsFormat.Ascii);

                            // Split the header and data strings into individual parts
                            string[] headerParts = string.Join(",", headerData).Split(',');
                            string[] dataParts = string.Join(",", rowData).Split(',');
                            string[] ppbody = new string[15500];
                            string[] array = new string[15500];

                            
                            // Number of "PPBody" entries matches the number of data parts
                            if (headerParts.Length == dataParts.Length)
                            {
                                // Loop through the header and data parts
                                for (int i = 0; i < headerParts.Length; i++)
                                {
                                    // Add "PPBody" as the header and the corresponding data part as data
                                    ppbody[i] = headerParts[i] + ": " + dataParts[i];
                                    array[i] = ppbody[i];


                                }
                                int startIndex = 0;
                                int count = headerParts.Length;

                                ArraySegment<string> segment = new ArraySegment<string>(array, startIndex, count);
                                string joinedString = String.Join("\r\n ", segment);
                                gemOPCController.LogInfo("JoinedString");
                                //s7f6.DataItem.AddList();
                                s7f6.DataItem[0].Add("PPBody", joinedString, SECsFormat.Ascii);

                                // Log a message indicating that the message is about to be sent
                                gemOPCController.LogInfo("BeforeSendReply");

                                // Send the constructed SECsMessage as a reply using the transaction ID
                                gemOPCController.GEMService.SendReply(s7f6, transaction.Id);
                            }
                            else
                            {
                                // Handle the case where the number of headers and data parts does not match
                                gemOPCController.LogInfo("Header and data parts count mismatch");
                                // You may want to add error handling here or send an error message as needed.
                            }
                        }
                    }
                    else
                    {
                        // Create an error message and send it as a reply
                        SECsMessage errorMessage = this.gemOPCController.GEMService.Services.CustomMessage.CreateMessage(1, 0, true);

                        // Add an error description to the message
                        //errorMessage.DataItem.Add();
                        s7f6.DataItem.AddList();

                        // Send the error message as a reply using the transaction ID
                        //gemOPCController.GEMService.SendReply(errorMessage, transaction.Id);
                        gemOPCController.GEMService.SendReply(s7f6, transaction.Id);
                    }
                }
            }
            else
            {
                // Log that the stream function is unknown
                gemOPCController.LogInfo("Unknown streamFunction: " + streamFunction);
            }
        }






        public string ServiceName
        {
            get { return "PPRUpload"; }
        }
    }
}
