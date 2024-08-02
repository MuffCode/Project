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
using OfficeOpenXml; //Open Excel 
using OfficeOpenXml.Table; //Open Excel table 
using System.IO;

namespace MyCustomHandler
{
    public class CurrentEPPDRequest : IGEMOPCService
    {
        private GEMOPCController gemOPCController;

        public void Initialize(GEMOPCController controller)
        {
            this.gemOPCController = controller;
            gemOPCController.LogInfo("Inside CurrentEPPDRequest");
            gemOPCController.AddGEMSubscription(this, 7, 19); // Add GEM event subscription for Carrier Action Request

        }

        public void OPCNotification(string opcItem, object value, string logicalName)
        {

        }


        public void SECsMessageNotification(string streamFunction, SECsTransaction transaction)
        {
            gemOPCController.LogInfo("Received: " + streamFunction);

            if (streamFunction == "S7F19")
            {
                gemOPCController.LogInfo("S7F19 processed");
                // Sending S7F20 
                //<L[2]

                SECsMessage s7f20 = this.gemOPCController.GEMService.Services.CustomMessage.CreateMessage(7, 20, false);
                s7f20.DataItem.AddList(); // Sending S7F20

                // Read data from excel
                using (var package = new ExcelPackage(new FileInfo("C:\\Users\\intern1\\Documents\\Programming\\RECIPE.xlsx")))
                {
                    var worksheet = package.Workbook.Worksheets[0]; // accessing the number of sheets, for this case will be the first sheet only
                    var startRow = 2; // Data starts from the second row

                    for (int row = startRow; row <= worksheet.Dimension.End.Row; row++) // using for loop to loop every row to ensure all rows are read
                    {
                        var recipeValue = worksheet.Cells[row, 1].Text.Trim(); // Data will be in column A

                        if (!string.IsNullOrEmpty(recipeValue)) // string = executed, empty string = not executed 
                        {
                            // Add the recipe value to the S7F20 message
                            s7f20.DataItem[0].Add("Recipe" + (row - startRow + 1), recipeValue, SECsFormat.Ascii);
                        }
                    }
                }

                gemOPCController.GEMService.SendReply(s7f20, transaction.Id); // send message to s7f20
                                      


            }

        }

        public string ServiceName
        {
            get { return "CurrentEPPDRequest"; }
        }
    }
}
