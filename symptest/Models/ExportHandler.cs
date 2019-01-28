using HH_client_manager.Models.Database;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using static symptest.Models.DataTableExtensions;

namespace symptest.Models
{
    public class ExportHandler
    {

        private static DataTable GenerateDummyClientsDataTable()
        {
            /// Returns a DataTable filled with dummy data on Clients.
            
            // random seeds and functions:
            Random r = new Random();
            string alpha = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz ";

            string randomString(Random rando, int length, string charbase)
            {
                var output = new char[length];
                for (int i = 0; i < output.Length; i++)
                {
                    output[i] = charbase[rando.Next(charbase.Length)];
                }
                return new String(output);
            }
            DateTime randomDateTime(Random rando)
            {
                DateTime start = new DateTime(1950, 1, 1);
                int range = (DateTime.Today - start).Days;
                return start.AddDays(rando.Next(range));
            }

            // create list of Clients with random field values
            List<Client> input = new List<Client>();
            for (int i = 0; i < r.Next(3, 12); i++)
            {
                input.Add(new Client("S" + r.Next(10000, 99999), randomDateTime(r), randomString(r, r.Next(3, 5), alpha), randomString(r, r.Next(7, 18), alpha), randomString(r, r.Next(20, 70), alpha), randomString(r, r.Next(4, 6), alpha), randomString(r, r.Next(4, 12), alpha), randomString(r, r.Next(4, 25), alpha), randomString(r, r.Next(4, 6), alpha)));
            }

            // convert list to DataTable, and return
            return input.ToDataTable();
        }

        public static async Task<byte[]> CreateExcelFileAsync(IHostingEnvironment hostingenviron, DataTable inputTable = null, string OutputFilename = "Export")
        {
            /// creates an Excel file, returning it to the user without changing the page
            /// currently creates a predefined dummy file
            /// 
            /// This code requires the following code (or something super-like it) in the relevant Controller:
            ///


            /*
            
            private readonly IHostingEnvironment _hostingEnvironment

            public HomeController(IHostingEnvironment environment)
            {
                _hostingEnvironment = environment;
            }
            
            public async Task<IActionResult> Export(DataTable inputTable, string OutputFilename = "Export")
            {// creates an Excel file, returning it to the user without changing the page
                byte[] memory = await ExportHandler.CreateExcelFileAsync(_hostingEnvironment, inputTable); // creates a dummy file if you don't include an inputTable (consider HH_client_manager.Models.Database.DataTableExtensions for an objectlist.ToDataTable<objecttype>() method)
                //send file in memory to user
                return File(memory, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", OutputFilename);
            }

            */


            // handle arguments
            if (inputTable != null){} else { // if no input table, generates dummy data on Clients to use instead
                inputTable = GenerateDummyClientsDataTable();
            }
            OutputFilename += ".xlsx"; // ensures filename ends in correct extension



            // begin actual method
            string sWebRootFolder = hostingenviron.WebRootPath; // the "wwwroot" folder location on server
            FileInfo file = new FileInfo(Path.Combine(sWebRootFolder, OutputFilename)); // create Excel file, place in wwwroot folder
            var memory = new MemoryStream(); // necessary in order to be able to delete file before giving file to user



            using (var fs = new FileStream(Path.Combine(sWebRootFolder, OutputFilename), FileMode.Create, FileAccess.Write))
            {
                using (var package = new ExcelPackage())
                {
                    // add sheet
                    ExcelWorksheet ws = package.Workbook.Worksheets.Add(inputTable.TableName); // uses name of DataTable as name of worksheet


                    //Excel file indices start on 1, not 0, in EPPlus 4.5.2. This is different in EPPlus 4.5.3, but 4.5.3 doesn't seem compatible with .NET Core 2.1, only 2.2
                    int row, col;
                    row = col = 1;
                    //\\

                    // table header values added to worksheet
                    foreach (DataColumn currentColumn in inputTable.Columns)
                    {
                        ws.Cells[row, col].Value = currentColumn.ColumnName;
                        col++;
                    }

                    // move current position in excel file to row after header row
                    row = 2;
                    col = 1;

                    // contents of each passed in instance
                    foreach (DataRow record in inputTable.Rows)
                    {
                        foreach (var field in record.ItemArray)
                        {
                            ws.Cells[row, col].Value = field;
                            col++;
                        }
                        col = 1;
                        row++;
                    }

                    //Autofit all
                    ws.Cells.AutoFitColumns(0);

                    // title file
                    package.Workbook.Properties.Title = "Generated by the Hubbard House Hope & Healing Dashboard System";

                    //finalize file
                    package.SaveAs(fs);
                }
            }

            // move file from saved-on-server to server-memory
            // TODO: cost/benefit analysis of current method vs finding way to delete file after returning it to browser
            using (var stream = new FileStream(Path.Combine(sWebRootFolder, OutputFilename), FileMode.Open))
            {
                await stream.CopyToAsync(memory);
            }

            //delete file on server after it's been placed in memory
            System.IO.File.Delete(Path.Combine(sWebRootFolder, OutputFilename));

            //send file in memory to user
            return memory.ToArray();

        }
    }
}
