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
        #region Dummy Table generation methods
        public static DataTable GenerateDummyClientsDataTable()
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

        public static DataTable GenerateDummyClientReportDataTable()
        {
            // random seeds and functions:
            Random r = new Random();

            // table and its column setup
            DataTable output_table = new DataTable("Assessments");
            output_table.Columns.Add(new DataColumn("Stay #"));
            output_table.Columns.Add(new DataColumn("Assessment 1", typeof(int)));
            output_table.Columns.Add(new DataColumn("Assessment 2", typeof(int)));

            // fill rows with random data
            DataRow row;
            for (int i = 0; i <= r.Next(2,5); i++)
            {
                row = output_table.NewRow();
                foreach(DataColumn col in output_table.Columns)
                {
                    if(col.ColumnName == "Stay #")
                    {
                        row[col] = "Stay " + (1+i);
                    }
                    else
                    {
                        row[col] = r.Next(4, 35);
                    }
                }
                output_table.Rows.Add(row);
            }
            return output_table;
        }

        #endregion

        public static async Task<byte[]> CreateExcelFileAsync(IHostingEnvironment hostingenviron, DataTable inputTable = null, string[] ChartAxes = null)
        {
            /// creates an Excel file, returning it to the user without changing the page
            /// if inputTable is left as null, it will create a dummy DataTable to use instead
            /// 
            /// Notes about the ChartAxes parameter:
            /// - No chart sheet if ChartAxes parameter is left as null.
            /// - The first listed string in its array is the Y axis set. Each subsequent string in the array is one of the sets that make up the X axis.
            /// 
            /// This code requires the following code (or something super-like it) in the relevant Controller:
            ///

            #region Copy/pastable code for Contollers

            /*
            
            private readonly IHostingEnvironment _hostingEnvironment;

            public HomeController(IHostingEnvironment environment)
            {
                _hostingEnvironment = environment;
            }
            
            public async Task<IActionResult> Export(DataTable inputTable, string OutputFilename = "Export")
            {// creates an Excel file, returning it to the user without changing the page
                byte[] memory = await ExportHandler.CreateExcelFileAsync(_hostingEnvironment, inputTable); // creates a dummy file if you don't include an inputTable (consider HH_client_manager.Models.Database.DataTableExtensions for an objectlist.ToDataTable<objecttype>() method)
                //send file in memory to user
                return File(memory, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", OutputFilename+".xlsx");
            }

            */

            #endregion

            #region Process parameters received

            // if no input table, generates dummy data on Clients to use instead
            if (inputTable != null){} else { 
                inputTable = GenerateDummyClientsDataTable();
            }

            #endregion 

            ///// begin actual method
            string OutputFilename = "ToExport.xlsx"; // this will not be the filename the person downloading will see, that's handled by the Controller method and doesn't involve this function
            string sWebRootFolder = hostingenviron.WebRootPath; // the "wwwroot" folder location on server
            FileInfo file = new FileInfo(Path.Combine(sWebRootFolder, OutputFilename)); // create Excel file, place in wwwroot folder
            var memory = new MemoryStream(); // necessary in order to be able to delete file before giving file to user
            const int INDEXBASE = 1; //Excel file indices start on 1, not 0, in EPPlus 4.5.2. This is different in EPPlus 4.5.3, but 4.5.3 doesn't seem compatible with .NET Core 2.1, only 2.2


            using (var fs = new FileStream(Path.Combine(sWebRootFolder, OutputFilename), FileMode.Create, FileAccess.Write))
            {
                using (var package = new ExcelPackage())
                {
                    #region generate data sheet
                    ExcelWorksheet ws = package.Workbook.Worksheets.Add(inputTable.TableName); // uses name of DataTable as name of worksheet

                    // starting position for per-cell runthrough (copy/paste datatable to worksheet)
                    int row, col;
                    row = col = INDEXBASE;
                    int column_count = 0;
                    //\\

                    // table header values added to worksheet
                    foreach (DataColumn currentColumn in inputTable.Columns)
                    {
                        ws.Cells[row, col].Value = currentColumn.ColumnName;
                        col++;
                    }
                    column_count = col - 1; // now we know the number of columns, for chart section later

                    // move current position in excel file to row after header row
                    row = INDEXBASE+1;
                    col = INDEXBASE;

                    // contents of each passed in instance
                    foreach (DataRow record in inputTable.Rows)
                    {
                        foreach (var field in record.ItemArray)
                        {
                            ws.Cells[row, col].Value = field;
                            col++;
                        }
                        col = INDEXBASE;
                        row++;
                    }

                    //Autofit all
                    ws.Cells.AutoFitColumns(0);
                    #endregion

                    #region generate chart sheet processes
                    if (ChartAxes != null) // so long as the string[] parameter for what axes the chart should have isn't still null
                    {
                        if (ChartAxes.Count() >= 2) // need at least 2 axes to make a chart
                        { 
                            ExcelWorksheet ws_chart = package.Workbook.Worksheets.Add("Chart"); // we need us a separate worksheet for the chart
                            var diagram = ws_chart.Drawings.AddChart("chart", OfficeOpenXml.Drawing.Chart.eChartType.ColumnClustered); // create chart in given worksheet. TODO: currently only does bar charts, will have line-over-time charts set next deliverable
                            char[] Letter = " ABCDEFGHIJKLMNOPQRSTUVWXYZ".ToCharArray(); // resource for address mapping. KNOWN BUG: if you go over 26 columns, incorrect data set range definitions (A1, Z2, etc.) will occur. Starts wil space because index base here is 1. Therefore, Letter[1] is 'A'

                            DataTable chartdata;
                            chartdata = inputTable.FilterByColumns(ChartAxes); // get rid of columns not on the graph, and order columns right
                            ws_chart.Cells["A1"].LoadFromDataTable(chartdata, true); // drop into the worksheet the data, as-is

                            for (int i = INDEXBASE + 1; i <= chartdata.Rows.Count + 1; i++) // for each row in the worksheet data that is not part of the header row
                            {

                                string row_dataseries = ""; // intended output like: "B2:D2", "B3:D3", "B4:D4", etc.



                                for (int j = 2; j <= chartdata.Columns.Count; j++)
                                {

                                    row_dataseries += Letter[j].ToString() + i;

                                    if (j != chartdata.Columns.Count)
                                    {
                                        row_dataseries += ":";
                                    }

                                }


                                string row_dataseries_titles = Letter[2].ToString() + "1"; // intended output like: "B1:D1"

                                if (ChartAxes.Count() > 2)
                                {

                                    row_dataseries_titles += ":" + Letter[ChartAxes.Count()].ToString() + "1";

                                }

                                var series = diagram.Series.Add(row_dataseries, row_dataseries_titles);
                                series.Header = ws_chart.Cells[$"A{i}"].Value.ToString();

                            }
                            diagram.Border.Fill.Color = System.Drawing.Color.Green; // green border


                            ws_chart.Cells.AutoFitColumns(0);


                        }
                    }
                    #endregion

                    // title file
                    package.Workbook.Properties.Title = "Generated by the Hubbard House Hope & Healing Dashboard System";

                    // finalize file
                    package.SaveAs(fs);
                }
            }

            // move file from saved-on-server to server-memory
            // TODO: cost/benefit analysis of current method vs finding way to delete file after returning it to browser
            using (var stream = new FileStream(Path.Combine(sWebRootFolder, OutputFilename), FileMode.Open))
            {
                await stream.CopyToAsync(memory);
            }

            // delete file on server after it's been moved to memory
            System.IO.File.Delete(Path.Combine(sWebRootFolder, OutputFilename));

            // ship out
            return memory.ToArray();

        } // end CreateExcelFileAsync method
    }
}
