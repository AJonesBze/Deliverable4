using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using symptest.Models;

namespace symptest.Controllers
{
    public class ExportDemoController : Controller
    {
        [AllowAnonymous] // don't require login to see this demo-only page
        public IActionResult Index()
        {
            return View();
        }

        private readonly IHostingEnvironment _hostingEnvironment;

        public ExportDemoController(IHostingEnvironment environment)
        {
            _hostingEnvironment = environment;
        }

        public async Task<IActionResult> Export(DataTable inputTable, string OutputFilename = "Text")
        {// creates an Excel file, returning it to the user without changing the page
            byte[] memory = await ExportHandler.CreateExcelFileAsync(_hostingEnvironment); // creates a dummy file if you don't include an inputTable argument (consider symptest.Models.DataTableExtensions for an objectlist.ToDataTable<objecttype>() method)
                                                                                                       //send file in memory to user
            return File(memory, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", OutputFilename+".xlsx");
        }

        public async Task<IActionResult> ExportChart(DataTable inputTable, string OutputFilename = "Chart")
        {// creates an Excel file, returning it to the user without changing the page
            string[] selected_columns_for_chart = { "Stay #", "Assessment 1", "Assessment 2" };
            byte[] memory = await ExportHandler.CreateExcelFileAsync(_hostingEnvironment, ExportHandler.GenerateDummyClientReportDataTable(), selected_columns_for_chart); // creates a dummy file if you don't include an inputTable argument (consider symptest.Models.DataTableExtensions for an objectlist.ToDataTable<objecttype>() method)
                                                                                                                                                                           //send file in memory to user
            return File(memory, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", OutputFilename + ".xlsx");
        }
    }
}