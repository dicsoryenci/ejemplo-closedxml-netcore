using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;
using waEjemploExcel.Models;

namespace waEjemploExcel.Controllers
{
    public class HomeController : Controller
    {
        public IActionResult Index()
        {
            return View();
        }

        public IActionResult Personas()
        {


            return View();
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }

        public FileResult descargar() {
            Stream str = new MemoryStream();

            using (var workbook = new XLWorkbook())
            {
                
                var worksheet = workbook.Worksheets.Add("Sample Sheet");
                int n = 8;

                for (int y= 1; y<= 10; y++) {
                    worksheet.Cell(y, "A").Value = n;
                    worksheet.Cell(y, "B").Value = "*";
                    worksheet.Cell(y, "C").Value = y;
                    worksheet.Cell(y, "D").Value = "=";
                    worksheet.Cell(y, "E").Value = n*y;

                }

                
                //worksheet.Cell("A1").Value = "Hello World!";
                //worksheet.Cell("A2").FormulaA1 = "=MID(A1, 7, 5)";
                workbook.SaveAs(str);
            }

            str.Position = 0;
            byte[] fileByte = ReadFully(str);

            return File(fileByte, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "archivo.xlsx");
        }

        static byte[] ReadFully(Stream input)
        {
            byte[] buffer = new byte[16 * 1024];
            using (MemoryStream ms = new MemoryStream())
            {
                int read;
                while ((read = input.Read(buffer, 0, buffer.Length)) > 0)
                {
                    ms.Write(buffer, 0, read);
                }
                return ms.ToArray();
            }
        }
    }
}
