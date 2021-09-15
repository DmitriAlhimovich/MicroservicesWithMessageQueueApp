using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.VisualBasic.FileIO;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace MicriservicesMQApp.FileProcessingService.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class FileController : ControllerBase
    {
        [HttpGet]
        public Task<double> GetAverageCultureValuesStars()
        {
            using TextFieldParser parser = new TextFieldParser("Content/employee_reviews.csv");
            parser.TextFieldType = FieldType.Delimited;
            parser.SetDelimiters(",");

            string[] headers = parser.ReadFields();

            var indexOfCultureValuesStars = headers.ToList().IndexOf("culture-values-stars");

            double sum = 0;
            int rowsCount = 1;

            while (!parser.EndOfData)
            {
                
                //Process row
                string[] fields = parser.ReadFields();

                if (double.TryParse(fields[indexOfCultureValuesStars], NumberStyles.AllowDecimalPoint,
                    new NumberFormatInfo() {CurrencyDecimalSeparator = "."}, out var value))
                {
                    sum += value;
                    rowsCount++;
                }
                else
                {
                    Console.WriteLine($"Failed parsing for {fields[indexOfCultureValuesStars]}, row {rowsCount}");
                }
            }
            return Task.FromResult(sum/rowsCount);
        }

        [HttpGet("excel")]
        public Task<string> ProcessExcel()
        {
            XSSFWorkbook xssfWorkbook;
            using (FileStream file = new FileStream("Content/airports.xlsx", FileMode.Open, FileAccess.Read))
            {
                xssfWorkbook = new XSSFWorkbook(file);
            }

            ISheet sheet = xssfWorkbook.GetSheetAt(0);
            StringBuilder sb = new StringBuilder();
            for (int row = 0; row <= sheet.LastRowNum; row++)
            {
                if (sheet.GetRow(row) != null) //null is when the row only contains empty cells 
                {
                    sb.Append( sheet.GetRow(row).GetCell(0) + "\n");
                }
            }

            return Task.FromResult(sb.ToString());
        }

    }
}
