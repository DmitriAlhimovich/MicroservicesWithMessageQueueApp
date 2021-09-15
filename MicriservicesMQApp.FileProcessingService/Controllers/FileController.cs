using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.VisualBasic.FileIO;

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
    }
}
