using ClosedXML.Excel;
using DemoClosedXML.Model;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace DemoClosedXML.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class WeatherForecastController : ControllerBase
    {
        private static readonly string[] Summaries = new[]
        {
            "Freezing", "Bracing", "Chilly", "Cool", "Mild", "Warm", "Balmy", "Hot", "Sweltering", "Scorching"
        };

        private readonly ILogger<WeatherForecastController> _logger;

        public WeatherForecastController(ILogger<WeatherForecastController> logger)
        {
            _logger = logger;
        }

        [HttpGet()]
        public IEnumerable<WeatherForecast> Get()
        {
            var rng = new Random();
            return Enumerable.Range(1, 5).Select(index => new WeatherForecast
            {
                Date = DateTime.Now.AddDays(index),
                TemperatureC = rng.Next(-20, 55),
                Summary = Summaries[rng.Next(Summaries.Length)]
            })
            .ToArray();
        }

        [HttpGet("ClosedXml")]
        public ActionResult ClosedXml()
        {
         
            List<Student> students = new List<Student>()
            {
                new Student(){ Id="5951071113",Name="Phạm Trọng Trường", Age=22, Major="Công nghệ thông tin"  },
                  new Student(){ Id="5951071113",Name="Phạm Trọng Trường", Age=22, Major="Công nghệ thông tin"  },
                    new Student(){ Id="5951071113",Name="Phạm Trọng Trường", Age=22, Major="Công nghệ thông tin"  },
                      new Student(){ Id="5951071113",Name="Phạm Trọng Trường", Age=22, Major="Công nghệ thông tin"  },
                        new Student(){ Id="5951071113",Name="Phạm Trọng Trường", Age=22, Major="Công nghệ thông tin"  },
                          new Student(){ Id="5951071113",Name="Phạm Trọng Trường", Age=22, Major="Công nghệ thông tin"  },
            };  

            using (var wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add("Demo XLWorkbook");

                ws.Cell(7, 6).Value = "From Query";
                ws.Range(7, 6, 7, 10).Merge().AddToNamed("Titles");
                ws.Cell(1, 1).InsertTable(students);
                // Prepare the style for the titles
                var titlesStyle = wb.Style;
                titlesStyle.Font.Bold = true;
                titlesStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                titlesStyle.Fill.BackgroundColor = XLColor.Cyan;

                // Format all titles in one shot
                wb.NamedRanges.NamedRange("Titles").Ranges.Style = titlesStyle;

                ws.Columns().AdjustToContents();

                wb.SaveAs("Test.xlsx");
            }

                return Ok("Tạo thành công");
        }
    }
}
