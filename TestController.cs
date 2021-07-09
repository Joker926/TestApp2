using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace CoreAPI.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    [Authorize]
    public class TestController : ControllerBase
    {
        private readonly ILogger _logger;

        public TestController(ILoggerFactory loggerFactory)
        {
            _logger = loggerFactory.CreateLogger<TestController>();
        }


        [HttpGet]
        [Route("Download")]
        public async Task<IActionResult> Download(int id)
        {
            var response = GetDataFromDB(id).Result;
            var ms =  DownloadMS(response);
            string fileName = "Шаблон для загрузки - " + response.Title;
            return File(ms.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
        }

        public async Task<Response> GetDataFromDB(int id)
        {
            // тут происходит обращение к БД и возврат данных, для упрощения сгенерирован пример данных

                Response response = new Response
                {
                    Title = "Test",
                    Fields = new List<Field>
                    {
                        new Field  {
                            Title = "Field1",
                            IsVisible = true
                        }
                    }
                };

            return response;
        }

        public MemoryStream DownloadMS(Response entity)
        {
            using (var ms = new MemoryStream())
            {
                using (var package = new ExcelPackage())
                {
                    var tableName = "Empty";
                    if (!String.IsNullOrEmpty(entity.Title))
                    {
                        tableName = entity.Title;
                    }

                    var worksheet = package.Workbook.Worksheets.Add(tableName); //Worksheet name
                    var visibleFields = entity.Fields.Where(x => x.IsVisible);

                    var i = 1;
                    foreach (var item in visibleFields)
                    {
                        worksheet.Column(i).Width = 35;
                        worksheet.Cells[1, i].Value = item.Title;
                        worksheet.Cells[1, i].Style.Font.Bold = true;
                        i++;
                    }

                    package.SaveAs(ms);
                    ms.Position = 0;
                    return ms;
                }
            }
        }
    }

    public class Response
    {
        public string Title { get; set; }
        public List<Field> Fields { get; set; }
    }

    public class Field
    {
        public string Title { get; set; }
        public bool IsVisible { get; set; }
    }
}
