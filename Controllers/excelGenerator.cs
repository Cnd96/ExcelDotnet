using System;
using System.Text;
using System.Data;
using System.Threading.Tasks;
using System.IO;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using  Newtonsoft.Json.Linq ;

namespace dotnetExcel.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class excelGenerator:ControllerBase
    { 
        [HttpPost()]
        public async Task<IActionResult> generateExcelFile([FromQuery]string appName)
        {
            using (StreamReader reader = new StreamReader(Request.Body, Encoding.UTF8))
            {  
                var data= await reader.ReadToEndAsync();
                if(data==""){
                    return BadRequest("Empty Data");
                }
                JArray json = JArray.Parse(data) as JArray;
                DataTable dataTable = new DataTable();
                int headersLength=0;
                int rowslength=0;

                foreach (var header in json[0]){
                    headersLength++;
                    dataTable.Columns.Add(header.ToString(), typeof(string));
                    // Console.WriteLine(header.ToString());
                }
                foreach (var obj in json)
                {
                    rowslength++;
                }
                for(int j=1;j<rowslength;j++)
                {
                    
                    DataRow dr = dataTable.NewRow();
                    for(int i=0;i<headersLength;i++){
                        dr[i]=json[j][i].ToString();
                        //  Console.WriteLine(json[j][i].ToString());
                    }
                    dataTable.Rows.Add(dr);
                }
                
                await Task.Yield();  
                var stream = new MemoryStream();  
            
                using (var package = new ExcelPackage(stream))  
                {  
                    var worksheet = package.Workbook.Worksheets.Add("Sheet1");
                    worksheet.Cells.LoadFromDataTable(dataTable, true);
                    package.Save();  
                }  
                stream.Position = 0;  
                string excelName = $"{DateTime.Now.ToString("yyyyMMddHHmmssfff")}.xlsx";  
                return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", excelName);  
            }
        }
         [HttpGet]
        public string getdata()
        { 
            return "scdcsd";
           
        }
        
    }
}