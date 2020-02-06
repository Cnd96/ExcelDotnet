using System;
using System.Data;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using Microsoft.AspNetCore.Http;
using  Newtonsoft.Json.Linq ;
namespace dotnetExcel.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class excelRead:ControllerBase
    { 
        [HttpPost("fromPath")]
         public async Task<ActionResult<IEnumerable<IEnumerable<string>>>> Get()
        {
            try{
              using (var reader = new StreamReader(Request.Body))
             {
                var body =await reader.ReadToEndAsync();
                var details = JObject.Parse(body);  
                string path =details.Value<string>("path");
                //  string path ="C:\\Users\\ChamalD\\Desktop\\Projects\\VMS\\aaa.xlsx";
                FileInfo fileInfo = new FileInfo(path);

                ExcelPackage package = new ExcelPackage(fileInfo);
                ExcelWorksheet worksheet =  package.Workbook.Worksheets.FirstOrDefault();
    
                List<string[]> excelDataList = new List<string[]>();
                // // get number of rows and columns in the sheet
                int rows = worksheet.Dimension.Rows;  
                int columns = worksheet.Dimension.Columns;

                // // loop through the worksheet rows and columns
                for (int i = 1; i <= rows; i++) {
                    List<string> rowData = new List<string>();
                    for (int j = 1; j <= columns; j++) {
                        if((worksheet.Cells[i, j]).Value!=null){
                        // string content = worksheet.Cells[i, j].Value.ToString();
                           rowData.Add((worksheet.Cells[i, j]).Value.ToString());
                        // Console.WriteLine(intList[0]);
                        /* Do something ...*/
                        }
                        else{
                            rowData.Add("");
                        }
                    }
                    excelDataList.Add(rowData.ToArray());
                }
                return excelDataList;
             }
            }catch(Exception exe){
                 return BadRequest("Error"+exe);
            }
              
        }
        [HttpPost("upload")]  
        public async Task<ActionResult<IEnumerable<IEnumerable<string>>>> Import(IFormFile formFile,[FromQuery]string sheetNumber)  
        {  
            try{
                if (formFile == null || formFile.Length <= 0)  
                {  
                    return BadRequest("Empty Data");
                }  
        
                if (!Path.GetExtension(formFile.FileName).Equals(".xlsx", StringComparison.OrdinalIgnoreCase))  
                {  
                    return BadRequest("Not an excel sheeet");
                }  
                List<string[]> excelDataList = new List<string[]>();
            
                using (var stream = new MemoryStream())  
                {  
                    await formFile.CopyToAsync(stream);  
            
                    using (var package = new ExcelPackage(stream))  
                    {  
                        ExcelWorksheet worksheet = package.Workbook.Worksheets[Int32.Parse(sheetNumber)];  
                        int rows = worksheet.Dimension.Rows;  
                        int columns = worksheet.Dimension.Columns; 
            
                        for (int i = 1; i <= rows; i++) {
                            List<string> rowData = new List<string>();
                            for (int j = 1; j <= columns; j++) {
                                if((worksheet.Cells[i, j]).Value!=null){
                                rowData.Add((worksheet.Cells[i, j]).Value.ToString());
                                }
                                else{
                                    rowData.Add("");
                                }
                            }
                            excelDataList.Add(rowData.ToArray());
                        }
                    }  
                }  
                return excelDataList;
            }catch(Exception exe){
                 return BadRequest("Error"+exe);
            }
           
        }  
    }
}