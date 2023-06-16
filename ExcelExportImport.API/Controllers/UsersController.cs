using ExcelExportImport.API.Entities;
using ExcelExportImport.API.Service;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;

namespace ExcelExportImport.API.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class UsersController : ControllerBase
    {

        [HttpGet("GetExcelList")]
        public IActionResult GetExcelList()
        {

            List<User> userList = new List<User>
            {
                new User {Id=Guid.NewGuid(), FirstName="Baran",LastName="Üçdağ",Email="baran.ucdag@warpiris.com",PasswordHash="Passowrd",PasswordSalt="Salt" ,Status=true},
                new User {Id=Guid.NewGuid(), FirstName="Baran",LastName="Üçdağ",Email="baran.ucdag@warpiris.com",PasswordHash="Passowrd",PasswordSalt="Salt" ,Status=true},
                new User {Id=Guid.NewGuid(), FirstName="Baran",LastName="Üçdağ",Email="baran.ucdag@warpiris.com",PasswordHash="Passowrd",PasswordSalt="Salt" ,Status=true},
                new User {Id=Guid.NewGuid(), FirstName="Baran",LastName="Üçdağ",Email="baran.ucdag@warpiris.com",PasswordHash="Passowrd",PasswordSalt="Salt" ,Status=true},
                new User {Id=Guid.NewGuid(), FirstName="Baran",LastName="Üçdağ",Email="baran.ucdag@warpiris.com",PasswordHash="Passowrd",PasswordSalt="Salt" ,Status=true}
            };

            Stream excelStream = ExcelProcessService.ExportToExcel(userList);


            // Excel dosyasını indirmek için gerekli adımlar
            byte[] excelBytes = ((MemoryStream)excelStream).ToArray();


            return Ok(File(excelBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"));
        }
    }
}
