using ContractorsWorkAPI.Model;
using ContractorsWorkAPI.Services;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;

namespace ContractorsWorkAPI.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class FilesController : ControllerBase
    {
        //private readonly IHttpContextAccessor _contextAccessor;
        private readonly IStorageService _storageService;
        
        public FilesController(IStorageService storageService)
        {
            _storageService = storageService;
        }

        [HttpPost]
        [Route("safefiles")]
        public async Task<IActionResult> SafeFilesAsync(IFormFile file)
        {
            try 
            {
                var a = await _storageService.SafeFiles(file);
                return Ok("Успешно");
            }
            catch(Exception ex) 
            {
                Console.WriteLine(ex.Message);
                return BadRequest("Ошибка сохранения");
            }
        }
    }
}
