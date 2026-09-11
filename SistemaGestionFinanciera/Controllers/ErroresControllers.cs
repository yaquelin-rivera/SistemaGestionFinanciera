using Microsoft.AspNetCore.Mvc;

namespace SistemaGestionFinanciera.Controllers
{
    public class ErroresController : Controller
    {
        [Route("Errores/404")]
        public IActionResult Error404()
        {
            return View();
        }
    }
}