using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using SistemaGestionFinanciera.Data;
using SistemaGestionFinanciera.Models.ViewModels;

namespace SistemaGestionFinanciera.ViewComponents;

public class AlertaCuentasPorPagarViewComponent : ViewComponent
{
    private readonly SistemaFinancieroContext _context;

    public AlertaCuentasPorPagarViewComponent(
        SistemaFinancieroContext context)
    {
        _context = context;
    }

    public async Task<IViewComponentResult> InvokeAsync()
    {
        string rol =
            HttpContext.Session.GetString("Rol") ?? "";

        var modelo =
            new AlertaPresupuestoMenuViewModel();

        bool puedeGestionarCobros =
            rol.Equals(
                "Administrador",
                StringComparison.OrdinalIgnoreCase) ||
            rol.Equals(
                "Contador",
                StringComparison.OrdinalIgnoreCase) ||
            rol.Equals(
                "Tesorero",
                StringComparison.OrdinalIgnoreCase);

        if (!puedeGestionarCobros)
        {
            return View(modelo);
        }

        int cantidadVencidas =
     await _context.CuentasPorPagars
         .AsNoTracking()
         .CountAsync(c =>
             c.Activo &&
             c.Estado == "Vencida" &&
             c.SaldoPendiente > 0);

        if (cantidadVencidas <= 0)
        {
            return View(modelo);
        }

        modelo.Mostrar = cantidadVencidas > 0;
        modelo.Tipo = "roja";
        modelo.Cantidad = cantidadVencidas;

        modelo.Mensaje =
            cantidadVencidas == 1
                ? "Hay 1 cuenta por pagar vencida pendiente de pago."
                : $"Hay {cantidadVencidas} cuentas por pagar vencidas pendientes de pago.";

        return View(modelo);
    }
}
