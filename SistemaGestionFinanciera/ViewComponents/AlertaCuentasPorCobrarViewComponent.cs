using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using SistemaGestionFinanciera.Data;
using SistemaGestionFinanciera.Models.ViewModels;

namespace SistemaGestionFinanciera.ViewComponents;

public class AlertaCuentasPorCobrarViewComponent : ViewComponent
{
    private readonly SistemaFinancieroContext _context;

    public AlertaCuentasPorCobrarViewComponent(
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
            await _context.CuentasPorCobrars
                .AsNoTracking()
                .CountAsync(c =>
                    c.Activo == true &&
                    c.Estado == "Vencida" &&
                    c.SaldoPendiente > 0);

        if (cantidadVencidas <= 0)
        {
            return View(modelo);
        }

        modelo.Mostrar = true;
        modelo.Tipo = "roja";
        modelo.Cantidad = cantidadVencidas;

        modelo.Mensaje =
            cantidadVencidas == 1
                ? "Hay 1 cuenta por cobrar vencida pendiente de cobro."
                : $"Hay {cantidadVencidas} cuentas por cobrar vencidas pendientes de cobro.";

        return View(modelo);
    }
}
