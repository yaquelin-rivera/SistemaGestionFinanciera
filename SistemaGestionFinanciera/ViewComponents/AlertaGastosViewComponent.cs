using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using SistemaGestionFinanciera.Data;
using SistemaGestionFinanciera.Models.ViewModels;

namespace SistemaGestionFinanciera.ViewComponents;

public class AlertaGastosViewComponent : ViewComponent
{
    private readonly SistemaFinancieroContext _context;

    public AlertaGastosViewComponent(
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

        bool puedeAutorizar =
            rol.Equals(
                "Administrador",
                StringComparison.OrdinalIgnoreCase) ||
            rol.Equals(
                "Contador",
                StringComparison.OrdinalIgnoreCase);

        // Solo quienes pueden aprobar o rechazar
        // reciben esta alerta.
        if (!puedeAutorizar)
        {
            return View(modelo);
        }

        int cantidadPendientes =
            await _context.Gastos
                .CountAsync(g =>
                    g.Activo &&
                    g.Estado ==
                        "Pendiente de autorización");

        if (cantidadPendientes <= 0)
        {
            return View(modelo);
        }

        modelo.Mostrar = true;
        modelo.Tipo = "roja";
        modelo.Cantidad = cantidadPendientes;

        modelo.Mensaje =
            cantidadPendientes == 1
                ? "Hay 1 gasto pendiente de autorización."
                : $"Hay {cantidadPendientes} gastos pendientes de autorización.";

        return View(modelo);
    }
}