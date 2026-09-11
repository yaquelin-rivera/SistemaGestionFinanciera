using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using SistemaGestionFinanciera.Data;
using SistemaGestionFinanciera.Models.ViewModels;

namespace SistemaGestionFinanciera.ViewComponents;

public class AlertaIngresosViewComponent : ViewComponent
{
    private readonly SistemaFinancieroContext _context;

    public AlertaIngresosViewComponent(
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

        bool puedeGestionar =
            rol.Equals(
                "Administrador",
                StringComparison.OrdinalIgnoreCase) ||
            rol.Equals(
                "Contador",
                StringComparison.OrdinalIgnoreCase) ||
            rol.Equals(
                "Auxiliar contable",
                StringComparison.OrdinalIgnoreCase);

        // Solo quienes pueden revisar o corregir
        // ingresos reciben esta alerta.
        if (!puedeGestionar)
        {
            return View(modelo);
        }

        int cantidadPendientes =
            await _context.Ingresos
                .AsNoTracking()
                .CountAsync(i =>
                    i.Activo &&
                    i.Estado != "Anulado" &&
                    i.PresupuestoMensualId == null &&
                    !string.IsNullOrWhiteSpace(
                        i.MotivoPendientePresupuestario));

        if (cantidadPendientes <= 0)
        {
            return View(modelo);
        }

        modelo.Mostrar = true;
        modelo.Tipo = "amarilla";
        modelo.Cantidad = cantidadPendientes;

        modelo.Mensaje =
            cantidadPendientes == 1
                ? "Hay 1 ingreso pendiente de asignación presupuestaria."
                : $"Hay {cantidadPendientes} ingresos pendientes de asignación presupuestaria.";

        return View(modelo);
    }
}