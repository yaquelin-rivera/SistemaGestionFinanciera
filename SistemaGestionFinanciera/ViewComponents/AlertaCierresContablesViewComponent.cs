using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using SistemaGestionFinanciera.Data;
using SistemaGestionFinanciera.Models.ViewModels;

namespace SistemaGestionFinanciera.ViewComponents
{
    public class AlertaCierresContablesViewComponent : ViewComponent
    {
        private readonly SistemaFinancieroContext _context;

        public AlertaCierresContablesViewComponent(
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

            // La alerta tiene sentido para quienes pueden
            // actuar sobre los cierres.
            bool puedeGestionar =
                rol.Equals(
                    "Administrador",
                    StringComparison.OrdinalIgnoreCase)
                ||
                rol.Equals(
                    "Contador",
                    StringComparison.OrdinalIgnoreCase);

            if (!puedeGestionar)
            {
                return View(modelo);
            }

            DateTime hoy =
                DateTime.Today;

            DateOnly inicioMesActual =
                new DateOnly(
                    hoy.Year,
                    hoy.Month,
                    1);

            // ============================================================
            // MESES ANTERIORES QUE REALMENTE TUVIERON MOVIMIENTOS
            // ============================================================
            var periodosConMovimientos =
                await _context.MovimientosContables
                    .AsNoTracking()
                    .Where(m =>
                        m.Fecha < inicioMesActual)
                    .Select(m =>
                        new
                        {
                            Anio = m.Fecha.Year,
                            Mes = m.Fecha.Month
                        })
                    .Distinct()
                    .ToListAsync();

            // ============================================================
            // CIERRES EXISTENTES
            // ============================================================
            var cierres =
                await _context.CierresContables
                    .AsNoTracking()
                    .Select(c =>
                        new
                        {
                            c.Anio,
                            c.Mes,
                            c.Estado
                        })
                    .ToListAsync();

            int cantidadPendientes =
                periodosConMovimientos.Count(p =>
                {
                    var cierre =
                        cierres.FirstOrDefault(c =>
                            c.Anio == p.Anio &&
                            c.Mes == p.Mes);

                    // Nunca cerrado.
                    if (cierre == null)
                    {
                        return true;
                    }

                    // Si fue reabierto, vuelve a estar pendiente.
                    return cierre.Estado != "Cerrado";
                });

            if (cantidadPendientes <= 0)
            {
                return View(modelo);
            }

            modelo.Mostrar = true;
            modelo.Tipo = "amarilla";
            modelo.Cantidad = cantidadPendientes;

            modelo.Mensaje =
                cantidadPendientes == 1
                    ? "Hay 1 período contable pendiente de cierre."
                    : $"Hay {cantidadPendientes} períodos contables pendientes de cierre.";

            return View(modelo);
        }
    }
}