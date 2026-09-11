using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Http;
using Microsoft.EntityFrameworkCore;
using SistemaGestionFinanciera.Data;
using SistemaGestionFinanciera.Models.ViewModels;

namespace SistemaGestionFinanciera.ViewComponents;

public class AlertaPresupuestosViewComponent
    : ViewComponent
{
    private readonly SistemaFinancieroContext
        _context;

    public AlertaPresupuestosViewComponent(
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

        bool esAdministrador =
            rol.Equals(
                "Administrador",
                StringComparison.OrdinalIgnoreCase);

        bool esContador =
            rol.Equals(
                "Contador",
                StringComparison.OrdinalIgnoreCase);

        bool esAuxiliar =
            rol.Equals(
                "Auxiliar Contable",
                StringComparison.OrdinalIgnoreCase);

        bool esTesorero =
            rol.Equals(
                "Tesorero",
                StringComparison.OrdinalIgnoreCase);

        /*
         * El Tesorero solamente consulta.
         * No recibe avisos de tareas pendientes.
         */
        if (esTesorero)
        {
            return View(modelo);
        }

        // ====================================================
        // ALERTA ROJA SEGÚN EL TRABAJO DEL ROL
        // ====================================================

        if (esAdministrador)
        {
            int borradores =
                await _context.PresupuestosMensuales
                    .CountAsync(p =>
                        p.EstadoAprobacion ==
                        "Borrador");

            if (borradores > 0)
            {
                modelo.Mostrar = true;
                modelo.Tipo = "roja";
                modelo.Cantidad = borradores;

                modelo.Mensaje =
                    borradores == 1
                        ? "Hay 1 presupuesto pendiente de aprobación."
                        : $"Hay {borradores} presupuestos pendientes de aprobación.";

                return View(modelo);
            }
        }

        if (esContador || esAuxiliar)
        {
            int rechazados =
                await _context.PresupuestosMensuales
                    .CountAsync(p =>
                        p.EstadoAprobacion ==
                        "Rechazado");

            if (rechazados > 0)
            {
                modelo.Mostrar = true;
                modelo.Tipo = "roja";
                modelo.Cantidad = rechazados;

                modelo.Mensaje =
                    rechazados == 1
                        ? "Hay 1 presupuesto rechazado pendiente de corrección."
                        : $"Hay {rechazados} presupuestos rechazados pendientes de corrección.";

                return View(modelo);
            }
        }

        // ====================================================
        // ALERTA AMARILLA DEL PRÓXIMO MES
        // Solo aparece cuando no hay alerta roja.
        // ====================================================

        if (!(esAdministrador ||
              esContador ||
              esAuxiliar))
        {
            return View(modelo);
        }

        DateOnly hoy =
            DateOnly.FromDateTime(DateTime.Today);

        int ultimoDiaMes =
            DateTime.DaysInMonth(
                hoy.Year,
                hoy.Month);

        int diasRestantes =
            ultimoDiaMes - hoy.Day;

        /*
         * Se muestra durante los últimos siete días
         * del mes, incluyendo el último día.
         */
        if (diasRestantes > 7)
        {
            return View(modelo);
        }

        DateOnly inicioProximoMes =
            new DateOnly(
                hoy.Year,
                hoy.Month,
                1)
            .AddMonths(1);

        DateOnly finProximoMes =
            inicioProximoMes
                .AddMonths(1)
                .AddDays(-1);

        int anioSiguiente =
            inicioProximoMes.Year;

        int mesSiguiente =
            inicioProximoMes.Month;

        /*
         * Proyectos activos que estarán vigentes
         * al menos un día del próximo mes.
         */
        int proyectosPendientes =
            await _context.Proyectos
                .Where(p =>
                    p.Activo &&
                    p.FechaInicio.HasValue &&
                    p.FechaInicio.Value <=
                        finProximoMes &&
                    (
                        !p.FechaFin.HasValue ||
                        p.FechaFin.Value >=
                            inicioProximoMes
                    ))
                .Where(p =>
                    !_context.PresupuestosMensuales
                        .Any(pm =>
                            pm.ProyectoId ==
                                p.IdProyecto &&
                            pm.Anio ==
                                anioSiguiente &&
                            pm.Mes ==
                                mesSiguiente))
                .CountAsync();

        if (proyectosPendientes <= 0)
        {
            return View(modelo);
        }

        string nombreMes =
            inicioProximoMes
                .ToDateTime(TimeOnly.MinValue)
                .ToString(
                    "MMMM",
                    new System.Globalization
                        .CultureInfo("es-CR"));

        nombreMes =
            char.ToUpper(nombreMes[0]) +
            nombreMes[1..];

        modelo.Mostrar = true;
        modelo.Tipo = "amarilla";
        modelo.Cantidad = proyectosPendientes;

        modelo.Mensaje =
            proyectosPendientes == 1
                ? $"Hay 1 proyecto que necesita presupuesto para {nombreMes} de {anioSiguiente}."
                : $"Hay {proyectosPendientes} proyectos que necesitan presupuesto para {nombreMes} de {anioSiguiente}.";

        return View(modelo);
    }
}