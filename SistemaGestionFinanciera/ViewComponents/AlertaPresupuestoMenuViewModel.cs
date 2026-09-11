using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using SistemaGestionFinanciera.Data;
using SistemaGestionFinanciera.Models.ViewModels;

namespace SistemaGestionFinanciera.ViewComponents;

public class AlertaProyectosViewComponent
    : ViewComponent
{
    private readonly SistemaFinancieroContext
        _context;

    public AlertaProyectosViewComponent(
        SistemaFinancieroContext context)
    {
        _context = context;
    }

    public async Task<IViewComponentResult> InvokeAsync()
    {
        var modelo =
            new AlertaPresupuestoMenuViewModel();

        string rol =
            HttpContext.Session.GetString("Rol") ?? "";

        bool puedePreparar =
            rol.Equals(
                "Administrador",
                StringComparison.OrdinalIgnoreCase) ||
            rol.Equals(
                "Contador",
                StringComparison.OrdinalIgnoreCase) ||
            rol.Equals(
                "Auxiliar Contable",
                StringComparison.OrdinalIgnoreCase);

        if (!puedePreparar)
        {
            return View(modelo);
        }

        DateOnly hoy =
            DateOnly.FromDateTime(
                DateTime.Today);

        DateOnly inicioMesActual =
            new DateOnly(
                hoy.Year,
                hoy.Month,
                1);

        DateOnly finMesActual =
            inicioMesActual
                .AddMonths(1)
                .AddDays(-1);

        // =================================================
        // ROJA: MES ACTUAL SIN PRESUPUESTO
        // =================================================

        int pendientesActuales =
            await _context.Proyectos
                .Where(p =>
                    p.Activo &&
                    p.FechaInicio.HasValue &&
                    p.FechaInicio.Value <=
                        finMesActual &&
                    (
                        !p.FechaFin.HasValue ||
                        p.FechaFin.Value >=
                            inicioMesActual
                    ))
                .Where(p =>
                    !_context.PresupuestosMensuales
                        .Any(pm =>
                            pm.ProyectoId ==
                                p.IdProyecto &&
                            pm.Anio ==
                                inicioMesActual.Year &&
                            pm.Mes ==
                                inicioMesActual.Month))
                .CountAsync();

        if (pendientesActuales > 0)
        {
            string mesActual =
                ObtenerNombreMes(
                    inicioMesActual.Month);

            modelo.Mostrar = true;
            modelo.Tipo = "roja";
            modelo.Cantidad =
                pendientesActuales;

            modelo.Mensaje =
                pendientesActuales == 1
                    ? $"Hay 1 proyecto vigente sin presupuesto para {mesActual} de {inicioMesActual.Year}."
                    : $"Hay {pendientesActuales} proyectos vigentes sin presupuesto para {mesActual} de {inicioMesActual.Year}.";

            return View(modelo);
        }

        // =================================================
        // AMARILLA: PRÓXIMO MES
        // =================================================

        int diasRestantes =
            DateTime.DaysInMonth(
                hoy.Year,
                hoy.Month) -
            hoy.Day;

        if (diasRestantes > 7)
        {
            return View(modelo);
        }

        DateOnly inicioProximoMes =
            inicioMesActual.AddMonths(1);

        DateOnly finProximoMes =
            inicioProximoMes
                .AddMonths(1)
                .AddDays(-1);

        int pendientesProximos =
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
                                inicioProximoMes.Year &&
                            pm.Mes ==
                                inicioProximoMes.Month))
                .CountAsync();

        if (pendientesProximos <= 0)
        {
            return View(modelo);
        }

        string mesProximo =
            ObtenerNombreMes(
                inicioProximoMes.Month);

        modelo.Mostrar = true;
        modelo.Tipo = "amarilla";
        modelo.Cantidad =
            pendientesProximos;

        modelo.Mensaje =
            pendientesProximos == 1
                ? $"Hay 1 proyecto que necesita presupuesto para {mesProximo} de {inicioProximoMes.Year}."
                : $"Hay {pendientesProximos} proyectos que necesitan presupuesto para {mesProximo} de {inicioProximoMes.Year}.";

        return View(modelo);
    }

    private static string ObtenerNombreMes(
        int mes)
    {
        string nombre =
            new DateTime(
                2000,
                mes,
                1)
            .ToString(
                "MMMM",
                new System.Globalization
                    .CultureInfo("es-CR"));

        return char.ToUpper(nombre[0]) +
               nombre[1..];
    }
}