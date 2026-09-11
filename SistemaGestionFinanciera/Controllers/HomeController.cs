using Microsoft.AspNetCore.Mvc;
using SistemaGestionFinanciera.Models;
using System.Diagnostics;
using Microsoft.AspNetCore.Http;
using SistemaGestionFinanciera.Data;
using SistemaGestionFinanciera.ViewModels;
using Microsoft.EntityFrameworkCore;
using SistemaGestionFinanciera.Models.ViewModels;
using System.Globalization;



namespace SistemaGestionFinanciera.Controllers
{
    public class HomeController : BaseController
    {
        private readonly ILogger<HomeController> _logger;
        private readonly SistemaFinancieroContext _context;

        public HomeController(ILogger<HomeController> logger, SistemaFinancieroContext context)
        {
            _logger = logger;
            _context = context;
        }
        public async Task<IActionResult> Index(int? anio,int?mes)
        {


            if (string.IsNullOrEmpty(HttpContext.Session.GetString("Usuario")))
            {
                return RedirectToAction("Login", "Acceso");
            }

            var hoy = DateOnly.FromDateTime(DateTime.Today);
            var mesActual = hoy.Month;
            var anioActual = hoy.Year;

            //filtros de año y mes 
            int anioConsulta = anio ?? anioActual;
            int? mesConsulta;

            if (!anio.HasValue && !mes.HasValue)
            {
                mesConsulta = mesActual;
            }
            else
            {
                mesConsulta = mes;
            }

            bool filtroAplicado =
                anio.HasValue || mes.HasValue;

            ViewBag.AnioSeleccionado = anioConsulta;
            ViewBag.MesSeleccionado = mesConsulta;
            ViewBag.FiltroAplicado = filtroAplicado;

            var aniosDisponibles = await _context.Ingresos
          .Select(i => i.Fecha.Year)
          .Union(_context.Gastos.Select(g => g.Fecha.Year))
          .Union(_context.PresupuestosMensuales.Select(p => p.Anio))
          .Distinct()
          .ToListAsync();

            if (!aniosDisponibles.Contains(anioActual))
            {
                aniosDisponibles.Add(anioActual);
            }

            ViewBag.Anios = aniosDisponibles
                .OrderByDescending(a => a)
                .ToList();

            var consultaIngresos = _context.Ingresos
    .Where(i =>
        i.Activo &&
        i.Fecha.Year == anioConsulta);

            var consultaGastos = _context.Gastos
    .Where(g =>
        g.Activo &&
        g.Estado == "Aprobado" &&
        g.Fecha.Year == anioConsulta);

            var consultaFacturas = _context.Facturas
                .Where(f =>
                    f.FechaEmision.Year == anioConsulta);

            var consultaMovimientos = _context.MovimientosContables
                .Where(m =>
                    m.Fecha.Year == anioConsulta &&
                    m.Anulado != true);

            if (mesConsulta.HasValue)
            {
                consultaIngresos = consultaIngresos
                    .Where(i => i.Fecha.Month == mesConsulta.Value);

                consultaGastos = consultaGastos
                    .Where(g => g.Fecha.Month == mesConsulta.Value);

                consultaFacturas = consultaFacturas
                    .Where(f => f.FechaEmision.Month == mesConsulta.Value);

                consultaMovimientos = consultaMovimientos
                    .Where(m => m.Fecha.Month == mesConsulta.Value);
            }


            var dashboard = new DashboardViewModel
            {
                Usuario = HttpContext.Session.GetString("Usuario"),
                Rol = HttpContext.Session.GetString("Rol"),

                IngresosMes = await consultaIngresos
    .SumAsync(i => i.MontoReal ?? 0),

                GastosMes = await consultaGastos
    .SumAsync(g => g.MontoTotal ?? 0),

                CuentasPorCobrarPendientes = await _context.CuentasPorCobrars
                    .Where(c => c.Activo && c.SaldoPendiente > 0 && c.Estado != "Anulada")
                    .SumAsync(c => c.SaldoPendiente),

                CuentasPorPagarPendientes = await _context.CuentasPorPagars
                    .Where(c => c.Activo && c.SaldoPendiente > 0 && c.Estado != "Anulada")
                    .SumAsync(c => c.SaldoPendiente),

                FacturasMes = await consultaFacturas
    .CountAsync(),

                MovimientosContablesMes = await consultaMovimientos
    .CountAsync(),
                CuentasPorCobrarVencidas = await _context.CuentasPorCobrars
                    .Where(c => c.Activo && c.SaldoPendiente > 0 && c.FechaVencimiento < hoy && c.Estado != "Anulada")
                    .CountAsync(),

                CuentasPorPagarVencidas = await _context.CuentasPorPagars
                    .Where(c => c.Activo && c.SaldoPendiente > 0 && c.FechaVencimiento < hoy && c.Estado != "Anulada")
                    .CountAsync(),

                AuditoriasHoy = await _context.Auditoria
                    .Where(a => a.Fecha.Date == DateTime.Today)
                    .CountAsync()
            };

            // todos los proyctos aprobados 
            var presupuestosPeriodo = _context.PresupuestosMensuales
        .Where(p =>
            p.Anio == anioConsulta &&
            p.EstadoAprobacion == "Aprobado");

            if (mesConsulta.HasValue)
            {
                presupuestosPeriodo = presupuestosPeriodo
                    .Where(p => p.Mes == mesConsulta.Value);
            }

            dashboard.PresupuestoIngresosPlanificado =
    await presupuestosPeriodo.SumAsync(
        p => p.MontoIngresadoPlanificado ?? 0);

            dashboard.PresupuestoGastosPlanificado =
                await presupuestosPeriodo.SumAsync(
                    p => p.MontoGastoPlanificado ?? 0);

            dashboard.PresupuestoIngresosReal =
                await presupuestosPeriodo.SumAsync(
                    p => p.TotalIngresoReal ?? 0);

            dashboard.PresupuestoGastosReal =
                await presupuestosPeriodo.SumAsync(
                    p => p.TotalGastoReal ?? 0);
            // ===============================
            // DATOS PARA LOS GRÁFICOS
            // Últimos 6 meses
            // ===============================

            dashboard.Meses = new List<string>();
            dashboard.IngresosMensuales = new List<decimal>();
            dashboard.GastosMensuales = new List<decimal>();

            if (mesConsulta.HasValue)
            {
                // Cuando se consulta un mes específico,
                // la gráfica muestra únicamente ese mes.
                var fechaConsulta = new DateTime(
                    anioConsulta,
                    mesConsulta.Value,
                    1);

                var nombreMes = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(
                    fechaConsulta.ToString("MMMM", new CultureInfo("es-CR"))
                );

                dashboard.Meses.Add(nombreMes);

                dashboard.IngresosMensuales.Add(
                    await consultaIngresos.SumAsync(
                        ingreso => ingreso.MontoReal ?? 0));

                dashboard.GastosMensuales.Add(
                    await consultaGastos.SumAsync(
                        gasto => gasto.MontoTotal ?? 0));
            }
            else
            {
                // Cuando se selecciona solamente el año,
                // la gráfica muestra los doce meses.
                for (int numeroMes = 1; numeroMes <= 12; numeroMes++)
                {
                    var fechaConsulta = new DateTime(
                        anioConsulta,
                        numeroMes,
                        1);

                    var nombreMes = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(
                        fechaConsulta.ToString("MMMM", new CultureInfo("es-CR"))
                    );

                    dashboard.Meses.Add(nombreMes);

                    var ingresosDelMes = await _context.Ingresos
                        .Where(ingreso =>
                            ingreso.Activo &&
                            ingreso.Fecha.Year == anioConsulta &&
                            ingreso.Fecha.Month == numeroMes)
                        .SumAsync(ingreso => ingreso.MontoReal ?? 0);


                    var gastosDelMes = await _context.Gastos
    .Where(gasto =>
        gasto.Activo &&
        gasto.Estado == "Aprobado" &&
        gasto.Fecha.Year == anioConsulta &&
        gasto.Fecha.Month == numeroMes)
    .SumAsync(gasto => gasto.MontoTotal ?? 0);

                    dashboard.IngresosMensuales.Add(ingresosDelMes);
                    dashboard.GastosMensuales.Add(gastosDelMes);
                }
            }
            //GRAFICA CIRCULAR en base a datos de cxc 
            // Pendientes no vencidas.
            dashboard.CuentasPorCobrarPendientesCantidad =
                await _context.CuentasPorCobrars.CountAsync(c =>
                    c.Activo &&
                    c.SaldoPendiente > 0 &&
                    c.FechaVencimiento >= hoy &&
                    c.Estado != "Anulada");

            // Pagadas o canceladas.
            dashboard.CuentasPorCobrarPagadasCantidad =
                await _context.CuentasPorCobrars.CountAsync(c =>
                    c.SaldoPendiente == 0 &&
                    (c.Estado == "Cancelada" || c.Estado == "Pagada"));

            // Vencidas con saldo pendiente.
            dashboard.CuentasPorCobrarVencidasCantidad =
                await _context.CuentasPorCobrars.CountAsync(c =>
                    c.Activo &&
                    c.SaldoPendiente > 0 &&
                    c.FechaVencimiento < hoy &&
                    c.Estado != "Anulada");

            // ÚLTIMOS MOVIMIENTOS

            dashboard.UltimosMovimientos = await _context.MovimientosContables
                .Where(m => m.Anulado != true)
                .OrderByDescending(m => m.Fecha)
                .ThenByDescending(m => m.IdMovimientoContable)
                .Take(5)
                .Select(m => new MovimientoDashboardItem
                {
                    Fecha = m.Fecha,
                    Modulo = m.OrigenModulo,
                    Descripcion = m.Descripcion,
                    Debe = m.Debe,
                    Haber = m.Haber
                })
                .ToListAsync();

            dashboard.UltimasAuditorias = await _context.Auditoria
                .Include(a => a.Usuario)
                .OrderByDescending(a => a.Fecha)
                .Take(5)
                .Select(a => new AuditoriaDashboardItem
                {
                    Fecha = a.Fecha,
                    Modulo = a.Tabla,
                    Accion = a.Accion,
                    Usuario = a.Usuario.Nombre
                })
                .ToListAsync();

            return View(dashboard);
        }
        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
