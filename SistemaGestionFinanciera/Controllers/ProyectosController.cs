using System;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using SistemaGestionFinanciera.Data;
using SistemaGestionFinanciera.Models;
using ClosedXML.Excel;
using System.IO;

namespace SistemaGestionFinanciera.Controllers
{
    public class ProyectosController : BaseController
    {
        private readonly SistemaFinancieroContext _context;

        public ProyectosController(SistemaFinancieroContext context)
        {
            _context = context;
        }

        public async Task<IActionResult> Index(
     string nombre,
     string estado,
     DateOnly? fechaInicio,
     DateOnly? fechaFin)
        {
            if (string.IsNullOrEmpty(
                HttpContext.Session.GetString("Usuario")))
            {
                return RedirectToAction(
                    "Login",
                    "Acceso");
            }

            var proyectos =
                _context.Proyectos
                    .AsQueryable();

            if (!string.IsNullOrWhiteSpace(nombre))
            {
                proyectos = proyectos.Where(p =>
                    p.Nombre.Contains(nombre));
            }

            if (!string.IsNullOrWhiteSpace(estado))
            {
                proyectos = proyectos.Where(p =>
                    p.Estado == estado);
            }

            if (fechaInicio.HasValue)
            {
                proyectos = proyectos.Where(p =>
                    p.FechaInicio >= fechaInicio.Value);
            }

            if (fechaFin.HasValue)
            {
                proyectos = proyectos.Where(p =>
                    p.FechaFin <= fechaFin.Value);
            }

            ViewBag.Nombre = nombre;
            ViewBag.Estado = estado;
            ViewBag.FechaInicio = fechaInicio;
            ViewBag.FechaFin = fechaFin;

            var lista =
                await proyectos
                    .Include(p =>
                        p.PresupuestosMensuales)
                    .OrderByDescending(p =>
                        p.IdProyecto)
                    .ToListAsync();

            // =====================================================
            // PERÍODO ACTUAL
            // =====================================================

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

            // =====================================================
            // PRÓXIMO PERÍODO
            // =====================================================

            DateOnly inicioProximoMes =
                inicioMesActual
                    .AddMonths(1);

            DateOnly finProximoMes =
                inicioProximoMes
                    .AddMonths(1)
                    .AddDays(-1);

            int diasRestantesMes =
                DateTime.DaysInMonth(
                    hoy.Year,
                    hoy.Month) -
                hoy.Day;

            bool mostrarAvisoProximoMes =
                diasRestantesMes <= 7;

            // =====================================================
            // PROYECTOS SIN PRESUPUESTO DEL MES ACTUAL
            // =====================================================

            var proyectosSinPresupuestoActual =
                lista
                    .Where(p =>
                        p.Activo &&
                        p.FechaInicio.HasValue &&
                        p.FechaInicio.Value <=
                            finMesActual &&
                        (
                            !p.FechaFin.HasValue ||
                            p.FechaFin.Value >=
                                inicioMesActual
                        ) &&
                        !p.PresupuestosMensuales.Any(pm =>
                            pm.Anio ==
                                inicioMesActual.Year &&
                            pm.Mes ==
                                inicioMesActual.Month))
                    .Select(p =>
                        p.IdProyecto)
                    .ToHashSet();

            // =====================================================
            // PROYECTOS SIN PRESUPUESTO DEL PRÓXIMO MES
            // =====================================================

            var proyectosSinPresupuestoProximo =
                mostrarAvisoProximoMes
                    ? lista
                        .Where(p =>
                            p.Activo &&
                            p.FechaInicio.HasValue &&
                            p.FechaInicio.Value <=
                                finProximoMes &&
                            (
                                !p.FechaFin.HasValue ||
                                p.FechaFin.Value >=
                                    inicioProximoMes
                            ) &&
                            !p.PresupuestosMensuales.Any(pm =>
                                pm.Anio ==
                                    inicioProximoMes.Year &&
                                pm.Mes ==
                                    inicioProximoMes.Month))
                        .Select(p =>
                            p.IdProyecto)
                        .ToHashSet()
                    : new HashSet<int>();

            // =====================================================
            // PROYECTOS VENCIDOS O PRÓXIMOS A VENCER
            // =====================================================

            DateOnly limiteVencimiento =
                hoy.AddDays(7);

            var proyectosVencidos =
                lista
                    .Where(p =>
                        p.Activo &&
                        p.FechaFin.HasValue &&
                        p.FechaFin.Value < hoy)
                    .Select(p =>
                        p.IdProyecto)
                    .ToHashSet();

            var proyectosPorVencer =
                lista
                    .Where(p =>
                        p.Activo &&
                        p.FechaFin.HasValue &&
                        p.FechaFin.Value >= hoy &&
                        p.FechaFin.Value <=
                            limiteVencimiento)
                    .Select(p =>
                        p.IdProyecto)
                    .ToHashSet();

            // =====================================================
            // INFORMACIÓN PARA LA VISTA
            // =====================================================

            ViewBag.ProyectosSinPresupuestoActual =
                proyectosSinPresupuestoActual;

            ViewBag.ProyectosSinPresupuestoProximo =
                proyectosSinPresupuestoProximo;

            ViewBag.ProyectosVencidos =
                proyectosVencidos;

            ViewBag.ProyectosPorVencer =
                proyectosPorVencer;

            ViewBag.MostrarAvisoProximoMes =
                mostrarAvisoProximoMes;

            ViewBag.AnioActual =
                inicioMesActual.Year;

            ViewBag.MesActual =
                inicioMesActual.Month;

            ViewBag.AnioProximo =
                inicioProximoMes.Year;

            ViewBag.MesProximo =
                inicioProximoMes.Month;

            ViewBag.NombreMesActual =
                ObtenerNombreMes(
                    inicioMesActual.Month);

            ViewBag.NombreMesProximo =
                ObtenerNombreMes(
                    inicioProximoMes.Month);

            return View(lista);
        }

        public async Task<IActionResult> Details(int? id)
        {
            if (id == null)
                return NotFound();

            var proyecto = await _context.Proyectos
                .FirstOrDefaultAsync(p => p.IdProyecto == id);

            if (proyecto == null)
                return NotFound();

            ViewBag.TotalIngresos = await _context.Ingresos
                .Where(i => i.ProyectoId == id && i.Estado != "Anulado")
                .SumAsync(i => i.MontoReal);

            ViewBag.TotalGastos = await _context.Gastos
                .Where(g => g.ProyectoId == id && g.Estado != "Anulado")
                .SumAsync(g => g.MontoTotal);

            ViewBag.TotalFacturas = await _context.Facturas
                .CountAsync(f => f.ProyectoId == id);

            ViewBag.TotalMovimientos = await _context.MovimientosContables
                .CountAsync(m => m.ProyectoId == id);

            ViewBag.TotalCxC = await _context.CuentasPorCobrars
    .CountAsync(c => c.ProyectoId == id);

            ViewBag.TotalCxP = await _context.CuentasPorPagars
                .CountAsync(c => c.ProyectoId == id);

            ViewBag.TotalPresupuestos = await _context.PresupuestosMensuales
                .CountAsync(p => p.ProyectoId == id);

            return View(proyecto);
        }
        // GET proyectos/ crear 
        public IActionResult Create()
        {
            if (!EsAdministrador())
            {
                TempData["MensajeError"] =
                    "Solo el Administrador puede registrar proyectos.";

                return RedirectToAction(nameof(Index));
            }

            return View();
        }
        // POST proyectos / crear 
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Create([Bind("IdProyecto,Nombre,Descripcion,FechaInicio,FechaFin,PresupuestoAsignado")] Proyecto proyecto)
        {
            if (!EsAdministrador())
            {
                TempData["MensajeError"] =
                    "Solo el Administrador puede registrar proyectos.";

                return RedirectToAction(nameof(Index));
            }

            proyecto.Estado = "Activo";
            proyecto.Activo = true;

            ModelState.Remove("Estado");
            ModelState.Remove("Activo");

            NormalizarProyecto(proyecto);
            await ValidarProyecto(proyecto, esEdicion: false);
            // ============================================================
            // NO PERMITIR CREAR UN PROYECTO EN UN PERÍODO YA CERRADO
            // ============================================================
            if (proyecto.FechaInicio.HasValue)
            {
                bool periodoInicioCerrado =
                    await PeriodoContableCerradoAsync(proyecto.FechaInicio.Value);

                if (periodoInicioCerrado)
                {
                    ModelState.AddModelError(
                        nameof(proyecto.FechaInicio),
                        $"No se puede registrar un proyecto con fecha de inicio " +
                        $"{proyecto.FechaInicio.Value:MM/yyyy} porque ese período contable se encuentra cerrado.");
                }
            }
            if (ModelState.IsValid)
            {
                _context.Proyectos.Add(proyecto);
                await _context.SaveChangesAsync();

                await RegistrarAuditoria(
                    "Proyectos",
                    proyecto.IdProyecto,
                    "Crear",
                    $"Se creó el proyecto {proyecto.Nombre} con presupuesto asignado de ₡{proyecto.PresupuestoAsignado:N2}."
                );

                TempData["MensajeExito"] = "Proyecto registrado correctamente.";
               

                return RedirectToAction(nameof(Index));
            }

            return View(proyecto);
        }
        // GET edit 
        public async Task<IActionResult> Edit(int? id)
        {
            if (!EsAdministrador())
            {
                TempData["MensajeError"] =
                    "Solo el Administrador puede modificar proyectos.";

                return RedirectToAction(nameof(Index));
            }

            if (id == null)
                return NotFound();

            var proyecto = await _context.Proyectos.FindAsync(id);

            if (proyecto == null)
                return NotFound();
            DateOnly hoy =
    DateOnly.FromDateTime(DateTime.Today);

            bool proyectoFinalizado =
                proyecto.FechaFin.HasValue &&
                proyecto.FechaFin.Value < hoy;

            if (!proyecto.Activo)
            {
                TempData["MensajeError"] =
                    "Debe reactivar el proyecto antes de modificarlo.";

                return RedirectToAction(nameof(Index));
            }

            if (proyectoFinalizado)
            {
                TempData["MensajeError"] =
                    "No se puede modificar un proyecto finalizado.";

                return RedirectToAction(nameof(Index));
            }

            bool tieneRegistros =
                await ProyectoTieneRegistros(
                    proyecto.IdProyecto);

            bool proyectoNoIniciado =
                proyecto.FechaInicio.HasValue &&
                proyecto.FechaInicio.Value > hoy;

            ViewBag.TieneRegistros =
                tieneRegistros;

            ViewBag.ProyectoNoIniciado =
                proyectoNoIniciado;

            ViewBag.BloquearEstructura =
                tieneRegistros ||
                proyectoNoIniciado;

            return View(proyecto);
          
        }
        // POST edit 
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Edit(int id, [Bind("IdProyecto,Nombre,Descripcion,FechaInicio,FechaFin,Estado," +"PresupuestoAsignado,MotivoAjustePresupuesto")] Proyecto proyectoFormulario)
        {
            if (!EsAdministrador())
            {
                TempData["MensajeError"] =
                    "Solo el Administrador puede modificar proyectos.";

                return RedirectToAction(nameof(Index));
            }


            if (id != proyectoFormulario.IdProyecto)
                return NotFound();

            var proyectoOriginal = await _context.Proyectos
                .FirstOrDefaultAsync(p => p.IdProyecto == id);

            if (proyectoOriginal == null)
                return NotFound();

            DateOnly hoy =
    DateOnly.FromDateTime(DateTime.Today);

            bool proyectoFinalizado =
                proyectoOriginal.FechaFin.HasValue &&
                proyectoOriginal.FechaFin.Value < hoy;

            if (!proyectoOriginal.Activo)
            {
                TempData["MensajeError"] =
                    "Debe reactivar el proyecto antes de modificarlo.";

                return RedirectToAction(nameof(Index));
            }

            if (proyectoFinalizado)
            {
                TempData["MensajeError"] =
                    "No se puede modificar un proyecto finalizado.";

                return RedirectToAction(nameof(Index));
            }

            bool tieneRegistros = await ProyectoTieneRegistros(id);


            bool proyectoNoIniciado =
    proyectoOriginal.FechaInicio.HasValue &&
    proyectoOriginal.FechaInicio.Value > hoy;

            bool bloquearEstructura =
                tieneRegistros ||
                proyectoNoIniciado;

            ViewBag.TieneRegistros =
                tieneRegistros;

            ViewBag.ProyectoNoIniciado =
                proyectoNoIniciado;

            ViewBag.BloquearEstructura =
                bloquearEstructura;
            NormalizarProyecto(proyectoFormulario);

            decimal presupuestoAnterior =
    proyectoOriginal.PresupuestoAsignado ?? 0m;

            decimal presupuestoNuevo =
                proyectoFormulario.PresupuestoAsignado ?? 0m;

            decimal presupuestoComprometido =
                await CalcularPresupuestoComprometido(id);

            bool cambioPresupuesto =
                presupuestoNuevo != presupuestoAnterior;

            if (bloquearEstructura)
            {
                proyectoFormulario.FechaInicio =
                    proyectoOriginal.FechaInicio;

                proyectoFormulario.FechaFin =
                    proyectoOriginal.FechaFin;

                proyectoFormulario.Estado =
                    proyectoOriginal.Estado;

                proyectoFormulario.Activo =
                    proyectoOriginal.Activo;

                ModelState.Remove(
                    nameof(proyectoFormulario.FechaInicio));

                ModelState.Remove(
                    nameof(proyectoFormulario.FechaFin));

                ModelState.Remove(
                    nameof(proyectoFormulario.Estado));
            }
            // ================================================================
            // BLOQUEO DE CAMBIO DE FECHA DE INICIO HACIA UN PERÍODO CERRADO
            // Solo aplica cuando el proyecto todavía permite modificar estructura.
            // ================================================================
            if (!bloquearEstructura &&
                proyectoFormulario.FechaInicio.HasValue &&
                proyectoFormulario.FechaInicio != proyectoOriginal.FechaInicio)
            {
                bool periodoInicioCerrado =
                    await PeriodoContableCerradoAsync(
                        proyectoFormulario.FechaInicio.Value);

                if (periodoInicioCerrado)
                {
                    ModelState.AddModelError(
                        nameof(proyectoFormulario.FechaInicio),
                        $"No se puede cambiar la fecha de inicio del proyecto a " +
                        $"{proyectoFormulario.FechaInicio.Value:MM/yyyy} porque ese " +
                        $"período contable se encuentra cerrado.");
                }
            }
            if (cambioPresupuesto)
            {
                if (presupuestoNuevo <= 0)
                {
                    ModelState.AddModelError(
                        nameof(proyectoFormulario.PresupuestoAsignado),
                        "El presupuesto asignado debe ser mayor que cero.");
                }

                if (presupuestoNuevo < presupuestoComprometido)
                {
                    ModelState.AddModelError(
                        nameof(proyectoFormulario.PresupuestoAsignado),
                        $"El presupuesto no puede reducirse a " +
                        $"₡{presupuestoNuevo:N2}, porque existen " +
                        $"₡{presupuestoComprometido:N2} comprometidos " +
                        $"en presupuestos mensuales en Borrador o Aprobados.");
                }

                /*
                 * Cuando ya hay relaciones financieras o presupuestarias,
                 * el motivo es obligatorio.
                 */
                if (tieneRegistros &&
                    string.IsNullOrWhiteSpace(
                        proyectoFormulario.MotivoAjustePresupuesto))
                {
                    ModelState.AddModelError(
                        nameof(
                            proyectoFormulario.MotivoAjustePresupuesto),
                        "Debe indicar el motivo del ajuste presupuestario.");
                }
            }
            else
            {
                /*
                 * Si no cambió el presupuesto, no exigimos motivo.
                 */
                ModelState.Remove(
                    nameof(
                        proyectoFormulario.MotivoAjustePresupuesto));
            }
            await ValidarProyecto(proyectoFormulario, esEdicion: true);

            if (ModelState.IsValid)
            {
                // Regla de negocio:
                // Si el proyecto posee registros asociados o todavía no ha iniciado,
                // solo se permite modificar nombre, descripción y presupuesto asignado.
                // Las fechas y el estado permanecen protegidos.
                if (bloquearEstructura)
                {
                    proyectoOriginal.Nombre =
                        proyectoFormulario.Nombre;

                    proyectoOriginal.Descripcion =
                        proyectoFormulario.Descripcion;

                    proyectoOriginal.PresupuestoAsignado =
                        proyectoFormulario.PresupuestoAsignado;
                }
                else
                {
                    proyectoOriginal.Nombre = proyectoFormulario.Nombre;
                    proyectoOriginal.Descripcion = proyectoFormulario.Descripcion;
                    proyectoOriginal.FechaInicio = proyectoFormulario.FechaInicio;
                    proyectoOriginal.FechaFin = proyectoFormulario.FechaFin;
                    proyectoOriginal.PresupuestoAsignado = proyectoFormulario.PresupuestoAsignado;
                    proyectoOriginal.Estado = proyectoFormulario.Estado;
                    proyectoOriginal.Activo = proyectoFormulario.Estado == "Activo";
                }

                _context.Update(proyectoOriginal);
                await _context.SaveChangesAsync();

                string descripcionAuditoria;

                if (cambioPresupuesto)
                {
                    decimal diferencia =
                        presupuestoNuevo - presupuestoAnterior;

                    string signo =
                        diferencia >= 0
                            ? "+"
                            : "-";

                    descripcionAuditoria =
                        $"Se modificó el proyecto " +
                        $"\"{proyectoOriginal.Nombre}\". " +
                        $"El presupuesto asignado cambió de " +
                        $"₡{presupuestoAnterior:N2} a " +
                        $"₡{presupuestoNuevo:N2}. " +
                        $"Diferencia: {signo}₡{Math.Abs(diferencia):N2}. " +
                        $"Presupuesto comprometido: " +
                        $"₡{presupuestoComprometido:N2}. " +
                        $"Motivo: " +
                        $"{proyectoFormulario.MotivoAjustePresupuesto}.";
                }
                else
                {
                    descripcionAuditoria =
                        $"Se modificaron los datos generales del proyecto " +
                        $"\"{proyectoOriginal.Nombre}\".";
                }

                await RegistrarAuditoria(
                    "Proyectos",
                    proyectoOriginal.IdProyecto,
                    "Modificar",
                    descripcionAuditoria);

                TempData["MensajeExito"] = "Proyecto modificado correctamente.";
              
                return RedirectToAction(nameof(Index));
            }

            return View(proyectoFormulario);
        }
        // GET proyectos/ desacvtivar 
        public async Task<IActionResult> Delete(int? id)
        {
            if (!EsAdministrador())
            {
                TempData["MensajeError"] =
                    "Solo el Administrador puede cambiar el estado de los proyectos.";

                return RedirectToAction(nameof(Index));
            }

            if (id == null)
                return NotFound();

            var proyecto = await _context.Proyectos
                .FirstOrDefaultAsync(p => p.IdProyecto == id);

            if (proyecto == null)
                return NotFound();

            ViewBag.TieneRegistros = await ProyectoTieneRegistros(proyecto.IdProyecto);

            return View(proyecto);
        }

        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> DeleteConfirmed(int id)
        {
            if (!EsAdministrador())
            {
                TempData["MensajeError"] =
                    "Solo el Administrador puede cambiar el estado de los proyectos.";

                return RedirectToAction(nameof(Index));
            }

            var proyecto = await _context.Proyectos.FindAsync(id);

            if (proyecto == null)
                return NotFound();

            bool tieneRegistros = await ProyectoTieneRegistros(id);

            if (tieneRegistros)
            {
                TempData["MensajeError"] = "No se puede desactivar este proyecto porque posee registros financieros asociados.";
              
                return RedirectToAction(nameof(Index));
            }

            proyecto.Activo = false;
            proyecto.Estado = "Inactivo";

            _context.Update(proyecto);
            await _context.SaveChangesAsync();

            await RegistrarAuditoria(
                "Proyectos",
                proyecto.IdProyecto,
                "Desactivar",
                $"Se desactivó el proyecto {proyecto.Nombre}."
            );

            TempData["MensajeExito"] = "Proyecto desactivado correctamente.";
           
            return RedirectToAction(nameof(Index));
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Reactivar(int id)
        {
            if (!EsAdministrador())
            {
                TempData["MensajeError"] =
                    "Solo el Administrador puede cambiar el estado de los proyectos.";

                return RedirectToAction(nameof(Index));
            }

            var proyecto = await _context.Proyectos.FindAsync(id);

            if (proyecto == null)
                return NotFound();

            proyecto.Activo = true;
            proyecto.Estado = "Activo";

            _context.Update(proyecto);
            await _context.SaveChangesAsync();

            await RegistrarAuditoria(
                "Proyectos",
                proyecto.IdProyecto,
                "Reactivar",
                $"Se reactivó el proyecto {proyecto.Nombre}."
            );

            TempData["MensajeExito"] = "Proyecto reactivado correctamente.";
         
            return RedirectToAction(nameof(Index));
        }
        // =====================================================
        // CONSULTA REUTILIZABLE PARA REPORTES DE PROYECTOS
        // =====================================================
        private IQueryable<Proyecto> ConstruirConsultaReporte(
            string nombre,
            string estado,
            DateOnly? fechaInicio,
            DateOnly? fechaFin)
        {
            var consulta = _context.Proyectos
                .AsNoTracking()
                .Include(p => p.PresupuestosMensuales)
                .AsQueryable();

            if (!string.IsNullOrWhiteSpace(nombre))
            {
                nombre = nombre.Trim();

                consulta = consulta.Where(p =>
                    p.Nombre.Contains(nombre));
            }

            if (!string.IsNullOrWhiteSpace(estado))
            {
                estado = estado.Trim();

                consulta = consulta.Where(p =>
                    p.Estado == estado);
            }

            if (fechaInicio.HasValue)
            {
                consulta = consulta.Where(p =>
                    p.FechaInicio.HasValue &&
                    p.FechaInicio.Value >= fechaInicio.Value);
            }

            if (fechaFin.HasValue)
            {
                consulta = consulta.Where(p =>
                    p.FechaFin.HasValue &&
                    p.FechaFin.Value <= fechaFin.Value);
            }

            return consulta;
        }


        // =====================================================
        // VISTA PREVIA DEL REPORTE
        // =====================================================
        [HttpGet]
        public async Task<IActionResult> Reporte(
            string nombre,
            string estado,
            DateOnly? fechaInicio,
            DateOnly? fechaFin)
        {
            if (string.IsNullOrEmpty(
                HttpContext.Session.GetString("Usuario")))
            {
                return RedirectToAction("Login", "Acceso");
            }

            if (fechaInicio.HasValue &&
                fechaFin.HasValue &&
                fechaInicio.Value > fechaFin.Value)
            {
                TempData["MensajeError"] =
                    "La fecha inicial no puede ser mayor que la fecha final.";

                return RedirectToAction(nameof(Index));
            }

            var proyectos = await ConstruirConsultaReporte(
                    nombre,
                    estado,
                    fechaInicio,
                    fechaFin)
                .OrderByDescending(p => p.FechaInicio)
                .ThenBy(p => p.Nombre)
                .ToListAsync();

            // Filtros utilizados
            ViewBag.Nombre = nombre;
            ViewBag.Estado = estado;
            ViewBag.FechaInicio = fechaInicio;
            ViewBag.FechaFin = fechaFin;

            // Resumen del reporte
            ViewBag.TotalProyectos = proyectos.Count;

            ViewBag.TotalActivos = proyectos.Count(p =>
                p.Estado == "Activo" || p.Activo);

            ViewBag.TotalInactivos = proyectos.Count(p =>
                p.Estado == "Inactivo" || !p.Activo);

            ViewBag.PresupuestoTotal = proyectos.Sum(p =>
                p.PresupuestoAsignado ?? 0);

            ViewBag.ProyectosConPresupuestoAprobado =
                proyectos.Count(p =>
                    p.PresupuestosMensuales != null &&
                    p.PresupuestosMensuales.Any(pm =>
                        pm.EstadoAprobacion == "Aprobado"));

            ViewBag.FechaGeneracion = DateTime.Now;
            await RegistrarAuditoria(
    "Proyectos",
    0,
    "Vista previa PDF",
    "Se generó la vista previa del reporte de proyectos."
);

            return View("Reporte", proyectos);
        }


        // =====================================================
        // EXPORTAR REPORTE DE PROYECTOS A EXCEL
        // =====================================================
        [HttpGet]
        public async Task<IActionResult> ExportarExcel(
            string nombre,
            string estado,
            DateOnly? fechaInicio,
            DateOnly? fechaFin)
        {
            if (string.IsNullOrEmpty(
                HttpContext.Session.GetString("Usuario")))
            {
                return RedirectToAction("Login", "Acceso");
            }

            if (fechaInicio.HasValue &&
                fechaFin.HasValue &&
                fechaInicio.Value > fechaFin.Value)
            {
                TempData["MensajeError"] =
                    "La fecha inicial no puede ser mayor que la fecha final.";

                return RedirectToAction(nameof(Index));
            }

            var proyectos = await ConstruirConsultaReporte(
                    nombre,
                    estado,
                    fechaInicio,
                    fechaFin)
                .OrderByDescending(p => p.FechaInicio)
                .ThenBy(p => p.Nombre)
                .ToListAsync();

            using var workbook = new XLWorkbook();

            var hoja = workbook.Worksheets.Add(
                "Reporte de proyectos");
           

            // Colores institucionales
            var colorInstitucional = XLColor.FromHtml("#0F6B73");
            var colorInstitucionalOscuro = XLColor.FromHtml("#0B4F55");
            var colorFondoSuave = XLColor.FromHtml("#E7F1F2");
            var colorFondoFiltros = XLColor.FromHtml("#F3F7F8");

            var colorVerde = XLColor.FromHtml("#198754");
            var colorVerdeSuave = XLColor.FromHtml("#D1E7DD");

            var colorGris = XLColor.FromHtml("#6C757D");
            var colorGrisSuave = XLColor.FromHtml("#E2E3E5");

            var colorAmarillo = XLColor.FromHtml("#FFC107");
            var colorAmarilloSuave = XLColor.FromHtml("#FFF3CD");

            var colorDorado = XLColor.FromHtml("#D39E00");
            var colorDoradoSuave = XLColor.FromHtml("#FFF3CD");
            // =================================================
            // TÍTULO
            // =================================================
            hoja.Cell("A1").Value =
                "UNIÓN CANTONAL DE ASOCIACIONES DE DESARROLLO DE SANTA ANA";

            hoja.Range("A1:H1").Merge();

            hoja.Cell("A2").Value =
                "REPORTE DE PROYECTOS";

            hoja.Range("A2:H2").Merge();

            hoja.Cell("A3").Value =
                $"Fecha de generación: {DateTime.Now:dd/MM/yyyy hh:mm tt}";

            hoja.Range("A3:H3").Merge();

            hoja.Range("A1:H3").Style.Alignment.Horizontal =
                XLAlignmentHorizontalValues.Center;

            hoja.Range("A1:H3").Style.Font.Bold = true;
            hoja.Range("A1:H1").Style.Fill.BackgroundColor =
    colorInstitucional;

            hoja.Range("A1:H2").Style.Font.FontColor =
                XLColor.White;

            hoja.Range("A1:H2").Style.Alignment.Horizontal =
                XLAlignmentHorizontalValues.Center;

            hoja.Range("A1:H2").Style.Alignment.Vertical =
                XLAlignmentVerticalValues.Center;

            hoja.Row(1).Height = 24;
            hoja.Row(2).Height = 22;

            hoja.Range("A3:H3").Style.Fill.BackgroundColor =
                colorFondoSuave;

            hoja.Range("A3:H3").Style.Font.FontColor =
                colorInstitucionalOscuro;
            hoja.Cell("A1").Style.Font.FontSize = 14;
            hoja.Cell("A2").Style.Font.FontSize = 13;

            // =================================================
            // FILTROS APLICADOS
            // =================================================
            hoja.Cell("A5").Value = "Filtros aplicados";
            hoja.Range("A5:H5").Merge();
            hoja.Range("A5:H5").Style.Font.Bold = true;
            hoja.Range("A5:H5").Style.Fill.BackgroundColor =
    colorInstitucional;

            hoja.Range("A5:H5").Style.Font.FontColor =
                XLColor.White;

            hoja.Range("A5:H5").Style.Font.Bold = true;

            hoja.Range("A5:H5").Style.Alignment.Horizontal =
                XLAlignmentHorizontalValues.Left;
            hoja.Cell("A6").Value = "Nombre:";
            hoja.Cell("B6").Value =
                string.IsNullOrWhiteSpace(nombre)
                    ? "Todos"
                    : nombre;

            hoja.Cell("C6").Value = "Estado:";
            hoja.Cell("D6").Value =
                string.IsNullOrWhiteSpace(estado)
                    ? "Todos"
                    : estado;

            hoja.Cell("E6").Value = "Fecha inicio desde:";
            hoja.Cell("F6").Value =
                fechaInicio.HasValue
                    ? fechaInicio.Value.ToString("dd/MM/yyyy")
                    : "Todas";

            hoja.Cell("G6").Value = "Fecha fin hasta:";
            hoja.Cell("H6").Value =
                fechaFin.HasValue
                    ? fechaFin.Value.ToString("dd/MM/yyyy")
                    : "Todas";

            hoja.Range("A6:H6").Style.Font.Bold = false;

            hoja.Cell("A6").Style.Font.Bold = true;
            hoja.Cell("C6").Style.Font.Bold = true;
            hoja.Cell("E6").Style.Font.Bold = true;
            hoja.Cell("G6").Style.Font.Bold = true;
            hoja.Range("A6:H6").Style.Fill.BackgroundColor =
    colorFondoFiltros;

            hoja.Range("A6:H6").Style.Border.OutsideBorder =
                XLBorderStyleValues.Thin;

            hoja.Range("A6:H6").Style.Border.InsideBorder =
                XLBorderStyleValues.Thin;

            hoja.Range("A6:H6").Style.Border.OutsideBorderColor =
                colorInstitucional;

            hoja.Range("A6:H6").Style.Border.InsideBorderColor =
                XLColor.LightGray;
            // =================================================
            // RESUMEN
            // =================================================
            hoja.Cell("A8").Value = "Resumen";
            hoja.Range("A8:H8").Merge();
            hoja.Range("A8:H8").Style.Font.Bold = true;
            hoja.Range("A8:H8").Style.Fill.BackgroundColor =
                colorInstitucional;

            hoja.Range("A8:H8").Style.Font.FontColor =
                XLColor.White;

            hoja.Range("A8:H8").Style.Font.Bold = true;
            hoja.Cell("A9").Value = "Total proyectos";
            hoja.Cell("B9").Value = proyectos.Count;

            hoja.Cell("C9").Value = "Activos";
            hoja.Cell("D9").Value = proyectos.Count(p =>
                p.Estado == "Activo" || p.Activo);

            hoja.Cell("E9").Value = "Inactivos";
            hoja.Cell("F9").Value = proyectos.Count(p =>
                p.Estado == "Inactivo" || !p.Activo);

            hoja.Cell("G9").Value = "Presupuesto total";
            hoja.Cell("H9").Value = proyectos.Sum(p =>
                p.PresupuestoAsignado ?? 0);

            hoja.Cell("H9").Style.NumberFormat.Format =
                "₡#,##0.00";

            hoja.Range("A9:H9").Style.Font.Bold = true;
            hoja.Range("A9:B9").Style.Fill.BackgroundColor =
    colorFondoSuave;

            hoja.Range("C9:D9").Style.Fill.BackgroundColor =
                colorVerdeSuave;

            hoja.Range("E9:F9").Style.Fill.BackgroundColor =
                colorGrisSuave;

            hoja.Range("G9:H9").Style.Fill.BackgroundColor =
                colorDoradoSuave;

            hoja.Range("A9:H9").Style.Border.OutsideBorder =
                XLBorderStyleValues.Thin;

            hoja.Range("A9:H9").Style.Border.InsideBorder =
                XLBorderStyleValues.Thin;

            hoja.Range("A9:H9").Style.Alignment.Vertical =
                XLAlignmentVerticalValues.Center;

            hoja.Range("A9:H9").Style.Font.Bold = true;

            hoja.Row(9).Height = 22;
            // =================================================
            // ENCABEZADOS DE LA TABLA
            // =================================================
            int filaEncabezado = 11;

            hoja.Cell(filaEncabezado, 1).Value = "Nombre";
            hoja.Cell(filaEncabezado, 2).Value = "Descripción";
            hoja.Cell(filaEncabezado, 3).Value = "Fecha inicio";
            hoja.Cell(filaEncabezado, 4).Value = "Fecha fin";
            hoja.Cell(filaEncabezado, 5).Value = "Estado";
            hoja.Cell(filaEncabezado, 6).Value =
                "Estado presupuestario";
            hoja.Cell(filaEncabezado, 7).Value =
                "Presupuesto asignado";
            hoja.Cell(filaEncabezado, 8).Value =
                "Presupuestos registrados";

            var encabezado = hoja.Range(
     filaEncabezado,
     1,
     filaEncabezado,
     8);

            encabezado.Style.Fill.BackgroundColor =
                colorInstitucional;

            encabezado.Style.Font.FontColor =
                XLColor.White;

            encabezado.Style.Font.Bold = true;

            encabezado.Style.Alignment.Horizontal =
                XLAlignmentHorizontalValues.Center;

            encabezado.Style.Alignment.Vertical =
                XLAlignmentVerticalValues.Center;

            encabezado.Style.Border.OutsideBorder =
                XLBorderStyleValues.Thin;

            encabezado.Style.Border.InsideBorder =
                XLBorderStyleValues.Thin;

            encabezado.Style.Border.OutsideBorderColor =
                colorInstitucionalOscuro;

            encabezado.Style.Border.InsideBorderColor =
                colorInstitucionalOscuro;

            hoja.Row(filaEncabezado).Height = 28;

            // =================================================
            // DATOS
            // =================================================
            int fila = filaEncabezado + 1;

            foreach (var proyecto in proyectos)
            {
                bool tieneAprobado =
                    proyecto.PresupuestosMensuales != null &&
                    proyecto.PresupuestosMensuales.Any(p =>
                        p.EstadoAprobacion == "Aprobado");

                bool tienePlanificacion =
                    proyecto.PresupuestosMensuales != null &&
                    proyecto.PresupuestosMensuales.Any();

                string estadoPresupuestario;

                if (tieneAprobado)
                {
                    estadoPresupuestario =
                        "Presupuesto aprobado";
                }
                else if (tienePlanificacion)
                {
                    estadoPresupuestario =
                        "En planificación";
                }
                else
                {
                    estadoPresupuestario =
                        "Sin presupuesto";
                }

                hoja.Cell(fila, 1).Value =
                    proyecto.Nombre;

                hoja.Cell(fila, 2).Value =
                    string.IsNullOrWhiteSpace(proyecto.Descripcion)
                        ? "Sin descripción"
                        : proyecto.Descripcion;

                hoja.Cell(fila, 3).Value =
                    proyecto.FechaInicio.HasValue
                        ? proyecto.FechaInicio.Value
                            .ToString("dd/MM/yyyy")
                        : "Sin fecha";

                hoja.Cell(fila, 4).Value =
                    proyecto.FechaFin.HasValue
                        ? proyecto.FechaFin.Value
                            .ToString("dd/MM/yyyy")
                        : "Sin fecha";

                hoja.Cell(fila, 5).Value =
                    proyecto.Activo
                        ? "Activo"
                        : "Inactivo";

                hoja.Cell(fila, 6).Value =
                    estadoPresupuestario;

                hoja.Cell(fila, 7).Value =
                    proyecto.PresupuestoAsignado ?? 0;

                hoja.Cell(fila, 7)
                    .Style.NumberFormat.Format =
                    "₡#,##0.00";

                hoja.Cell(fila, 8).Value =
                    proyecto.PresupuestosMensuales?.Count ?? 0;
                // Color alternado de filas
                if (fila % 2 == 0)
                {
                    hoja.Range(fila, 1, fila, 8)
                        .Style.Fill.BackgroundColor =
                        XLColor.FromHtml("#F7F9FA");
                }

                // Estado del proyecto
                if (proyecto.Activo)
                {
                    hoja.Cell(fila, 5).Style.Fill.BackgroundColor =
                        colorVerdeSuave;

                    hoja.Cell(fila, 5).Style.Font.FontColor =
                        colorVerde;

                    hoja.Cell(fila, 5).Style.Font.Bold = true;
                }
                else
                {
                    hoja.Cell(fila, 5).Style.Fill.BackgroundColor =
                        colorGrisSuave;

                    hoja.Cell(fila, 5).Style.Font.FontColor =
                        colorGris;

                    hoja.Cell(fila, 5).Style.Font.Bold = true;
                }

                // Estado presupuestario
                if (tieneAprobado)
                {
                    hoja.Cell(fila, 6).Style.Fill.BackgroundColor =
                        colorVerdeSuave;

                    hoja.Cell(fila, 6).Style.Font.FontColor =
                        colorVerde;
                }
                else if (tienePlanificacion)
                {
                    hoja.Cell(fila, 6).Style.Fill.BackgroundColor =
                        colorAmarilloSuave;

                    hoja.Cell(fila, 6).Style.Font.FontColor =
                        XLColor.FromHtml("#856404");
                }
                else
                {
                    hoja.Cell(fila, 6).Style.Fill.BackgroundColor =
                        colorGrisSuave;

                    hoja.Cell(fila, 6).Style.Font.FontColor =
                        colorGris;
                }

                hoja.Cell(fila, 6).Style.Font.Bold = true;
                fila++;
            }

            // =================================================
            // MENSAJE CUANDO NO HAY REGISTROS
            // =================================================
            if (!proyectos.Any())
            {
                hoja.Cell(fila, 1).Value =
                    "No se encontraron proyectos con los filtros aplicados.";

                hoja.Range(fila, 1, fila, 8).Merge();

                hoja.Range(fila, 1, fila, 8)
                    .Style.Alignment.Horizontal =
                    XLAlignmentHorizontalValues.Center;

                hoja.Range(fila, 1, fila, 8)
                    .Style.Font.Italic = true;
            }

            // =================================================
            // TOTAL FINAL
            // =================================================
            int filaTotal = fila + 1;

            hoja.Cell(filaTotal, 6).Value =
                "Presupuesto total:";

            hoja.Cell(filaTotal, 7).Value =
                proyectos.Sum(p =>
                    p.PresupuestoAsignado ?? 0);

            hoja.Cell(filaTotal, 7)
                .Style.NumberFormat.Format =
                "₡#,##0.00";

            hoja.Range(filaTotal, 6, filaTotal, 7)
                .Style.Font.Bold = true;

            // =================================================
            // FORMATO GENERAL
            // =================================================
            int ultimaFila = Math.Max(
                filaTotal,
                filaEncabezado + 1);
            hoja.Range(filaTotal, 1, filaTotal, 8)
    .Style.Fill.BackgroundColor =
    colorFondoSuave;

            hoja.Range(filaTotal, 1, filaTotal, 8)
                .Style.Border.TopBorder =
                XLBorderStyleValues.Medium;

            hoja.Range(filaTotal, 1, filaTotal, 8)
                .Style.Border.TopBorderColor =
                colorInstitucional;

            hoja.Cell(filaTotal, 6).Style.Font.FontColor =
                colorInstitucionalOscuro;

            hoja.Cell(filaTotal, 7).Style.Font.FontColor =
                colorInstitucionalOscuro;
            hoja.Range(
                    filaEncabezado,
                    1,
                    ultimaFila,
                    8)
                .Style.Border.OutsideBorder =
                XLBorderStyleValues.Thin;

            hoja.Range(
                    filaEncabezado,
                    1,
                    ultimaFila,
                    8)
                .Style.Border.InsideBorder =
                XLBorderStyleValues.Thin;

            hoja.Range(
                    filaEncabezado + 1,
                    3,
                    ultimaFila,
                    6)
                .Style.Alignment.Horizontal =
                XLAlignmentHorizontalValues.Center;

            hoja.Range(
                    filaEncabezado + 1,
                    7,
                    ultimaFila,
                    8)
                .Style.Alignment.Horizontal =
                XLAlignmentHorizontalValues.Right;

            hoja.Columns().AdjustToContents();

            hoja.Column(1).Width = 28;
            hoja.Column(2).Width = 45;
            hoja.Column(6).Width = 25;

            hoja.SheetView.FreezeRows(filaEncabezado);

            hoja.PageSetup.PageOrientation =
                XLPageOrientation.Landscape;

            hoja.PageSetup.FitToPages(1, 0);

            using var stream = new MemoryStream();

            workbook.SaveAs(stream);

            string nombreArchivo =
                $"Reporte_Proyectos_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";
            await RegistrarAuditoria(
    "Proyectos",
    0,
    "Exportar Excel",
    "Se exportó el reporte de proyectos a Excel."
);

            return File(
                stream.ToArray(),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                nombreArchivo);
        }
        private bool EsAdministrador()
        {
            string rol = HttpContext.Session.GetString("Rol") ?? "";
            return rol == "Administrador";
        }

        private void NormalizarProyecto(Proyecto proyecto)
        {
            proyecto.Nombre =
                string.IsNullOrWhiteSpace(proyecto.Nombre)
                    ? string.Empty
                    : proyecto.Nombre.Trim();

            proyecto.Descripcion =
                string.IsNullOrWhiteSpace(proyecto.Descripcion)
                    ? string.Empty
                    : proyecto.Descripcion.Trim();

            proyecto.Estado =
                string.IsNullOrWhiteSpace(proyecto.Estado)
                    ? string.Empty
                    : proyecto.Estado.Trim();

            proyecto.MotivoAjustePresupuesto =
                string.IsNullOrWhiteSpace(
                    proyecto.MotivoAjustePresupuesto)
                    ? null
                    : proyecto.MotivoAjustePresupuesto.Trim();
        }
        private async Task<decimal> CalcularPresupuestoComprometido(
    int proyectoId)
        {
            return await _context.PresupuestosMensuales
                .Where(p =>
                    p.ProyectoId == proyectoId &&
                    (
                        p.EstadoAprobacion == "Borrador" ||
                        p.EstadoAprobacion == "Aprobado"
                    ))
                .SumAsync(p =>
                    p.MontoGastoPlanificado ?? 0m);
        }
        private async Task ValidarProyecto(Proyecto proyecto, bool esEdicion)
        {
            if (string.IsNullOrWhiteSpace(proyecto.Nombre))
                ModelState.AddModelError("Nombre", "El nombre del proyecto es obligatorio.");
            else
            {
                if (proyecto.Nombre.Length < 3)
                    ModelState.AddModelError("Nombre", "El nombre debe tener al menos 3 caracteres reales.");

                if (!TextoValido(proyecto.Nombre, permitirNumeros: true))
                    ModelState.AddModelError("Nombre", "El nombre contiene caracteres no permitidos o no es válido.");

                if (SoloNumeros(proyecto.Nombre))
                    ModelState.AddModelError("Nombre", "El nombre no puede contener solo números.");

                if (TextoRepetitivo(proyecto.Nombre))
                    ModelState.AddModelError("Nombre", "El nombre no puede estar formado por caracteres repetidos.");

                bool nombreExiste = await _context.Proyectos.AnyAsync(p =>
                    p.Nombre == proyecto.Nombre &&
                    p.IdProyecto != proyecto.IdProyecto);

                if (nombreExiste)
                    ModelState.AddModelError("Nombre", "Ya existe un proyecto con este nombre.");
            }

            if (!string.IsNullOrWhiteSpace(proyecto.Descripcion))
            {
                if (proyecto.Descripcion.Length < 5)
                    ModelState.AddModelError("Descripcion", "La descripción debe tener al menos 5 caracteres reales.");

                if (!TextoValido(proyecto.Descripcion, permitirNumeros: true))
                    ModelState.AddModelError("Descripcion", "La descripción contiene caracteres no permitidos.");

                if (TextoRepetitivo(proyecto.Descripcion))
                    ModelState.AddModelError("Descripcion", "La descripción no puede estar formada por caracteres repetidos.");
            }

            if (proyecto.FechaFin.HasValue && proyecto.FechaFin < proyecto.FechaInicio)
                ModelState.AddModelError("FechaFin", "La fecha final no puede ser menor que la fecha de inicio.");

            if (proyecto.PresupuestoAsignado <= 0)
            {
                ModelState.AddModelError(
                    "PresupuestoAsignado",
                    "El presupuesto asignado debe ser mayor que cero.");
            }

            if (string.IsNullOrWhiteSpace(proyecto.Estado))
                ModelState.AddModelError("Estado", "Debe seleccionar el estado del proyecto.");

            if (proyecto.Estado != "Activo" && proyecto.Estado != "Inactivo")
                ModelState.AddModelError("Estado", "El estado seleccionado no es válido.");
        }

        private async Task<bool> ProyectoTieneRegistros(int id)
        {
            bool tieneIngresos =
                await _context.Ingresos
                    .AnyAsync(i =>
                        i.ProyectoId == id &&
                        i.Estado != "Anulado");

            bool tieneGastos =
                await _context.Gastos
                    .AnyAsync(g =>
                        g.ProyectoId == id &&
                        g.Estado != "Anulado" &&
                        g.Estado != "Rechazado");

            bool tieneFacturas =
                await _context.Facturas
                    .AnyAsync(f =>
                        f.ProyectoId == id &&
                        f.Estado != "Anulada");

            bool tieneCxC =
                await _context.CuentasPorCobrars
                    .AnyAsync(c =>
                        c.ProyectoId == id &&
                        c.Estado != "Anulada");

            bool tieneCxP =
                await _context.CuentasPorPagars
                    .AnyAsync(c =>
                        c.ProyectoId == id &&
                        c.Estado != "Anulada");

            bool tieneMovimientos =
                await _context.MovimientosContables
                    .AnyAsync(m =>
                        m.ProyectoId == id &&
                        m.Anulado == false);

            bool tieneRelacionPresupuestaria =
        await _context.PresupuestosMensuales
            .AnyAsync(p =>
                p.ProyectoId == id &&
                (
                    p.EstadoAprobacion == "Borrador" ||
                    p.EstadoAprobacion == "Aprobado"
                ));

            return tieneIngresos ||
                   tieneGastos ||
                   tieneFacturas ||
                   tieneCxC ||
                   tieneCxP ||
                   tieneMovimientos ||
                   tieneRelacionPresupuestaria;
        }

        private bool TextoValido(string texto, bool permitirNumeros)
        {
            if (string.IsNullOrWhiteSpace(texto))
                return false;

            string patron = permitirNumeros
                ? @"^[a-zA-ZáéíóúÁÉÍÓÚñÑ0-9\s.,()#/-]+$"
                : @"^[a-zA-ZáéíóúÁÉÍÓÚñÑ\s]+$";

            return Regex.IsMatch(texto, patron);
        }

        private bool SoloNumeros(string texto)
        {
            return Regex.IsMatch(texto.Replace(" ", ""), @"^\d+$");
        }

        private bool TextoRepetitivo(string texto)
        {
            var limpio = new string(texto
                .Where(c => !char.IsWhiteSpace(c))
                .ToArray())
                .ToLower();

            if (limpio.Length < 4)
                return false;

            return limpio.Distinct().Count() == 1;
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
        private async Task RegistrarAuditoria(string tabla, int registroId, string accion, string descripcion)
        {
            int usuarioId = HttpContext.Session.GetInt32("UsuarioId") ?? 1;

            var auditoria = new Auditorium
            {
                UsuarioId = usuarioId,
                Tabla = tabla,
                RegistroId = registroId,
                Accion = accion,
                Descripcion = descripcion,
                Fecha = DateTime.Now
            };

            _context.Auditoria.Add(auditoria);
            await _context.SaveChangesAsync();
        }
        private async Task<bool> PeriodoContableCerradoAsync(DateOnly fecha)
        {
            return await _context.CierresContables
                .AsNoTracking()
                .AnyAsync(c =>
                    c.Anio == fecha.Year &&
                    c.Mes == fecha.Month &&
                    c.Estado == "Cerrado");
        }
        private bool ProyectoExists(int id)
        {
            return _context.Proyectos.Any(e => e.IdProyecto == id);
        }
    }
}