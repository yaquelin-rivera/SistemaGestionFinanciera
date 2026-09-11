using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Http;
using Microsoft.EntityFrameworkCore;
using SistemaGestionFinanciera.Data;
using SistemaGestionFinanciera.Models;
using ClosedXML.Excel;
using System.IO;

namespace SistemaGestionFinanciera.Controllers
{
    public class AuditoriumsController : BaseController
    {
        private readonly SistemaFinancieroContext _context;

        public AuditoriumsController(SistemaFinancieroContext context)
        {
            _context = context;
        }
        // INDEX - AUDITORÍA
        // ================================================================
        [HttpGet]
        public async Task<IActionResult> Index(
            string? buscar,
            string? modulo,
            string? accion,
            DateTime? desde,
            DateTime? hasta)
        {
            var consulta = ConstruirConsultaReporte(
                buscar,
                modulo,
                accion,
                desde,
                hasta);

            var auditorias = await consulta
                .OrderByDescending(a => a.Fecha)
                .ToListAsync();

            // Listas completas para los filtros.
            ViewBag.Modulos = await _context.Auditoria
                .Where(a => a.Tabla != null && a.Tabla != "")
                .Select(a => a.Tabla)
                .Distinct()
                .OrderBy(a => a)
                .ToListAsync();

            ViewBag.Acciones = await _context.Auditoria
                .Where(a => a.Accion != null && a.Accion != "")
                .Select(a => a.Accion)
                .Distinct()
                .OrderBy(a => a)
                .ToListAsync();

            ViewBag.Buscar = buscar;
            ViewBag.Modulo = modulo;
            ViewBag.Accion = accion;
            ViewBag.Desde = desde?.ToString("yyyy-MM-dd");
            ViewBag.Hasta = hasta?.ToString("yyyy-MM-dd");

            return View(auditorias);
        }
        // ================================================================
        // REPORTE DE AUDITORÍA
        // ================================================================
        [HttpGet]
        public async Task<IActionResult> Reporte(
            string? buscar,
            string? modulo,
            string? accion,
            DateTime? desde,
            DateTime? hasta)
        {
            var auditorias = await ConstruirConsultaReporte(
                    buscar,
                    modulo,
                    accion,
                    desde,
                    hasta)
                .OrderByDescending(a => a.Fecha)
                .ToListAsync();

            ViewBag.TotalRegistros =
                auditorias.Count;

            ViewBag.RegistrosHoy =
                auditorias.Count(a =>
                    a.Fecha.Date == DateTime.Today);

            ViewBag.UsuariosDistintos =
                auditorias
                    .Select(a => a.UsuarioId)
                    .Distinct()
                    .Count();

            ViewBag.ModulosAuditados =
                auditorias
                    .Where(a => !string.IsNullOrWhiteSpace(a.Tabla))
                    .Select(a => a.Tabla)
                    .Distinct()
                    .Count();

            ViewBag.Buscar = buscar;
            ViewBag.Modulo = modulo;
            ViewBag.Accion = accion;
            ViewBag.Desde = desde;
            ViewBag.Hasta = hasta;
            ViewBag.FechaGeneracion = DateTime.Now;
            await RegistrarAuditoria(
                "Vista previa PDF",
                "Auditoria",
                0,
                "Se generó la vista previa del reporte de auditoría del sistema."
            );

            return View(auditorias);
        }
        // ================================================================
        // EXPORTAR AUDITORÍA A EXCEL
        // ================================================================
        [HttpGet]
        public async Task<IActionResult> ExportarExcel(
            string? buscar,
            string? modulo,
            string? accion,
            DateTime? desde,
            DateTime? hasta)
        {
            var auditorias = await ConstruirConsultaReporte(
                    buscar,
                    modulo,
                    accion,
                    desde,
                    hasta)
                .OrderByDescending(a => a.Fecha)
                .ToListAsync();

            using var workbook = new XLWorkbook();

            var hoja = workbook.Worksheets.Add("Auditoría");

            const int totalColumnas = 6;

            // ============================================================
            // COLORES
            // ============================================================
            var colorInstitucional =
                XLColor.FromHtml("#0F6B73");

            var colorInstitucionalOscuro =
                XLColor.FromHtml("#0B4F55");

            var colorFondoSuave =
                XLColor.FromHtml("#E7F1F2");

            var colorFondoFiltros =
                XLColor.FromHtml("#F3F7F8");

            var colorAzulSuave =
                XLColor.FromHtml("#D9EAF7");

            var colorVerdeSuave =
                XLColor.FromHtml("#D1E7DD");

            var colorGrisSuave =
                XLColor.FromHtml("#E2E3E5");

            var colorDoradoSuave =
                XLColor.FromHtml("#FFF3CD");

            // ============================================================
            // ENCABEZADO INSTITUCIONAL
            // ============================================================
            hoja.Cell("A1").Value =
                "Unión Cantonal de Asociaciones de Desarrollo de Santa Ana";

            hoja.Range(
                1,
                1,
                1,
                totalColumnas).Merge();

            hoja.Cell("A2").Value =
                "Reporte de Auditoría del Sistema";

            hoja.Range(
                2,
                1,
                2,
                totalColumnas).Merge();

            hoja.Cell("A3").Value =
                $"Fecha de generación: {DateTime.Now:dd/MM/yyyy hh:mm tt}";

            hoja.Range(
                3,
                1,
                3,
                totalColumnas).Merge();

            var rangoInstitucion = hoja.Range(
                1,
                1,
                1,
                totalColumnas);

            rangoInstitucion.Style.Font.Bold = true;
            rangoInstitucion.Style.Font.FontSize = 16;

            rangoInstitucion.Style.Font.FontColor =
                XLColor.FromHtml("#0F5C64");

            rangoInstitucion.Style.Alignment.Horizontal =
                XLAlignmentHorizontalValues.Center;

            rangoInstitucion.Style.Alignment.Vertical =
                XLAlignmentVerticalValues.Center;

            var rangoTitulo = hoja.Range(
                2,
                1,
                2,
                totalColumnas);

            rangoTitulo.Style.Font.Bold = true;
            rangoTitulo.Style.Font.FontSize = 13;
            rangoTitulo.Style.Font.FontColor = XLColor.Black;

            rangoTitulo.Style.Alignment.Horizontal =
                XLAlignmentHorizontalValues.Center;

            rangoTitulo.Style.Alignment.Vertical =
                XLAlignmentVerticalValues.Center;

            var rangoFecha = hoja.Range(
                3,
                1,
                3,
                totalColumnas);

            rangoFecha.Style.Font.FontColor =
                XLColor.FromHtml("#666666");

            rangoFecha.Style.Alignment.Horizontal =
                XLAlignmentHorizontalValues.Center;

            rangoFecha.Style.Alignment.Vertical =
                XLAlignmentVerticalValues.Center;

            hoja.Row(1).Height = 24;
            hoja.Row(2).Height = 22;
            hoja.Row(3).Height = 20;

            // ============================================================
            // FILTROS APLICADOS
            // ============================================================
            hoja.Cell("A5").Value =
                "Filtros aplicados";

            hoja.Range(
                5,
                1,
                5,
                totalColumnas).Merge();

            var tituloFiltros = hoja.Range(
                5,
                1,
                5,
                totalColumnas);

            tituloFiltros.Style.Fill.BackgroundColor =
                colorInstitucional;

            tituloFiltros.Style.Font.FontColor =
                XLColor.White;

            tituloFiltros.Style.Font.Bold = true;

            hoja.Cell("A6").Value = "Buscar:";
            hoja.Cell("B6").Value =
                string.IsNullOrWhiteSpace(buscar)
                    ? "Todos"
                    : buscar;

            hoja.Cell("C6").Value = "Módulo:";
            hoja.Cell("D6").Value =
                string.IsNullOrWhiteSpace(modulo)
                    ? "Todos"
                    : modulo;

            hoja.Cell("E6").Value = "Acción:";
            hoja.Cell("F6").Value =
                string.IsNullOrWhiteSpace(accion)
                    ? "Todas"
                    : accion;

            hoja.Cell("A7").Value = "Desde:";
            hoja.Cell("B7").Value =
                desde.HasValue
                    ? desde.Value.ToString("dd/MM/yyyy")
                    : "Sin límite";

            hoja.Cell("C7").Value = "Hasta:";
            hoja.Cell("D7").Value =
                hasta.HasValue
                    ? hasta.Value.ToString("dd/MM/yyyy")
                    : "Sin límite";

            hoja.Cell("E7").Value = "Orden:";
            hoja.Cell("F7").Value =
                "Más reciente primero";

            foreach (int columna in new[] { 1, 3, 5 })
            {
                hoja.Cell(6, columna).Style.Font.Bold = true;
                hoja.Cell(7, columna).Style.Font.Bold = true;
            }

            var rangoFiltros = hoja.Range(
                6,
                1,
                7,
                totalColumnas);

            rangoFiltros.Style.Fill.BackgroundColor =
                colorFondoFiltros;

            rangoFiltros.Style.Border.OutsideBorder =
                XLBorderStyleValues.Thin;

            rangoFiltros.Style.Border.InsideBorder =
                XLBorderStyleValues.Thin;

            rangoFiltros.Style.Border.OutsideBorderColor =
                colorInstitucional;

            rangoFiltros.Style.Border.InsideBorderColor =
                XLColor.LightGray;

            rangoFiltros.Style.Alignment.Vertical =
                XLAlignmentVerticalValues.Center;

            // ============================================================
            // RESUMEN
            // ============================================================
            hoja.Cell("A9").Value =
                "Resumen";

            hoja.Range(
                9,
                1,
                9,
                totalColumnas).Merge();

            var tituloResumen = hoja.Range(
                9,
                1,
                9,
                totalColumnas);

            tituloResumen.Style.Fill.BackgroundColor =
                colorInstitucional;

            tituloResumen.Style.Font.FontColor =
                XLColor.White;

            tituloResumen.Style.Font.Bold = true;

            int totalEventos =
                auditorias.Count;

            int eventosHoy =
                auditorias.Count(a =>
                    a.Fecha.Date == DateTime.Today);

            int usuariosResponsables =
                auditorias
                    .Select(a => a.UsuarioId)
                    .Distinct()
                    .Count();

            int modulosAuditados =
                auditorias
                    .Where(a => !string.IsNullOrWhiteSpace(a.Tabla))
                    .Select(a => a.Tabla)
                    .Distinct()
                    .Count();

            hoja.Cell("A10").Value =
                "Total de eventos";

            hoja.Cell("B10").Value =
                totalEventos;

            hoja.Cell("C10").Value =
                "Eventos de hoy";

            hoja.Cell("D10").Value =
                eventosHoy;

            hoja.Cell("E10").Value =
                "Usuarios responsables";

            hoja.Cell("F10").Value =
                usuariosResponsables;

            hoja.Cell("A11").Value =
                "Módulos auditados";

            hoja.Cell("B11").Value =
                modulosAuditados;

            hoja.Range("A10:B10")
                .Style.Fill.BackgroundColor =
                colorAzulSuave;

            hoja.Range("C10:D10")
                .Style.Fill.BackgroundColor =
                colorVerdeSuave;

            hoja.Range("E10:F10")
                .Style.Fill.BackgroundColor =
                colorDoradoSuave;

            hoja.Range("A11:B11")
                .Style.Fill.BackgroundColor =
                colorGrisSuave;

            hoja.Range("A10:F11").Style.Font.Bold = true;

            hoja.Range("A10:F11")
                .Style.Border.OutsideBorder =
                XLBorderStyleValues.Thin;

            hoja.Range("A10:F11")
                .Style.Border.InsideBorder =
                XLBorderStyleValues.Thin;

            hoja.Range("A10:F11")
                .Style.Alignment.Vertical =
                XLAlignmentVerticalValues.Center;

            // ============================================================
            // ENCABEZADO DE LA TABLA
            // ============================================================
            const int filaEncabezado = 13;

            string[] encabezados =
            {
        "Fecha y hora",
        "Responsable",
        "Módulo",
        "Acción",
        "Registro",
        "Descripción"
    };

            for (int columna = 0;
                 columna < encabezados.Length;
                 columna++)
            {
                hoja.Cell(
                    filaEncabezado,
                    columna + 1).Value =
                    encabezados[columna];
            }

            var encabezadoTabla = hoja.Range(
                filaEncabezado,
                1,
                filaEncabezado,
                totalColumnas);

            encabezadoTabla.Style.Fill.BackgroundColor =
                colorInstitucional;

            encabezadoTabla.Style.Font.FontColor =
                XLColor.White;

            encabezadoTabla.Style.Font.Bold = true;

            encabezadoTabla.Style.Alignment.Horizontal =
                XLAlignmentHorizontalValues.Center;

            encabezadoTabla.Style.Alignment.Vertical =
                XLAlignmentVerticalValues.Center;

            encabezadoTabla.Style.Border.OutsideBorder =
                XLBorderStyleValues.Thin;

            encabezadoTabla.Style.Border.InsideBorder =
                XLBorderStyleValues.Thin;

            encabezadoTabla.Style.Border.OutsideBorderColor =
                colorInstitucionalOscuro;

            encabezadoTabla.Style.Border.InsideBorderColor =
                colorInstitucionalOscuro;

            hoja.Row(filaEncabezado).Height = 28;

            // ============================================================
            // DATOS
            // ============================================================
            int fila = filaEncabezado + 1;

            foreach (var auditoria in auditorias)
            {
                string responsable =
                    auditoria.Usuario != null
                        ? auditoria.Usuario.Nombre
                        : "Usuario no disponible";

                if (auditoria.Usuario?.Rol != null)
                {
                    responsable +=
                        $" - {auditoria.Usuario.Rol.Nombre}";
                }

                hoja.Cell(fila, 1).Value =
                    auditoria.Fecha;

                hoja.Cell(fila, 1)
                    .Style.DateFormat.Format =
                    "dd/MM/yyyy hh:mm AM/PM";

                hoja.Cell(fila, 2).Value =
                    responsable;

                hoja.Cell(fila, 3).Value =
                    auditoria.Tabla ?? "";

                hoja.Cell(fila, 4).Value =
                    auditoria.Accion ?? "";

                hoja.Cell(fila, 5).Value =
                    auditoria.RegistroId;

                hoja.Cell(fila, 6).Value =
                    string.IsNullOrWhiteSpace(auditoria.Descripcion)
                        ? "Sin descripción"
                        : auditoria.Descripcion;

                if (fila % 2 == 0)
                {
                    hoja.Range(
                            fila,
                            1,
                            fila,
                            totalColumnas)
                        .Style.Fill.BackgroundColor =
                        XLColor.FromHtml("#F7F9FA");
                }

                fila++;
            }

            // ============================================================
            // TABLA Y BORDES
            // ============================================================
            if (auditorias.Any())
            {
                int filaFinDatos = fila - 1;

                var rangoTabla = hoja.Range(
                    filaEncabezado,
                    1,
                    filaFinDatos,
                    totalColumnas);

                var tabla = rangoTabla.CreateTable(
                    "TablaAuditoria");

                tabla.Theme = XLTableTheme.None;
                tabla.ShowAutoFilter = true;
                tabla.ShowRowStripes = false;

                var rangoDatos = hoja.Range(
                    filaEncabezado + 1,
                    1,
                    filaFinDatos,
                    totalColumnas);

                rangoDatos.Style.Alignment.Vertical =
                    XLAlignmentVerticalValues.Center;

                rangoDatos.Style.Border.BottomBorder =
                    XLBorderStyleValues.Thin;

                rangoDatos.Style.Border.BottomBorderColor =
                    XLColor.LightGray;

                hoja.Range(
                        filaEncabezado + 1,
                        1,
                        filaFinDatos,
                        totalColumnas)
                    .Style.Alignment.WrapText = true;

                hoja.Range(
                        filaEncabezado + 1,
                        5,
                        filaFinDatos,
                        5)
                    .Style.Alignment.Horizontal =
                    XLAlignmentHorizontalValues.Center;
            }
            else
            {
                hoja.Cell(fila, 1).Value =
                    "No se encontraron registros de auditoría con los filtros aplicados.";

                hoja.Range(
                    fila,
                    1,
                    fila,
                    totalColumnas).Merge();

                hoja.Range(
                        fila,
                        1,
                        fila,
                        totalColumnas)
                    .Style.Alignment.Horizontal =
                    XLAlignmentHorizontalValues.Center;

                hoja.Range(
                        fila,
                        1,
                        fila,
                        totalColumnas)
                    .Style.Font.Italic = true;

                hoja.Range(
                        fila,
                        1,
                        fila,
                        totalColumnas)
                    .Style.Fill.BackgroundColor =
                    colorFondoSuave;

                fila++;
            }

            // ============================================================
            // TOTAL FINAL
            // ============================================================
            int filaTotal = fila + 1;

            hoja.Cell(filaTotal, 5).Value =
                "Total de registros:";

            hoja.Cell(filaTotal, 6).Value =
                auditorias.Count;

            var rangoTotal = hoja.Range(
                filaTotal,
                1,
                filaTotal,
                totalColumnas);

            rangoTotal.Style.Fill.BackgroundColor =
                colorFondoSuave;

            rangoTotal.Style.Border.TopBorder =
                XLBorderStyleValues.Medium;

            rangoTotal.Style.Border.TopBorderColor =
                colorInstitucional;

            hoja.Cell(filaTotal, 5).Style.Font.Bold = true;
            hoja.Cell(filaTotal, 6).Style.Font.Bold = true;

            hoja.Cell(filaTotal, 5)
                .Style.Font.FontColor =
                colorInstitucionalOscuro;

            hoja.Cell(filaTotal, 6)
                .Style.Font.FontColor =
                colorInstitucionalOscuro;

            hoja.Cell(filaTotal, 6)
                .Style.Alignment.Horizontal =
                XLAlignmentHorizontalValues.Center;

            // ============================================================
            // CONFIGURACIÓN FINAL
            // ============================================================
            hoja.Column(1).Width = 22;
            hoja.Column(2).Width = 34;
            hoja.Column(3).Width = 22;
            hoja.Column(4).Width = 20;
            hoja.Column(5).Width = 14;
            hoja.Column(6).Width = 55;

            hoja.Columns(1, totalColumnas)
                .Style.Alignment.WrapText = true;

            hoja.SheetView.FreezeRows(
                filaEncabezado);

            hoja.PageSetup.PageOrientation =
                XLPageOrientation.Landscape;

            hoja.PageSetup.PaperSize =
                XLPaperSize.A4Paper;

            hoja.PageSetup.FitToPages(1, 0);

            hoja.PageSetup.CenterHorizontally = true;

            hoja.PageSetup.SetRowsToRepeatAtTop(
                filaEncabezado,
                filaEncabezado);

            using var stream = new MemoryStream();

            workbook.SaveAs(stream);

            string nombreArchivo =
                $"Reporte_Auditoria_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";
            await RegistrarAuditoria(
    "Exportar Excel",
    "Auditoria",
    0,
    "Se exportó el reporte de auditoría del sistema a Excel."
);
            return File(
                stream.ToArray(),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                nombreArchivo);
        }
        // GET: Auditoria / detalles 
        public async Task<IActionResult> Details(int? id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var auditorium = await _context.Auditoria
    .Include(a => a.Usuario)
    .ThenInclude(u => u.Rol)
    .FirstOrDefaultAsync(m => m.IdAuditoria == id);
            if (auditorium == null)
            {
                return NotFound();
            }

            return View(auditorium);
        }
        // Auditoría es solo lectura. No se permite crear manualmente.
        public IActionResult Create()
        {
            TempData["MensajeError"] = "La auditoría se genera automáticamente por el sistema y no puede crearse manualmente.";
            return RedirectToAction(nameof(Index));
        }

        // Auditoría es solo lectura. No se permite modificar.
        public IActionResult Edit(int? id)
        {
            TempData["MensajeError"] = "Los registros de auditoría no pueden modificarse.";
            return RedirectToAction(nameof(Index));
        }

        // Auditoría es solo lectura. No se permite eliminar.
        public IActionResult Delete(int? id)
        {
            TempData["MensajeError"] = "Los registros de auditoría no pueden eliminarse.";
            return RedirectToAction(nameof(Index));
        }

        // Auditoría es solo lectura. No se permite eliminar por POST.
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public IActionResult DeleteConfirmed(int id)
        {
            TempData["MensajeError"] = "Los registros de auditoría no pueden eliminarse.";
            return RedirectToAction(nameof(Index));
        }
        // ================================================================
        // CONSULTA COMPARTIDA PARA INDEX, REPORTE Y EXCEL
        // ================================================================
        private IQueryable<Auditorium> ConstruirConsultaReporte(
            string? buscar,
            string? modulo,
            string? accion,
            DateTime? desde,
            DateTime? hasta)
        {
            var consulta = _context.Auditoria
                .Include(a => a.Usuario)
                    .ThenInclude(u => u.Rol)
                .AsNoTracking()
                .AsQueryable();

            if (!string.IsNullOrWhiteSpace(buscar))
            {
                buscar = buscar.Trim();

                consulta = consulta.Where(a =>
                    (a.Usuario != null &&
                     a.Usuario.Nombre.Contains(buscar)) ||

                    (a.Usuario != null &&
                     a.Usuario.Rol != null &&
                     a.Usuario.Rol.Nombre.Contains(buscar)) ||

                    (a.Tabla != null &&
                     a.Tabla.Contains(buscar)) ||

                    (a.Accion != null &&
                     a.Accion.Contains(buscar)) ||

                    (a.Descripcion != null &&
                     a.Descripcion.Contains(buscar)) ||

                    a.RegistroId.ToString().Contains(buscar));
            }

            if (!string.IsNullOrWhiteSpace(modulo))
            {
                consulta = consulta.Where(a =>
                    a.Tabla == modulo);
            }

            if (!string.IsNullOrWhiteSpace(accion))
            {
                consulta = consulta.Where(a =>
                    a.Accion == accion);
            }

            if (desde.HasValue)
            {
                DateTime fechaDesde =
                    desde.Value.Date;

                consulta = consulta.Where(a =>
                    a.Fecha >= fechaDesde);
            }

            if (hasta.HasValue)
            {
                DateTime fechaHastaExclusiva =
                    hasta.Value.Date.AddDays(1);

                consulta = consulta.Where(a =>
                    a.Fecha < fechaHastaExclusiva);
            }

            return consulta;
        }
        private async Task RegistrarAuditoria(
    string accion,
    string tabla,
    int registroId,
    string descripcion)
        {
            int usuarioId =
                HttpContext.Session.GetInt32("UsuarioId") ?? 1;

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
    }

}

