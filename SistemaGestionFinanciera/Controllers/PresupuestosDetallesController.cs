using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.AspNetCore.Http;
using Microsoft.EntityFrameworkCore;
using SistemaGestionFinanciera.Data;
using SistemaGestionFinanciera.Models;
using SistemaGestionFinanciera.Models.ViewModels;
using ClosedXML.Excel;
using System.IO;


namespace SistemaGestionFinanciera.Controllers
{
    public class PresupuestosDetallesController : Controller
    {
        private readonly SistemaFinancieroContext _context;

        public PresupuestosDetallesController(SistemaFinancieroContext context)
        {
            _context = context;
        }
        private IQueryable<PresupuestosMensuale> ConstruirConsultaReporte(
    int? proyectoId,
    int? anio,
    int? mes,
    DateOnly? desde,
    DateOnly? hasta)
        {
            var consulta = _context.PresupuestosMensuales
                .Include(p => p.Proyecto)
                .AsQueryable();

            if (proyectoId.HasValue && proyectoId.Value > 0)
            {
                consulta = consulta.Where(p =>
                    p.ProyectoId == proyectoId.Value);
            }

            if (anio.HasValue && anio.Value > 0)
            {
                consulta = consulta.Where(p =>
                    p.Anio == anio.Value);
            }

            if (mes.HasValue && mes.Value > 0)
            {
                consulta = consulta.Where(p =>
                    p.Mes == mes.Value);
            }
            if (desde.HasValue)
            {
                int desdeAnio = desde.Value.Year;
                int desdeMes = desde.Value.Month;

                consulta = consulta.Where(p =>
     p.Anio > desdeAnio ||
     (p.Anio == desdeAnio && p.Mes >= desdeMes));
            }

            if (hasta.HasValue)
            {
                int hastaAnio = hasta.Value.Year;
                int hastaMes = hasta.Value.Month;
                consulta = consulta.Where(p =>
    p.Anio < hastaAnio ||
    (p.Anio == hastaAnio && p.Mes <= hastaMes));
            }

            return consulta;
        }
        // GET: PresupuestosDetalles
        public async Task<IActionResult> Index(
      int? proyectoId,
      int? anio,
      int? mes,
      DateOnly? desde,
      DateOnly? hasta)
        {
            var presupuestos = _context.PresupuestosMensuales
                .Include(p => p.Proyecto)
                .AsQueryable();

            if (proyectoId.HasValue && proyectoId.Value > 0)
            {
                presupuestos = presupuestos.Where(p => p.ProyectoId == proyectoId.Value);
            }

            if (anio.HasValue && anio.Value > 0)
            {
                presupuestos = presupuestos.Where(p => p.Anio == anio.Value);
            }

            if (mes.HasValue && mes.Value > 0)
            {
                presupuestos = presupuestos.Where(p => p.Mes == mes.Value);
            }

            if (desde.HasValue)
            {
                int desdeAnio = desde.Value.Year;
                int desdeMes = desde.Value.Month;

                presupuestos = presupuestos.Where(p =>
                    p.Anio > desdeAnio ||
                    (p.Anio == desdeAnio && p.Mes >= desdeMes));
            }

            if (hasta.HasValue)
            {
                int hastaAnio = hasta.Value.Year;
                int hastaMes = hasta.Value.Month;

                presupuestos = presupuestos.Where(p =>
                    p.Anio < hastaAnio ||
                    (p.Anio == hastaAnio && p.Mes <= hastaMes));
            }

            var lista = await presupuestos
                .OrderByDescending(p => p.Anio)
                .ThenByDescending(p => p.Mes)
                .ThenBy(p => p.Proyecto.Nombre)
                .ToListAsync();

            ViewBag.ProyectoId = proyectoId;
            ViewBag.Anio = anio;
            ViewBag.Mes = mes;
            ViewBag.Desde = desde?.ToString("yyyy-MM-dd");
            ViewBag.Hasta = hasta?.ToString("yyyy-MM-dd");

            ViewBag.Proyectos = new SelectList(
                await _context.Proyectos
                    .Where(p => p.Activo)
                    .OrderBy(p => p.Nombre)
                    .ToListAsync(),
                "IdProyecto",
                "Nombre",
                proyectoId);
            // AÑOS Q ESTEN GUARDADOS EN BD 
            ViewBag.Anios = await _context.PresupuestosMensuales
    .Select(p => p.Anio)
    .Distinct()
    .OrderByDescending(a => a)
    .ToListAsync();

            ViewBag.FiltroAplicado =
                proyectoId.HasValue || anio.HasValue || mes.HasValue || desde.HasValue || hasta.HasValue;

            return View(lista);
        }

        public async Task<IActionResult> Reporte(
    int? proyectoId,
    int? anio,
    int? mes,
    DateOnly? desde,
    DateOnly? hasta)
        {
            var consulta = ConstruirConsultaReporte(
                proyectoId,
                anio,
                mes,
                desde,
                hasta);

            var lista = await consulta
                .OrderByDescending(p => p.Anio)
                .ThenByDescending(p => p.Mes)
                .ThenBy(p => p.Proyecto.Nombre)
                .ToListAsync();

            decimal totalIngresos = lista.Sum(
                p => p.TotalIngresoReal ?? 0);

            decimal totalGastos = lista.Sum(
                p => p.TotalGastoReal ?? 0);

            ViewBag.TotalRegistros = lista.Count;
            ViewBag.TotalIngresos = totalIngresos;
            ViewBag.TotalGastos = totalGastos;
            ViewBag.Balance = totalIngresos - totalGastos;

            ViewBag.SinMovimientos = lista.Count(
                p => p.Estado == "Sin movimientos");

            ViewBag.EnEjecucion = lista.Count(
                p => p.Estado == "En ejecución");

            ViewBag.IngresosBajos = lista.Count(
                p => p.Estado == "Ingresos bajos");

            ViewBag.GastoExcedido = lista.Count(
                p => p.Estado == "Gasto excedido");

            ViewBag.ProyectoId = proyectoId;
            ViewBag.Anio = anio;
            ViewBag.Mes = mes;
            ViewBag.Desde = desde?.ToString("yyyy-MM-dd");
            ViewBag.Hasta = hasta?.ToString("yyyy-MM-dd");

            ViewBag.NombreProyecto = proyectoId.HasValue
                ? await _context.Proyectos
                    .Where(p => p.IdProyecto == proyectoId.Value)
                    .Select(p => p.Nombre)
                    .FirstOrDefaultAsync()
                : null;

            ViewBag.FechaGeneracion = DateTime.Now;
            await RegistrarAuditoria(
    "Vista previa PDF",
    0,
    "Se generó la vista previa del reporte de detalle de presupuesto."
);

            
            return View(lista);
        }

        public async Task<IActionResult> ExportarExcel(
      int? proyectoId,
      int? anio,
      int? mes,
      DateOnly? desde,
      DateOnly? hasta)
        {
            var consulta = ConstruirConsultaReporte(
                proyectoId,
                anio,
                mes,
                desde,
                hasta);

            var lista = await consulta
                .OrderByDescending(p => p.Anio)
                .ThenByDescending(p => p.Mes)
                .ThenBy(p => p.Proyecto.Nombre)
                .ToListAsync();

            string nombreProyecto = proyectoId.HasValue
                ? await _context.Proyectos
                    .Where(p => p.IdProyecto == proyectoId.Value)
                    .Select(p => p.Nombre)
                    .FirstOrDefaultAsync() ?? "No encontrado"
                : "Todos";

            string nombreMes = mes switch
            {
                1 => "Enero",
                2 => "Febrero",
                3 => "Marzo",
                4 => "Abril",
                5 => "Mayo",
                6 => "Junio",
                7 => "Julio",
                8 => "Agosto",
                9 => "Septiembre",
                10 => "Octubre",
                11 => "Noviembre",
                12 => "Diciembre",
                _ => "Todos"
            };

            using var workbook = new XLWorkbook();

            var hoja = workbook.Worksheets.Add(
                "Detalle Presupuesto");

            // =====================================================
            // ENCABEZADO PRINCIPAL
            // =====================================================

            hoja.Cell("A1").Value =
                "Unión Cantonal de Asociaciones de Desarrollo de Santa Ana";

            hoja.Range("A1:H1").Merge();

            hoja.Cell("A2").Value =
                "Reporte de Detalle de Presupuesto";

            hoja.Range("A2:H2").Merge();

            hoja.Cell("A3").Value =
                $"Fecha de generación: {DateTime.Now:dd/MM/yyyy HH:mm}";

            hoja.Range("A3:H3").Merge();

            hoja.Range("A1:H3").Style.Alignment.Horizontal =
                XLAlignmentHorizontalValues.Center;

            hoja.Range("A1:H3").Style.Alignment.Vertical =
                XLAlignmentVerticalValues.Center;

            hoja.Range("A1:H1").Style.Font.Bold = true;
            hoja.Range("A1:H1").Style.Font.FontSize = 16;
            hoja.Range("A1:H1").Style.Font.FontColor =
                XLColor.FromHtml("#0F5C64");

            hoja.Range("A2:H2").Style.Font.Bold = true;
            hoja.Range("A2:H2").Style.Font.FontSize = 13;
            hoja.Range("A2:H2").Style.Font.FontColor =
                XLColor.FromHtml("#15454D");

            hoja.Range("A3:H3").Style.Font.FontSize = 10;
            hoja.Range("A3:H3").Style.Font.FontColor =
                XLColor.FromHtml("#6C757D");

            hoja.Row(1).Height = 24;
            hoja.Row(2).Height = 20;
            hoja.Row(3).Height = 18;

            // =====================================================
            // FILTROS
            // =====================================================

            hoja.Cell("A5").Value = "Filtros aplicados";
            hoja.Range("A5:H5").Merge();

            hoja.Range("A5:H5").Style.Fill.BackgroundColor =
                XLColor.FromHtml("#EAF3F4");

            hoja.Range("A5:H5").Style.Font.Bold = true;
            hoja.Range("A5:H5").Style.Font.FontColor =
                XLColor.FromHtml("#0F5C64");

            hoja.Range("A5:H5").Style.Border.BottomBorder =
                XLBorderStyleValues.Thin;

            hoja.Range("A5:H5").Style.Border.BottomBorderColor =
                XLColor.FromHtml("#0F7C90");

            hoja.Cell("A6").Value = "Proyecto";
            hoja.Cell("B6").Value = nombreProyecto;

            hoja.Cell("C6").Value = "Año";
            hoja.Cell("D6").Value =
                anio?.ToString() ?? "Todos";

            hoja.Cell("E6").Value = "Mes";
            hoja.Cell("F6").Value = nombreMes;

            hoja.Cell("G6").Value = "Desde / Hasta";
            hoja.Cell("H6").Value =
                $"{desde?.ToString("dd/MM/yyyy") ?? "Sin límite"} - " +
                $"{hasta?.ToString("dd/MM/yyyy") ?? "Sin límite"}";

            hoja.Range("A6:H6").Style.Fill.BackgroundColor =
                XLColor.FromHtml("#F7FAFA");

            hoja.Range("A6:H6").Style.Border.OutsideBorder =
                XLBorderStyleValues.Thin;

            hoja.Range("A6:H6").Style.Border.InsideBorder =
                XLBorderStyleValues.Thin;

            hoja.Range("A6:H6").Style.Border.OutsideBorderColor =
                XLColor.FromHtml("#D5E2E4");

            hoja.Range("A6:H6").Style.Border.InsideBorderColor =
                XLColor.FromHtml("#D5E2E4");

            foreach (int columna in new[] { 1, 3, 5, 7 })
            {
                hoja.Cell(6, columna).Style.Font.Bold = true;
                hoja.Cell(6, columna).Style.Font.FontColor =
                    XLColor.FromHtml("#0F5C64");
            }

            // =====================================================
            // ENCABEZADOS DE TABLA
            // =====================================================

            int filaEncabezado = 8;

            string[] encabezados =
            {
        "Proyecto",
        "Año",
        "Mes",
        "Ingresos",
        "Gastos",
        "Balance",
        "Estado",
        "Periodo"
    };

            for (int columna = 0;
                 columna < encabezados.Length;
                 columna++)
            {
                hoja.Cell(filaEncabezado, columna + 1).Value =
                    encabezados[columna];
            }

            var rangoEncabezado =
                hoja.Range(
                    filaEncabezado,
                    1,
                    filaEncabezado,
                    8);

            rangoEncabezado.Style.Font.Bold = true;
            rangoEncabezado.Style.Font.FontColor =
                XLColor.White;

            rangoEncabezado.Style.Fill.BackgroundColor =
                XLColor.FromHtml("#0F5C64");

            rangoEncabezado.Style.Alignment.Horizontal =
                XLAlignmentHorizontalValues.Center;

            rangoEncabezado.Style.Alignment.Vertical =
                XLAlignmentVerticalValues.Center;

            rangoEncabezado.Style.Border.OutsideBorder =
                XLBorderStyleValues.Thin;

            rangoEncabezado.Style.Border.InsideBorder =
                XLBorderStyleValues.Thin;

            rangoEncabezado.Style.Border.OutsideBorderColor =
                XLColor.White;

            rangoEncabezado.Style.Border.InsideBorderColor =
                XLColor.White;

            hoja.Row(filaEncabezado).Height = 22;

            // =====================================================
            // REGISTROS
            // =====================================================

            int fila = filaEncabezado + 1;

            foreach (var item in lista)
            {
                decimal ingresos =
                    item.TotalIngresoReal ?? 0;

                decimal gastos =
                    item.TotalGastoReal ?? 0;

                decimal balance =
                    ingresos - gastos;

                hoja.Cell(fila, 1).Value =
                    item.Proyecto?.Nombre ?? "Sin proyecto";

                hoja.Cell(fila, 2).Value =
                    item.Anio;

                hoja.Cell(fila, 3).Value =
                    item.Mes;

                hoja.Cell(fila, 4).Value =
                    ingresos;

                hoja.Cell(fila, 5).Value =
                    gastos;

                hoja.Cell(fila, 6).Value =
                    balance;

                hoja.Cell(fila, 7).Value =
                    item.Estado;

                hoja.Cell(fila, 8).Value =
                    $"{item.Mes:D2}/{item.Anio}";

                var rangoFila =
                    hoja.Range(fila, 1, fila, 8);

                rangoFila.Style.Border.OutsideBorder =
                    XLBorderStyleValues.Thin;

                rangoFila.Style.Border.InsideBorder =
                    XLBorderStyleValues.Thin;

                rangoFila.Style.Border.OutsideBorderColor =
                    XLColor.FromHtml("#D9E1E5");

                rangoFila.Style.Border.InsideBorderColor =
                    XLColor.FromHtml("#D9E1E5");

                rangoFila.Style.Alignment.Vertical =
                    XLAlignmentVerticalValues.Center;

                // Colores suaves según estado
                switch (item.Estado)
                {
                    case "Gasto excedido":
                        rangoFila.Style.Fill.BackgroundColor =
                            XLColor.FromHtml("#FBE5E8");

                        hoja.Cell(fila, 7)
                            .Style.Fill.BackgroundColor =
                            XLColor.FromHtml("#EE9CA7");

                        hoja.Cell(fila, 7)
                            .Style.Font.FontColor =
                            XLColor.FromHtml("#7A1C2A");

                        hoja.Cell(fila, 7)
                            .Style.Font.Bold = true;

                        break;

                    case "Ingresos bajos":
                        rangoFila.Style.Fill.BackgroundColor =
                            XLColor.FromHtml("#FDF4D6");

                        hoja.Cell(fila, 7)
                            .Style.Fill.BackgroundColor =
                            XLColor.FromHtml("#F3D36B");

                        hoja.Cell(fila, 7)
                            .Style.Font.FontColor =
                            XLColor.FromHtml("#5F4700");

                        hoja.Cell(fila, 7)
                            .Style.Font.Bold = true;

                        break;

                    case "Sin movimientos":
                        rangoFila.Style.Fill.BackgroundColor =
                            XLColor.FromHtml("#EEF1F4");

                        hoja.Cell(fila, 7)
                            .Style.Fill.BackgroundColor =
                            XLColor.FromHtml("#C3CCD6");

                        hoja.Cell(fila, 7)
                            .Style.Font.FontColor =
                            XLColor.FromHtml("#364152");

                        hoja.Cell(fila, 7)
                            .Style.Font.Bold = true;

                        break;

                    default:
                        rangoFila.Style.Fill.BackgroundColor =
                            XLColor.FromHtml("#EDF8EF");

                        hoja.Cell(fila, 7)
                            .Style.Fill.BackgroundColor =
                            XLColor.FromHtml("#8FD3A4");

                        hoja.Cell(fila, 7)
                            .Style.Font.FontColor =
                            XLColor.FromHtml("#14532D");

                        hoja.Cell(fila, 7)
                            .Style.Font.Bold = true;

                        break;
                }

                // Colores de montos
                hoja.Cell(fila, 4).Style.Font.FontColor =
                    XLColor.FromHtml("#198754");

                hoja.Cell(fila, 5).Style.Font.FontColor =
                    XLColor.FromHtml("#DC3545");

                if (balance < 0)
                {
                    hoja.Cell(fila, 6).Style.Font.FontColor =
                        XLColor.FromHtml("#DC3545");
                }

                hoja.Cell(fila, 6).Style.Font.Bold = true;

                hoja.Cell(fila, 2).Style.Alignment.Horizontal =
                    XLAlignmentHorizontalValues.Center;

                hoja.Cell(fila, 3).Style.Alignment.Horizontal =
                    XLAlignmentHorizontalValues.Center;

                hoja.Cell(fila, 7).Style.Alignment.Horizontal =
                    XLAlignmentHorizontalValues.Center;

                hoja.Cell(fila, 8).Style.Alignment.Horizontal =
                    XLAlignmentHorizontalValues.Center;

                fila++;
            }

            // =====================================================
            // TOTALES
            // =====================================================

            int filaTotal = fila;

            decimal totalIngresos =
                lista.Sum(p => p.TotalIngresoReal ?? 0);

            decimal totalGastos =
                lista.Sum(p => p.TotalGastoReal ?? 0);

            decimal balanceTotal =
                totalIngresos - totalGastos;

            hoja.Cell(filaTotal, 1).Value =
                "TOTALES";

            hoja.Range(filaTotal, 1, filaTotal, 3)
                .Merge();

            hoja.Cell(filaTotal, 4).Value =
                totalIngresos;

            hoja.Cell(filaTotal, 5).Value =
                totalGastos;

            hoja.Cell(filaTotal, 6).Value =
                balanceTotal;

            hoja.Cell(filaTotal, 7).Value =
                $"{lista.Count} registros";

            hoja.Range(filaTotal, 7, filaTotal, 8)
                .Merge();

            var rangoTotales =
                hoja.Range(filaTotal, 1, filaTotal, 8);

            rangoTotales.Style.Font.Bold = true;
            rangoTotales.Style.Fill.BackgroundColor =
                XLColor.FromHtml("#D9EDEF");

            rangoTotales.Style.Font.FontColor =
                XLColor.FromHtml("#123F46");

            rangoTotales.Style.Border.TopBorder =
                XLBorderStyleValues.Medium;

            rangoTotales.Style.Border.TopBorderColor =
                XLColor.FromHtml("#0F5C64");

            hoja.Cell(filaTotal, 1)
                .Style.Alignment.Horizontal =
                XLAlignmentHorizontalValues.Right;

            hoja.Cell(filaTotal, 7)
                .Style.Alignment.Horizontal =
                XLAlignmentHorizontalValues.Center;

            // =====================================================
            // FORMATO MONETARIO
            // =====================================================

            hoja.Range(
                filaEncabezado + 1,
                4,
                filaTotal,
                6)
                .Style.NumberFormat.Format =
                "₡#,##0.00;[Red]-₡#,##0.00";

            hoja.Range(
                filaEncabezado + 1,
                4,
                filaTotal,
                6)
                .Style.Alignment.Horizontal =
                XLAlignmentHorizontalValues.Right;

            // =====================================================
            // ANCHOS DE COLUMNAS
            // =====================================================

            hoja.Column(1).Width = 38;
            hoja.Column(2).Width = 10;
            hoja.Column(3).Width = 12;
            hoja.Column(4).Width = 17;
            hoja.Column(5).Width = 17;
            hoja.Column(6).Width = 17;
            hoja.Column(7).Width = 21;
            hoja.Column(8).Width = 14;

            hoja.Columns(1, 8).Style.Alignment.WrapText =
                false;

            // =====================================================
            // CONFIGURACIÓN DE LA HOJA
            // =====================================================

            hoja.SheetView.FreezeRows(filaEncabezado);

            hoja.PageSetup.PageOrientation =
                XLPageOrientation.Landscape;

            hoja.PageSetup.PaperSize =
                XLPaperSize.A4Paper;

            hoja.PageSetup.FitToPages(1, 0);

            hoja.PageSetup.Margins.Top = 0.4;
            hoja.PageSetup.Margins.Bottom = 0.4;
            hoja.PageSetup.Margins.Left = 0.3;
            hoja.PageSetup.Margins.Right = 0.3;

            hoja.PageSetup.CenterHorizontally = true;

            hoja.PageSetup.PrintAreas.Clear();

            hoja.PageSetup.PrintAreas.Add(
                $"A1:H{filaTotal}");

            hoja.Range(
                1,
                1,
                filaTotal,
                8)
                .Style.Font.FontName =
                "Arial";

            hoja.Range(
                1,
                1,
                filaTotal,
                8)
                .Style.Font.FontSize =
                10;

            using var stream =
                new MemoryStream();

            workbook.SaveAs(stream);

            string nombreArchivo =
                $"Detalle_Presupuesto_" +
                $"{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";
            await RegistrarAuditoria(
    "Exportar Excel",
    0,
    "Se exportó el reporte de detalle de presupuesto a Excel."
);
            return File(
                stream.ToArray(),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                nombreArchivo);
        }
        // GET: PresupuestosDetalles/Detalles
        public async Task<IActionResult> Details(int? id, string? origen)
        {
            if (id == null)
            {
                return NotFound();
            }

            var presupuestoDetalle = await _context.PresupuestoDetalles
                .Include(p => p.CategoriaGasto)
                .Include(p => p.CategoriaIngreso)
                .Include(p => p.PresupuestoMensual)
                .ThenInclude(pm => pm.Proyecto)
                .FirstOrDefaultAsync(m => m.IdPresupuestoDetalle == id);


            if (presupuestoDetalle == null)
            {
                return NotFound();
            }
            ViewBag.Origen = origen;
            return View(presupuestoDetalle);
        }

        // GET: PresupuestosDetalles/Crear
 
        public IActionResult Create()
        {
            CargarCombosPresupuestoDetalle();
            return View();
        }

        // POST: PresupuestosDetalles/Create
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Create([Bind("PresupuestoMensualId,Tipo,CategoriaIngresoId,CategoriaGastoId,MontoPlanificado")] PresupuestoDetalle detalle)
        {
            // Validación para seleccionar presupuesto mensual.
            if (detalle.PresupuestoMensualId <= 0)
            {
                ModelState.AddModelError("PresupuestoMensualId", "Debe seleccionar un presupuesto mensual.");
            }

            // Validación para seleccionar tipo.
            if (string.IsNullOrWhiteSpace(detalle.Tipo))
            {
                ModelState.AddModelError("Tipo", "Debe seleccionar el tipo de detalle.");
            }

            // Validación para monto planificado.
            if (detalle.MontoPlanificado <= 0)
            {
                ModelState.AddModelError("MontoPlanificado", "El monto planificado debe ser mayor a cero.");
            }

            // Validación de categoría según tipo.
            if (detalle.Tipo == "Ingreso" && detalle.CategoriaIngresoId == null)
            {
                ModelState.AddModelError("CategoriaIngresoId", "Debe seleccionar una categoría de ingreso.");
            }

            if (detalle.Tipo == "Gasto" && detalle.CategoriaGastoId == null)
            {
                ModelState.AddModelError("CategoriaGastoId", "Debe seleccionar una categoría de gasto.");
            }

            // Si es ingreso, se limpia categoría gasto.
            if (detalle.Tipo == "Ingreso")
            {
                detalle.CategoriaGastoId = null;
            }

            // Si es gasto, se limpia categoría ingreso.
            if (detalle.Tipo == "Gasto")
            {
                detalle.CategoriaIngresoId = null;
            }

            // Se eliminan del ModelState los campos calculados por el sistema.
            ModelState.Remove("MontoEjecutado");
            ModelState.Remove("Diferencia");
            ModelState.Remove("PresupuestoMensual");
            ModelState.Remove("CategoriaIngreso");
            ModelState.Remove("CategoriaGasto");

            if (!ModelState.IsValid)
            {
                CargarCombosPresupuestoDetalle(detalle);
                return View(detalle);
            }

            // Se calculan monto ejecutado y diferencia.
            await RecalcularDetallePresupuesto(detalle);

            _context.Add(detalle);
            await _context.SaveChangesAsync();

            await ActualizarPresupuestoMensual(detalle.PresupuestoMensualId);

            TempData["MensajeExito"] = "El detalle de presupuesto fue registrado correctamente.";
            return RedirectToAction(nameof(Index));
        }

        // Método para cargar los combos del formulario.
        private void CargarCombosPresupuestoDetalle(PresupuestoDetalle? detalle = null)
        {
            ViewData["PresupuestoMensualId"] = new SelectList(
                _context.PresupuestosMensuales
                    .Include(p => p.Proyecto)
                    .OrderByDescending(p => p.Anio)
                    .ThenByDescending(p => p.Mes)
                    .Select(p => new
                    {
                        p.IdPresupuestoMensual,
                        Nombre = p.Proyecto.Nombre + " - " + p.Mes + "/" + p.Anio
                    }),
                "IdPresupuestoMensual",
                "Nombre",
                detalle?.PresupuestoMensualId
            );

            ViewData["CategoriaIngresoId"] = new SelectList(
                _context.CategoriasIngresos.Where(c => c.Activo == true),
                "IdCategoriaIngreso",
                "Nombre",
                detalle?.CategoriaIngresoId
            );

            ViewData["CategoriaGastoId"] = new SelectList(
                _context.CategoriasGastos.Where(c => c.Activo == true),
                "IdCategoriaGasto",
                "Nombre",
                detalle?.CategoriaGastoId
            );
        }
        // Método para calcular el monto ejecutado y la diferencia del detalle.
        private async Task RecalcularDetallePresupuesto(PresupuestoDetalle detalle)
        {
           
            var presupuesto = await _context.PresupuestosMensuales
                .FirstOrDefaultAsync(p => p.IdPresupuestoMensual == detalle.PresupuestoMensualId);

            if (presupuesto == null)
            {
                detalle.MontoEjecutado = 0;
                detalle.Diferencia = detalle.MontoPlanificado;
                return;
            }

            var fechaInicio = new DateOnly(presupuesto.Anio, presupuesto.Mes, 1);
            var fechaFin = fechaInicio.AddMonths(1).AddDays(-1);

            if (detalle.Tipo == "Ingreso")
            {
                detalle.MontoEjecutado =
                    await _context.Ingresos
                        .Where(i =>
                            i.Activo == true &&
                            i.Estado != "Anulado" &&
                            i.CategoriaIngresoId ==
                                detalle.CategoriaIngresoId &&
                            (
                                i.PresupuestoMensualId ==
                                    presupuesto.IdPresupuestoMensual

                                ||

                                (
                                    i.PresupuestoMensualId == null &&
                                    string.IsNullOrWhiteSpace(
                                        i.MotivoPendientePresupuestario) &&
                                    i.ProyectoId ==
                                        presupuesto.ProyectoId &&
                                    i.Fecha >= fechaInicio &&
                                    i.Fecha <= fechaFin
                                )
                            ))
                        .SumAsync(i =>
                            i.MontoReal ?? 0);

                detalle.Diferencia =
                    detalle.MontoPlanificado -
                    (detalle.MontoEjecutado ?? 0);
            }

            if (detalle.Tipo == "Gasto")
            {
                detalle.MontoEjecutado =
                    await _context.Gastos
                        .Where(g =>
                            g.CategoriaGastoId ==
                                detalle.CategoriaGastoId &&
                            g.Activo == true &&
                            g.Estado == "Aprobado" &&
                            (
                                g.PresupuestoMensualId ==
                                    presupuesto.IdPresupuestoMensual
                                ||
                                (
                                    g.PresupuestoMensualId == null &&
                                    g.ProyectoId ==
                                        presupuesto.ProyectoId &&
                                    g.Fecha >= fechaInicio &&
                                    g.Fecha <= fechaFin
                                )
                            ))
                        .SumAsync(g =>
                            g.MontoTotal ?? 0);

                detalle.Diferencia =
                    detalle.MontoPlanificado -
                    (detalle.MontoEjecutado ?? 0);
            }
        }
        // Método para actualizar automáticamente el presupuesto mensual
        // con base en los detalles registrados.
        private async Task ActualizarPresupuestoMensual(int presupuestoMensualId)
        {
            // Se busca el presupuesto mensual correspondiente.
            var presupuesto = await _context.PresupuestosMensuales
                .FirstOrDefaultAsync(p => p.IdPresupuestoMensual == presupuestoMensualId);

            if (presupuesto == null)
            {
                return;
            }

            // Se suma el total planificado de ingresos desde los detalles.
            presupuesto.MontoIngresadoPlanificado = await _context.PresupuestoDetalles
                .Where(d =>
                    d.PresupuestoMensualId == presupuestoMensualId &&
                    d.Tipo == "Ingreso")
                .SumAsync(d => d.MontoPlanificado);

            // Se suma el total planificado de gastos desde los detalles.
            presupuesto.MontoGastoPlanificado = await _context.PresupuestoDetalles
                .Where(d =>
                    d.PresupuestoMensualId == presupuestoMensualId &&
                    d.Tipo == "Gasto")
                .SumAsync(d => d.MontoPlanificado);

            // Se obtiene la fecha inicial y final del mes presupuestado.
            var fechaInicio = new DateOnly(presupuesto.Anio, presupuesto.Mes, 1);
            var fechaFin = fechaInicio.AddMonths(1).AddDays(-1);

            // Se calculan los ingresos reales correspondientes
            // al período presupuestario.
            // Registros nuevos:
            // se relacionan directamente mediante PresupuestoMensualId.
            // Registros históricos:
            // si todavía no tenían PresupuestoMensualId y tampoco
            // están marcados como pendientes de asignación presupuestaria,
            // conservan la lógica anterior por proyecto + fecha.
            presupuesto.TotalIngresoReal =
                await _context.Ingresos
                    .Where(i =>
                        i.Activo == true &&
                        i.Estado != "Anulado" &&
                        (
                            i.PresupuestoMensualId ==
                                presupuesto.IdPresupuestoMensual

                            ||

                            (
                                i.PresupuestoMensualId == null &&
                                string.IsNullOrWhiteSpace(
                                    i.MotivoPendientePresupuestario) &&
                                i.ProyectoId ==
                                    presupuesto.ProyectoId &&
                                i.Fecha >= fechaInicio &&
                                i.Fecha <= fechaFin
                            )
                        ))
                    .SumAsync(i =>
                        i.MontoReal ?? 0);

            // Solo los gastos aprobados se toman como gasto real.
            presupuesto.TotalGastoReal =
      await _context.Gastos
          .Where(g =>
              g.Activo == true &&
              g.Estado == "Aprobado" &&
              (
                  g.PresupuestoMensualId ==
                      presupuesto.IdPresupuestoMensual
                  ||
                  (
                      g.PresupuestoMensualId == null &&
                      g.ProyectoId ==
                          presupuesto.ProyectoId &&
                      g.Fecha >= fechaInicio &&
                      g.Fecha <= fechaFin
                  )
              ))
          .SumAsync(g =>
              g.MontoTotal ?? 0);

            // Se calcula la diferencia de ingreso.
            presupuesto.DiferenciaIngreso =
                (presupuesto.MontoIngresadoPlanificado ?? 0) -
                (presupuesto.TotalIngresoReal ?? 0);

            // Se calcula el disponible para gasto.
            presupuesto.DiferenciaGasto =
                (presupuesto.MontoGastoPlanificado ?? 0) -
                (presupuesto.TotalGastoReal ?? 0);

            // Se calcula el balance real del período.
            presupuesto.SaldoDisponible =
                (presupuesto.TotalIngresoReal ?? 0) -
                (presupuesto.TotalGastoReal ?? 0);

            // Se asigna automáticamente el estado del presupuesto.
            if ((presupuesto.TotalIngresoReal ?? 0) == 0 &&
                (presupuesto.TotalGastoReal ?? 0) == 0)
            {
                presupuesto.Estado = "Sin movimientos";
            }
            else if ((presupuesto.TotalGastoReal ?? 0) >
                     (presupuesto.MontoGastoPlanificado ?? 0))
            {
                presupuesto.Estado = "Gasto excedido";
            }
            else if ((presupuesto.TotalIngresoReal ?? 0) <
                     (presupuesto.MontoIngresadoPlanificado ?? 0))
            {
                presupuesto.Estado = "Ingresos bajos";
            }
            else
            {
                presupuesto.Estado = "En ejecución";
            }

            // Se actualiza el presupuesto mensual..
            _context.Update(presupuesto);
            await _context.SaveChangesAsync();
        }

        private async Task<(PresupuestosMensuale? Presupuesto,List<MovimientoPresupuestoViewModel> Movimientos)>ObtenerMovimientosPresupuesto(int id)
        {
            var presupuesto = await _context.PresupuestosMensuales
                .Include(p => p.Proyecto)
                .FirstOrDefaultAsync(
                    p => p.IdPresupuestoMensual == id);

            if (presupuesto == null)
            {
                return ( null, new List<MovimientoPresupuestoViewModel>() );
            }

            var fechaInicio =
                new DateOnly(
                    presupuesto.Anio,
                    presupuesto.Mes,
                    1
                );

            var fechaFin =
                fechaInicio
                    .AddMonths(1)
                    .AddDays(-1);

            // =====================================================
            // INGRESOS
            // =====================================================

            var ingresos = await _context.Ingresos
                .Include(i => i.CategoriaIngreso)
              .Where(i =>
    i.Activo == true &&
    i.Estado != "Anulado" &&
    (
        i.PresupuestoMensualId ==
            presupuesto.IdPresupuestoMensual

        ||

        (
            i.PresupuestoMensualId == null &&
            string.IsNullOrWhiteSpace(
                i.MotivoPendientePresupuestario) &&
            i.ProyectoId ==
                presupuesto.ProyectoId &&
            i.Fecha >= fechaInicio &&
            i.Fecha <= fechaFin
        )
    ))
                .Select(i => new MovimientoPresupuestoViewModel
                {
                    Fecha = i.Fecha,
                    Tipo = "Ingreso",

                    Categoria =
                        i.CategoriaIngreso != null
                            ? i.CategoriaIngreso.Nombre
                            : "Sin categoría",

                    Documento =
                        i.Comprobante ??
                        i.Fuente ??
                        "Sin documento",

                    Descripcion =
                        i.Descripcion ??
                        "Sin descripción",

                    Monto =
                        i.MontoReal ?? 0,

                    Origen =
                        i.FacturaId != null
                            ? "Factura"
                            : "Ingreso manual"
                })
                .ToListAsync();

            // =====================================================
            // GASTOS
            // =====================================================

            var gastos = await _context.Gastos
     .Include(g => g.CategoriaGasto)
     .Where(g =>
         g.Activo == true &&
         g.Estado == "Aprobado" &&
         (
             // Registros nuevos: relación presupuestaria directa
             g.PresupuestoMensualId ==
                 presupuesto.IdPresupuestoMensual

             ||

             // Compatibilidad con registros históricos
             (
                 g.PresupuestoMensualId == null &&
                 g.ProyectoId == presupuesto.ProyectoId &&
                 g.Fecha >= fechaInicio &&
                 g.Fecha <= fechaFin
             )
         ))
     .Select(g => new MovimientoPresupuestoViewModel
     {
         Fecha = g.Fecha,
         Tipo = "Gasto",

         Categoria =
             g.CategoriaGasto != null
                 ? g.CategoriaGasto.Nombre
                 : "Sin categoría",

         Documento =
             g.NumeroFactura ??
             "Sin documento",

         Descripcion =
             g.Concepto ??
             "Sin descripción",

         Monto =
             g.MontoTotal ?? 0,

         Origen = "Gasto"
     })
     .ToListAsync();

            var movimientos = ingresos
                .Concat(gastos)
                .OrderBy(m => m.Fecha)
                .ThenBy(m => m.Tipo)
                .ToList();

            return (
                presupuesto,
                movimientos
            );
        }
        // GET: PresupuestosDetalles/Movimientos/5
        public async Task<IActionResult> Movimientos(
            int id,
            string? origen)
        {
            var resultado =
                await ObtenerMovimientosPresupuesto(id);

            if (resultado.Presupuesto == null)
            {
                return NotFound();
            }

            var presupuesto =
                resultado.Presupuesto;

            var movimientos =
                resultado.Movimientos;

            decimal totalIngresos =
                movimientos
                    .Where(m => m.Tipo == "Ingreso")
                    .Sum(m => m.Monto);

            decimal totalGastos =
                movimientos
                    .Where(m => m.Tipo == "Gasto")
                    .Sum(m => m.Monto);

            ViewBag.Presupuesto =
                presupuesto;

            ViewBag.Proyecto =
                presupuesto.Proyecto?.Nombre ??
                "Sin proyecto";

            ViewBag.Anio =
                presupuesto.Anio;

            ViewBag.Mes =
                presupuesto.Mes;

            ViewBag.TotalIngresos =
                totalIngresos;

            ViewBag.TotalGastos =
                totalGastos;

            ViewBag.Balance =
                totalIngresos - totalGastos;

            ViewBag.Origen =
                origen;

            ViewBag.IdPresupuesto =
                id;
            await RegistrarAuditoria(
    "Vista previa PDF",
    id,
    $"Se generó la vista previa del reporte de movimientos del presupuesto ID {id}."
);

           
            return View(movimientos);
        }
        // GET: PresupuestosDetalles/ReporteMovimientos/5
        public async Task<IActionResult> ReporteMovimientos(
            int id,
            string? origen)
        {
            var resultado =
                await ObtenerMovimientosPresupuesto(id);

            if (resultado.Presupuesto == null)
            {
                return NotFound();
            }

            var presupuesto =
                resultado.Presupuesto;

            var movimientos =
                resultado.Movimientos;

            decimal totalIngresos =
                movimientos
                    .Where(m => m.Tipo == "Ingreso")
                    .Sum(m => m.Monto);

            decimal totalGastos =
                movimientos
                    .Where(m => m.Tipo == "Gasto")
                    .Sum(m => m.Monto);

            ViewBag.Proyecto =
                presupuesto.Proyecto?.Nombre ??
                "Sin proyecto";

            ViewBag.Anio =
                presupuesto.Anio;

            ViewBag.Mes =
                presupuesto.Mes;

            ViewBag.TotalIngresos =
                totalIngresos;

            ViewBag.TotalGastos =
                totalGastos;

            ViewBag.Balance =
                totalIngresos - totalGastos;

            ViewBag.TotalMovimientos =
                movimientos.Count;

            ViewBag.TotalMovimientosIngreso =
                movimientos.Count(
                    m => m.Tipo == "Ingreso"
                );

            ViewBag.TotalMovimientosGasto =
                movimientos.Count(
                    m => m.Tipo == "Gasto"
                );

            ViewBag.FechaGeneracion =
                DateTime.Now;

            ViewBag.IdPresupuesto =
                id;

            ViewBag.Origen =
                origen;

            return View(movimientos);
        }
        // GET: PresupuestosDetalles/ExportarMovimientosExcel/5
        public async Task<IActionResult> ExportarMovimientosExcel(
            int id,
            string? origen)
        {
            var resultado =
                await ObtenerMovimientosPresupuesto(id);

            if (resultado.Presupuesto == null)
            {
                return NotFound();
            }

            var presupuesto =
                resultado.Presupuesto;

            var movimientos =
                resultado.Movimientos;

            decimal totalIngresos =
                movimientos
                    .Where(m => m.Tipo == "Ingreso")
                    .Sum(m => m.Monto);

            decimal totalGastos =
                movimientos
                    .Where(m => m.Tipo == "Gasto")
                    .Sum(m => m.Monto);

            decimal balance =
                totalIngresos - totalGastos;

            string nombreProyecto =
                presupuesto.Proyecto?.Nombre ??
                "Sin proyecto";

            string nombreMes =
                presupuesto.Mes switch
                {
                    1 => "Enero",
                    2 => "Febrero",
                    3 => "Marzo",
                    4 => "Abril",
                    5 => "Mayo",
                    6 => "Junio",
                    7 => "Julio",
                    8 => "Agosto",
                    9 => "Septiembre",
                    10 => "Octubre",
                    11 => "Noviembre",
                    12 => "Diciembre",
                    _ => presupuesto.Mes.ToString()
                };

            using var workbook =
                new XLWorkbook();

            var hoja =
                workbook.Worksheets.Add(
                    "Movimientos Presupuesto"
                );

            // =====================================================
            // ENCABEZADO PRINCIPAL
            // =====================================================

            hoja.Cell("A1").Value =
                "Unión Cantonal de Asociaciones de Desarrollo de Santa Ana";

            hoja.Range("A1:G1").Merge();

            hoja.Cell("A2").Value =
                "Reporte de Movimientos del Presupuesto";

            hoja.Range("A2:G2").Merge();

            hoja.Cell("A3").Value =
                $"Fecha de generación: {DateTime.Now:dd/MM/yyyy HH:mm}";

            hoja.Range("A3:G3").Merge();

            hoja.Range("A1:G3")
                .Style.Alignment.Horizontal =
                XLAlignmentHorizontalValues.Center;

            hoja.Range("A1:G3")
                .Style.Alignment.Vertical =
                XLAlignmentVerticalValues.Center;

            hoja.Range("A1:G1")
                .Style.Font.Bold = true;

            hoja.Range("A1:G1")
                .Style.Font.FontSize = 16;

            hoja.Range("A1:G1")
                .Style.Font.FontColor =
                XLColor.FromHtml("#0F5C64");

            hoja.Range("A2:G2")
                .Style.Font.Bold = true;

            hoja.Range("A2:G2")
                .Style.Font.FontSize = 13;

            hoja.Range("A2:G2")
                .Style.Font.FontColor =
                XLColor.FromHtml("#15454D");

            hoja.Range("A3:G3")
                .Style.Font.FontColor =
                XLColor.FromHtml("#6C757D");

            hoja.Row(1).Height = 24;
            hoja.Row(2).Height = 20;
            hoja.Row(3).Height = 18;

            // =====================================================
            // INFORMACIÓN DEL PRESUPUESTO
            // =====================================================

            hoja.Cell("A5").Value =
                "Información del presupuesto";

            hoja.Range("A5:G5").Merge();

            hoja.Range("A5:G5")
                .Style.Fill.BackgroundColor =
                XLColor.FromHtml("#EAF3F4");

            hoja.Range("A5:G5")
                .Style.Font.Bold = true;

            hoja.Range("A5:G5")
                .Style.Font.FontColor =
                XLColor.FromHtml("#0F5C64");

            hoja.Cell("A6").Value =
                "Proyecto";

            hoja.Cell("B6").Value =
                nombreProyecto;

            hoja.Range("B6:C6").Merge();

            hoja.Cell("D6").Value =
                "Año";

            hoja.Cell("E6").Value =
                presupuesto.Anio;

            hoja.Cell("F6").Value =
                "Mes";

            hoja.Cell("G6").Value =
                nombreMes;

            foreach (int columna in new[] { 1, 4, 6 })
            {
                hoja.Cell(6, columna)
                    .Style.Font.Bold = true;

                hoja.Cell(6, columna)
                    .Style.Font.FontColor =
                    XLColor.FromHtml("#0F5C64");
            }

            hoja.Range("A6:G6")
                .Style.Fill.BackgroundColor =
                XLColor.FromHtml("#F7FAFA");

            hoja.Range("A6:G6")
                .Style.Border.OutsideBorder =
                XLBorderStyleValues.Thin;

            hoja.Range("A6:G6")
                .Style.Border.InsideBorder =
                XLBorderStyleValues.Thin;

            hoja.Range("A6:G6")
                .Style.Border.OutsideBorderColor =
                XLColor.FromHtml("#D5E2E4");

            hoja.Range("A6:G6")
                .Style.Border.InsideBorderColor =
                XLColor.FromHtml("#D5E2E4");

            // =====================================================
            // RESUMEN FINANCIERO
            // =====================================================

            hoja.Cell("A8").Value =
                "Total ingresos";

            hoja.Cell("B8").Value =
                totalIngresos;

            hoja.Cell("C8").Value =
                "Total gastos";

            hoja.Cell("D8").Value =
                totalGastos;

            hoja.Cell("E8").Value =
                "Balance";

            hoja.Cell("F8").Value =
                balance;

            hoja.Cell("G8").Value =
                $"{movimientos.Count} movimientos";

            hoja.Range("A8:B8")
                .Style.Fill.BackgroundColor =
                XLColor.FromHtml("#D8F3DC");

            hoja.Range("C8:D8")
                .Style.Fill.BackgroundColor =
                XLColor.FromHtml("#F8D7DA");

            hoja.Range("E8:F8")
                .Style.Fill.BackgroundColor =
                balance >= 0
                    ? XLColor.FromHtml("#D9EDEF")
                    : XLColor.FromHtml("#F8D7DA");

            hoja.Cell("G8")
                .Style.Fill.BackgroundColor =
                XLColor.FromHtml("#EEF1F4");

            hoja.Range("A8:G8")
                .Style.Font.Bold = true;

            hoja.Range("A8:G8")
                .Style.Border.OutsideBorder =
                XLBorderStyleValues.Thin;

            hoja.Range("A8:G8")
                .Style.Border.InsideBorder =
                XLBorderStyleValues.Thin;

            hoja.Range("B8:F8")
                .Style.NumberFormat.Format =
                "₡#,##0.00;[Red]-₡#,##0.00";

            // =====================================================
            // ENCABEZADOS DE LA TABLA
            // =====================================================

            int filaEncabezado = 10;

            string[] encabezados =
            {
        "Fecha",
        "Tipo",
        "Categoría",
        "Documento",
        "Descripción",
        "Monto",
        "Origen"
    };

            for (int columna = 0;
                 columna < encabezados.Length;
                 columna++)
            {
                hoja.Cell(
                    filaEncabezado,
                    columna + 1
                ).Value = encabezados[columna];
            }

            var encabezado =
                hoja.Range(
                    filaEncabezado,
                    1,
                    filaEncabezado,
                    7
                );

            encabezado.Style.Fill.BackgroundColor =
                XLColor.FromHtml("#0F5C64");

            encabezado.Style.Font.FontColor =
                XLColor.White;

            encabezado.Style.Font.Bold =
                true;

            encabezado.Style.Alignment.Horizontal =
                XLAlignmentHorizontalValues.Center;

            encabezado.Style.Alignment.Vertical =
                XLAlignmentVerticalValues.Center;

            encabezado.Style.Border.OutsideBorder =
                XLBorderStyleValues.Thin;

            encabezado.Style.Border.InsideBorder =
                XLBorderStyleValues.Thin;

            encabezado.Style.Border.OutsideBorderColor =
                XLColor.White;

            encabezado.Style.Border.InsideBorderColor =
                XLColor.White;

            hoja.Row(filaEncabezado).Height =
                22;

            // =====================================================
            // MOVIMIENTOS
            // =====================================================

            int fila =
                filaEncabezado + 1;

            foreach (var movimiento in movimientos)
            {
                hoja.Cell(fila, 1).Value =
                    movimiento.Fecha.ToDateTime(
                        TimeOnly.MinValue
                    );

                hoja.Cell(fila, 2).Value =
                    movimiento.Tipo;

                hoja.Cell(fila, 3).Value =
                    movimiento.Categoria;

                hoja.Cell(fila, 4).Value =
                    movimiento.Documento;

                hoja.Cell(fila, 5).Value =
                    movimiento.Descripcion;

                hoja.Cell(fila, 6).Value =
                    movimiento.Monto;

                hoja.Cell(fila, 7).Value =
                    movimiento.Origen;

                hoja.Cell(fila, 1)
                    .Style.DateFormat.Format =
                    "dd/MM/yyyy";

                hoja.Cell(fila, 6)
                    .Style.NumberFormat.Format =
                    "₡#,##0.00;[Red]-₡#,##0.00";

                var rangoFila =
                    hoja.Range(fila, 1, fila, 7);

                rangoFila.Style.Border.OutsideBorder =
                    XLBorderStyleValues.Thin;

                rangoFila.Style.Border.InsideBorder =
                    XLBorderStyleValues.Thin;

                rangoFila.Style.Border.OutsideBorderColor =
                    XLColor.FromHtml("#D9E1E5");

                rangoFila.Style.Border.InsideBorderColor =
                    XLColor.FromHtml("#D9E1E5");

                rangoFila.Style.Alignment.Vertical =
                    XLAlignmentVerticalValues.Center;

                if (movimiento.Tipo == "Ingreso")
                {
                    rangoFila.Style.Fill.BackgroundColor =
                        XLColor.FromHtml("#EDF8EF");

                    hoja.Cell(fila, 2)
                        .Style.Fill.BackgroundColor =
                        XLColor.FromHtml("#8FD3A4");

                    hoja.Cell(fila, 2)
                        .Style.Font.FontColor =
                        XLColor.FromHtml("#14532D");

                    hoja.Cell(fila, 6)
                        .Style.Font.FontColor =
                        XLColor.FromHtml("#198754");
                }
                else
                {
                    rangoFila.Style.Fill.BackgroundColor =
                        XLColor.FromHtml("#FBE5E8");

                    hoja.Cell(fila, 2)
                        .Style.Fill.BackgroundColor =
                        XLColor.FromHtml("#EE9CA7");

                    hoja.Cell(fila, 2)
                        .Style.Font.FontColor =
                        XLColor.FromHtml("#7A1C2A");

                    hoja.Cell(fila, 6)
                        .Style.Font.FontColor =
                        XLColor.FromHtml("#DC3545");
                }

                hoja.Cell(fila, 2)
                    .Style.Font.Bold = true;

                hoja.Cell(fila, 2)
                    .Style.Alignment.Horizontal =
                    XLAlignmentHorizontalValues.Center;

                hoja.Cell(fila, 6)
                    .Style.Font.Bold = true;

                hoja.Cell(fila, 6)
                    .Style.Alignment.Horizontal =
                    XLAlignmentHorizontalValues.Right;

                hoja.Cell(fila, 7)
                    .Style.Alignment.Horizontal =
                    XLAlignmentHorizontalValues.Center;

                hoja.Cell(fila, 5)
                    .Style.Alignment.WrapText = true;

                hoja.Row(fila).AdjustToContents();

                fila++;
            }

            // =====================================================
            // TOTALES
            // =====================================================

            int filaTotal = fila;

            hoja.Cell(filaTotal, 1).Value =
                "TOTALES";

            hoja.Range(
                filaTotal,
                1,
                filaTotal,
                5
            ).Merge();

            hoja.Cell(filaTotal, 6).Value =
                totalIngresos - totalGastos;

            hoja.Cell(filaTotal, 7).Value =
                $"{movimientos.Count} movimientos";

            var rangoTotales =
                hoja.Range(
                    filaTotal,
                    1,
                    filaTotal,
                    7
                );

            rangoTotales.Style.Fill.BackgroundColor =
                XLColor.FromHtml("#D9EDEF");

            rangoTotales.Style.Font.Bold =
                true;

            rangoTotales.Style.Font.FontColor =
                XLColor.FromHtml("#123F46");

            rangoTotales.Style.Border.TopBorder =
                XLBorderStyleValues.Medium;

            rangoTotales.Style.Border.TopBorderColor =
                XLColor.FromHtml("#0F5C64");

            hoja.Cell(filaTotal, 1)
                .Style.Alignment.Horizontal =
                XLAlignmentHorizontalValues.Right;

            hoja.Cell(filaTotal, 6)
                .Style.NumberFormat.Format =
                "₡#,##0.00;[Red]-₡#,##0.00";

            hoja.Cell(filaTotal, 7)
                .Style.Alignment.Horizontal =
                XLAlignmentHorizontalValues.Center;

            // =====================================================
            // ANCHOS
            // =====================================================

            hoja.Column(1).Width = 13;
            hoja.Column(2).Width = 13;
            hoja.Column(3).Width = 24;
            hoja.Column(4).Width = 20;
            hoja.Column(5).Width = 58;
            hoja.Column(6).Width = 18;
            hoja.Column(7).Width = 20;

            // =====================================================
            // CONFIGURACIÓN GENERAL
            // =====================================================

            hoja.Range(
                1,
                1,
                filaTotal,
                7
            ).Style.Font.FontName =
                "Arial";

            hoja.Range(
                1,
                1,
                filaTotal,
                7
            ).Style.Font.FontSize =
                10;

            hoja.SheetView.FreezeRows(
                filaEncabezado
            );

            hoja.PageSetup.PageOrientation =
                XLPageOrientation.Landscape;

            hoja.PageSetup.PaperSize =
                XLPaperSize.A4Paper;

            hoja.PageSetup.FitToPages(
                1,
                0
            );

            hoja.PageSetup.Margins.Top =
                0.4;

            hoja.PageSetup.Margins.Bottom =
                0.4;

            hoja.PageSetup.Margins.Left =
                0.3;

            hoja.PageSetup.Margins.Right =
                0.3;

            hoja.PageSetup.CenterHorizontally =
                true;

            hoja.PageSetup.PrintAreas.Clear();

            hoja.PageSetup.PrintAreas.Add(
                $"A1:G{filaTotal}"
            );

            using var stream =
                new MemoryStream();

            workbook.SaveAs(stream);

            string nombreSeguro =
                string.Concat(
                    nombreProyecto
                        .Where(c =>
                            !Path
                                .GetInvalidFileNameChars()
                                .Contains(c))
                )
                .Replace(" ", "_");

            string nombreArchivo =
                $"Movimientos_Presupuesto_" +
                $"{nombreSeguro}_" +
                $"{presupuesto.Mes:D2}_" +
                $"{presupuesto.Anio}_" +
                $"{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";

            await RegistrarAuditoria(
    "Exportar Excel",
    id,
    $"Se exportó el reporte de movimientos del presupuesto ID {id} a Excel."
);
            return File(
                stream.ToArray(),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                nombreArchivo
            );
        }
        // GET: PresupuestosDetalles/Editar/5
        public async Task<IActionResult> Edit(int? id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var presupuestoDetalle = await _context.PresupuestoDetalles.FindAsync(id);
            if (presupuestoDetalle == null)
            {
                return NotFound();
            }
        

            CargarCombosPresupuestoDetalle(presupuestoDetalle);
            
            return View(presupuestoDetalle);
        }

        // POST: PresupuestosDetalles/Edit/5
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Edit(int id, [Bind("IdPresupuestoDetalle,MontoPlanificado")] PresupuestoDetalle detalle)
        {
            // Valida que el registro sea el correcto.
            if (id != detalle.IdPresupuestoDetalle)
            {
                return NotFound();
            }
            // Se eliminan las validaciones de los campos que ya no se modifican.
            ModelState.Remove("Tipo");
            ModelState.Remove("PresupuestoMensualId");
            ModelState.Remove("PresupuestoMensual");
            ModelState.Remove("CategoriaIngresoId");
            ModelState.Remove("CategoriaIngreso");
            ModelState.Remove("CategoriaGastoId");
            ModelState.Remove("CategoriaGasto");

            // Valida el monto planificado.
            if (detalle.MontoPlanificado <= 0)
            {
                ModelState.AddModelError("MontoPlanificado", "El monto planificado debe ser mayor a cero.");
            }

            // Se busca el detalle original.
            var detalleActual = await _context.PresupuestoDetalles
                .FirstOrDefaultAsync(d => d.IdPresupuestoDetalle == id);

            if (detalleActual == null)
            {
                return NotFound();
            }
            // Se recalcula primero para saber si ya tiene movimientos reales.
            await RecalcularDetallePresupuesto(detalleActual);

            // Si ya existe monto ejecutado, no se permite modificar el monto planificado.
            if ((detalleActual.MontoEjecutado ?? 0) > 0)
            {
                TempData["MensajeError"] = "No se puede modificar este detalle porque ya tiene movimientos reales registrados.";
                return RedirectToAction(nameof(Index));
            }

            if (!ModelState.IsValid)
            {
                CargarCombosPresupuestoDetalle(detalleActual);
                return View(detalleActual);
            }

            // Solo se permite modificar el monto planificado.
            detalleActual.MontoPlanificado = detalle.MontoPlanificado;

            // Se recalcula el ejecutado y la diferencia.
            await RecalcularDetallePresupuesto(detalleActual);

            _context.Update(detalleActual);
            await _context.SaveChangesAsync();

            // Se actualiza el presupuesto mensual relacionado.
            await ActualizarPresupuestoMensual(detalleActual.PresupuestoMensualId);

            TempData["MensajeExito"] = "El detalle de presupuesto fue modificado correctamente.";
            return RedirectToAction(nameof(Index));
        }

        // GET: PresupuestosDetalles/Delete/5
        public async Task<IActionResult> Delete(int? id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var presupuestoDetalle = await _context.PresupuestoDetalles
                .Include(p => p.CategoriaGasto)
                .Include(p => p.CategoriaIngreso)
                .Include(p => p.PresupuestoMensual)
                .FirstOrDefaultAsync(m => m.IdPresupuestoDetalle == id);
            if (presupuestoDetalle == null)
            {
                return NotFound();
            }

            return View(presupuestoDetalle);
        }

        // POST: PresupuestosDetalles/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> DeleteConfirmed(int id)
        {
            var presupuestoDetalle = await _context.PresupuestoDetalles.FindAsync(id);
            if (presupuestoDetalle != null)
            {
                _context.PresupuestoDetalles.Remove(presupuestoDetalle);
            }

            await _context.SaveChangesAsync();
            return RedirectToAction(nameof(Index));
        }
        private async Task RegistrarAuditoria(
    string accion,
    int registroId,
    string descripcion)
        {
            var usuarioSesion =
                HttpContext.Session.GetString("Usuario");

            var usuario = await _context.Usuarios
                .FirstOrDefaultAsync(u => u.Nombre == usuarioSesion);

            if (usuario != null)
            {
                var auditoria = new Auditorium
                {
                    UsuarioId = usuario.IdUsuario,
                    Tabla = "PresupuestoDetalles",
                    RegistroId = registroId,
                    Accion = accion,
                    Descripcion = descripcion,
                    Fecha = DateTime.Now
                };

                _context.Auditoria.Add(auditoria);
                await _context.SaveChangesAsync();
            }
        }
        private bool PresupuestoDetalleExists(int id)
        {
            return _context.PresupuestoDetalles.Any(e => e.IdPresupuestoDetalle == id);
        }
       
    }
}
