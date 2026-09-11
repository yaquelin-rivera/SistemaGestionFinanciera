using System;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.AspNetCore.Http;
using Microsoft.EntityFrameworkCore;
using SistemaGestionFinanciera.Data;
using SistemaGestionFinanciera.Models;
using ClosedXML.Excel;
using System.IO;
using SistemaGestionFinanciera.Models.ViewModels;

namespace SistemaGestionFinanciera.Controllers
{
    public class CategoriasGastosController : BaseController
    {
        private readonly SistemaFinancieroContext _context;

        public CategoriasGastosController(SistemaFinancieroContext context)
        {
            _context = context;
        }
        // INDEX - LISTADO Y CONTROL MENSUAL DE CATEGORÍAS DE GASTOS
        public async Task<IActionResult> Index(
            string? buscar,
            string? cuenta,
            string? estado,
            string? limite,
            int? anio,
            int? mes,
            string? estadoLimite)
        {
            // Filtros normales de las categorías.
            var categorias = await ConstruirConsultaReporte(
                    buscar,
                    cuenta,
                    estado,
                    limite)
                .OrderByDescending(c => c.IdCategoriaGasto)
                .ToListAsync();

            // Calcula el consumo del período seleccionado.
            var categoriasControl =
                await CalcularControlMensualCategorias(
                    categorias,
                    anio,
                    mes);

            // Alertas antes de aplicar el filtro rápido.
            var categoriasProximas = categoriasControl
                .Where(x =>
                    x.EstadoLimite == "Próxima al límite")
                .ToList();

            var categoriasExcedidas = categoriasControl
                .Where(x =>
                    x.EstadoLimite == "Excedida")
                .ToList();

            ViewBag.CantidadProximas =
                categoriasProximas.Count;

            ViewBag.MontoConsumidoProximas =
                categoriasProximas.Sum(x =>
                    x.GastoPeriodo);

            ViewBag.CantidadExcedidas =
                categoriasExcedidas.Count;

            ViewBag.ExcesoAcumulado =
                categoriasExcedidas.Sum(x =>
                    x.Exceso);

            // Conserva filtros.
            ViewBag.Buscar = buscar;
            ViewBag.Cuenta = cuenta;
            ViewBag.Estado = estado;
            ViewBag.Limite = limite;
            ViewBag.Anio = anio;
            ViewBag.Mes = mes;
            ViewBag.EstadoLimite = estadoLimite;

            var hoy = DateOnly.FromDateTime(DateTime.Today);

            int anioAplicado =
                anio ?? hoy.Year;

            int mesAplicado =
                mes ?? hoy.Month;

            ViewBag.AnioAplicado = anioAplicado;
            ViewBag.MesAplicado = mesAplicado;

            ViewBag.PeriodoTexto =
                new DateTime(anioAplicado, mesAplicado, 1)
                    .ToString(
                        "MMMM yyyy",
                        new System.Globalization.CultureInfo("es-CR"));

            // Filtro rápido activado desde las alertas.
            categoriasControl =
                AplicarFiltroEstadoLimite(
                    categoriasControl,
                    estadoLimite);

            return View(
                categoriasControl
                    .OrderBy(x => x.Categoria.Nombre)
                    .ToList());
        }
        // ================================================================
        // REPORTE DE CATEGORÍAS DE GASTOS
        // ================================================================
        [HttpGet]
        public async Task<IActionResult> Reporte(
            string? buscar,
            string? cuenta,
            string? estado,
            string? limite,
            int? anio,
            int? mes,
            string? estadoLimite)
        {
            var categorias = await ConstruirConsultaReporte(
                    buscar,
                    cuenta,
                    estado,
                    limite)
                .OrderBy(c => c.Nombre)
                .ToListAsync();

            var categoriasControl =
                await CalcularControlMensualCategorias(
                    categorias,
                    anio,
                    mes);

            // Las alertas se calculan antes del filtro especial.
            var categoriasProximas = categoriasControl
                .Where(x =>
                    x.EstadoLimite == "Próxima al límite")
                .ToList();

            var categoriasExcedidas = categoriasControl
                .Where(x =>
                    x.EstadoLimite == "Excedida")
                .ToList();

            ViewBag.CantidadProximas =
                categoriasProximas.Count;

            ViewBag.MontoConsumidoProximas =
                categoriasProximas.Sum(x =>
                    x.GastoPeriodo);

            ViewBag.CantidadExcedidas =
                categoriasExcedidas.Count;

            ViewBag.ExcesoAcumulado =
                categoriasExcedidas.Sum(x =>
                    x.Exceso);

            // Aplica el filtro activado desde la alerta.
            categoriasControl =
                AplicarFiltroEstadoLimite(
                    categoriasControl,
                    estadoLimite);

            // Resumen de los registros que aparecerán en el PDF.
            ViewBag.TotalRegistros =
                categoriasControl.Count;

            ViewBag.TotalActivas =
                categoriasControl.Count(x =>
                    x.Categoria.Activo);

            ViewBag.TotalInactivas =
                categoriasControl.Count(x =>
                    !x.Categoria.Activo);

            ViewBag.TotalLimiteMensual =
                categoriasControl
                    .Where(x => x.Categoria.Activo)
                    .Sum(x => x.Categoria.LimiteMensual);

            ViewBag.TotalGastoPeriodo =
                categoriasControl.Sum(x =>
                    x.GastoPeriodo);

            var hoy =
                DateOnly.FromDateTime(DateTime.Today);

            int anioAplicado =
                anio ?? hoy.Year;

            int mesAplicado =
                mes ?? hoy.Month;

            ViewBag.Anio = anio;
            ViewBag.Mes = mes;
            ViewBag.AnioAplicado = anioAplicado;
            ViewBag.MesAplicado = mesAplicado;

            ViewBag.PeriodoTexto =
                new DateTime(
                    anioAplicado,
                    mesAplicado,
                    1)
                .ToString(
                    "MMMM yyyy",
                    new System.Globalization.CultureInfo("es-CR"));

            ViewBag.Buscar = buscar;
            ViewBag.Cuenta = cuenta;
            ViewBag.Estado = estado;
            ViewBag.Limite = limite;
            ViewBag.EstadoLimite = estadoLimite;
            ViewBag.FechaGeneracion = DateTime.Now;

            await RegistrarAuditoria(
                "Vista previa PDF",
                "CategoriasGastos",
                0,
                $"Se generó la vista previa del reporte de categorías de gastos. " +
                $"Período: {ViewBag.PeriodoTexto}.");

            return View(categoriasControl);
        }
        // ================================================================
        // EXPORTAR REPORTE DE CATEGORÍAS DE GASTOS A EXCEL
        // ================================================================
        [HttpGet]
        public async Task<IActionResult> ExportarExcel(
            string? buscar,
            string? cuenta,
            string? estado,
            string? limite,
            int? anio,
            int? mes,
            string? estadoLimite)
        {
            var categorias = await ConstruirConsultaReporte(
                    buscar,
                    cuenta,
                    estado,
                    limite)
                .OrderBy(c => c.Nombre)
                .ToListAsync();

            var categoriasControl =
                await CalcularControlMensualCategorias(
                    categorias,
                    anio,
                    mes);

            var categoriasProximas = categoriasControl
                .Where(x =>
                    x.EstadoLimite == "Próxima al límite")
                .ToList();

            var categoriasExcedidas = categoriasControl
                .Where(x =>
                    x.EstadoLimite == "Excedida")
                .ToList();

            int cantidadProximas =
                categoriasProximas.Count;

            decimal montoConsumidoProximas =
                categoriasProximas.Sum(x =>
                    x.GastoPeriodo);

            int cantidadExcedidas =
                categoriasExcedidas.Count;

            decimal excesoAcumulado =
                categoriasExcedidas.Sum(x =>
                    x.Exceso);

            categoriasControl =
                AplicarFiltroEstadoLimite(
                    categoriasControl,
                    estadoLimite);

            var hoy =
                DateOnly.FromDateTime(DateTime.Today);

            int anioAplicado =
                anio ?? hoy.Year;

            int mesAplicado =
                mes ?? hoy.Month;

            string periodoTexto =
                new DateTime(
                    anioAplicado,
                    mesAplicado,
                    1)
                .ToString(
                    "MMMM yyyy",
                    new System.Globalization.CultureInfo("es-CR"));

            using var workbook =
                new XLWorkbook();

            var hoja =
                workbook.Worksheets.Add(
                    "Categorías de gastos");

            const int totalColumnas = 8;

            var colorInstitucional =
                XLColor.FromHtml("#0F6B73");

            var colorInstitucionalOscuro =
                XLColor.FromHtml("#0B4F55");

            var colorFondoSuave =
                XLColor.FromHtml("#E7F1F2");

            var colorFondoFiltros =
                XLColor.FromHtml("#F3F7F8");

            var colorVerdeSuave =
                XLColor.FromHtml("#D1E7DD");

            var colorAmarilloSuave =
                XLColor.FromHtml("#FFF3CD");

            var colorRojoSuave =
                XLColor.FromHtml("#F8D7DA");

            var colorGrisSuave =
                XLColor.FromHtml("#E2E3E5");

            // ============================================================
            // ENCABEZADO
            // ============================================================

            hoja.Cell("A1").Value =
                "Unión Cantonal de Asociaciones de Desarrollo de Santa Ana";

            hoja.Range(1, 1, 1, totalColumnas).Merge();

            hoja.Cell("A2").Value =
                "Reporte de Control Mensual de Categorías de Gastos";

            hoja.Range(2, 1, 2, totalColumnas).Merge();

            hoja.Cell("A3").Value =
                $"Fecha de generación: {DateTime.Now:dd/MM/yyyy hh:mm tt}";

            hoja.Range(3, 1, 3, totalColumnas).Merge();

            hoja.Range(1, 1, 1, totalColumnas)
                .Style.Font.Bold = true;

            hoja.Range(1, 1, 1, totalColumnas)
                .Style.Font.FontSize = 16;

            hoja.Range(1, 1, 3, totalColumnas)
                .Style.Alignment.Horizontal =
                XLAlignmentHorizontalValues.Center;

            hoja.Range(1, 1, 3, totalColumnas)
                .Style.Alignment.Vertical =
                XLAlignmentVerticalValues.Center;

            // ============================================================
            // FILTROS APLICADOS
            // ============================================================

            hoja.Cell("A5").Value =
                "Filtros aplicados";

            hoja.Range(5, 1, 5, totalColumnas).Merge();

            hoja.Range(5, 1, 5, totalColumnas)
                .Style.Fill.BackgroundColor =
                colorInstitucional;

            hoja.Range(5, 1, 5, totalColumnas)
                .Style.Font.FontColor =
                XLColor.White;

            hoja.Range(5, 1, 5, totalColumnas)
                .Style.Font.Bold = true;

            hoja.Cell("A6").Value = "Buscar:";
            hoja.Cell("B6").Value =
                string.IsNullOrWhiteSpace(buscar)
                    ? "Todos"
                    : buscar;

            hoja.Cell("C6").Value = "Cuenta:";
            hoja.Cell("D6").Value =
                string.IsNullOrWhiteSpace(cuenta)
                    ? "Todas"
                    : cuenta;

            hoja.Cell("E6").Value = "Estado:";
            hoja.Cell("F6").Value =
                string.IsNullOrWhiteSpace(estado)
                    ? "Todos"
                    : estado;

            hoja.Cell("G6").Value = "Límite:";
            hoja.Cell("H6").Value =
                limite switch
                {
                    "ConLimite" => "Con límite",
                    "SinLimite" => "Sin límite",
                    _ => "Todos"
                };

            hoja.Cell("A7").Value = "Período:";
            hoja.Cell("B7").Value = periodoTexto;

            hoja.Cell("C7").Value = "Año:";
            hoja.Cell("D7").Value = anioAplicado;

            hoja.Cell("E7").Value = "Mes:";
            hoja.Cell("F7").Value =
                new DateTime(
                    anioAplicado,
                    mesAplicado,
                    1)
                .ToString(
                    "MMMM",
                    new System.Globalization.CultureInfo("es-CR"));

            hoja.Cell("G7").Value =
                "Control mensual:";

            hoja.Cell("H7").Value =
                estadoLimite switch
                {
                    "Proxima" =>
                        "Próximas al límite",

                    "Excedida" =>
                        "Límite excedido",

                    _ => "Todos"
                };

            hoja.Range(6, 1, 7, totalColumnas)
                .Style.Fill.BackgroundColor =
                colorFondoFiltros;

            hoja.Range(6, 1, 7, totalColumnas)
                .Style.Border.OutsideBorder =
                XLBorderStyleValues.Thin;

            hoja.Range(6, 1, 7, totalColumnas)
                .Style.Border.InsideBorder =
                XLBorderStyleValues.Thin;

            foreach (string celda in new[]
            {
        "A6", "C6", "E6", "G6",
        "A7", "C7", "E7", "G7"
    })
            {
                hoja.Cell(celda).Style.Font.Bold = true;
            }

            // ============================================================
            // ALERTAS DEL PERÍODO
            // ============================================================

            hoja.Cell("A9").Value =
                "Alertas del período";

            hoja.Range(9, 1, 9, totalColumnas).Merge();

            hoja.Range(9, 1, 9, totalColumnas)
                .Style.Fill.BackgroundColor =
                colorInstitucional;

            hoja.Range(9, 1, 9, totalColumnas)
                .Style.Font.FontColor =
                XLColor.White;

            hoja.Range(9, 1, 9, totalColumnas)
                .Style.Font.Bold = true;

            hoja.Cell("A10").Value =
                "Categorías próximas al límite";

            hoja.Cell("B10").Value =
                cantidadProximas;

            hoja.Cell("C10").Value =
                "Monto consumido";

            hoja.Cell("D10").Value =
                montoConsumidoProximas;

            hoja.Cell("D10")
                .Style.NumberFormat.Format =
                "₡#,##0.00";

            hoja.Range("A10:D10")
                .Style.Fill.BackgroundColor =
                colorAmarilloSuave;

            hoja.Cell("E10").Value =
                "Categorías excedidas";

            hoja.Cell("F10").Value =
                cantidadExcedidas;

            hoja.Cell("G10").Value =
                "Exceso acumulado";

            hoja.Cell("H10").Value =
                excesoAcumulado;

            hoja.Cell("H10")
                .Style.NumberFormat.Format =
                "₡#,##0.00";

            hoja.Range("E10:H10")
                .Style.Fill.BackgroundColor =
                colorRojoSuave;

            hoja.Range("A10:H10")
                .Style.Font.Bold = true;

            hoja.Range("A10:H10")
                .Style.Border.OutsideBorder =
                XLBorderStyleValues.Thin;

            hoja.Range("A10:H10")
                .Style.Border.InsideBorder =
                XLBorderStyleValues.Thin;

            // ============================================================
            // RESUMEN
            // ============================================================

            hoja.Cell("A12").Value =
                "Resumen";

            hoja.Range(12, 1, 12, totalColumnas).Merge();

            hoja.Range(12, 1, 12, totalColumnas)
                .Style.Fill.BackgroundColor =
                colorInstitucional;

            hoja.Range(12, 1, 12, totalColumnas)
                .Style.Font.FontColor =
                XLColor.White;

            hoja.Range(12, 1, 12, totalColumnas)
                .Style.Font.Bold = true;

            hoja.Cell("A13").Value =
                "Total de categorías";

            hoja.Cell("B13").Value =
                categoriasControl.Count;

            hoja.Cell("C13").Value =
                "Activas";

            hoja.Cell("D13").Value =
                categoriasControl.Count(x =>
                    x.Categoria.Activo);

            hoja.Cell("E13").Value =
                "Inactivas";

            hoja.Cell("F13").Value =
                categoriasControl.Count(x =>
                    !x.Categoria.Activo);

            hoja.Cell("G13").Value =
                "Gasto del período";

            hoja.Cell("H13").Value =
                categoriasControl.Sum(x =>
                    x.GastoPeriodo);

            hoja.Cell("H13")
                .Style.NumberFormat.Format =
                "₡#,##0.00";

            hoja.Range("A13:H13")
                .Style.Fill.BackgroundColor =
                colorFondoSuave;

            hoja.Range("A13:H13")
                .Style.Font.Bold = true;

            // ============================================================
            // TABLA
            // ============================================================

            const int filaEncabezado = 15;

            string[] encabezados =
            {
        "Nombre",
        "Descripción",
        "Límite mensual",
        "Gasto del período",
        "% consumido",
        "Estado del límite",
        "Cuenta contable",
        "Estado"
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

            var encabezadoTabla =
                hoja.Range(
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

            int fila =
                filaEncabezado + 1;

            foreach (var item in categoriasControl)
            {
                var categoria =
                    item.Categoria;

                string cuentaContable =
                    categoria.CuentaContable != null
                        ? $"{categoria.CuentaContable.Codigo} - " +
                          $"{categoria.CuentaContable.Nombre}"
                        : "Sin cuenta asociada";

                hoja.Cell(fila, 1).Value =
                    categoria.Nombre;

                hoja.Cell(fila, 2).Value =
                    string.IsNullOrWhiteSpace(
                        categoria.Descripcion)
                        ? "Sin descripción"
                        : categoria.Descripcion;

                if (categoria.LimiteMensual > 0)
                {
                    hoja.Cell(fila, 3).Value =
                        categoria.LimiteMensual;

                    hoja.Cell(fila, 3)
                        .Style.NumberFormat.Format =
                        "₡#,##0.00";
                }
                else
                {
                    hoja.Cell(fila, 3).Value =
                        "Sin límite";
                }

                hoja.Cell(fila, 4).Value =
                    item.GastoPeriodo;

                hoja.Cell(fila, 4)
                    .Style.NumberFormat.Format =
                    "₡#,##0.00";

                if (categoria.LimiteMensual > 0)
                {
                    hoja.Cell(fila, 5).Value =
                        item.PorcentajeConsumido / 100m;

                    hoja.Cell(fila, 5)
                        .Style.NumberFormat.Format =
                        "0.00%";
                }
                else
                {
                    hoja.Cell(fila, 5).Value =
                        "No aplica";
                }

                hoja.Cell(fila, 6).Value =
                    item.EstadoLimite;

                hoja.Cell(fila, 7).Value =
                    cuentaContable;

                hoja.Cell(fila, 8).Value =
                    categoria.Activo
                        ? "Activo"
                        : "Inactivo";

                switch (item.EstadoLimite)
                {
                    case "Excedida":
                        hoja.Cell(fila, 6)
                            .Style.Fill.BackgroundColor =
                            colorRojoSuave;
                        break;

                    case "Próxima al límite":
                        hoja.Cell(fila, 6)
                            .Style.Fill.BackgroundColor =
                            colorAmarilloSuave;
                        break;

                    case "Dentro del límite":
                        hoja.Cell(fila, 6)
                            .Style.Fill.BackgroundColor =
                            colorVerdeSuave;
                        break;

                    default:
                        hoja.Cell(fila, 6)
                            .Style.Fill.BackgroundColor =
                            colorGrisSuave;
                        break;
                }

                hoja.Cell(fila, 6)
                    .Style.Font.Bold = true;

                fila++;
            }

            if (categoriasControl.Count > 0)
            {
                int filaFinDatos =
                    fila - 1;

                var rangoTabla =
                    hoja.Range(
                        filaEncabezado,
                        1,
                        filaFinDatos,
                        totalColumnas);

                var tabla =
                    rangoTabla.CreateTable(
                        "TablaControlCategoriasGastos");

                tabla.Theme =
                    XLTableTheme.None;

                tabla.ShowAutoFilter =
                    true;

                tabla.ShowRowStripes =
                    false;

                hoja.Range(
                        filaEncabezado + 1,
                        1,
                        filaFinDatos,
                        totalColumnas)
                    .Style.Border.BottomBorder =
                    XLBorderStyleValues.Thin;
            }
            else
            {
                hoja.Cell(fila, 1).Value =
                    "No se encontraron categorías con los filtros aplicados.";

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
            }

            hoja.Column(1).Width = 28;
            hoja.Column(2).Width = 45;
            hoja.Column(3).Width = 18;
            hoja.Column(4).Width = 20;
            hoja.Column(5).Width = 16;
            hoja.Column(6).Width = 24;
            hoja.Column(7).Width = 40;
            hoja.Column(8).Width = 14;

            hoja.Columns(1, totalColumnas)
                .Style.Alignment.WrapText =
                true;

            hoja.SheetView.FreezeRows(
                filaEncabezado);

            hoja.PageSetup.PageOrientation =
                XLPageOrientation.Landscape;

            hoja.PageSetup.PaperSize =
                XLPaperSize.A4Paper;

            hoja.PageSetup.FitToPages(
                1,
                0);

            hoja.PageSetup.CenterHorizontally =
                true;

            using var stream =
                new MemoryStream();

            workbook.SaveAs(stream);

            string nombreArchivo =
                $"Reporte_Control_Categorias_Gastos_" +
                $"{anioAplicado}_{mesAplicado:D2}_" +
                $"{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";

            await RegistrarAuditoria(
                "Exportar Excel",
                "CategoriasGastos",
                0,
                $"Se exportó el control mensual de categorías de gastos. " +
                $"Período: {periodoTexto}.");

            return File(
                stream.ToArray(),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                nombreArchivo);
        }
        public async Task<IActionResult> Details(int? id)
        {
            if (id == null) return NotFound();

            var categoriasGasto = await _context.CategoriasGastos
                .Include(c => c.CuentaContable)
                .FirstOrDefaultAsync(c => c.IdCategoriaGasto == id);

            if (categoriasGasto == null) return NotFound();

            return View(categoriasGasto);
        }

        public IActionResult Create()
        {
            CargarCuentasContables();
            return View();
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Create(
            [Bind("IdCategoriaGasto,Nombre,Descripcion,LimiteMensual,CuentaContableId")]
            CategoriasGasto categoriasGasto)
        {
            NormalizarDatos(categoriasGasto);
            await ValidarCategoriaGasto(categoriasGasto);

            if (ModelState.IsValid)
            {
                categoriasGasto.Activo = true;

                _context.CategoriasGastos.Add(categoriasGasto);
                await _context.SaveChangesAsync();

                await RegistrarAuditoria(
                    "Crear",
                    "CategoriasGastos",
                    categoriasGasto.IdCategoriaGasto,
                    $"Se creó la categoría de gasto: {categoriasGasto.Nombre}.");

                TempData["MensajeExito"] = "Categoría de gasto creada correctamente.";
                return RedirectToAction(nameof(Index));
            }

            CargarCuentasContables(categoriasGasto.CuentaContableId);
            return View(categoriasGasto);
        }

        public async Task<IActionResult> Edit(int? id)
        {
            if (id == null) return NotFound();

            var categoriasGasto = await _context.CategoriasGastos
                .FirstOrDefaultAsync(c => c.IdCategoriaGasto == id);

            if (categoriasGasto == null) return NotFound();

            ViewBag.TieneMovimientosAsociados =
                await TieneRegistrosAsociados(categoriasGasto.IdCategoriaGasto);

            CargarCuentasContables(categoriasGasto.CuentaContableId);
            return View(categoriasGasto);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Edit(
            int id,
            [Bind("IdCategoriaGasto,Nombre,Descripcion,LimiteMensual,CuentaContableId,Activo")]
            CategoriasGasto categoriasGasto)
        {
            if (id != categoriasGasto.IdCategoriaGasto) return NotFound();

            var categoriaActual = await _context.CategoriasGastos
                .AsNoTracking()
                .FirstOrDefaultAsync(c => c.IdCategoriaGasto == id);

            if (categoriaActual == null) return NotFound();

            bool tieneRegistrosAsociados = await TieneRegistrosAsociados(id);

            if (!categoriaActual.Activo)
            {
                TempData["MensajeError"] =
                    "La categoría de gasto está inactiva. Para modificarla primero debe activarla.";

                return RedirectToAction(nameof(Index));
            }

            NormalizarDatos(categoriasGasto);

            if (tieneRegistrosAsociados)
            {
                categoriasGasto.CuentaContableId = categoriaActual.CuentaContableId;
            }

            await ValidarCategoriaGasto(categoriasGasto);

            if (ModelState.IsValid)
            {
                try
                {
                    _context.CategoriasGastos.Update(categoriasGasto);
                    await _context.SaveChangesAsync();

                    await RegistrarAuditoria(
                        "Modificar",
                        "CategoriasGastos",
                        categoriasGasto.IdCategoriaGasto,
                        $"Se modificó la categoría de gasto: {categoriasGasto.Nombre}.");

                    TempData["MensajeExito"] = "Categoría de gasto modificada correctamente.";
                    return RedirectToAction(nameof(Index));
                }
                catch (DbUpdateConcurrencyException)
                {
                    if (!CategoriasGastoExists(categoriasGasto.IdCategoriaGasto))
                    {
                        return NotFound();
                    }

                    throw;
                }
            }

            ViewBag.TieneMovimientosAsociados = tieneRegistrosAsociados;
            CargarCuentasContables(categoriasGasto.CuentaContableId);
            return View(categoriasGasto);
        }

        public async Task<IActionResult> Delete(int? id)
        {
            if (id == null) return NotFound();

            var categoriasGasto = await _context.CategoriasGastos
                .Include(c => c.CuentaContable)
                .FirstOrDefaultAsync(c => c.IdCategoriaGasto == id);

            if (categoriasGasto == null) return NotFound();

            return View(categoriasGasto);
        }

        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> DeleteConfirmed(int id)
        {
            var categoriasGasto = await _context.CategoriasGastos
                .FirstOrDefaultAsync(c => c.IdCategoriaGasto == id);

            if (categoriasGasto == null) return NotFound();

            if (!categoriasGasto.Activo)
            {
                TempData["MensajeError"] = "La categoría de gasto ya se encuentra inactiva.";
                return RedirectToAction(nameof(Index));
            }

            categoriasGasto.Activo = false;
            _context.CategoriasGastos.Update(categoriasGasto);
            await _context.SaveChangesAsync();

            await RegistrarAuditoria(
                "Desactivar",
                "CategoriasGastos",
                categoriasGasto.IdCategoriaGasto,
                $"Se desactivó la categoría de gasto: {categoriasGasto.Nombre}.");

            TempData["MensajeExito"] = "Categoría de gasto desactivada correctamente.";
            return RedirectToAction(nameof(Index));
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Activar(int id)
        {
            var categoriasGasto = await _context.CategoriasGastos
                .FirstOrDefaultAsync(c => c.IdCategoriaGasto == id);

            if (categoriasGasto == null) return NotFound();

            if (categoriasGasto.Activo)
            {
                TempData["MensajeError"] = "La categoría de gasto ya se encuentra activa.";
                return RedirectToAction(nameof(Index));
            }

            categoriasGasto.Activo = true;
            _context.CategoriasGastos.Update(categoriasGasto);
            await _context.SaveChangesAsync();

            await RegistrarAuditoria(
                "Activar",
                "CategoriasGastos",
                categoriasGasto.IdCategoriaGasto,
                $"Se activó la categoría de gasto: {categoriasGasto.Nombre}.");

            TempData["MensajeExito"] = "Categoría de gasto activada correctamente.";
            return RedirectToAction(nameof(Index));
        }
        // ================================================================
        // CONSULTA COMPARTIDA PARA INDEX, REPORTE Y EXCEL
        // ================================================================
        private IQueryable<CategoriasGasto> ConstruirConsultaReporte(
            string? buscar,
            string? cuenta,
            string? estado,
            string? limite)
        {
            var consulta = _context.CategoriasGastos
                .Include(c => c.CuentaContable)
                .AsNoTracking()
                .AsQueryable();

            if (!string.IsNullOrWhiteSpace(buscar))
            {
                buscar = buscar.Trim();

                consulta = consulta.Where(c =>
                    c.Nombre.Contains(buscar) ||
                    (c.Descripcion != null &&
                     c.Descripcion.Contains(buscar)));
            }

            if (!string.IsNullOrWhiteSpace(cuenta))
            {
                cuenta = cuenta.Trim();

                consulta = consulta.Where(c =>
                    c.CuentaContable != null &&
                    (
                        c.CuentaContable.Codigo.Contains(cuenta) ||
                        c.CuentaContable.Nombre.Contains(cuenta)
                    ));
            }

            if (!string.IsNullOrWhiteSpace(estado))
            {
                if (estado == "Activo")
                {
                    consulta = consulta.Where(c => c.Activo);
                }
                else if (estado == "Inactivo")
                {
                    consulta = consulta.Where(c => !c.Activo);
                }
            }

            if (!string.IsNullOrWhiteSpace(limite))
            {
                if (limite == "ConLimite")
                {
                    consulta = consulta.Where(c =>
                        c.LimiteMensual > 0);
                }
                else if (limite == "SinLimite")
                {
                    consulta = consulta.Where(c =>
                        c.LimiteMensual <= 0);
                }
            }

            return consulta;
        }


        // ================================================================
        // CALCULAR CONSUMO DE CATEGORÍAS PARA EL AÑO Y MES SELECCIONADOS
        // ================================================================
        private async Task<List<CategoriaGastoControlViewModel>>
            CalcularControlMensualCategorias(
                List<CategoriasGasto> categorias,
                int? anio,
                int? mes)
        {
            var hoy =
                DateOnly.FromDateTime(DateTime.Today);

            int anioAplicado =
                anio ?? hoy.Year;

            int mesAplicado =
                mes ?? hoy.Month;

            var inicioPeriodo =
                new DateOnly(
                    anioAplicado,
                    mesAplicado,
                    1);

            var finPeriodo =
                inicioPeriodo
                    .AddMonths(1)
                    .AddDays(-1);

            var resultado =
                new List<CategoriaGastoControlViewModel>();

            foreach (var categoria in categorias)
            {
                /*
                 * Se consideran gastos comprometidos:
                 * - Aprobados.
                 * - Pendientes de autorización.
                 *
                 * Se excluyen:
                 * - Anulados.
                 * - Rechazados.
                 * - Inactivos.
                 */
                decimal gastoPeriodo =
                    await _context.Gastos
                        .AsNoTracking()
                        .Where(g =>
                            g.Activo &&
                            g.CategoriaGastoId ==
                                categoria.IdCategoriaGasto &&
                            g.Fecha >= inicioPeriodo &&
                            g.Fecha <= finPeriodo &&
                            g.Estado != "Anulado" &&
                            g.Estado != "Rechazado")
                        .SumAsync(g =>
                            g.MontoTotal ?? 0m);

                decimal porcentajeConsumido =
                    categoria.LimiteMensual > 0
                        ? gastoPeriodo /
                          categoria.LimiteMensual *
                          100m
                        : 0m;

                decimal exceso =
                    categoria.LimiteMensual > 0
                        ? Math.Max(
                            0m,
                            gastoPeriodo -
                            categoria.LimiteMensual)
                        : 0m;

                string estadoControl;

                if (!categoria.Activo)
                {
                    estadoControl = "Inactiva";
                }
                else if (categoria.LimiteMensual <= 0)
                {
                    estadoControl = "Sin límite";
                }
                else if (porcentajeConsumido > 100m)
                {
                    estadoControl = "Excedida";
                }
                else if (porcentajeConsumido >= 80m)
                {
                    estadoControl = "Próxima al límite";
                }
                else if (gastoPeriodo > 0)
                {
                    estadoControl = "Dentro del límite";
                }
                else
                {
                    estadoControl = "Sin consumo";
                }

                resultado.Add(
                    new CategoriaGastoControlViewModel
                    {
                        Categoria = categoria,
                        GastoPeriodo = gastoPeriodo,
                        PorcentajeConsumido =
                            porcentajeConsumido,
                        Exceso = exceso,
                        EstadoLimite = estadoControl
                    });
            }

            return resultado;
        }
        // ================================================================
        // FILTRO RÁPIDO ACTIVADO DESDE LAS ALERTAS
        // ================================================================
        private static List<CategoriaGastoControlViewModel>
            AplicarFiltroEstadoLimite(
                List<CategoriaGastoControlViewModel>
                    categoriasControl,
                string? estadoLimite)
        {
            if (estadoLimite == "Proxima")
            {
                return categoriasControl
                    .Where(x =>
                        x.EstadoLimite ==
                        "Próxima al límite")
                    .ToList();
            }

            if (estadoLimite == "Excedida")
            {
                return categoriasControl
                    .Where(x =>
                        x.EstadoLimite ==
                        "Excedida")
                    .ToList();
            }

            return categoriasControl;
        }
        private void NormalizarDatos(CategoriasGasto categoria)
        {
            categoria.Nombre = categoria.Nombre?.Trim() ?? "";
            categoria.Descripcion = categoria.Descripcion?.Trim();
        }

        private async Task ValidarCategoriaGasto(CategoriasGasto categoria)
        {
            if (string.IsNullOrWhiteSpace(categoria.Nombre))
            {
                ModelState.AddModelError("Nombre", "El nombre es obligatorio.");
            }
            else if (categoria.Nombre.Length < 3)
            {
                ModelState.AddModelError("Nombre", "El nombre debe tener al menos 3 caracteres.");
            }
            else if (categoria.Nombre.Length > 100)
            {
                ModelState.AddModelError("Nombre", "El nombre no puede superar los 100 caracteres.");
            }
            else if (!System.Text.RegularExpressions.Regex.IsMatch(
                categoria.Nombre,
                @"^[A-Za-zÁÉÍÓÚáéíóúÑñ\s]+$"))
            {
                ModelState.AddModelError("Nombre", "El nombre solo puede contener letras y espacios.");
            }

            bool nombreDuplicado = await _context.CategoriasGastos
                .AnyAsync(c =>
                    c.Nombre.ToLower() == categoria.Nombre.ToLower()
                    && c.IdCategoriaGasto != categoria.IdCategoriaGasto);

            if (nombreDuplicado)
            {
                ModelState.AddModelError("Nombre", "Ya existe una categoría de gasto con ese nombre.");
            }

            if (!string.IsNullOrWhiteSpace(categoria.Descripcion))
            {
                if (categoria.Descripcion.Length < 3)
                {
                    ModelState.AddModelError("Descripcion", "La descripción debe tener al menos 3 caracteres.");
                }
                else if (categoria.Descripcion.Length > 250)
                {
                    ModelState.AddModelError("Descripcion", "La descripción no puede superar los 250 caracteres.");
                }
                else if (!System.Text.RegularExpressions.Regex.IsMatch(
                    categoria.Descripcion,
                    @"^[A-Za-zÁÉÍÓÚáéíóúÑñ0-9\s.,;:()/$#°-]+$"))
                {
                    ModelState.AddModelError("Descripcion", "La descripción contiene caracteres no permitidos.");
                }
                else if (!System.Text.RegularExpressions.Regex.IsMatch(
                    categoria.Descripcion,
                    @"[A-Za-zÁÉÍÓÚáéíóúÑñ0-9]"))
                {
                    ModelState.AddModelError("Descripcion", "La descripción debe contener al menos una letra o un número.");
                }
            }

            if (categoria.LimiteMensual < 0)
            {
                ModelState.AddModelError("LimiteMensual", "El límite mensual debe ser mayor o igual a 0.");
            }

            bool cuentaExiste = await _context.CuentasContables
                .AnyAsync(c =>
                    c.IdCuentaContable == categoria.CuentaContableId
                    && c.Activo);

            if (!cuentaExiste)
            {
                ModelState.AddModelError("CuentaContableId", "Debe seleccionar una cuenta contable activa.");
            }
        }

        private async Task<bool> TieneRegistrosAsociados(int idCategoriaGasto)
        {
            bool tieneGastos = await _context.Gastos
                .AnyAsync(g => g.CategoriaGastoId == idCategoriaGasto);

            bool tienePresupuesto = await _context.PresupuestoDetalles
                .AnyAsync(p => p.CategoriaGastoId == idCategoriaGasto);

            return tieneGastos || tienePresupuesto;
        }

        private void CargarCuentasContables(int? cuentaSeleccionada = null)
        {
            ViewData["CuentaContableId"] = new SelectList(
                _context.CuentasContables
                    .Where(c => c.Activo || c.IdCuentaContable == cuentaSeleccionada)
                    .OrderBy(c => c.Codigo)
                    .Select(c => new
                    {
                        c.IdCuentaContable,
                        Nombre = c.Codigo + " - " + c.Nombre
                    }),
                "IdCuentaContable",
                "Nombre",
                cuentaSeleccionada);
        }

        private async Task RegistrarAuditoria(
            string accion,
            string tabla,
            int registroId,
            string descripcion)
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

        private bool CategoriasGastoExists(int id)
        {
            return _context.CategoriasGastos.Any(e => e.IdCategoriaGasto == id);
        }
    }
}