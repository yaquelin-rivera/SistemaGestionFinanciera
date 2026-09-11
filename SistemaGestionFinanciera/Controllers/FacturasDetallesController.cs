using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Http;
using Microsoft.EntityFrameworkCore;
using SistemaGestionFinanciera.Data;
using SistemaGestionFinanciera.Models;
using ClosedXML.Excel;
using System.IO;

namespace SistemaGestionFinanciera.Controllers
{
    public class FacturasDetallesController : BaseController
    {
        private readonly SistemaFinancieroContext _context;

        public FacturasDetallesController(SistemaFinancieroContext context)
        {
            _context = context;
        }
       
        // INDEX - CATÁLOGO DE SERVICIOS O PRODUCTOS
        // ================================================================
        public async Task<IActionResult> Index(
            string? buscar,
            string? estado)
        {
            var catalogo = ConstruirConsultaReporte(
                buscar,
                estado);

            ViewBag.Buscar = buscar;
            ViewBag.Estado = estado;

            return View(await catalogo
                .OrderByDescending(d => d.IdFacturaDetalle)
                .ToListAsync());
        }
        // ================================================================
        // REPORTE DEL CATÁLOGO DE SERVICIOS O PRODUCTOS
        // ================================================================
        [HttpGet]
        public async Task<IActionResult> Reporte(
            string? buscar,
            string? estado)
        {
            var servicios = await ConstruirConsultaReporte(
                    buscar,
                    estado)
                .OrderByDescending(d => d.IdFacturaDetalle)
                .ToListAsync();

            ViewBag.TotalRegistros =
                servicios.Count;

            ViewBag.TotalActivos =
                servicios.Count(d => d.Activo);

            ViewBag.TotalInactivos =
                servicios.Count(d => !d.Activo);

            ViewBag.PrecioPromedio =
                servicios.Any()
                    ? servicios.Average(d => d.PrecioUnitario)
                    : 0;

            ViewBag.Buscar = buscar;
            ViewBag.Estado = estado;
            ViewBag.FechaGeneracion = DateTime.Now;
            await RegistrarAuditoria(
    "Vista previa PDF",
    "FacturaDetalles",
    0,
    $"Se generó la vista previa del catálogo de servicios y productos. " +
    ConstruirDescripcionFiltros(buscar, estado));
            return View(servicios);
        }
        // ================================================================
        // EXPORTAR CATÁLOGO DE SERVICIOS O PRODUCTOS A EXCEL
        // ================================================================
        [HttpGet]
        public async Task<IActionResult> ExportarExcel(
            string? buscar,
            string? estado)
        {
            var servicios = await ConstruirConsultaReporte(
                    buscar,
                    estado)
                .OrderByDescending(d => d.IdFacturaDetalle)
                .ToListAsync();

            using var workbook = new XLWorkbook();

            var hoja = workbook.Worksheets.Add(
                "Servicios y productos");

            const int totalColumnas = 4;

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

            var colorVerde =
                XLColor.FromHtml("#198754");

            var colorVerdeSuave =
                XLColor.FromHtml("#D1E7DD");

            var colorGris =
                XLColor.FromHtml("#6C757D");

            var colorGrisSuave =
                XLColor.FromHtml("#E2E3E5");

            var colorAzulSuave =
                XLColor.FromHtml("#D9EAF7");

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
                "Reporte del Catálogo de Servicios y Productos";

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

            hoja.Cell("C6").Value = "Estado:";

            hoja.Cell("D6").Value =
                string.IsNullOrWhiteSpace(estado)
                    ? "Todos"
                    : estado;

            hoja.Cell("A6").Style.Font.Bold = true;
            hoja.Cell("C6").Style.Font.Bold = true;

            var rangoFiltros = hoja.Range(
                6,
                1,
                6,
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

            hoja.Row(6).Height = 22;

            // ============================================================
            // RESUMEN
            // ============================================================
            hoja.Cell("A8").Value =
                "Resumen";

            hoja.Range(
                8,
                1,
                8,
                totalColumnas).Merge();

            var tituloResumen = hoja.Range(
                8,
                1,
                8,
                totalColumnas);

            tituloResumen.Style.Fill.BackgroundColor =
                colorInstitucional;

            tituloResumen.Style.Font.FontColor =
                XLColor.White;

            tituloResumen.Style.Font.Bold = true;

            hoja.Cell("A9").Value =
                "Total de servicios";

            hoja.Cell("B9").Value =
                servicios.Count;

            hoja.Cell("C9").Value =
                "Activos";

            hoja.Cell("D9").Value =
                servicios.Count(d => d.Activo);

            hoja.Cell("A10").Value =
                "Inactivos";

            hoja.Cell("B10").Value =
                servicios.Count(d => !d.Activo);

            hoja.Cell("C10").Value =
                "Precio promedio";

            hoja.Cell("D10").Value =
                servicios.Any()
                    ? servicios.Average(d => d.PrecioUnitario)
                    : 0;

            hoja.Cell("D10").Style.NumberFormat.Format =
                "₡#,##0.00";

            hoja.Range("A9:B9")
                .Style.Fill.BackgroundColor =
                colorAzulSuave;

            hoja.Range("C9:D9")
                .Style.Fill.BackgroundColor =
                colorVerdeSuave;

            hoja.Range("A10:B10")
                .Style.Fill.BackgroundColor =
                colorGrisSuave;

            hoja.Range("C10:D10")
                .Style.Fill.BackgroundColor =
                colorDoradoSuave;

            hoja.Range("A9:D10").Style.Font.Bold = true;

            hoja.Range("A9:D10")
                .Style.Border.OutsideBorder =
                XLBorderStyleValues.Thin;

            hoja.Range("A9:D10")
                .Style.Border.InsideBorder =
                XLBorderStyleValues.Thin;

            hoja.Range("A9:D10")
                .Style.Alignment.Vertical =
                XLAlignmentVerticalValues.Center;

            hoja.Cell("B9")
                .Style.Alignment.Horizontal =
                XLAlignmentHorizontalValues.Right;

            hoja.Cell("D9")
                .Style.Alignment.Horizontal =
                XLAlignmentHorizontalValues.Right;

            hoja.Cell("B10")
                .Style.Alignment.Horizontal =
                XLAlignmentHorizontalValues.Right;

            hoja.Cell("D10")
                .Style.Alignment.Horizontal =
                XLAlignmentHorizontalValues.Right;

            hoja.Row(9).Height = 22;
            hoja.Row(10).Height = 22;

            // ============================================================
            // ENCABEZADOS DE LA TABLA
            // ============================================================
            const int filaEncabezado = 12;

            string[] encabezados =
            {
        "Nombre",
        "Descripción",
        "Precio unitario",
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

            foreach (var servicio in servicios)
            {
                hoja.Cell(fila, 1).Value =
                    servicio.Nombre;

                hoja.Cell(fila, 2).Value =
                    string.IsNullOrWhiteSpace(servicio.Descripcion)
                        ? "Sin descripción"
                        : servicio.Descripcion;

                hoja.Cell(fila, 3).Value =
                    servicio.PrecioUnitario;

                hoja.Cell(fila, 3)
                    .Style.NumberFormat.Format =
                    "₡#,##0.00";

                hoja.Cell(fila, 4).Value =
                    servicio.Activo
                        ? "Activo"
                        : "Inactivo";

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

                if (servicio.Activo)
                {
                    hoja.Cell(fila, 4)
                        .Style.Fill.BackgroundColor =
                        colorVerdeSuave;

                    hoja.Cell(fila, 4)
                        .Style.Font.FontColor =
                        colorVerde;
                }
                else
                {
                    hoja.Cell(fila, 4)
                        .Style.Fill.BackgroundColor =
                        colorGrisSuave;

                    hoja.Cell(fila, 4)
                        .Style.Font.FontColor =
                        colorGris;
                }

                hoja.Cell(fila, 4)
                    .Style.Font.Bold = true;

                fila++;
            }

            // ============================================================
            // CONVERTIR RANGO EN TABLA
            // ============================================================
            if (servicios.Any())
            {
                int filaFinDatos = fila - 1;

                var rangoTabla = hoja.Range(
                    filaEncabezado,
                    1,
                    filaFinDatos,
                    totalColumnas);

                var tabla = rangoTabla.CreateTable(
                    "TablaServiciosProductos");

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
                        3,
                        filaFinDatos,
                        3)
                    .Style.Alignment.Horizontal =
                    XLAlignmentHorizontalValues.Right;

                hoja.Range(
                        filaEncabezado + 1,
                        4,
                        filaFinDatos,
                        4)
                    .Style.Alignment.Horizontal =
                    XLAlignmentHorizontalValues.Center;

                hoja.Range(
                        filaEncabezado + 1,
                        1,
                        filaFinDatos,
                        2)
                    .Style.Alignment.WrapText = true;
            }
            else
            {
                hoja.Cell(fila, 1).Value =
                    "No se encontraron servicios o productos con los filtros aplicados.";

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

            hoja.Cell(filaTotal, 3).Value =
                "Total de registros:";

            hoja.Cell(filaTotal, 4).Value =
                servicios.Count;

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

            hoja.Cell(filaTotal, 3).Style.Font.Bold = true;
            hoja.Cell(filaTotal, 4).Style.Font.Bold = true;

            hoja.Cell(filaTotal, 3)
                .Style.Font.FontColor =
                colorInstitucionalOscuro;

            hoja.Cell(filaTotal, 4)
                .Style.Font.FontColor =
                colorInstitucionalOscuro;

            hoja.Cell(filaTotal, 4)
                .Style.Alignment.Horizontal =
                XLAlignmentHorizontalValues.Center;

            // ============================================================
            // CONFIGURACIÓN FINAL
            // ============================================================
            hoja.Column(1).Width = 34;
            hoja.Column(2).Width = 52;
            hoja.Column(3).Width = 20;
            hoja.Column(4).Width = 14;

            hoja.Column(1).Style.Alignment.WrapText = true;
            hoja.Column(2).Style.Alignment.WrapText = true;

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
                $"Reporte_Servicios_Productos_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";
            await RegistrarAuditoria(
    "Exportar Excel",
    "FacturaDetalles",
    0,
    $"Se exportó el catálogo de servicios y productos a Excel. " +
    ConstruirDescripcionFiltros(buscar, estado));
            return File(
                stream.ToArray(),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                nombreArchivo);
        }
        public async Task<IActionResult> Details(int? id)
        {
            if (id == null) return NotFound();

            var detalle = await _context.FacturaDetalles
                .FirstOrDefaultAsync(d => d.IdFacturaDetalle == id && d.FacturaId == null);

            if (detalle == null) return NotFound();

            return View(detalle);
        }

        public IActionResult Create()
        {
            return View(new FacturaDetalle
            {
                Cantidad = 1,
                Descuento = 0,
                Impuesto = 0,
                TotalLinea = 0,
                Activo = true
            });
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Create(
            [Bind("Nombre,Descripcion,PrecioUnitario")]
            FacturaDetalle detalle)
        {
            NormalizarDatos(detalle);
            await ValidarDetalleCatalogo(detalle);

            if (ModelState.IsValid)
            {
                detalle.FacturaId = null;
                detalle.Cantidad = 1;
                detalle.Descuento = 0;
                detalle.Impuesto = 0;
                detalle.TotalLinea = detalle.PrecioUnitario;
                detalle.Activo = true;

                _context.FacturaDetalles.Add(detalle);
                await _context.SaveChangesAsync();

                await RegistrarAuditoria(
                    "Crear",
                    "FacturaDetalles",
                    detalle.IdFacturaDetalle,
                    $"Se creó el servicio o producto: {detalle.Nombre}.");

                TempData["MensajeExito"] = "Servicio o producto registrado correctamente.";
                return RedirectToAction(nameof(Index));
            }

            return View(detalle);
        }

        public async Task<IActionResult> Edit(int? id)
        {
            if (id == null) return NotFound();

            var detalle = await _context.FacturaDetalles
                .FirstOrDefaultAsync(d => d.IdFacturaDetalle == id && d.FacturaId == null);

            if (detalle == null) return NotFound();

            if (!detalle.Activo)
            {
                TempData["MensajeError"] =
                    "El servicio o producto está inactivo. Para modificarlo primero debe activarlo.";

                return RedirectToAction(nameof(Index));
            }

            return View(detalle);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Edit(
            int id,
            [Bind("IdFacturaDetalle,Nombre,Descripcion,PrecioUnitario,Activo")]
            FacturaDetalle detalle)
        {
            if (id != detalle.IdFacturaDetalle) return NotFound();

            var detalleActual = await _context.FacturaDetalles
                .AsNoTracking()
                .FirstOrDefaultAsync(d => d.IdFacturaDetalle == id && d.FacturaId == null);

            if (detalleActual == null) return NotFound();

            if (!detalleActual.Activo)
            {
                TempData["MensajeError"] =
                    "El servicio o producto está inactivo. Para modificarlo primero debe activarlo.";

                return RedirectToAction(nameof(Index));
            }

            NormalizarDatos(detalle);
            await ValidarDetalleCatalogo(detalle);

            if (ModelState.IsValid)
            {
                try
                {
                    detalle.FacturaId = null;
                    detalle.Cantidad = 1;
                    detalle.Descuento = 0;
                    detalle.Impuesto = 0;
                    detalle.TotalLinea = detalle.PrecioUnitario;
                    detalle.Activo = detalleActual.Activo;

                    _context.FacturaDetalles.Update(detalle);
                    await _context.SaveChangesAsync();

                    await RegistrarAuditoria(
                        "Modificar",
                        "FacturaDetalles",
                        detalle.IdFacturaDetalle,
                        $"Se modificó el servicio o producto: {detalle.Nombre}.");

                    TempData["MensajeExito"] = "Servicio o producto modificado correctamente.";
                    return RedirectToAction(nameof(Index));
                }
                catch (DbUpdateConcurrencyException)
                {
                    if (!FacturaDetalleExists(detalle.IdFacturaDetalle))
                        return NotFound();

                    throw;
                }
            }

            return View(detalle);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> CambiarEstado(int id)
        {
            var detalle = await _context.FacturaDetalles
                .FirstOrDefaultAsync(d => d.IdFacturaDetalle == id && d.FacturaId == null);

            if (detalle == null) return NotFound();

            detalle.Activo = !detalle.Activo;

            _context.FacturaDetalles.Update(detalle);
            await _context.SaveChangesAsync();

            await RegistrarAuditoria(
                detalle.Activo ? "Activar" : "Desactivar",
                "FacturaDetalles",
                detalle.IdFacturaDetalle,
                detalle.Activo
                    ? $"Se activó el servicio o producto: {detalle.Nombre}."
                    : $"Se desactivó el servicio o producto: {detalle.Nombre}.");

            TempData["MensajeExito"] = detalle.Activo
                ? "Servicio o producto activado correctamente."
                : "Servicio o producto desactivado correctamente.";

            return RedirectToAction(nameof(Index));
        }
        // ================================================================
        // CONSULTA COMPARTIDA PARA INDEX, REPORTE Y EXCEL
        // ================================================================
        private IQueryable<FacturaDetalle> ConstruirConsultaReporte(
            string? buscar,
            string? estado)
        {
            var consulta = _context.FacturaDetalles
                .Where(d => d.FacturaId == null)
                .AsNoTracking()
                .AsQueryable();

            if (!string.IsNullOrWhiteSpace(buscar))
            {
                buscar = buscar.Trim();

                consulta = consulta.Where(d =>
                    d.Nombre.Contains(buscar) ||
                    (
                        d.Descripcion != null &&
                        d.Descripcion.Contains(buscar)
                    ));
            }

            if (!string.IsNullOrWhiteSpace(estado))
            {
                if (estado == "Activo")
                {
                    consulta = consulta.Where(d =>
                        d.Activo);
                }
                else if (estado == "Inactivo")
                {
                    consulta = consulta.Where(d =>
                        !d.Activo);
                }
            }

            return consulta;
        }
        private string ConstruirDescripcionFiltros(
    string? buscar,
    string? estado)
        {
            var filtros = new List<string>();

            if (!string.IsNullOrWhiteSpace(buscar))
            {
                filtros.Add($"Búsqueda: {buscar.Trim()}");
            }

            if (!string.IsNullOrWhiteSpace(estado))
            {
                filtros.Add($"Estado: {estado}");
            }

            if (!filtros.Any())
            {
                return "Sin filtros aplicados.";
            }

            return "Filtros aplicados: " +
                   string.Join(" | ", filtros) +
                   ".";
        }
        private void NormalizarDatos(FacturaDetalle detalle)
        {
            detalle.Nombre = detalle.Nombre?.Trim() ?? "";
            detalle.Descripcion = detalle.Descripcion?.Trim();
        }

        private async Task ValidarDetalleCatalogo(FacturaDetalle detalle)
        {
            if (string.IsNullOrWhiteSpace(detalle.Nombre))
            {
                ModelState.AddModelError("Nombre", "El nombre del servicio o producto es obligatorio.");
            }
            else if (detalle.Nombre.Length < 3)
            {
                ModelState.AddModelError("Nombre", "El nombre debe tener al menos 3 caracteres.");
            }
            else if (detalle.Nombre.Length > 150)
            {
                ModelState.AddModelError("Nombre", "El nombre no puede superar los 150 caracteres.");
            }
            else if (!System.Text.RegularExpressions.Regex.IsMatch(
                detalle.Nombre,
                @"^[A-Za-zÁÉÍÓÚáéíóúÑñ0-9\s.,()/-]+$"))
            {
                ModelState.AddModelError("Nombre", "El nombre contiene caracteres no permitidos.");
            }

            bool nombreDuplicado = await _context.FacturaDetalles
                .AnyAsync(d =>
                    d.FacturaId == null &&
                    d.Nombre.ToLower() == detalle.Nombre.ToLower() &&
                    d.IdFacturaDetalle != detalle.IdFacturaDetalle);

            if (nombreDuplicado)
            {
                ModelState.AddModelError("Nombre", "Ya existe un servicio o producto con ese nombre.");
            }

            if (!string.IsNullOrWhiteSpace(detalle.Descripcion))
            {
                if (detalle.Descripcion.Length > 250)
                {
                    ModelState.AddModelError("Descripcion", "La descripción no puede superar los 250 caracteres.");
                }
                else if (!System.Text.RegularExpressions.Regex.IsMatch(
                    detalle.Descripcion,
                    @"^[A-Za-zÁÉÍÓÚáéíóúÑñ0-9\s.,;:()/-]+$"))
                {
                    ModelState.AddModelError("Descripcion", "La descripción contiene caracteres no permitidos.");
                }
            }

            if (detalle.PrecioUnitario <= 0)
            {
                ModelState.AddModelError("PrecioUnitario", "El precio unitario debe ser mayor a cero.");
            }

            if (detalle.PrecioUnitario != Math.Round(detalle.PrecioUnitario, 2))
            {
                ModelState.AddModelError("PrecioUnitario", "El precio unitario solo puede tener dos decimales.");
            }
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

        private bool FacturaDetalleExists(int id)
        {
            return _context.FacturaDetalles.Any(e => e.IdFacturaDetalle == id);
        }
    }
}