using System;
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
    public class CentrosCostosController : BaseController
    {
        private readonly SistemaFinancieroContext _context;

        public CentrosCostosController(SistemaFinancieroContext context)
        {
            _context = context;
        }
        // INDEX - LISTADO DE CENTROS DE COSTO
        // ================================================================
        public async Task<IActionResult> Index(
            string? buscar,
            string? estado)
        {
            var centros = ConstruirConsultaReporte(
                buscar,
                estado);

            ViewBag.Buscar = buscar;
            ViewBag.Estado = estado;

            return View(await centros
                .OrderByDescending(c => c.IdCentroCosto)
                .ToListAsync());
        }
        // ================================================================
        // REPORTE DE CENTROS DE COSTO
        // ================================================================
        [HttpGet]
        public async Task<IActionResult> Reporte(
            string? buscar,
            string? estado)
        {
            var centros = await ConstruirConsultaReporte(
                    buscar,
                    estado)
                .OrderByDescending(c => c.IdCentroCosto)
                .ToListAsync();

            ViewBag.TotalRegistros =
                centros.Count;

            ViewBag.TotalActivos =
                centros.Count(c => c.Activo);

            ViewBag.TotalInactivos =
                centros.Count(c => !c.Activo);

            ViewBag.TotalConMovimientos =
                centros.Count(c =>
                    c.MovimientosContables != null &&
                    c.MovimientosContables.Any());

            ViewBag.Buscar = buscar;
            ViewBag.Estado = estado;
            ViewBag.FechaGeneracion = DateTime.Now;
            await RegistrarAuditoria(
    "Vista previa PDF",
    "CentrosCosto",
    0,
    "Se generó la vista previa del reporte de centros de costo."
);
            return View(centros);
        }
        // ================================================================
        // EXPORTAR REPORTE DE CENTROS DE COSTO A EXCEL
        // ================================================================
        [HttpGet]
        public async Task<IActionResult> ExportarExcel(
            string? buscar,
            string? estado)
        {
            var centros = await ConstruirConsultaReporte(
                    buscar,
                    estado)
                .OrderByDescending(c => c.IdCentroCosto)
                .ToListAsync();

            using var workbook = new XLWorkbook();

            var hoja = workbook.Worksheets.Add(
                "Centros de costo");

            const int totalColumnas = 5;

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
            // ENCABEZADO
            // ============================================================
            hoja.Cell("A1").Value =
                "Unión Cantonal de Asociaciones de Desarrollo de Santa Ana";

            hoja.Range(
                1,
                1,
                1,
                totalColumnas).Merge();

            hoja.Cell("A2").Value =
                "Reporte de Centros de Costo";

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

            rangoFecha.Style.Alignment.Horizontal =
                XLAlignmentHorizontalValues.Center;

            rangoFecha.Style.Alignment.Vertical =
                XLAlignmentVerticalValues.Center;

            rangoFecha.Style.Font.FontColor =
                XLColor.FromHtml("#666666");

            rangoFecha.Style.Fill.BackgroundColor =
                XLColor.White;

            hoja.Row(1).Height = 24;
            hoja.Row(2).Height = 22;
            hoja.Row(3).Height = 20;

            // ============================================================
            // FILTROS
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

            hoja.Cell("A9").Value = "Total";
            hoja.Cell("B9").Value = centros.Count;

            hoja.Cell("C9").Value = "Activos";
            hoja.Cell("D9").Value =
                centros.Count(c => c.Activo);

            hoja.Cell("A10").Value = "Inactivos";
            hoja.Cell("B10").Value =
                centros.Count(c => !c.Activo);

            hoja.Cell("C10").Value =
                "Con movimientos";

            hoja.Cell("D10").Value =
                centros.Count(c =>
                    c.MovimientosContables != null &&
                    c.MovimientosContables.Any());

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

            hoja.Cell("E9").Value = "";
            hoja.Range("E9:E10").Merge();

            hoja.Range("E9:E10")
                .Style.Fill.BackgroundColor =
                colorFondoSuave;

            hoja.Range("A9:E10").Style.Font.Bold = true;

            hoja.Range("A9:E10")
                .Style.Border.OutsideBorder =
                XLBorderStyleValues.Thin;

            hoja.Range("A9:E10")
                .Style.Border.InsideBorder =
                XLBorderStyleValues.Thin;

            hoja.Range("A9:E10")
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
            // ENCABEZADOS
            // ============================================================
            const int filaEncabezado = 12;

            string[] encabezados =
            {
        "Código",
        "Nombre",
        "Descripción",
        "Movimientos",
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

            foreach (var centro in centros)
            {
                int cantidadMovimientos =
                    centro.MovimientosContables?.Count ?? 0;

                hoja.Cell(fila, 1).Value =
                    centro.Codigo;

                hoja.Cell(fila, 2).Value =
                    centro.Nombre;

                hoja.Cell(fila, 3).Value =
                    string.IsNullOrWhiteSpace(centro.Descripcion)
                        ? "Sin descripción"
                        : centro.Descripcion;

                hoja.Cell(fila, 4).Value =
                    cantidadMovimientos > 0
                        ? $"{cantidadMovimientos} movimiento(s)"
                        : "Sin movimientos";

                hoja.Cell(fila, 5).Value =
                    centro.Activo
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

                if (centro.Activo)
                {
                    hoja.Cell(fila, 5)
                        .Style.Fill.BackgroundColor =
                        colorVerdeSuave;

                    hoja.Cell(fila, 5)
                        .Style.Font.FontColor =
                        colorVerde;
                }
                else
                {
                    hoja.Cell(fila, 5)
                        .Style.Fill.BackgroundColor =
                        colorGrisSuave;

                    hoja.Cell(fila, 5)
                        .Style.Font.FontColor =
                        colorGris;
                }

                hoja.Cell(fila, 5)
                    .Style.Font.Bold = true;

                fila++;
            }

            // ============================================================
            // TABLA
            // ============================================================
            if (centros.Any())
            {
                int filaFinDatos = fila - 1;

                var rangoTabla = hoja.Range(
                    filaEncabezado,
                    1,
                    filaFinDatos,
                    totalColumnas);

                var tabla = rangoTabla.CreateTable(
                    "TablaCentrosCosto");

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
                        1)
                    .Style.Alignment.Horizontal =
                    XLAlignmentHorizontalValues.Center;

                hoja.Range(
                        filaEncabezado + 1,
                        4,
                        filaFinDatos,
                        5)
                    .Style.Alignment.Horizontal =
                    XLAlignmentHorizontalValues.Center;

                hoja.Range(
                        filaEncabezado + 1,
                        2,
                        filaFinDatos,
                        3)
                    .Style.Alignment.WrapText = true;
            }
            else
            {
                hoja.Cell(fila, 1).Value =
                    "No se encontraron centros de costo con los filtros aplicados.";

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

            hoja.Cell(filaTotal, 4).Value =
                "Total de registros:";

            hoja.Cell(filaTotal, 5).Value =
                centros.Count;

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

            hoja.Cell(filaTotal, 4).Style.Font.Bold = true;
            hoja.Cell(filaTotal, 5).Style.Font.Bold = true;

            hoja.Cell(filaTotal, 4)
                .Style.Font.FontColor =
                colorInstitucionalOscuro;

            hoja.Cell(filaTotal, 5)
                .Style.Font.FontColor =
                colorInstitucionalOscuro;

            hoja.Cell(filaTotal, 5)
                .Style.Alignment.Horizontal =
                XLAlignmentHorizontalValues.Center;

            // ============================================================
            // CONFIGURACIÓN FINAL
            // ============================================================
            hoja.Column(1).Width = 18;
            hoja.Column(2).Width = 32;
            hoja.Column(3).Width = 48;
            hoja.Column(4).Width = 22;
            hoja.Column(5).Width = 14;

            hoja.Column(2).Style.Alignment.WrapText = true;
            hoja.Column(3).Style.Alignment.WrapText = true;

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
                $"Reporte_Centros_Costo_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";
            await RegistrarAuditoria(
    "Exportar Excel",
    "CentrosCosto",
    0,
    "Se exportó el reporte de centros de costo a Excel."
);

            return File(
                stream.ToArray(),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                nombreArchivo);
        }

        // DETAILS
        // Consulta la información completa del centro de costo.

        public async Task<IActionResult> Details(int? id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var centrosCosto = await _context.CentrosCostos
                .FirstOrDefaultAsync(c => c.IdCentroCosto == id);

            if (centrosCosto == null)
            {
                return NotFound();
            }

            return View(centrosCosto);
        }

    
        // CREATE GET
        // Carga el formulario para registrar un nuevo centro de costo.
    
        public IActionResult Create()
        {
            return View(new CentrosCosto { Activo = true });
        }

    
        // CREATE POST
        // Registra un centro de costo validando reglas de negocio.
    
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Create(
            [Bind("IdCentroCosto,Codigo,Nombre,Descripcion")]
            CentrosCosto centrosCosto)
        {
            NormalizarDatos(centrosCosto);
            await ValidarCentroCosto(centrosCosto);

            if (ModelState.IsValid)
            {
                centrosCosto.Activo = true;

                _context.CentrosCostos.Add(centrosCosto);
                await _context.SaveChangesAsync();

                await RegistrarAuditoria(
                    "Crear",
                    "CentrosCosto",
                    centrosCosto.IdCentroCosto,
                    $"Se creó el centro de costo: {centrosCosto.Codigo} - {centrosCosto.Nombre}.");

                TempData["MensajeExito"] = "Centro de costo creado correctamente.";
                return RedirectToAction(nameof(Index));
            }

            return View(centrosCosto);
        }

      
        // EDIT GET
        // Carga el centro de costo para modificación. Si tiene movimientos asociados, el código queda bloqueado.
     
        public async Task<IActionResult> Edit(int? id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var centrosCosto = await _context.CentrosCostos
                .FirstOrDefaultAsync(c => c.IdCentroCosto == id);

            if (centrosCosto == null)
            {
                return NotFound();
            }

            ViewBag.TieneMovimientosAsociados =
                await TieneRegistrosAsociados(centrosCosto.IdCentroCosto);

            return View(centrosCosto);
        }

        // EDIT POST
        // Modifica el centro de costo aplicando reglas de negocio. Si tiene movimientos asociados, no permite cambiar el código.
   
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Edit(
            int id,
            [Bind("IdCentroCosto,Codigo,Nombre,Descripcion,Activo")]
            CentrosCosto centrosCosto)
        {
            if (id != centrosCosto.IdCentroCosto)
            {
                return NotFound();
            }

            var centroActual = await _context.CentrosCostos
                .AsNoTracking()
                .FirstOrDefaultAsync(c => c.IdCentroCosto == id);

            if (centroActual == null)
            {
                return NotFound();
            }

            bool tieneRegistrosAsociados =
                await TieneRegistrosAsociados(id);

            if (!centroActual.Activo)
            {
                TempData["MensajeError"] =
                    "El centro de costo está inactivo. Para modificarlo primero debe activarlo.";

                return RedirectToAction(nameof(Index));
            }

            NormalizarDatos(centrosCosto);

            if (tieneRegistrosAsociados)
            {
                centrosCosto.Codigo = centroActual.Codigo;
            }

            await ValidarCentroCosto(centrosCosto);

            if (ModelState.IsValid)
            {
                try
                {
                    _context.CentrosCostos.Update(centrosCosto);
                    await _context.SaveChangesAsync();

                    await RegistrarAuditoria(
                        "Modificar",
                        "CentrosCosto",
                        centrosCosto.IdCentroCosto,
                        $"Se modificó el centro de costo: {centrosCosto.Codigo} - {centrosCosto.Nombre}.");

                    TempData["MensajeExito"] = "Centro de costo modificado correctamente.";
                    return RedirectToAction(nameof(Index));
                }
                catch (DbUpdateConcurrencyException)
                {
                    if (!CentrosCostoExists(centrosCosto.IdCentroCosto))
                    {
                        return NotFound();
                    }

                    throw;
                }
            }

            ViewBag.TieneMovimientosAsociados = tieneRegistrosAsociados;
            return View(centrosCosto);
        }

        // DELETE POST
        // Desactiva el centro de costo.
        // No elimina físicamente para conservar el historial.
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> DeleteConfirmed(int id)
        {
            var centrosCosto = await _context.CentrosCostos
                .FirstOrDefaultAsync(c => c.IdCentroCosto == id);

            if (centrosCosto == null)
            {
                return NotFound();
            }

            if (!centrosCosto.Activo)
            {
                TempData["MensajeError"] = "El centro de costo ya se encuentra inactivo.";
                return RedirectToAction(nameof(Index));
            }

            centrosCosto.Activo = false;

            _context.CentrosCostos.Update(centrosCosto);
            await _context.SaveChangesAsync();

            await RegistrarAuditoria(
                "Desactivar",
                "CentrosCosto",
                centrosCosto.IdCentroCosto,
                $"Se desactivó el centro de costo: {centrosCosto.Codigo} - {centrosCosto.Nombre}.");

            TempData["MensajeExito"] = "Centro de costo desactivado correctamente.";
            return RedirectToAction(nameof(Index));
        }

        // ACTIVAR
        // Reactiva un centro de costo inactivo.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Activar(int id)
        {
            var centrosCosto = await _context.CentrosCostos
                .FirstOrDefaultAsync(c => c.IdCentroCosto == id);

            if (centrosCosto == null)
            {
                return NotFound();
            }

            if (centrosCosto.Activo)
            {
                TempData["MensajeError"] = "El centro de costo ya se encuentra activo.";
                return RedirectToAction(nameof(Index));
            }

            centrosCosto.Activo = true;

            _context.CentrosCostos.Update(centrosCosto);
            await _context.SaveChangesAsync();

            await RegistrarAuditoria(
                "Activar",
                "CentrosCosto",
                centrosCosto.IdCentroCosto,
                $"Se activó el centro de costo: {centrosCosto.Codigo} - {centrosCosto.Nombre}.");

            TempData["MensajeExito"] = "Centro de costo activado correctamente.";
            return RedirectToAction(nameof(Index));
        }

        // MÉTODOS PRIVADOS DE APOYO CONSULTA COMPARTIDA PARA INDEX, REPORTE Y EXCEL
   
        private IQueryable<CentrosCosto> ConstruirConsultaReporte(
            string? buscar,
            string? estado)
        {
            var consulta = _context.CentrosCostos
                .Include(c => c.MovimientosContables)
                .AsNoTracking()
                .AsQueryable();

            if (!string.IsNullOrWhiteSpace(buscar))
            {
                buscar = buscar.Trim();

                consulta = consulta.Where(c =>
                    c.Codigo.Contains(buscar) ||
                    c.Nombre.Contains(buscar) ||
                    (c.Descripcion != null &&
                     c.Descripcion.Contains(buscar)));
            }

            if (!string.IsNullOrWhiteSpace(estado))
            {
                if (estado == "Activo")
                {
                    consulta = consulta.Where(c =>
                        c.Activo);
                }
                else if (estado == "Inactivo")
                {
                    consulta = consulta.Where(c =>
                        !c.Activo);
                }
            }

            return consulta;
        }
        private void NormalizarDatos(CentrosCosto centro)
        {
            centro.Codigo = centro.Codigo?.Trim().ToUpper() ?? "";
            centro.Nombre = centro.Nombre?.Trim() ?? "";
            centro.Descripcion = centro.Descripcion?.Trim();
        }

        private async Task ValidarCentroCosto(CentrosCosto centro)
        {
            if (string.IsNullOrWhiteSpace(centro.Codigo))
            {
                ModelState.AddModelError("Codigo", "El código es obligatorio.");
            }
            else if (centro.Codigo.Length > 50)
            {
                ModelState.AddModelError("Codigo", "El código no puede superar los 50 caracteres.");
            }
            else if (!System.Text.RegularExpressions.Regex.IsMatch(
                centro.Codigo,
                @"^[A-Z0-9-]+$"))
            {
                ModelState.AddModelError(
                    "Codigo",
                    "El código solo puede contener letras, números y guiones.");
            }

            bool codigoDuplicado = await _context.CentrosCostos
                .AnyAsync(c =>
                    c.Codigo.ToLower() == centro.Codigo.ToLower()
                    && c.IdCentroCosto != centro.IdCentroCosto);

            if (codigoDuplicado)
            {
                ModelState.AddModelError("Codigo", "Ya existe un centro de costo con ese código.");
            }

            if (string.IsNullOrWhiteSpace(centro.Nombre))
            {
                ModelState.AddModelError("Nombre", "El nombre es obligatorio.");
            }
            else if (centro.Nombre.Length < 3)
            {
                ModelState.AddModelError("Nombre", "El nombre debe tener al menos 3 caracteres.");
            }
            else if (centro.Nombre.Length > 150)
            {
                ModelState.AddModelError("Nombre", "El nombre no puede superar los 150 caracteres.");
            }
            else if (!System.Text.RegularExpressions.Regex.IsMatch(
                centro.Nombre,
                @"^[A-Za-zÁÉÍÓÚáéíóúÑñ0-9\s]+$"))
            {
                ModelState.AddModelError(
                    "Nombre",
                    "El nombre solo puede contener letras, números y espacios.");
            }

            bool nombreDuplicado = await _context.CentrosCostos
                .AnyAsync(c =>
                    c.Nombre.ToLower() == centro.Nombre.ToLower()
                    && c.IdCentroCosto != centro.IdCentroCosto);

            if (nombreDuplicado)
            {
                ModelState.AddModelError("Nombre", "Ya existe un centro de costo con ese nombre.");
            }

            if (!string.IsNullOrWhiteSpace(centro.Descripcion))
            {
                if (centro.Descripcion.Length > 300)
                {
                    ModelState.AddModelError(
                        "Descripcion",
                        "La descripción no puede superar los 300 caracteres.");
                }
                else if (!System.Text.RegularExpressions.Regex.IsMatch(
                    centro.Descripcion,
                    @"^[A-Za-zÁÉÍÓÚáéíóúÑñ0-9\s.,;:()/-]+$"))
                {
                    ModelState.AddModelError(
                        "Descripcion",
                        "La descripción contiene caracteres no permitidos.");
                }
                else if (!System.Text.RegularExpressions.Regex.IsMatch(
                    centro.Descripcion,
                    @"[A-Za-zÁÉÍÓÚáéíóúÑñ0-9]"))
                {
                    ModelState.AddModelError(
                        "Descripcion",
                        "La descripción debe contener al menos una letra o un número.");
                }
            }
        }

        private async Task<bool> TieneRegistrosAsociados(int idCentroCosto)
        {
            bool tieneMovimientos = await _context.MovimientosContables
                .AnyAsync(m => m.CentroCostoId == idCentroCosto);

            return tieneMovimientos;
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

        private bool CentrosCostoExists(int id)
        {
            return _context.CentrosCostos.Any(e => e.IdCentroCosto == id);
        }
    }
}