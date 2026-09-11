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

namespace SistemaGestionFinanciera.Controllers
{
    public class CategoriasIngresosController : BaseController
    {
        private readonly SistemaFinancieroContext _context;

        public CategoriasIngresosController(SistemaFinancieroContext context)
        {
            _context = context;
        }

        // INDEX - LISTADO DE CATEGORÍAS DE INGRESOS
        // ================================================================
        public async Task<IActionResult> Index(
            string? buscar,
            string? cuenta,
            string? estado)
        {
            var categorias = ConstruirConsultaReporte(
                buscar,
                cuenta,
                estado);

            ViewBag.Buscar = buscar;
            ViewBag.Cuenta = cuenta;
            ViewBag.Estado = estado;

            return View(await categorias
                .OrderByDescending(c => c.IdCategoriaIngreso)
                .ToListAsync());
        }
        // ================================================================
        // REPORTE DE CATEGORÍAS DE INGRESOS
        // ================================================================
        [HttpGet]
        public async Task<IActionResult> Reporte(
            string? buscar,
            string? cuenta,
            string? estado)
        {
            var categorias = await ConstruirConsultaReporte(
                    buscar,
                    cuenta,
                    estado)
                .OrderByDescending(c => c.IdCategoriaIngreso)
                .ToListAsync();

            ViewBag.TotalRegistros =
                categorias.Count;

            ViewBag.TotalActivas =
                categorias.Count(c => c.Activo);

            ViewBag.TotalInactivas =
                categorias.Count(c => !c.Activo);

            ViewBag.TotalConCuenta =
                categorias.Count(c => c.CuentaContable != null);

            ViewBag.Buscar = buscar;
            ViewBag.Cuenta = cuenta;
            ViewBag.Estado = estado;
            ViewBag.FechaGeneracion = DateTime.Now;
            await RegistrarAuditoria(
    "Vista previa PDF",
    "CategoriasIngresos",
    0,
    "Se generó la vista previa del reporte de categorías de ingresos."
);

            return View(categorias);
        }
        // ================================================================
        // EXPORTAR REPORTE DE CATEGORÍAS DE INGRESOS A EXCEL
        // ================================================================
        [HttpGet]
        public async Task<IActionResult> ExportarExcel(
            string? buscar,
            string? cuenta,
            string? estado)
        {
            var categorias = await ConstruirConsultaReporte(
                    buscar,
                    cuenta,
                    estado)
                .OrderByDescending(c => c.IdCategoriaIngreso)
                .ToListAsync();

            using var workbook = new XLWorkbook();

            var hoja = workbook.Worksheets.Add(
                "Categorías de ingresos");

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
                "Reporte de Categorías de Ingresos";

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

            hoja.Cell("C6").Value = "Cuenta:";

            hoja.Cell("D6").Value =
                string.IsNullOrWhiteSpace(cuenta)
                    ? "Todas"
                    : cuenta;

            hoja.Cell("A7").Value = "Estado:";

            hoja.Cell("B7").Value =
                string.IsNullOrWhiteSpace(estado)
                    ? "Todos"
                    : estado;

            hoja.Range("C7:D7").Merge();

            hoja.Cell("A6").Style.Font.Bold = true;
            hoja.Cell("C6").Style.Font.Bold = true;
            hoja.Cell("A7").Style.Font.Bold = true;

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

            hoja.Row(6).Height = 22;
            hoja.Row(7).Height = 22;

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

            hoja.Cell("A10").Value = "Total";
            hoja.Cell("B10").Value = categorias.Count;

            hoja.Cell("C10").Value = "Activas";
            hoja.Cell("D10").Value =
                categorias.Count(c => c.Activo);

            hoja.Cell("A11").Value = "Inactivas";
            hoja.Cell("B11").Value =
                categorias.Count(c => !c.Activo);

            hoja.Cell("C11").Value =
                "Con cuenta contable";

            hoja.Cell("D11").Value =
                categorias.Count(c =>
                    c.CuentaContable != null);

            hoja.Range("A10:B10")
                .Style.Fill.BackgroundColor =
                colorAzulSuave;

            hoja.Range("C10:D10")
                .Style.Fill.BackgroundColor =
                colorVerdeSuave;

            hoja.Range("A11:B11")
                .Style.Fill.BackgroundColor =
                colorGrisSuave;

            hoja.Range("C11:D11")
                .Style.Fill.BackgroundColor =
                colorDoradoSuave;

            hoja.Range("A10:D11").Style.Font.Bold = true;

            hoja.Range("A10:D11")
                .Style.Border.OutsideBorder =
                XLBorderStyleValues.Thin;

            hoja.Range("A10:D11")
                .Style.Border.InsideBorder =
                XLBorderStyleValues.Thin;

            hoja.Range("A10:D11")
                .Style.Alignment.Vertical =
                XLAlignmentVerticalValues.Center;

            hoja.Cell("B10")
                .Style.Alignment.Horizontal =
                XLAlignmentHorizontalValues.Right;

            hoja.Cell("D10")
                .Style.Alignment.Horizontal =
                XLAlignmentHorizontalValues.Right;

            hoja.Cell("B11")
                .Style.Alignment.Horizontal =
                XLAlignmentHorizontalValues.Right;

            hoja.Cell("D11")
                .Style.Alignment.Horizontal =
                XLAlignmentHorizontalValues.Right;

            hoja.Row(10).Height = 22;
            hoja.Row(11).Height = 22;

            // ============================================================
            // ENCABEZADOS DE LA TABLA
            // ============================================================
            const int filaEncabezado = 13;

            string[] encabezados =
            {
        "Nombre",
        "Descripción",
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

            foreach (var categoria in categorias)
            {
                string cuentaContable =
                    categoria.CuentaContable != null
                        ? $"{categoria.CuentaContable.Codigo} - " +
                          $"{categoria.CuentaContable.Nombre}"
                        : "Sin cuenta asociada";

                hoja.Cell(fila, 1).Value =
                    categoria.Nombre;

                hoja.Cell(fila, 2).Value =
                    string.IsNullOrWhiteSpace(categoria.Descripcion)
                        ? "Sin descripción"
                        : categoria.Descripcion;

                hoja.Cell(fila, 3).Value =
                    cuentaContable;

                hoja.Cell(fila, 4).Value =
                    categoria.Activo
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

                if (categoria.Activo)
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
            // TABLA
            // ============================================================
            if (categorias.Any())
            {
                int filaFinDatos = fila - 1;

                var rangoTabla = hoja.Range(
                    filaEncabezado,
                    1,
                    filaFinDatos,
                    totalColumnas);

                var tabla = rangoTabla.CreateTable(
                    "TablaCategoriasIngresos");

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
                        4,
                        filaFinDatos,
                        4)
                    .Style.Alignment.Horizontal =
                    XLAlignmentHorizontalValues.Center;

                hoja.Range(
                        filaEncabezado + 1,
                        1,
                        filaFinDatos,
                        3)
                    .Style.Alignment.WrapText = true;
            }
            else
            {
                hoja.Cell(fila, 1).Value =
                    "No se encontraron categorías de ingresos con los filtros aplicados.";

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
                categorias.Count;

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
            hoja.Column(1).Width = 30;
            hoja.Column(2).Width = 48;
            hoja.Column(3).Width = 45;
            hoja.Column(4).Width = 14;

            hoja.Column(1).Style.Alignment.WrapText = true;
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
                $"Reporte_Categorias_Ingresos_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";
            await RegistrarAuditoria(
    "Exportar Excel",
    "CategoriasIngresos",
    0,
    "Se exportó el reporte de categorías de ingresos a Excel."
);
            return File(
                stream.ToArray(),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                nombreArchivo);
        }
        // DETAILS
        // Consulta la información completa de la categoría.

        public async Task<IActionResult> Details(int? id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var categoriasIngreso = await _context.CategoriasIngresos
                .Include(c => c.CuentaContable)
                .FirstOrDefaultAsync(c => c.IdCategoriaIngreso == id);

            if (categoriasIngreso == null)
            {
                return NotFound();
            }

            return View(categoriasIngreso);
        }

      
        // CREATE GET
        // Carga el formulario para crear una nueva categoría.
        // Solo permite seleccionar cuentas contables activas.
      
        public IActionResult Create()
        {
            CargarCuentasContables();
            return View();
        }

        // CREATE POST
        // Crea una categoría de ingreso validando reglas de negocio.
        // ============================================================
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Create(
            [Bind("IdCategoriaIngreso,Nombre,Descripcion,CuentaContableId")]
            CategoriasIngreso categoriasIngreso)
        {
            NormalizarDatos(categoriasIngreso);
            await ValidarCategoriaIngreso(categoriasIngreso, esEdicion: false);

            if (ModelState.IsValid)
            {
                categoriasIngreso.Activo = true;

                _context.CategoriasIngresos.Add(categoriasIngreso);
                await _context.SaveChangesAsync();
                //auditoria
                await RegistrarAuditoria(
    "Crear",
    "CategoriasIngresos",
    categoriasIngreso.IdCategoriaIngreso,
    $"Se creó la categoría de ingreso: {categoriasIngreso.Nombre}.");

                TempData["MensajeExito"] = "Categoría de ingreso creada correctamente.";
                return RedirectToAction(nameof(Index));
            }

            CargarCuentasContables(categoriasIngreso.CuentaContableId);
            return View(categoriasIngreso);
        }

    
        // EDIT GET
        // Carga la categoría para modificación.
        // Si ya tiene ingresos asociados, la cuenta contable queda bloqueada.
   
        public async Task<IActionResult> Edit(int? id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var categoriasIngreso = await _context.CategoriasIngresos
                .FirstOrDefaultAsync(c => c.IdCategoriaIngreso == id);

            if (categoriasIngreso == null)
            {
                return NotFound();
            }

            ViewBag.TieneMovimientosAsociados =
                await TieneRegistrosAsociados(categoriasIngreso.IdCategoriaIngreso);

            CargarCuentasContables(categoriasIngreso.CuentaContableId);
            return View(categoriasIngreso);
        }

 
        // EDIT POST
        // Modifica la categoría aplicando restricciones de negocio.
        // Si tiene ingresos asociados, no permite cambiar la cuenta contable.
        // ============================================================
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Edit(
            int id,
            [Bind("IdCategoriaIngreso,Nombre,Descripcion,CuentaContableId,Activo")]
            CategoriasIngreso categoriasIngreso)
        {
            if (id != categoriasIngreso.IdCategoriaIngreso)
            {
                return NotFound();
            }

            var categoriaActual = await _context.CategoriasIngresos
                .AsNoTracking()
                .FirstOrDefaultAsync(c => c.IdCategoriaIngreso == id);

            if (categoriaActual == null)
            {
                return NotFound();
            }

            bool tieneRegistrosAsociados =
                await TieneRegistrosAsociados(id);
            // Si la categoría está inactiva primero debe activarse.
            if (!categoriaActual.Activo)
            {
                TempData["MensajeError"] =
                    "La categoría de ingreso está inactiva. Para modificarla primero debe activarla.";

                return RedirectToAction(nameof(Index));
            }

            // Si ya existen ingresos o presupuesto asociados,
            // la cuenta contable queda protegida y conserva la original.
            if (tieneRegistrosAsociados)
            {
                categoriasIngreso.CuentaContableId = categoriaActual.CuentaContableId;
            }

            NormalizarDatos(categoriasIngreso);

            await ValidarCategoriaIngreso(categoriasIngreso, esEdicion: true);

            if (ModelState.IsValid)
            {
                try
                {
                    _context.CategoriasIngresos.Update(categoriasIngreso);
                    await _context.SaveChangesAsync();
                    //auditoria
                    await RegistrarAuditoria(
     "Modificar",
     "CategoriasIngresos",
     categoriasIngreso.IdCategoriaIngreso,
     $"Se modificó la categoría de ingreso: {categoriasIngreso.Nombre}.");

                    TempData["MensajeExito"] = "Categoría de ingreso modificada correctamente.";
                    return RedirectToAction(nameof(Index));
                }
                catch (DbUpdateConcurrencyException)
                {
                    if (!CategoriasIngresoExists(categoriasIngreso.IdCategoriaIngreso))
                    {
                        return NotFound();
                    }

                    throw;
                }
            }

            ViewBag.TieneMovimientosAsociados = tieneRegistrosAsociados;
            CargarCuentasContables(categoriasIngreso.CuentaContableId);
            return View(categoriasIngreso);
        }

       
        // DELETE GET
        // No elimina físicamente.
        // Solo muestra confirmación para desactivar.
      
        public async Task<IActionResult> Delete(int? id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var categoriasIngreso = await _context.CategoriasIngresos
                .Include(c => c.CuentaContable)
                .FirstOrDefaultAsync(c => c.IdCategoriaIngreso == id);

            if (categoriasIngreso == null)
            {
                return NotFound();
            }

            return View(categoriasIngreso);
        }

    
        // DELETE POST
        // Desactiva la categoría.
        // No se elimina físicamente para conservar historial.
        // ============================================================
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> DeleteConfirmed(int id)
        {
            var categoriasIngreso = await _context.CategoriasIngresos
                .FirstOrDefaultAsync(c => c.IdCategoriaIngreso == id);

            if (categoriasIngreso == null)
            {
                return NotFound();
            }

            if (!categoriasIngreso.Activo)
            {
                TempData["MensajeError"] = "La categoría de ingreso ya se encuentra inactiva.";
                return RedirectToAction(nameof(Index));
            }

            categoriasIngreso.Activo = false;
            _context.CategoriasIngresos.Update(categoriasIngreso);
            await _context.SaveChangesAsync();
            //AUDITORIA
            await RegistrarAuditoria(
    "Desactivar",
    "CategoriasIngresos",
    categoriasIngreso.IdCategoriaIngreso,
    $"Se desactivó la categoría de ingreso: {categoriasIngreso.Nombre}.");

            TempData["MensajeExito"] = "Categoría de ingreso desactivada correctamente.";
            return RedirectToAction(nameof(Index));
        }

        // ACTIVAR
        // Reactiva una categoría previamente desactivada.
      
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Activar(int id)
        {
            var categoriasIngreso = await _context.CategoriasIngresos
                .FirstOrDefaultAsync(c => c.IdCategoriaIngreso == id);

            if (categoriasIngreso == null)
            {
                return NotFound();
            }

            if (categoriasIngreso.Activo)
            {
                TempData["MensajeError"] = "La categoría de ingreso ya se encuentra activa.";
                return RedirectToAction(nameof(Index));
            }

            categoriasIngreso.Activo = true;
            _context.CategoriasIngresos.Update(categoriasIngreso);
            await _context.SaveChangesAsync();
            //auditoria
            await RegistrarAuditoria(
    "Activar",
    "CategoriasIngresos",
    categoriasIngreso.IdCategoriaIngreso,
    $"Se activó la categoría de ingreso: {categoriasIngreso.Nombre}.");

            TempData["MensajeExito"] = "Categoría de ingreso activada correctamente.";
            return RedirectToAction(nameof(Index));
        }


        // MÉTODOS PRIVADOS DE APOYO
        // ================================================================
        // CONSULTA COMPARTIDA PARA INDEX, REPORTE Y EXCEL
        // ================================================================
        private IQueryable<CategoriasIngreso> ConstruirConsultaReporte(
            string? buscar,
            string? cuenta,
            string? estado)
        {
            var consulta = _context.CategoriasIngresos
                .Include(c => c.CuentaContable)
                .AsNoTracking()
                .AsQueryable();

            if (!string.IsNullOrWhiteSpace(buscar))
            {
                buscar = buscar.Trim();

                consulta = consulta.Where(c =>
                    c.Nombre.Contains(buscar) ||
                    (
                        c.Descripcion != null &&
                        c.Descripcion.Contains(buscar)
                    ));
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

        private void NormalizarDatos(CategoriasIngreso categoria)
        {
            categoria.Nombre = categoria.Nombre?.Trim() ?? "";
            categoria.Descripcion = categoria.Descripcion?.Trim();
        }

        private async Task ValidarCategoriaIngreso(
            CategoriasIngreso categoria,
            bool esEdicion)
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

            bool nombreDuplicado = await _context.CategoriasIngresos
                .AnyAsync(c =>
                    c.Nombre.ToLower() == categoria.Nombre.ToLower()
                    && c.IdCategoriaIngreso != categoria.IdCategoriaIngreso);

            if (nombreDuplicado)
            {
                ModelState.AddModelError("Nombre", "Ya existe una categoría de ingreso con ese nombre.");
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
                    @"^[A-Za-zÁÉÍÓÚáéíóúÑñ0-9\s.,;:()/-]+$"))
                {
                    ModelState.AddModelError("Descripcion", "La descripción contiene caracteres no permitidos.");
                }
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

        private async Task<bool> TieneRegistrosAsociados(int idCategoriaIngreso)
        {
            bool tieneIngresos = await _context.Ingresos
                .AnyAsync(i => i.CategoriaIngresoId == idCategoriaIngreso);

            bool tienePresupuesto = await _context.PresupuestoDetalles
                .AnyAsync(p => p.CategoriaIngresoId == idCategoriaIngreso);

            return tieneIngresos || tienePresupuesto;
        }

        private void CargarCuentasContables(int? cuentaSeleccionada = null)
        {
            ViewData["CuentaContableId"] = new SelectList(
                _context.CuentasContables
                    .Where(c => c.Activo)
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

        private bool CategoriasIngresoExists(int id)
        {
            return _context.CategoriasIngresos.Any(e => e.IdCategoriaIngreso == id);
        }
    }
}