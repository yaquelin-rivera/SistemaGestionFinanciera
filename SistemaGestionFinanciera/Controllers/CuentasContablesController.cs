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
using System.Collections.Generic;

namespace SistemaGestionFinanciera.Controllers
{
    public class CuentasContablesController : BaseController
    {
        private readonly SistemaFinancieroContext _context;

        public CuentasContablesController(SistemaFinancieroContext context)
        {
            _context = context;
        }

        // GET: CuentasContables
        public async Task<IActionResult> Index(string codigo, string nombre, string tipoCuenta, string estado)
        {
            var cuentas = _context.CuentasContables.AsQueryable();

            if (!string.IsNullOrWhiteSpace(codigo))
                cuentas = cuentas.Where(c => c.Codigo.Contains(codigo));

            if (!string.IsNullOrWhiteSpace(nombre))
                cuentas = cuentas.Where(c => c.Nombre.Contains(nombre));

            if (!string.IsNullOrWhiteSpace(tipoCuenta))
                cuentas = cuentas.Where(c => c.TipoCuenta == tipoCuenta);

            if (!string.IsNullOrWhiteSpace(estado))
            {
                if (estado == "Activa")
                    cuentas = cuentas.Where(c => c.Activo);
                else if (estado == "Inactiva")
                    cuentas = cuentas.Where(c => !c.Activo);
            }

            ViewBag.Codigo = codigo;
            ViewBag.Nombre = nombre;
            ViewBag.TipoCuenta = tipoCuenta;
            ViewBag.Estado = estado;

            decimal totalActivos = await _context.CuentasContables
     .Where(c => c.TipoCuenta == "Activo" && c.Activo)
     .SumAsync(c => c.SaldoActual);

            decimal totalPasivos = await _context.CuentasContables
                .Where(c => c.TipoCuenta == "Pasivo" && c.Activo)
                .SumAsync(c => c.SaldoActual);

            decimal totalIngresos = await _context.CuentasContables
                .Where(c => c.TipoCuenta == "Ingreso" && c.Activo)
                .SumAsync(c => c.SaldoActual);

            decimal totalGastos = await _context.CuentasContables
                .Where(c => c.TipoCuenta == "Gasto" && c.Activo)
                .SumAsync(c => c.SaldoActual);

            // Patrimonio financiero general:
            decimal totalPatrimonio = totalActivos - totalPasivos;

            // Resultado del periodo:
            decimal resultadoPeriodo = totalIngresos - totalGastos;

            ViewBag.TotalActivos = totalActivos;
            ViewBag.TotalPasivos = totalPasivos;
            ViewBag.TotalPatrimonio = totalPatrimonio;
            ViewBag.TotalIngresos = totalIngresos;
            ViewBag.TotalGastos = totalGastos;
            ViewBag.ResultadoPeriodo = resultadoPeriodo;

            return View(await cuentas.OrderBy(c => c.Codigo).ToListAsync());
        }
        // ===========================================================
        // REPORTE DE CUENTAS CONTABLES
        // ===========================================================

        public async Task<IActionResult> Reporte(
            string? codigo,
            string? nombre,
            string? tipoCuenta,
            string? estado)
        {
            var consulta = ConstruirConsultaReporte(
                codigo,
                nombre,
                tipoCuenta,
                estado);

            var cuentas = await consulta
                .OrderBy(c => c.Codigo)
                .ToListAsync();

            // Solo las cuentas activas participan en los totales
            // financieros del reporte.
            var cuentasActivas = cuentas
                .Where(c => c.Activo)
                .ToList();

            decimal totalActivos = cuentasActivas
                .Where(c => c.TipoCuenta == "Activo")
                .Sum(c => c.SaldoActual);

            decimal totalPasivos = cuentasActivas
                .Where(c => c.TipoCuenta == "Pasivo")
                .Sum(c => c.SaldoActual);

            decimal totalPatrimonio =
                totalActivos - totalPasivos;

            decimal totalIngresos = cuentasActivas
                .Where(c => c.TipoCuenta == "Ingreso")
                .Sum(c => c.SaldoActual);

            decimal totalGastos = cuentasActivas
                .Where(c => c.TipoCuenta == "Gasto")
                .Sum(c => c.SaldoActual);

            decimal resultadoPeriodo =
                totalIngresos - totalGastos;

            ViewBag.TotalRegistros = cuentas.Count;

            ViewBag.CuentasActivas =
                cuentas.Count(c => c.Activo);

            ViewBag.CuentasInactivas =
                cuentas.Count(c => !c.Activo);

            ViewBag.TotalActivos = totalActivos;
            ViewBag.TotalPasivos = totalPasivos;
            ViewBag.TotalPatrimonio = totalPatrimonio;
            ViewBag.TotalIngresos = totalIngresos;
            ViewBag.TotalGastos = totalGastos;
            ViewBag.ResultadoPeriodo = resultadoPeriodo;

            ViewBag.Codigo = codigo;
            ViewBag.Nombre = nombre;
            ViewBag.TipoCuenta = tipoCuenta;
            ViewBag.Estado = estado;

            ViewBag.FechaGeneracion = DateTime.Now;
            await RegistrarAuditoria(
    "CuentasContables",
    0,
    "Vista previa PDF",
    "Se generó la vista previa del reporte de cuentas contables."
);
            return View(cuentas);
        }
        // ===========================================================
        // EXPORTAR REPORTE DE CUENTAS CONTABLES A EXCEL
        // ===========================================================

        public async Task<IActionResult> ExportarExcel(
            string? codigo,
            string? nombre,
            string? tipoCuenta,
            string? estado)
        {
            var consulta = ConstruirConsultaReporte(
                codigo,
                nombre,
                tipoCuenta,
                estado);

            var cuentas = await consulta
                .OrderBy(c => c.Codigo)
                .ToListAsync();

            using var workbook = new XLWorkbook();

            var hoja = workbook.Worksheets.Add(
                "Cuentas Contables");

            const int totalColumnas = 8;

            // =======================================================
            // ENCABEZADO
            // =======================================================

            hoja.Cell("A1").Value =
                "Unión Cantonal de Asociaciones de Desarrollo de Santa Ana";

            hoja.Range(1, 1, 1, totalColumnas).Merge();

            hoja.Cell("A2").Value =
                "Reporte de Cuentas Contables";

            hoja.Range(2, 1, 2, totalColumnas).Merge();

            hoja.Cell("A3").Value =
                $"Generado el {DateTime.Now:dd/MM/yyyy HH:mm}";

            hoja.Range(3, 1, 3, totalColumnas).Merge();

            var encabezadoInstitucion = hoja.Range(
                1,
                1,
                1,
                totalColumnas);

            encabezadoInstitucion.Style.Font.Bold = true;
            encabezadoInstitucion.Style.Font.FontSize = 16;

            encabezadoInstitucion.Style.Font.FontColor =
                XLColor.FromHtml("#0F5C64");

            encabezadoInstitucion.Style.Alignment.Horizontal =
                XLAlignmentHorizontalValues.Center;

            var tituloReporte = hoja.Range(
                2,
                1,
                2,
                totalColumnas);

            tituloReporte.Style.Font.Bold = true;
            tituloReporte.Style.Font.FontSize = 13;

            tituloReporte.Style.Alignment.Horizontal =
                XLAlignmentHorizontalValues.Center;

            var fechaReporte = hoja.Range(
                3,
                1,
                3,
                totalColumnas);

            fechaReporte.Style.Alignment.Horizontal =
                XLAlignmentHorizontalValues.Center;

            fechaReporte.Style.Font.FontColor =
                XLColor.FromHtml("#666666");

            // =======================================================
            // FILTROS APLICADOS
            // =======================================================

            var filtros = new List<string>();

            if (!string.IsNullOrWhiteSpace(codigo))
            {
                filtros.Add($"Código: {codigo}");
            }

            if (!string.IsNullOrWhiteSpace(nombre))
            {
                filtros.Add($"Nombre: {nombre}");
            }

            if (!string.IsNullOrWhiteSpace(tipoCuenta))
            {
                filtros.Add($"Tipo: {tipoCuenta}");
            }

            if (!string.IsNullOrWhiteSpace(estado))
            {
                filtros.Add($"Estado: {estado}");
            }

            hoja.Cell("A5").Value = "Filtros aplicados:";
            hoja.Cell("A5").Style.Font.Bold = true;

            hoja.Cell("B5").Value =
                filtros.Any()
                    ? string.Join(" | ", filtros)
                    : "Sin filtros. Se muestran todas las cuentas contables.";

            hoja.Range(5, 2, 5, totalColumnas).Merge();

            var rangoFiltros = hoja.Range(
                5,
                1,
                5,
                totalColumnas);

            rangoFiltros.Style.Fill.BackgroundColor =
                XLColor.FromHtml("#F2F7F8");

            rangoFiltros.Style.Border.LeftBorder =
                XLBorderStyleValues.Medium;

            rangoFiltros.Style.Border.LeftBorderColor =
                XLColor.FromHtml("#0F7C90");

            rangoFiltros.Style.Alignment.Vertical =
                XLAlignmentVerticalValues.Center;

            rangoFiltros.Style.Alignment.WrapText = true;

            hoja.Row(5).Height = 22;

            // =======================================================
            // ENCABEZADOS DE LA TABLA
            // =======================================================

            const int filaEncabezado = 7;

            string[] encabezados =
            {
        "Código",
        "Nombre",
        "Tipo de cuenta",
        "Naturaleza",
        "Saldo inicial",
        "Saldo actual",
        "Estado",
        "Fecha de creación"
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

            var rangoEncabezado = hoja.Range(
                filaEncabezado,
                1,
                filaEncabezado,
                totalColumnas);

            rangoEncabezado.Style.Font.Bold = true;
            rangoEncabezado.Style.Font.FontColor = XLColor.White;

            rangoEncabezado.Style.Fill.BackgroundColor =
                XLColor.FromHtml("#0F5C64");

            rangoEncabezado.Style.Alignment.Horizontal =
                XLAlignmentHorizontalValues.Center;

            rangoEncabezado.Style.Alignment.Vertical =
                XLAlignmentVerticalValues.Center;

            // =======================================================
            // DATOS
            // =======================================================

            int fila = filaEncabezado + 1;

            foreach (var cuenta in cuentas)
            {
                hoja.Cell(fila, 1).Value =
                    cuenta.Codigo;

                hoja.Cell(fila, 2).Value =
                    cuenta.Nombre;

                hoja.Cell(fila, 3).Value =
                    cuenta.TipoCuenta;

                hoja.Cell(fila, 4).Value =
                    cuenta.Naturaleza;

                hoja.Cell(fila, 5).Value =
                    cuenta.SaldoInicial;

                hoja.Cell(fila, 6).Value =
                    cuenta.SaldoActual;

                hoja.Cell(fila, 7).Value =
                    cuenta.Activo
                        ? "Activa"
                        : "Inactiva";

                if (cuenta.FechaCreacion.HasValue)
                {
                    hoja.Cell(fila, 8).Value =
                        cuenta.FechaCreacion.Value;
                }

                fila++;
            }

            // =======================================================
            // FORMATO DE LA TABLA
            // =======================================================

            if (cuentas.Any())
            {
                int filaInicioDatos =
                    filaEncabezado + 1;

                int filaFinDatos =
                    fila - 1;

                var rangoTabla = hoja.Range(
                    filaEncabezado,
                    1,
                    filaFinDatos,
                    totalColumnas);

                var tabla = rangoTabla.CreateTable(
                    "TablaCuentasContables");

                tabla.Theme =
                    XLTableTheme.TableStyleMedium2;

                tabla.ShowAutoFilter = true;
                tabla.ShowRowStripes = false;

                var rangoDatos = hoja.Range(
                    filaInicioDatos,
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
                        filaInicioDatos,
                        5,
                        filaFinDatos,
                        6)
                    .Style.NumberFormat.Format =
                        "₡ #,##0.00";

                hoja.Range(
                        filaInicioDatos,
                        5,
                        filaFinDatos,
                        6)
                    .Style.Alignment.Horizontal =
                        XLAlignmentHorizontalValues.Right;

                hoja.Range(
                        filaInicioDatos,
                        7,
                        filaFinDatos,
                        8)
                    .Style.Alignment.Horizontal =
                        XLAlignmentHorizontalValues.Center;

                hoja.Range(
                        filaInicioDatos,
                        8,
                        filaFinDatos,
                        8)
                    .Style.DateFormat.Format =
                        "dd/MM/yyyy";

                // Cuentas inactivas en gris.
                for (int indice = 0;
                     indice < cuentas.Count;
                     indice++)
                {
                    var cuenta = cuentas[indice];

                    if (!cuenta.Activo)
                    {
                        int filaCuenta =
                            filaInicioDatos + indice;

                        var rangoFila = hoja.Range(
                            filaCuenta,
                            1,
                            filaCuenta,
                            totalColumnas);

                        rangoFila.Style.Fill.BackgroundColor =
                            XLColor.FromHtml("#E9ECEF");

                        rangoFila.Style.Font.FontColor =
                            XLColor.FromHtml("#495057");
                    }
                }
            }

            // =======================================================
            // TOTALES FINANCIEROS
            // =======================================================

            var cuentasActivas = cuentas
                .Where(c => c.Activo)
                .ToList();

            decimal totalActivos = cuentasActivas
                .Where(c => c.TipoCuenta == "Activo")
                .Sum(c => c.SaldoActual);

            decimal totalPasivos = cuentasActivas
                .Where(c => c.TipoCuenta == "Pasivo")
                .Sum(c => c.SaldoActual);

            decimal totalPatrimonio =
                totalActivos - totalPasivos;

            decimal totalIngresos = cuentasActivas
                .Where(c => c.TipoCuenta == "Ingreso")
                .Sum(c => c.SaldoActual);

            decimal totalGastos = cuentasActivas
                .Where(c => c.TipoCuenta == "Gasto")
                .Sum(c => c.SaldoActual);

            decimal resultadoPeriodo =
                totalIngresos - totalGastos;

            int filaResumen = fila + 2;

            hoja.Cell(filaResumen, 1).Value =
                "RESUMEN FINANCIERO";

            hoja.Range(
                    filaResumen,
                    1,
                    filaResumen,
                    6)
                .Merge();

            var tituloResumen = hoja.Range(
                filaResumen,
                1,
                filaResumen,
                6);

            tituloResumen.Style.Fill.BackgroundColor =
                XLColor.FromHtml("#0F5C64");

            tituloResumen.Style.Font.Bold = true;
            tituloResumen.Style.Font.FontColor = XLColor.White;

            tituloResumen.Style.Alignment.Horizontal =
                XLAlignmentHorizontalValues.Center;

            string[] nombresResumen =
            {
        "Activos",
        "Pasivos",
        "Patrimonio",
        "Ingresos",
        "Gastos",
        "Resultado"
    };

            decimal[] valoresResumen =
            {
        totalActivos,
        totalPasivos,
        totalPatrimonio,
        totalIngresos,
        totalGastos,
        resultadoPeriodo
    };

            for (int columna = 0;
                 columna < nombresResumen.Length;
                 columna++)
            {
                hoja.Cell(
                    filaResumen + 1,
                    columna + 1).Value =
                        nombresResumen[columna];

                hoja.Cell(
                    filaResumen + 2,
                    columna + 1).Value =
                        valoresResumen[columna];
            }

            var encabezadosResumen = hoja.Range(
                filaResumen + 1,
                1,
                filaResumen + 1,
                6);

            encabezadosResumen.Style.Fill.BackgroundColor =
                XLColor.FromHtml("#D9EDEF");

            encabezadosResumen.Style.Font.Bold = true;

            encabezadosResumen.Style.Alignment.Horizontal =
                XLAlignmentHorizontalValues.Center;

            var montosResumen = hoja.Range(
                filaResumen + 2,
                1,
                filaResumen + 2,
                6);

            montosResumen.Style.NumberFormat.Format =
                "₡ #,##0.00";

            montosResumen.Style.Font.Bold = true;

            montosResumen.Style.Alignment.Horizontal =
                XLAlignmentHorizontalValues.Center;

            var rangoResumen = hoja.Range(
                filaResumen + 1,
                1,
                filaResumen + 2,
                6);

            rangoResumen.Style.Border.OutsideBorder =
                XLBorderStyleValues.Thin;

            rangoResumen.Style.Border.InsideBorder =
                XLBorderStyleValues.Thin;

            // =======================================================
            // RESUMEN DE ESTADOS
            // =======================================================

            int filaEstados = filaResumen + 5;

            hoja.Cell(filaEstados, 1).Value =
                "RESUMEN DE CUENTAS";

            hoja.Range(
                    filaEstados,
                    1,
                    filaEstados,
                    3)
                .Merge();

            var tituloEstados = hoja.Range(
                filaEstados,
                1,
                filaEstados,
                3);

            tituloEstados.Style.Fill.BackgroundColor =
                XLColor.FromHtml("#0F5C64");

            tituloEstados.Style.Font.Bold = true;
            tituloEstados.Style.Font.FontColor = XLColor.White;

            tituloEstados.Style.Alignment.Horizontal =
                XLAlignmentHorizontalValues.Center;

            hoja.Cell(filaEstados + 1, 1).Value =
                "Activas";

            hoja.Cell(filaEstados + 1, 2).Value =
                "Inactivas";

            hoja.Cell(filaEstados + 1, 3).Value =
                "Total";

            hoja.Cell(filaEstados + 2, 1).Value =
                cuentas.Count(c => c.Activo);

            hoja.Cell(filaEstados + 2, 2).Value =
                cuentas.Count(c => !c.Activo);

            hoja.Cell(filaEstados + 2, 3).Value =
                cuentas.Count;

            var encabezadosEstados = hoja.Range(
                filaEstados + 1,
                1,
                filaEstados + 1,
                3);

            encabezadosEstados.Style.Fill.BackgroundColor =
                XLColor.FromHtml("#D9EDEF");

            encabezadosEstados.Style.Font.Bold = true;

            var rangoEstados = hoja.Range(
                filaEstados + 1,
                1,
                filaEstados + 2,
                3);

            rangoEstados.Style.Alignment.Horizontal =
                XLAlignmentHorizontalValues.Center;

            rangoEstados.Style.Border.OutsideBorder =
                XLBorderStyleValues.Thin;

            rangoEstados.Style.Border.InsideBorder =
                XLBorderStyleValues.Thin;

            // =======================================================
            // ANCHOS Y CONFIGURACIÓN FINAL
            // =======================================================

            hoja.Columns().AdjustToContents();

            hoja.Column(1).Width = 14;
            hoja.Column(2).Width = 34;
            hoja.Column(3).Width = 18;
            hoja.Column(4).Width = 16;
            hoja.Column(5).Width = 18;
            hoja.Column(6).Width = 18;
            hoja.Column(7).Width = 14;
            hoja.Column(8).Width = 18;

            hoja.Column(2).Style.Alignment.WrapText = true;

            hoja.SheetView.FreezeRows(filaEncabezado);

            hoja.PageSetup.PageOrientation =
                XLPageOrientation.Landscape;

            hoja.PageSetup.FitToPages(1, 0);

            hoja.PageSetup.SetRowsToRepeatAtTop(
                filaEncabezado,
                filaEncabezado);

            hoja.PageSetup.CenterHorizontally = true;

            using var stream = new MemoryStream();

            workbook.SaveAs(stream);

            string nombreArchivo =
                $"CuentasContables_{DateTime.Now:yyyyMMdd_HHmm}.xlsx";
            await RegistrarAuditoria(
    "CuentasContables",
    0,
    "Exportar Excel",
    "Se exportó el reporte de cuentas contables a Excel."
);

            return File(
                stream.ToArray(),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                nombreArchivo);
        }
        // GET: CuentasContables/Details/5
        public async Task<IActionResult> Details(int? id)
        {
            if (id == null)
                return NotFound();

            var cuenta = await _context.CuentasContables
                .FirstOrDefaultAsync(c => c.IdCuentaContable == id);

            if (cuenta == null)
                return NotFound();

            var movimientos = await _context.MovimientosContables
                .Include(m => m.Proyecto)
                .Include(m => m.CentroCosto)
                .Where(m => m.CuentaContableId == id)
                .OrderByDescending(m => m.Fecha)
                .ThenByDescending(m => m.IdMovimientoContable)
                .Take(5)
                .ToListAsync();

            ViewBag.TotalDebe = await _context.MovimientosContables
      .Where(m => m.CuentaContableId == id
               && m.Anulado != true
               && m.Estado == "Registrado")
      .SumAsync(m => m.Debe);

            ViewBag.TotalHaber = await _context.MovimientosContables
                .Where(m => m.CuentaContableId == id
                         && m.Anulado != true
                         && m.Estado == "Registrado")
                .SumAsync(m => m.Haber);

            ViewBag.CantidadMovimientos = await _context.MovimientosContables
                .CountAsync(m => m.CuentaContableId == id);

            ViewBag.UltimosMovimientos = movimientos;

            return View(cuenta);
        }

        // GET: CuentasContables/Create
        public IActionResult Create()
        {
            return View();
        }

        // POST: CuentasContables/Create
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Create([Bind("IdCuentaContable,Codigo,Nombre,TipoCuenta,Naturaleza,SaldoInicial")] CuentasContable cuenta)
        {
            cuenta.Codigo = cuenta.Codigo?.Trim();
            cuenta.Nombre = cuenta.Nombre?.Trim();

            var codigoExiste = await _context.CuentasContables
                .AnyAsync(c => c.Codigo == cuenta.Codigo);

            if (codigoExiste)
                ModelState.AddModelError("Codigo", "Ya existe una cuenta contable con este código.");

            var nombreExiste = await _context.CuentasContables
                .AnyAsync(c => c.Nombre == cuenta.Nombre);

            if (nombreExiste)
                ModelState.AddModelError("Nombre", "Ya existe una cuenta contable con este nombre.");

            ModelState.Remove("SaldoActual");
            ModelState.Remove("Activo");
            ModelState.Remove("FechaCreacion");
            ModelState.Remove("MovimientosContables");

            if (ModelState.IsValid)
            {
                cuenta.Activo = true;
                cuenta.FechaCreacion = DateTime.Now;
                cuenta.SaldoActual = cuenta.SaldoInicial;

                _context.CuentasContables.Add(cuenta);
                await _context.SaveChangesAsync();

                await RegistrarAuditoria(
                    "CuentasContables",
                    cuenta.IdCuentaContable,
                    "Crear",
                    "Se creó la cuenta contable " + cuenta.Codigo + " - " + cuenta.Nombre
                );

                TempData["MensajeExito"] = "Cuenta contable creada correctamente.";
                return RedirectToAction(nameof(Index));
            }

            return View(cuenta);
        }

        // GET: CuentasContables/Edit/5
        public async Task<IActionResult> Edit(int? id)
        {
            if (id == null)
                return NotFound();

            var cuenta = await _context.CuentasContables.FindAsync(id);

            if (cuenta == null)
                return NotFound();

            bool tieneMovimientos = await _context.MovimientosContables
                .AnyAsync(m => m.CuentaContableId == id);

            ViewBag.TieneMovimientos = tieneMovimientos;

            return View(cuenta);
        }

        // POST: CuentasContables/Edit/5
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Edit(int id, [Bind("IdCuentaContable,Codigo,Nombre,TipoCuenta,Naturaleza,SaldoInicial")] CuentasContable cuentaFormulario)
        {
            if (id != cuentaFormulario.IdCuentaContable)
                return NotFound();

            cuentaFormulario.Codigo = cuentaFormulario.Codigo?.Trim();
            cuentaFormulario.Nombre = cuentaFormulario.Nombre?.Trim();

            var cuentaOriginal = await _context.CuentasContables
                .FirstOrDefaultAsync(c => c.IdCuentaContable == id);

            if (cuentaOriginal == null)
                return NotFound();

            bool tieneMovimientos = await _context.MovimientosContables
                .AnyAsync(m => m.CuentaContableId == id);

            ViewBag.TieneMovimientos = tieneMovimientos;

            var nombreExiste = await _context.CuentasContables
                .AnyAsync(c => c.Nombre == cuentaFormulario.Nombre && c.IdCuentaContable != id);

            if (nombreExiste)
                ModelState.AddModelError("Nombre", "Ya existe una cuenta contable con este nombre.");

            if (!tieneMovimientos)
            {
                var codigoExiste = await _context.CuentasContables
                    .AnyAsync(c => c.Codigo == cuentaFormulario.Codigo && c.IdCuentaContable != id);

                if (codigoExiste)
                    ModelState.AddModelError("Codigo", "Ya existe una cuenta contable con este código.");
            }

            ModelState.Remove("SaldoActual");
            ModelState.Remove("Activo");
            ModelState.Remove("FechaCreacion");
            ModelState.Remove("MovimientosContables");

            if (ModelState.IsValid)
            {
                if (tieneMovimientos)
                {
                    // Regla de negocio:
                    // Si la cuenta ya posee movimientos, solo se permite cambiar el nombre.
                    cuentaOriginal.Nombre = cuentaFormulario.Nombre;
                }
                else
                {
                    // Si no posee movimientos, se permite modificar la estructura completa.
                    cuentaOriginal.Codigo = cuentaFormulario.Codigo;
                    cuentaOriginal.Nombre = cuentaFormulario.Nombre;
                    cuentaOriginal.TipoCuenta = cuentaFormulario.TipoCuenta;
                    cuentaOriginal.Naturaleza = cuentaFormulario.Naturaleza;
                    cuentaOriginal.SaldoInicial = cuentaFormulario.SaldoInicial;
                    cuentaOriginal.SaldoActual = cuentaFormulario.SaldoInicial;
                }

                _context.Update(cuentaOriginal);
                await _context.SaveChangesAsync();

                await RegistrarAuditoria(
                    "CuentasContables",
                    cuentaOriginal.IdCuentaContable,
                    "Modificar",
                    "Se modificó la cuenta contable " + cuentaOriginal.Codigo + " - " + cuentaOriginal.Nombre
                );

                TempData["MensajeExito"] = tieneMovimientos
                    ? "Cuenta contable modificada correctamente. Como tiene movimientos, solo se actualizó el nombre."
                    : "Cuenta contable modificada correctamente.";

                return RedirectToAction(nameof(Index));
            }

            return View(cuentaFormulario);
        }

        // GET: CuentasContables/Delete/5
        public async Task<IActionResult> Delete(int? id)
        {
            if (id == null)
                return NotFound();

            var cuenta = await _context.CuentasContables
                .FirstOrDefaultAsync(c => c.IdCuentaContable == id);

            if (cuenta == null)
                return NotFound();

            return View(cuenta);
        }

        // POST: CuentasContables/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> DeleteConfirmed(int id)
        {
            var cuenta = await _context.CuentasContables.FindAsync(id);

            if (cuenta == null)
                return NotFound();

            bool tieneMovimientos = await _context.MovimientosContables
                .AnyAsync(m => m.CuentaContableId == id);

            if (tieneMovimientos)
            {
                TempData["MensajeError"] = "No se puede desactivar esta cuenta porque tiene movimientos contables asociados.";
                return RedirectToAction(nameof(Index));
            }

            if (cuenta.SaldoActual != 0)
            {
                TempData["MensajeError"] = "No se puede desactivar esta cuenta porque tiene saldo actual diferente de cero.";
                return RedirectToAction(nameof(Index));
            }

            cuenta.Activo = false;

            _context.Update(cuenta);
            await _context.SaveChangesAsync();

            await RegistrarAuditoria(
                "CuentasContables",
                cuenta.IdCuentaContable,
                "Desactivar",
                "Se desactivó la cuenta contable " + cuenta.Codigo + " - " + cuenta.Nombre
            );

            TempData["MensajeExito"] = "Cuenta contable desactivada correctamente.";
            return RedirectToAction(nameof(Index));
        }

        // POST: CuentasContables/Reactivar/5
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Reactivar(int id)
        {
            var cuenta = await _context.CuentasContables.FindAsync(id);

            if (cuenta == null)
                return NotFound();

            cuenta.Activo = true;

            _context.Update(cuenta);
            await _context.SaveChangesAsync();

            await RegistrarAuditoria(
                "CuentasContables",
                cuenta.IdCuentaContable,
                "Reactivar",
                "Se reactivó la cuenta contable " + cuenta.Codigo + " - " + cuenta.Nombre
            );

            TempData["MensajeExito"] = "Cuenta contable reactivada correctamente.";
            return RedirectToAction(nameof(Index));
        }

             // LIBRO MAYOR
        // ===========================================================
        public async Task<IActionResult> LibroMayor(
            int id,
            DateOnly? fechaInicio,
            DateOnly? fechaFin,
            string? origen,
            int? proyectoId,
            int? centroCostoId,
            string? estado,
            string? origenVista)
        {
            var cuenta = await _context.CuentasContables
                .AsNoTracking()
                .FirstOrDefaultAsync(c =>
                    c.IdCuentaContable == id);

            if (cuenta == null)
            {
                return NotFound();
            }

            var consulta = ConstruirConsultaLibroMayor(
                id,
                fechaInicio,
                fechaFin,
                origen,
                proyectoId,
                centroCostoId,
                estado);

            var movimientos = await consulta
                .OrderBy(m => m.Fecha)
                .ThenBy(m => m.IdMovimientoContable)
                .ToListAsync();

            decimal saldoAnterior =
                await ObtenerSaldoAnteriorLibroMayor(
                    cuenta,
                    fechaInicio);

            CalcularSaldosLibroMayor(
                movimientos,
                cuenta,
                saldoAnterior);

            var movimientosVigentes = movimientos
     .Where(m =>
         m.Anulado == false &&
         m.Estado != "Reversado")
     .ToList();

            ViewBag.Cuenta = cuenta;

            ViewBag.FechaInicio =
                fechaInicio?.ToString("yyyy-MM-dd");

            ViewBag.FechaFin =
                fechaFin?.ToString("yyyy-MM-dd");

            ViewBag.Origen = origen;
            ViewBag.ProyectoId = proyectoId;
            ViewBag.CentroCostoId = centroCostoId;
            ViewBag.Estado = estado;
            ViewBag.OrigenVista = origenVista;

            ViewBag.SaldoAnterior = saldoAnterior;

            ViewBag.TotalDebe =
                movimientosVigentes.Sum(m => m.Debe);

            ViewBag.TotalHaber =
                movimientosVigentes.Sum(m => m.Haber);

            ViewBag.SaldoActual = movimientos.Any()
                ? movimientos.Last().Monto
                : saldoAnterior;

            ViewBag.CantidadMovimientos =
                movimientos.Count;

            ViewBag.CantidadVigentes =
                movimientosVigentes.Count;

            ViewBag.CantidadAnulados =
                movimientos.Count(m =>
                    m.Anulado == true ||
                    m.Estado == "Anulado");

            ViewBag.CantidadReversados =
                movimientos.Count(m =>
                    m.TipoMovimiento == "Reversión" ||
                    m.Estado == "Reversado");

            await CargarFiltrosLibroMayor(id);

            return View(movimientos);
        }
        // ===========================================================
        // CARGAR FILTROS DEL LIBRO MAYOR
        // ===========================================================
        private async Task CargarFiltrosLibroMayor(
            int cuentaContableId)
        {
            ViewBag.Origenes = await _context.MovimientosContables
                .AsNoTracking()
                .Where(m =>
                    m.CuentaContableId == cuentaContableId &&
                    m.OrigenModulo != null)
                .Select(m => m.OrigenModulo!)
                .Distinct()
                .OrderBy(o => o)
                .ToListAsync();

            ViewBag.Proyectos = await _context.Proyectos
    .AsNoTracking()
    .Where(p => p.Activo)
    .OrderBy(p => p.Nombre)
    .ToListAsync();

            ViewBag.CentrosCosto = await _context.CentrosCostos
                .AsNoTracking()
                .Where(c => c.Activo)
                .OrderBy(c => c.Nombre)
                .ToListAsync();
        }
        // ===========================================================
        // CONSULTA BASE DEL LIBRO MAYOR
        // ===========================================================
        private IQueryable<MovimientosContable> ConstruirConsultaLibroMayor(
            int cuentaContableId,
            DateOnly? fechaInicio,
            DateOnly? fechaFin,
            string? origen,
            int? proyectoId,
            int? centroCostoId,
            string? estado)
        {
            var consulta = _context.MovimientosContables
                .AsNoTracking()
                .Include(m => m.Proyecto)
                .Include(m => m.CentroCosto)
                .Where(m => m.CuentaContableId == cuentaContableId)
                .AsQueryable();

            if (fechaInicio.HasValue)
            {
                consulta = consulta.Where(m =>
                    m.Fecha >= fechaInicio.Value);
            }

            if (fechaFin.HasValue)
            {
                consulta = consulta.Where(m =>
                    m.Fecha <= fechaFin.Value);
            }

            if (!string.IsNullOrWhiteSpace(origen))
            {
                origen = origen.Trim();

                consulta = consulta.Where(m =>
                    m.OrigenModulo == origen);
            }

            if (proyectoId.HasValue)
            {
                consulta = consulta.Where(m =>
                    m.ProyectoId == proyectoId.Value);
            }

            if (centroCostoId.HasValue)
            {
                consulta = consulta.Where(m =>
                    m.CentroCostoId == centroCostoId.Value);
            }

            if (!string.IsNullOrWhiteSpace(estado))
            {
                switch (estado)
                {
                    case "Registrado":
                        consulta = consulta.Where(m =>
                            m.Anulado != true &&
                            m.Estado == "Registrado" &&
                            m.TipoMovimiento != "Reversión");
                        break;

                    case "Anulado":
                        consulta = consulta.Where(m =>
                            m.Anulado == true ||
                            m.Estado == "Anulado");
                        break;

                    case "Reversado":
                        consulta = consulta.Where(m =>
                            m.TipoMovimiento == "Reversión" ||
                            m.Estado == "Reversado");
                        break;
                }
            }

            return consulta;
        }
        // ===========================================================
        // CALCULAR SALDO ACUMULADO DEL LIBRO MAYOR
        // ===========================================================
        private void CalcularSaldosLibroMayor(
            List<MovimientosContable> movimientos,
            CuentasContable cuenta,
            decimal saldoInicial)
        {
            decimal saldoAcumulado = saldoInicial;

            foreach (var movimiento in movimientos)
            {
                bool movimientoVigente =
     movimiento.Anulado == false &&
     movimiento.Estado != "Reversado";

                if (movimientoVigente)
                {
                    if (cuenta.Naturaleza == "Deudora")
                    {
                        saldoAcumulado +=
                            movimiento.Debe - movimiento.Haber;
                    }
                    else
                    {
                        saldoAcumulado +=
                            movimiento.Haber - movimiento.Debe;
                    }
                }

                // La vista ya utiliza Monto como saldo acumulado.
                movimiento.Monto = saldoAcumulado;
            }
        }
        // ===========================================================
        // SALDO ANTERIOR AL PERIODO FILTRADO
        // ===========================================================
        private async Task<decimal> ObtenerSaldoAnteriorLibroMayor(
            CuentasContable cuenta,
            DateOnly? fechaInicio)
        {
            decimal saldoAnterior = cuenta.SaldoInicial;

            if (!fechaInicio.HasValue)
            {
                return saldoAnterior;
            }

            var movimientosAnteriores = await _context.MovimientosContables
                .AsNoTracking()
                .Where(m =>
                    m.CuentaContableId == cuenta.IdCuentaContable &&
                    m.Fecha < fechaInicio.Value &&
                    m.Anulado == false &&
                    m.Estado != "Reversado")
                .OrderBy(m => m.Fecha)
                .ThenBy(m => m.IdMovimientoContable)
                .ToListAsync();

            foreach (var movimiento in movimientosAnteriores)
            {
                if (cuenta.Naturaleza == "Deudora")
                {
                    saldoAnterior +=
                        movimiento.Debe - movimiento.Haber;
                }
                else
                {
                    saldoAnterior +=
                        movimiento.Haber - movimiento.Debe;
                }
            }

            return saldoAnterior;
        }
        // ===========================================================
        // VISTA PREVIA DEL REPORTE DE LIBRO MAYOR
        // ===========================================================
        public async Task<IActionResult> ReporteLibroMayor(
            int id,
            DateOnly? fechaInicio,
            DateOnly? fechaFin,
            string? origen,
            int? proyectoId,
            int? centroCostoId,
            string? estado)
        {
            var cuenta = await _context.CuentasContables
                .AsNoTracking()
                .FirstOrDefaultAsync(c =>
                    c.IdCuentaContable == id);

            if (cuenta == null)
            {
                return NotFound();
            }

            var consulta = ConstruirConsultaLibroMayor(
                id,
                fechaInicio,
                fechaFin,
                origen,
                proyectoId,
                centroCostoId,
                estado);

            var movimientos = await consulta
                .OrderBy(m => m.Fecha)
                .ThenBy(m => m.IdMovimientoContable)
                .ToListAsync();

            decimal saldoAnterior =
                await ObtenerSaldoAnteriorLibroMayor(
                    cuenta,
                    fechaInicio);

            CalcularSaldosLibroMayor(
                movimientos,
                cuenta,
                saldoAnterior);

            var movimientosVigentes = movimientos
                .Where(m =>
                    m.Anulado != true &&
                    m.Estado != "Anulado")
                .ToList();

            ViewBag.Cuenta = cuenta;

            ViewBag.FechaInicio =
                fechaInicio?.ToString("yyyy-MM-dd");

            ViewBag.FechaFin =
                fechaFin?.ToString("yyyy-MM-dd");

            ViewBag.Origen = origen;
            ViewBag.ProyectoId = proyectoId;
            ViewBag.CentroCostoId = centroCostoId;
            ViewBag.Estado = estado;

            ViewBag.SaldoAnterior = saldoAnterior;

            ViewBag.TotalDebe =
                movimientosVigentes.Sum(m => m.Debe);

            ViewBag.TotalHaber =
                movimientosVigentes.Sum(m => m.Haber);

            ViewBag.SaldoFinal = movimientos.Any()
                ? movimientos.Last().Monto
                : saldoAnterior;

            ViewBag.CantidadMovimientos =
                movimientos.Count;

            ViewBag.CantidadVigentes =
                movimientosVigentes.Count;

            ViewBag.CantidadAnulados =
                movimientos.Count(m =>
                    m.Anulado == true ||
                    m.Estado == "Anulado");

            ViewBag.CantidadReversados =
                movimientos.Count(m =>
                    m.TipoMovimiento == "Reversión" ||
                    m.Estado == "Reversado");

            ViewBag.FechaGeneracion = DateTime.Now;
            await RegistrarAuditoria(
    "CuentasContables",
    id,
    "Vista previa PDF",
    $"Se generó la vista previa del Libro Mayor de la cuenta {cuenta.Codigo} - {cuenta.Nombre}."
);
            return View(movimientos);
        }
        // ===========================================================
        // EXPORTAR LIBRO MAYOR A EXCEL
        // ===========================================================
        public async Task<IActionResult> ExportarLibroMayorExcel(
            int id,
            DateOnly? fechaInicio,
            DateOnly? fechaFin,
            string? origen,
            int? proyectoId,
            int? centroCostoId,
            string? estado)
        {
            var cuenta = await _context.CuentasContables
                .AsNoTracking()
                .FirstOrDefaultAsync(c =>
                    c.IdCuentaContable == id);

            if (cuenta == null)
            {
                return NotFound();
            }

            var consulta = ConstruirConsultaLibroMayor(
                id,
                fechaInicio,
                fechaFin,
                origen,
                proyectoId,
                centroCostoId,
                estado);

            var movimientos = await consulta
                .OrderBy(m => m.Fecha)
                .ThenBy(m => m.IdMovimientoContable)
                .ToListAsync();

            decimal saldoAnterior =
                await ObtenerSaldoAnteriorLibroMayor(
                    cuenta,
                    fechaInicio);

            CalcularSaldosLibroMayor(
                movimientos,
                cuenta,
                saldoAnterior);

            var movimientosVigentes = movimientos
                .Where(m =>
                    m.Anulado != true &&
                    m.Estado != "Anulado")
                .ToList();

            decimal totalDebe =
                movimientosVigentes.Sum(m => m.Debe);

            decimal totalHaber =
                movimientosVigentes.Sum(m => m.Haber);

            decimal saldoFinal = movimientos.Any()
                ? movimientos.Last().Monto
                : saldoAnterior;

            using var workbook = new XLWorkbook();

            var hoja = workbook.Worksheets.Add(
                "Libro Mayor");

            const int totalColumnas = 9;

            // =======================================================
            // ENCABEZADO
            // =======================================================
            hoja.Cell("A1").Value =
                "Unión Cantonal de Asociaciones de Desarrollo de Santa Ana";

            hoja.Range(1, 1, 1, totalColumnas).Merge();

            hoja.Cell("A2").Value =
                "Reporte de Libro Mayor";

            hoja.Range(2, 1, 2, totalColumnas).Merge();

            hoja.Cell("A3").Value =
                $"{cuenta.Codigo} - {cuenta.Nombre}";

            hoja.Range(3, 1, 3, totalColumnas).Merge();

            hoja.Cell("A4").Value =
                $"Generado el {DateTime.Now:dd/MM/yyyy HH:mm}";

            hoja.Range(4, 1, 4, totalColumnas).Merge();

            hoja.Range(1, 1, 1, totalColumnas)
                .Style.Font.Bold = true;

            hoja.Range(1, 1, 1, totalColumnas)
                .Style.Font.FontSize = 16;

            hoja.Range(1, 1, 1, totalColumnas)
                .Style.Font.FontColor =
                    XLColor.FromHtml("#0F5C64");

            hoja.Range(1, 1, 4, totalColumnas)
                .Style.Alignment.Horizontal =
                    XLAlignmentHorizontalValues.Center;

            hoja.Range(2, 1, 3, totalColumnas)
                .Style.Font.Bold = true;

            // =======================================================
            // INFORMACIÓN DE LA CUENTA
            // =======================================================
            hoja.Cell("A6").Value = "Código:";
            hoja.Cell("B6").Value = cuenta.Codigo;

            hoja.Cell("C6").Value = "Tipo:";
            hoja.Cell("D6").Value = cuenta.TipoCuenta;

            hoja.Cell("E6").Value = "Naturaleza:";
            hoja.Cell("F6").Value = cuenta.Naturaleza;

            hoja.Cell("G6").Value = fechaInicio.HasValue
      ? "Saldo anterior:"
      : "Saldo inicial:";

            hoja.Cell("H6").Value = saldoAnterior;

            hoja.Range("A6:H6").Style.Fill.BackgroundColor =
                XLColor.FromHtml("#F2F7F8");

            hoja.Range("A6:H6").Style.Border.OutsideBorder =
                XLBorderStyleValues.Thin;

            hoja.Cell("A6").Style.Font.Bold = true;
            hoja.Cell("C6").Style.Font.Bold = true;
            hoja.Cell("E6").Style.Font.Bold = true;
            hoja.Cell("G6").Style.Font.Bold = true;

            hoja.Cell("H6").Style.NumberFormat.Format =
                "₡ #,##0.00";

            // =======================================================
            // FILTROS
            // =======================================================
            var filtros = new List<string>();

            if (fechaInicio.HasValue)
                filtros.Add(
                    $"Desde: {fechaInicio.Value:dd/MM/yyyy}");

            if (fechaFin.HasValue)
                filtros.Add(
                    $"Hasta: {fechaFin.Value:dd/MM/yyyy}");

            if (!string.IsNullOrWhiteSpace(origen))
                filtros.Add($"Origen: {origen}");

            if (proyectoId.HasValue)
            {
                var proyecto = await _context.Proyectos
                    .AsNoTracking()
                    .Where(p =>
                        p.IdProyecto == proyectoId.Value)
                    .Select(p => p.Nombre)
                    .FirstOrDefaultAsync();

                filtros.Add(
                    $"Proyecto: {proyecto ?? "No disponible"}");
            }

            if (centroCostoId.HasValue)
            {
                var centroCosto = await _context.CentrosCostos
                    .AsNoTracking()
                    .Where(c =>
                        c.IdCentroCosto == centroCostoId.Value)
                    .Select(c => c.Nombre)
                    .FirstOrDefaultAsync();

                filtros.Add(
                    $"Centro de costo: {centroCosto ?? "No disponible"}");
            }

            if (!string.IsNullOrWhiteSpace(estado))
                filtros.Add($"Estado: {estado}");

            hoja.Cell("A8").Value =
                "Filtros aplicados:";

            hoja.Cell("A8").Style.Font.Bold = true;

            hoja.Cell("B8").Value = filtros.Any()
                ? string.Join(" | ", filtros)
                : "Sin filtros. Se muestran todos los movimientos.";

            hoja.Range(8, 2, 8, totalColumnas).Merge();

            hoja.Range(8, 1, 8, totalColumnas)
                .Style.Fill.BackgroundColor =
                    XLColor.FromHtml("#F2F7F8");

            hoja.Range(8, 1, 8, totalColumnas)
                .Style.Border.LeftBorder =
                    XLBorderStyleValues.Medium;

            hoja.Range(8, 1, 8, totalColumnas)
                .Style.Border.LeftBorderColor =
                    XLColor.FromHtml("#0F7C90");

            // =======================================================
            // TABLA
            // =======================================================
            const int filaEncabezado = 10;

            string[] encabezados =
            {
        "Fecha",
        "Referencia",
        "Origen",
        "Proyecto",
        "Centro de costo",
        "Debe",
        "Haber",
        "Saldo acumulado",
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

            var rangoEncabezado = hoja.Range(
                filaEncabezado,
                1,
                filaEncabezado,
                totalColumnas);

            rangoEncabezado.Style.Font.Bold = true;
            rangoEncabezado.Style.Font.FontColor =
                XLColor.White;

            rangoEncabezado.Style.Fill.BackgroundColor =
                XLColor.FromHtml("#0F5C64");

            rangoEncabezado.Style.Alignment.Horizontal =
                XLAlignmentHorizontalValues.Center;

            int fila = filaEncabezado + 1;

            foreach (var movimiento in movimientos)
            {
                string estadoMovimiento;

                if (movimiento.Anulado == true ||
                    movimiento.Estado == "Anulado")
                {
                    estadoMovimiento = "Anulado";
                }
                else if (
                    movimiento.TipoMovimiento == "Reversión" ||
                    movimiento.Estado == "Reversado")
                {
                    estadoMovimiento = "Reversión";
                }
                else
                {
                    estadoMovimiento = "Registrado";
                }

                hoja.Cell(fila, 1).Value =
                    movimiento.Fecha.ToDateTime(
                        TimeOnly.MinValue);

                hoja.Cell(fila, 2).Value =
                    movimiento.Referencia ??
                    "Sin referencia";

                hoja.Cell(fila, 3).Value =
                    movimiento.OrigenModulo ??
                    "Sin origen";

                hoja.Cell(fila, 4).Value =
                    movimiento.Proyecto?.Nombre ??
                    "Sin proyecto";

                hoja.Cell(fila, 5).Value =
                    movimiento.CentroCosto?.Nombre ??
                    "Sin centro";

                hoja.Cell(fila, 6).Value =
                    movimiento.Debe;

                hoja.Cell(fila, 7).Value =
                    movimiento.Haber;

                hoja.Cell(fila, 8).Value =
                    movimiento.Monto;

                hoja.Cell(fila, 9).Value =
                    estadoMovimiento;

                fila++;
            }

            if (movimientos.Any())
            {
                int filaInicioDatos =
                    filaEncabezado + 1;

                int filaFinDatos =
                    fila - 1;

                var rangoTabla = hoja.Range(
                    filaEncabezado,
                    1,
                    filaFinDatos,
                    totalColumnas);

                var tabla = rangoTabla.CreateTable(
                    "TablaLibroMayor");

                tabla.Theme =
                    XLTableTheme.TableStyleMedium2;

                tabla.ShowAutoFilter = true;
                tabla.ShowRowStripes = false;

                hoja.Range(
                        filaInicioDatos,
                        1,
                        filaFinDatos,
                        1)
                    .Style.DateFormat.Format =
                        "dd/MM/yyyy";

                hoja.Range(
                        filaInicioDatos,
                        6,
                        filaFinDatos,
                        8)
                    .Style.NumberFormat.Format =
                        "₡ #,##0.00";

                hoja.Range(
                        filaInicioDatos,
                        6,
                        filaFinDatos,
                        8)
                    .Style.Alignment.Horizontal =
                        XLAlignmentHorizontalValues.Right;
            }

            // =======================================================
            // TOTALES
            // =======================================================
            int filaTotales = fila + 1;

            hoja.Cell(filaTotales, 5).Value =
                "TOTALES:";

            hoja.Cell(filaTotales, 6).Value =
                totalDebe;

            hoja.Cell(filaTotales, 7).Value =
                totalHaber;

            hoja.Cell(filaTotales, 8).Value =
                saldoFinal;

            var rangoTotales = hoja.Range(
                filaTotales,
                5,
                filaTotales,
                8);

            rangoTotales.Style.Font.Bold = true;

            rangoTotales.Style.Fill.BackgroundColor =
                XLColor.FromHtml("#D9EDEF");

            hoja.Range(
                    filaTotales,
                    6,
                    filaTotales,
                    8)
                .Style.NumberFormat.Format =
                    "₡ #,##0.00";

            // =======================================================
            // RESUMEN
            // =======================================================
            int filaResumen = filaTotales + 3;

            hoja.Cell(filaResumen, 1).Value =
                "RESUMEN DEL LIBRO MAYOR";

            hoja.Range(
                    filaResumen,
                    1,
                    filaResumen,
                    5)
                .Merge();

            hoja.Range(
                    filaResumen,
                    1,
                    filaResumen,
                    5)
                .Style.Fill.BackgroundColor =
                    XLColor.FromHtml("#0F5C64");

            hoja.Range(
                    filaResumen,
                    1,
                    filaResumen,
                    5)
                .Style.Font.FontColor =
                    XLColor.White;

            hoja.Range(
                    filaResumen,
                    1,
                    filaResumen,
                    5)
                .Style.Font.Bold = true;

            string[] resumenNombres =
            {
        "Saldo anterior",
        "Total Debe",
        "Total Haber",
        "Saldo final",
        "Movimientos"
    };

            object[] resumenValores =
            {
        saldoAnterior,
        totalDebe,
        totalHaber,
        saldoFinal,
        movimientos.Count
    };

            for (int columna = 0;
                 columna < resumenNombres.Length;
                 columna++)
            {
                hoja.Cell(
                    filaResumen + 1,
                    columna + 1).Value =
                        resumenNombres[columna];

                if (columna < 4)
                {
                    hoja.Cell(
                        filaResumen + 2,
                        columna + 1).Value =
                            Convert.ToDecimal(
                                resumenValores[columna]);
                }
                else
                {
                    hoja.Cell(
                        filaResumen + 2,
                        columna + 1).Value =
                            Convert.ToInt32(
                                resumenValores[columna]);
                }
            }

            hoja.Range(
                    filaResumen + 1,
                    1,
                    filaResumen + 1,
                    5)
                .Style.Fill.BackgroundColor =
                    XLColor.FromHtml("#D9EDEF");

            hoja.Range(
                    filaResumen + 1,
                    1,
                    filaResumen + 1,
                    5)
                .Style.Font.Bold = true;

            hoja.Range(
                    filaResumen + 2,
                    1,
                    filaResumen + 2,
                    4)
                .Style.NumberFormat.Format =
                    "₡ #,##0.00";

            hoja.Range(
                    filaResumen + 1,
                    1,
                    filaResumen + 2,
                    5)
                .Style.Alignment.Horizontal =
                    XLAlignmentHorizontalValues.Center;

            hoja.Range(
                    filaResumen + 1,
                    1,
                    filaResumen + 2,
                    5)
                .Style.Border.OutsideBorder =
                    XLBorderStyleValues.Thin;

            hoja.Range(
                    filaResumen + 1,
                    1,
                    filaResumen + 2,
                    5)
                .Style.Border.InsideBorder =
                    XLBorderStyleValues.Thin;

            // =======================================================
            // CONFIGURACIÓN FINAL
            // =======================================================
            hoja.SheetView.FreezeRows(filaEncabezado);
            hoja.PageSetup.PageOrientation =
                XLPageOrientation.Landscape;

            hoja.PageSetup.FitToPages(1, 0);

            hoja.Column(1).Width = 14;
            hoja.Column(2).Width = 22;
            hoja.Column(3).Width = 18;
            hoja.Column(4).Width = 24;
            hoja.Column(5).Width = 24;
            hoja.Column(6).Width = 16;
            hoja.Column(7).Width = 16;
            hoja.Column(8).Width = 18;
            hoja.Column(9).Width = 15;

            using var stream = new MemoryStream();

            workbook.SaveAs(stream);

            string nombreArchivo =
                $"LibroMayor_{cuenta.Codigo}_" +
                $"{DateTime.Now:yyyyMMdd_HHmm}.xlsx";
            await RegistrarAuditoria(
    "CuentasContables",
    id,
    "Exportar Excel",
    $"Se exportó a Excel el Libro Mayor de la cuenta {cuenta.Codigo} - {cuenta.Nombre}."
);
            return File(
                stream.ToArray(),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                nombreArchivo);
        }
        // ===========================================================
        // CONSULTA BASE PARA REPORTES DE CUENTAS CONTABLES
        // ===========================================================

        private IQueryable<CuentasContable> ConstruirConsultaReporte(
            string? codigo,
            string? nombre,
            string? tipoCuenta,
            string? estado)
        {
            var consulta =
                _context.CuentasContables
                    .AsNoTracking()
                    .AsQueryable();

            if (!string.IsNullOrWhiteSpace(codigo))
            {
                codigo = codigo.Trim();

                consulta = consulta.Where(c =>
                    c.Codigo.Contains(codigo));
            }

            if (!string.IsNullOrWhiteSpace(nombre))
            {
                nombre = nombre.Trim();

                consulta = consulta.Where(c =>
                    c.Nombre.Contains(nombre));
            }

            if (!string.IsNullOrWhiteSpace(tipoCuenta))
            {
                consulta = consulta.Where(c =>
                    c.TipoCuenta == tipoCuenta);
            }

            if (!string.IsNullOrWhiteSpace(estado))
            {
                if (estado == "Activa")
                {
                    consulta = consulta.Where(c => c.Activo);
                }
                else if (estado == "Inactiva")
                {
                    consulta = consulta.Where(c => !c.Activo);
                }
            }

            return consulta;
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

        private bool CuentasContableExists(int id)
        {
            return _context.CuentasContables.Any(e => e.IdCuentaContable == id);
        }
    }
}