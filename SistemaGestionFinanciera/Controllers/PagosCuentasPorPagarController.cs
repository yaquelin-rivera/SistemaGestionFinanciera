using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.EntityFrameworkCore;
using SistemaGestionFinanciera.Data;
using SistemaGestionFinanciera.Models;
using System.Text.Json;
using ClosedXML.Excel;
using System.Globalization;
using System.IO;


namespace SistemaGestionFinanciera.Controllers
{
    public class PagosCuentasPorPagarController : BaseController
    {
        private readonly SistemaFinancieroContext _context;

        public PagosCuentasPorPagarController(SistemaFinancieroContext context)
        {
            _context = context;
        }

        // LISTADO DE PAGOS DE CUENTAS POR PAGAR
        // ===========================================================
        public async Task<IActionResult> Index(
            string? buscar,
            string? tipoDocumento,
            string? estado,
            string? metodoPago,
            int? proyectoId,
            int? anio,
            int? mes,
            DateOnly? fechaInicio,
            DateOnly? fechaFin)
        {
            bool hayFiltros =
                !string.IsNullOrWhiteSpace(buscar) ||
                !string.IsNullOrWhiteSpace(tipoDocumento) ||
                !string.IsNullOrWhiteSpace(estado) ||
                !string.IsNullOrWhiteSpace(metodoPago) ||
                (proyectoId.HasValue && proyectoId.Value > 0) ||
                anio.HasValue ||
                mes.HasValue ||
                fechaInicio.HasValue ||
                fechaFin.HasValue;

            var listaPagos = await ConstruirConsultaReporte(
                buscar,
                tipoDocumento,
                estado,
                metodoPago,
                proyectoId,
                anio,
                mes,
                fechaInicio,
                fechaFin);

            var pagosParaTotales = listaPagos;

            // Sin filtros, las cards muestran solamente el mes actual.
            if (!hayFiltros)
            {
                var hoy = DateOnly.FromDateTime(DateTime.Today);

                var inicioMes =
                    new DateOnly(hoy.Year, hoy.Month, 1);

                var finMes =
                    inicioMes.AddMonths(1).AddDays(-1);

                pagosParaTotales = listaPagos
                    .Where(p =>
                        p.FechaPago >= inicioMes &&
                        p.FechaPago <= finMes)
                    .ToList();

                ViewBag.TipoResumen = "Mes actual";
            }
            else
            {
                ViewBag.TipoResumen = "Filtro aplicado";
            }

            // Solo los pagos activos participan en el total monetario.
            ViewBag.TotalPagosActivos = pagosParaTotales
                .Where(p => p.Anulado == false)
                .Sum(p => p.Monto ?? 0m);

            ViewBag.CantidadPagos =
                pagosParaTotales.Count;

            ViewBag.PagosAnulados = pagosParaTotales
                .Count(p => p.Anulado == true);

            ViewBag.MetodosUtilizados = string.Join(
                " | ",
                pagosParaTotales
                    .Where(p =>
                        !string.IsNullOrWhiteSpace(p.MetodoPago))
                    .GroupBy(p => p.MetodoPago)
                    .OrderBy(g => g.Key)
                    .Select(g => $"{g.Key}: {g.Count()}"));

            // Conserva los filtros en pantalla.
            ViewBag.Buscar = buscar;
            ViewBag.TipoDocumento = tipoDocumento;
            ViewBag.Estado = estado;
            ViewBag.MetodoPago = metodoPago;
            ViewBag.ProyectoIdValor = proyectoId;
            ViewBag.Anio = anio;
            ViewBag.Mes = mes;

            ViewBag.FechaInicio =
                fechaInicio?.ToString("yyyy-MM-dd");

            ViewBag.FechaFin =
                fechaFin?.ToString("yyyy-MM-dd");

            // Proyectos activos.
            ViewBag.ProyectoId = new SelectList(
                await _context.Proyectos
                    .Where(p => p.Activo == true)
                    .OrderBy(p => p.Nombre)
                    .ToListAsync(),
                "IdProyecto",
                "Nombre",
                proyectoId);

            // Años disponibles.
            ViewBag.Anios = await _context
                .PagosCuentaPorPagars
                .Select(p => p.FechaPago.Year)
                .Distinct()
                .OrderByDescending(a => a)
                .ToListAsync();

            return View(listaPagos);
        }
        // ===========================================================
        // CONSULTA COMPARTIDA PARA INDEX, REPORTE Y EXCEL
        // ===========================================================
        private async Task<List<PagosCuentaPorPagar>>
            ConstruirConsultaReporte(
                string? buscar,
                string? tipoDocumento,
                string? estado,
                string? metodoPago,
                int? proyectoId,
                int? anio,
                int? mes,
                DateOnly? fechaInicio,
                DateOnly? fechaFin)
        {
            var consulta = _context.PagosCuentaPorPagars
                .Include(p => p.CuentaPorPagar)
                    .ThenInclude(c => c.Proveedor)
                .Include(p => p.CuentaPorPagar)
                    .ThenInclude(c => c.Proyecto)
                .AsQueryable();

            // Búsqueda por referencia, documento o proveedor.
            if (!string.IsNullOrWhiteSpace(buscar))
            {
                buscar = buscar.Trim();

                consulta = consulta.Where(p =>
                    (p.Referencia != null &&
                     p.Referencia.Contains(buscar)) ||

                    (p.CuentaPorPagar != null &&
                     p.CuentaPorPagar.NumeroDocumento != null &&
                     p.CuentaPorPagar.NumeroDocumento
                        .Contains(buscar)) ||

                    (p.CuentaPorPagar != null &&
                     p.CuentaPorPagar.Proveedor != null &&
                     p.CuentaPorPagar.Proveedor.Nombre
                        .Contains(buscar)));
            }

            // Tipo de documento.
            if (!string.IsNullOrWhiteSpace(tipoDocumento))
            {
                consulta = consulta.Where(p =>
                    p.CuentaPorPagar != null &&
                    p.CuentaPorPagar.TipoDocumento ==
                        tipoDocumento);
            }

            // Estado del pago.
            if (!string.IsNullOrWhiteSpace(estado))
            {
                if (estado == "Activo")
                {
                    consulta = consulta.Where(p =>
                        p.Anulado == false);
                }
                else if (estado == "Anulado")
                {
                    consulta = consulta.Where(p =>
                        p.Anulado == true);
                }
            }

            // Método de pago.
            if (!string.IsNullOrWhiteSpace(metodoPago))
            {
                consulta = consulta.Where(p =>
                    p.MetodoPago == metodoPago);
            }

            // Proyecto.
            if (proyectoId.HasValue &&
                proyectoId.Value > 0)
            {
                consulta = consulta.Where(p =>
                    p.CuentaPorPagar != null &&
                    p.CuentaPorPagar.ProyectoId ==
                        proyectoId.Value);
            }

            // Año.
            if (anio.HasValue)
            {
                consulta = consulta.Where(p =>
                    p.FechaPago.Year == anio.Value);
            }

            // Mes.
            if (mes.HasValue)
            {
                consulta = consulta.Where(p =>
                    p.FechaPago.Month == mes.Value);
            }

            // Fecha inicial.
            if (fechaInicio.HasValue)
            {
                consulta = consulta.Where(p =>
                    p.FechaPago >= fechaInicio.Value);
            }

            // Fecha final.
            if (fechaFin.HasValue)
            {
                consulta = consulta.Where(p =>
                    p.FechaPago <= fechaFin.Value);
            }

            // Pagos más recientes primero.
            return await consulta
                .OrderByDescending(p => p.FechaPago)
                .ThenByDescending(p =>
                    p.IdPagoCuentaPorPagar)
                .ToListAsync();
        }
        // ===========================================================
        // VISTA PREVIA DEL HISTORIAL DE PAGOS CXP
        // ===========================================================
        public async Task<IActionResult> Reporte(
            string? buscar,
            string? tipoDocumento,
            string? estado,
            string? metodoPago,
            int? proyectoId,
            int? anio,
            int? mes,
            DateOnly? fechaInicio,
            DateOnly? fechaFin)
        {
            var pagos = await ConstruirConsultaReporte(
                buscar,
                tipoDocumento,
                estado,
                metodoPago,
                proyectoId,
                anio,
                mes,
                fechaInicio,
                fechaFin);

            var pagosActivos = pagos
                .Where(p => p.Anulado == false)
                .ToList();

            var pagosAnulados = pagos
                .Where(p => p.Anulado == true)
                .ToList();

            ViewBag.TotalRegistros =
                pagos.Count;

            ViewBag.CantidadActivos =
                pagosActivos.Count;

            ViewBag.CantidadAnulados =
                pagosAnulados.Count;

            ViewBag.TotalPagosActivos = pagosActivos
                .Sum(p => p.Monto ?? 0m);

            ViewBag.TotalPagosAnulados = pagosAnulados
                .Sum(p => p.Monto ?? 0m);

            ViewBag.MetodosUtilizados = string.Join(
                " | ",
                pagosActivos
                    .Where(p =>
                        !string.IsNullOrWhiteSpace(p.MetodoPago))
                    .GroupBy(p => p.MetodoPago)
                    .OrderBy(g => g.Key)
                    .Select(g => $"{g.Key}: {g.Count()}"));

            // Resumen por método de pago.
            ViewBag.CantidadEfectivo = pagosActivos
                .Count(p => p.MetodoPago == "Efectivo");

            ViewBag.CantidadTransferencia = pagosActivos
                .Count(p => p.MetodoPago == "Transferencia");

            ViewBag.CantidadDeposito = pagosActivos
                .Count(p => p.MetodoPago == "Depósito");

            ViewBag.CantidadSinpe = pagosActivos
                .Count(p => p.MetodoPago == "SINPE Móvil");

            ViewBag.CantidadCheque = pagosActivos
                .Count(p => p.MetodoPago == "Cheque");

            ViewBag.FechaGeneracion =
                DateTime.Now;

            // Conserva filtros.
            ViewBag.Buscar = buscar;
            ViewBag.TipoDocumento = tipoDocumento;
            ViewBag.Estado = estado;
            ViewBag.MetodoPago = metodoPago;
            ViewBag.ProyectoId = proyectoId;
            ViewBag.Anio = anio;
            ViewBag.Mes = mes;

            ViewBag.FechaInicio =
                fechaInicio?.ToString("yyyy-MM-dd");

            ViewBag.FechaFin =
                fechaFin?.ToString("yyyy-MM-dd");

            ViewBag.NombreProyecto =
                proyectoId.HasValue &&
                proyectoId.Value > 0
                    ? await _context.Proyectos
                        .Where(p =>
                            p.IdProyecto == proyectoId.Value)
                        .Select(p => p.Nombre)
                        .FirstOrDefaultAsync()
                    : null;

            return View(pagos);
        }
        // ===========================================================
        // EXPORTAR HISTORIAL DE PAGOS CXP A EXCEL
        // ===========================================================
        public async Task<IActionResult> ExportarExcel(
            string? buscar,
            string? tipoDocumento,
            string? estado,
            string? metodoPago,
            int? proyectoId,
            int? anio,
            int? mes,
            DateOnly? fechaInicio,
            DateOnly? fechaFin)
        {
            var pagos = await ConstruirConsultaReporte(
                buscar,
                tipoDocumento,
                estado,
                metodoPago,
                proyectoId,
                anio,
                mes,
                fechaInicio,
                fechaFin);

            pagos = pagos
                .OrderByDescending(p => p.FechaPago)
                .ThenByDescending(p =>
                    p.IdPagoCuentaPorPagar)
                .ToList();

            var pagosActivos = pagos
                .Where(p => p.Anulado == false)
                .ToList();

            using var workbook =
                new XLWorkbook();

            var hoja = workbook.Worksheets.Add(
                "Pagos CxP");

            const int totalColumnas = 11;

            // =======================================================
            // ENCABEZADO INSTITUCIONAL
            // =======================================================
            hoja.Cell("A1").Value =
                "Unión Cantonal de Asociaciones de Desarrollo de Santa Ana";

            hoja.Range(
                1,
                1,
                1,
                totalColumnas).Merge();

            hoja.Cell("A2").Value =
                "Historial de Pagos de Cuentas por Pagar";

            hoja.Range(
                2,
                1,
                2,
                totalColumnas).Merge();

            hoja.Cell("A3").Value =
                $"Generado el {DateTime.Now:dd/MM/yyyy HH:mm}";

            hoja.Range(
                3,
                1,
                3,
                totalColumnas).Merge();

            hoja.Range(
                    1,
                    1,
                    1,
                    totalColumnas)
                .Style.Font.Bold = true;

            hoja.Range(
                    1,
                    1,
                    1,
                    totalColumnas)
                .Style.Font.FontSize = 16;

            hoja.Range(
                    1,
                    1,
                    1,
                    totalColumnas)
                .Style.Font.FontColor =
                    XLColor.FromHtml("#0F5C64");

            hoja.Range(
                    2,
                    1,
                    2,
                    totalColumnas)
                .Style.Font.Bold = true;

            hoja.Range(
                    2,
                    1,
                    2,
                    totalColumnas)
                .Style.Font.FontSize = 13;

            hoja.Range(
                    1,
                    1,
                    3,
                    totalColumnas)
                .Style.Alignment.Horizontal =
                    XLAlignmentHorizontalValues.Center;

            hoja.Range(
                    1,
                    1,
                    3,
                    totalColumnas)
                .Style.Alignment.Vertical =
                    XLAlignmentVerticalValues.Center;

            // =======================================================
            // FILTROS APLICADOS
            // =======================================================
            var filtros = new List<string>();

            if (!string.IsNullOrWhiteSpace(buscar))
            {
                filtros.Add($"Búsqueda: {buscar}");
            }

            if (!string.IsNullOrWhiteSpace(tipoDocumento))
            {
                filtros.Add(
                    $"Tipo de documento: {tipoDocumento}");
            }

            if (!string.IsNullOrWhiteSpace(estado))
            {
                filtros.Add($"Estado: {estado}");
            }

            if (!string.IsNullOrWhiteSpace(metodoPago))
            {
                filtros.Add($"Método: {metodoPago}");
            }

            if (proyectoId.HasValue &&
                proyectoId.Value > 0)
            {
                string? nombreProyecto =
                    await _context.Proyectos
                        .Where(p =>
                            p.IdProyecto ==
                            proyectoId.Value)
                        .Select(p => p.Nombre)
                        .FirstOrDefaultAsync();

                filtros.Add(
                    $"Proyecto: " +
                    $"{nombreProyecto ?? proyectoId.Value.ToString()}");
            }

            if (anio.HasValue)
            {
                filtros.Add($"Año: {anio.Value}");
            }

            if (mes.HasValue)
            {
                string nombreMes = new DateTime(
                    2000,
                    mes.Value,
                    1)
                    .ToString(
                        "MMMM",
                        new CultureInfo("es-CR"));

                nombreMes =
                    char.ToUpper(nombreMes[0]) +
                    nombreMes.Substring(1);

                filtros.Add($"Mes: {nombreMes}");
            }

            if (fechaInicio.HasValue)
            {
                filtros.Add(
                    $"Desde: {fechaInicio.Value:dd/MM/yyyy}");
            }

            if (fechaFin.HasValue)
            {
                filtros.Add(
                    $"Hasta: {fechaFin.Value:dd/MM/yyyy}");
            }

            hoja.Cell("A5").Value =
                "Filtros aplicados:";

            hoja.Cell("A5")
                .Style.Font.Bold = true;

            hoja.Cell("B5").Value = filtros.Any()
                ? string.Join(" | ", filtros)
                : "Sin filtros. Se muestran todos los pagos de cuentas por pagar.";

            hoja.Range(
                    5,
                    2,
                    5,
                    totalColumnas)
                .Merge();

            hoja.Range(
                    5,
                    2,
                    5,
                    totalColumnas)
                .Style.Alignment.WrapText = true;

            // =======================================================
            // ENCABEZADOS DE LA TABLA
            // =======================================================
            const int filaEncabezado = 7;

            string[] encabezados =
            {
        "Fecha de pago",
        "N.º documento",
        "Tipo de documento",
        "Origen",
        "Proveedor",
        "Proyecto",
        "Monto",
        "Método",
        "Referencia",
        "Observación",
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

            rangoEncabezado.Style.Alignment.Vertical =
                XLAlignmentVerticalValues.Center;

            rangoEncabezado.Style.Alignment.WrapText = true;

            rangoEncabezado.Style.Border.OutsideBorder =
                XLBorderStyleValues.Thin;

            rangoEncabezado.Style.Border.InsideBorder =
                XLBorderStyleValues.Thin;

            hoja.Row(filaEncabezado).Height = 25;

            // =======================================================
            // DATOS
            // =======================================================
            int fila =
                filaEncabezado + 1;

            foreach (var pago in pagos)
            {
                hoja.Cell(fila, 1).Value =
                    pago.FechaPago
                        .ToDateTime(TimeOnly.MinValue);

                hoja.Cell(fila, 2).Value =
                    pago.CuentaPorPagar?.NumeroDocumento
                    ?? "Sin documento";

                hoja.Cell(fila, 3).Value =
                    pago.CuentaPorPagar?.TipoDocumento
                    ?? "Sin tipo";

                hoja.Cell(fila, 4).Value =
                    pago.CuentaPorPagar?.Origen
                    ?? "Sin origen";

                hoja.Cell(fila, 5).Value =
                    pago.CuentaPorPagar?.Proveedor?.Nombre
                    ?? "Sin proveedor";

                hoja.Cell(fila, 6).Value =
                    pago.CuentaPorPagar?.Proyecto?.Nombre
                    ?? "Sin proyecto";

                hoja.Cell(fila, 7).Value =
                    pago.Monto ?? 0m;

                hoja.Cell(fila, 8).Value =
                    string.IsNullOrWhiteSpace(pago.MetodoPago)
                        ? "Sin método"
                        : pago.MetodoPago;

                hoja.Cell(fila, 9).Value =
                    string.IsNullOrWhiteSpace(pago.Referencia)
                        ? "Sin referencia"
                        : pago.Referencia;

                hoja.Cell(fila, 10).Value =
                    string.IsNullOrWhiteSpace(pago.Observacion)
                        ? "Sin observación"
                        : pago.Observacion;

                hoja.Cell(fila, 11).Value =
                    pago.Anulado
                        ? "Anulado"
                        : "Registrado";

                fila++;
            }

            int filaInicioDatos =
                filaEncabezado + 1;

            int filaFinDatos =
                fila - 1;

            if (pagos.Any())
            {
                var rangoDatos = hoja.Range(
                    filaInicioDatos,
                    1,
                    filaFinDatos,
                    totalColumnas);

                rangoDatos.Style.Border.OutsideBorder =
                    XLBorderStyleValues.Thin;

                rangoDatos.Style.Border.InsideBorder =
                    XLBorderStyleValues.Thin;

                rangoDatos.Style.Border.OutsideBorderColor =
                    XLColor.LightGray;

                rangoDatos.Style.Border.InsideBorderColor =
                    XLColor.LightGray;

                rangoDatos.Style.Alignment.Vertical =
                    XLAlignmentVerticalValues.Center;

                // Evita que las filas se deformen.
                rangoDatos.Style.Alignment.WrapText = false;

                // Fechas.
                hoja.Range(
                        filaInicioDatos,
                        1,
                        filaFinDatos,
                        1)
                    .Style.DateFormat.Format =
                        "dd/MM/yyyy";

                // Montos.
                hoja.Range(
                        filaInicioDatos,
                        7,
                        filaFinDatos,
                        7)
                    .Style.NumberFormat.Format =
                        "₡ #,##0.00";

                hoja.Range(
                        filaInicioDatos,
                        1,
                        filaFinDatos,
                        1)
                    .Style.Alignment.Horizontal =
                        XLAlignmentHorizontalValues.Center;

                hoja.Range(
                        filaInicioDatos,
                        7,
                        filaFinDatos,
                        7)
                    .Style.Alignment.Horizontal =
                        XLAlignmentHorizontalValues.Right;

                hoja.Range(
                        filaInicioDatos,
                        11,
                        filaFinDatos,
                        11)
                    .Style.Alignment.Horizontal =
                        XLAlignmentHorizontalValues.Center;

                // Autofiltros.
                hoja.Range(
                        filaEncabezado,
                        1,
                        filaFinDatos,
                        totalColumnas)
                    .SetAutoFilter();

                // Pagos anulados en rojo tenue.
                for (int indice = 0;
                     indice < pagos.Count;
                     indice++)
                {
                    if (pagos[indice].Anulado)
                    {
                        int filaPago =
                            filaInicioDatos + indice;

                        hoja.Range(
                                filaPago,
                                1,
                                filaPago,
                                totalColumnas)
                            .Style.Fill.BackgroundColor =
                                XLColor.FromHtml("#F8D7DA");

                        hoja.Range(
                                filaPago,
                                1,
                                filaPago,
                                totalColumnas)
                            .Style.Font.FontColor =
                                XLColor.FromHtml("#842029");
                    }
                }
            }
            else
            {
                hoja.Cell(filaInicioDatos, 1).Value =
                    "No se encontraron pagos con los filtros aplicados.";

                hoja.Range(
                        filaInicioDatos,
                        1,
                        filaInicioDatos,
                        totalColumnas)
                    .Merge();

                hoja.Range(
                        filaInicioDatos,
                        1,
                        filaInicioDatos,
                        totalColumnas)
                    .Style.Alignment.Horizontal =
                        XLAlignmentHorizontalValues.Center;

                hoja.Range(
                        filaInicioDatos,
                        1,
                        filaInicioDatos,
                        totalColumnas)
                    .Style.Font.Italic = true;

                hoja.Range(
                        filaInicioDatos,
                        1,
                        filaInicioDatos,
                        totalColumnas)
                    .Style.Fill.BackgroundColor =
                        XLColor.FromHtml("#F2F2F2");

                filaFinDatos =
                    filaInicioDatos;

                fila =
                    filaInicioDatos + 1;
            }

            // =======================================================
            // TOTALES
            // =======================================================
            decimal totalPagado = pagosActivos
                .Sum(p => p.Monto ?? 0m);

            int filaTotales =
                fila + 1;

            hoja.Cell(filaTotales, 6).Value =
                "TOTAL PAGADO:";

            hoja.Cell(filaTotales, 7).Value =
                totalPagado;

            hoja.Range(
                    filaTotales,
                    6,
                    filaTotales,
                    7)
                .Style.Fill.BackgroundColor =
                    XLColor.FromHtml("#D9EDEF");

            hoja.Range(
                    filaTotales,
                    6,
                    filaTotales,
                    7)
                .Style.Font.Bold = true;

            hoja.Cell(filaTotales, 7)
                .Style.NumberFormat.Format =
                    "₡ #,##0.00";

            hoja.Cell(filaTotales, 10).Value =
                "REGISTROS:";

            hoja.Cell(filaTotales, 11).Value =
                pagos.Count == 1
                    ? "1 registro"
                    : $"{pagos.Count} registros";

            hoja.Range(
                    filaTotales,
                    10,
                    filaTotales,
                    11)
                .Style.Fill.BackgroundColor =
                    XLColor.FromHtml("#D9EDEF");

            hoja.Range(
                    filaTotales,
                    10,
                    filaTotales,
                    11)
                .Style.Font.Bold = true;

            // =======================================================
            // RESUMEN INFERIOR
            // =======================================================
            int filaResumen =
                filaTotales + 3;

            hoja.Cell(filaResumen, 1).Value =
                "Resumen del reporte";

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

            hoja.Range(
                    filaResumen,
                    1,
                    filaResumen,
                    5)
                .Style.Alignment.Horizontal =
                    XLAlignmentHorizontalValues.Center;

            string[] resumenTitulos =
            {
        "Registrados",
        "Anulados",
        "Efectivo",
        "Transferencias",
        "SINPE Móvil"
    };

            int[] resumenValores =
            {
        pagosActivos.Count,

        pagos.Count(p =>
            p.Anulado),

        pagosActivos.Count(p =>
            p.MetodoPago == "Efectivo"),

        pagosActivos.Count(p =>
            p.MetodoPago == "Transferencia"),

        pagosActivos.Count(p =>
            p.MetodoPago == "SINPE Móvil")
    };

            for (int columna = 0;
                 columna < 5;
                 columna++)
            {
                hoja.Cell(
                    filaResumen + 1,
                    columna + 1).Value =
                        resumenTitulos[columna];

                hoja.Cell(
                    filaResumen + 2,
                    columna + 1).Value =
                        resumenValores[columna];
            }

            var rangoTitulosResumen = hoja.Range(
                filaResumen + 1,
                1,
                filaResumen + 1,
                5);

            rangoTitulosResumen.Style.Font.Bold = true;

            rangoTitulosResumen.Style.Fill.BackgroundColor =
                XLColor.FromHtml("#D9EDEF");

            var rangoResumen = hoja.Range(
                filaResumen + 1,
                1,
                filaResumen + 2,
                5);

            rangoResumen.Style.Alignment.Horizontal =
                XLAlignmentHorizontalValues.Center;

            rangoResumen.Style.Border.OutsideBorder =
                XLBorderStyleValues.Thin;

            rangoResumen.Style.Border.InsideBorder =
                XLBorderStyleValues.Thin;

            // =======================================================
            // ANCHOS DE COLUMNAS
            // =======================================================
            hoja.Column(1).Width = 16;
            hoja.Column(2).Width = 18;
            hoja.Column(3).Width = 20;
            hoja.Column(4).Width = 13;
            hoja.Column(5).Width = 30;
            hoja.Column(6).Width = 28;
            hoja.Column(7).Width = 18;
            hoja.Column(8).Width = 18;
            hoja.Column(9).Width = 22;
            hoja.Column(10).Width = 32;
            hoja.Column(11).Width = 15;

            if (pagos.Any())
            {
                for (int numeroFila = filaInicioDatos;
                     numeroFila <= filaFinDatos;
                     numeroFila++)
                {
                    hoja.Row(numeroFila).Height = 20;
                }
            }

            hoja.SheetView.FreezeRows(
                filaEncabezado);

            // =======================================================
            // CONFIGURACIÓN DE IMPRESIÓN
            // =======================================================
            hoja.PageSetup.PageOrientation =
                XLPageOrientation.Landscape;

            hoja.PageSetup.PaperSize =
                XLPaperSize.A4Paper;

            hoja.PageSetup.FitToPages(1, 0);

            hoja.PageSetup.Margins.Top = 0.4;
            hoja.PageSetup.Margins.Bottom = 0.4;
            hoja.PageSetup.Margins.Left = 0.25;
            hoja.PageSetup.Margins.Right = 0.25;

            hoja.PageSetup.CenterHorizontally = true;

            hoja.PageSetup.SetRowsToRepeatAtTop(
                filaEncabezado,
                filaEncabezado);

            hoja.Range(
                    1,
                    1,
                    filaResumen + 2,
                    totalColumnas)
                .Style.Font.FontName =
                    "Arial";

            // =======================================================
            // GENERAR ARCHIVO
            // =======================================================
            using var stream =
                new MemoryStream();

            workbook.SaveAs(stream);

            string nombreArchivo =
                $"Historial_Pagos_CxP_" +
                $"{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";

            return File(
                stream.ToArray(),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                nombreArchivo);
        }
        // GET detalles
        public async Task<IActionResult> Details(int? id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var pago = await _context.PagosCuentaPorPagars
                .Include(p => p.CuentaPorPagar)
                    .ThenInclude(c => c.Proveedor)
                .FirstOrDefaultAsync(p => p.IdPagoCuentaPorPagar == id);

            if (pago == null)
            {
                return NotFound();
            }

            return View(pago);
        }
        // GET: Crear pago de cuenta por pagar
        // Este método solo permite abrir el formulario si el pago viene desde una CxP específica.
        public async Task<IActionResult> Create(int? cuentaPorPagarId)
        {
            if (cuentaPorPagarId == null || cuentaPorPagarId == 0)
            {
                TempData["MensajeError"] = "El pago debe registrarse desde una cuenta por pagar específica.";
                return RedirectToAction("Index", "CuentasPorPagar");
            }

            var cuenta = await _context.CuentasPorPagars
                .Include(c => c.Proveedor)
                .FirstOrDefaultAsync(c => c.IdCuentaPorPagar == cuentaPorPagarId);

            if (cuenta == null)
            {
                return NotFound();
            }

            // No se permiten pagos a cuentas anuladas, rechazadas, canceladas o sin saldo.
            if (cuenta.Activo == false ||
                cuenta.Estado == "Anulada" ||
                cuenta.Estado == "Rechazada" ||
                cuenta.Estado == "Cancelada" ||
                cuenta.SaldoPendiente <= 0)
            {
                TempData["MensajeError"] = "No se puede registrar un pago para esta cuenta por pagar.";
                return RedirectToAction("Index", "CuentasPorPagar");
            }

            // Si la CxP nació manualmente, primero debe tener su gasto aprobado.
            // Si nació desde factura, se permite pagar directamente.
            bool esOrigenManual = cuenta.Origen == "Manual";

            if (esOrigenManual)
            {
                var gasto = await _context.Gastos
                   .FirstOrDefaultAsync(g =>
        g.NumeroFactura == cuenta.NumeroDocumento ||
        g.Concepto.Contains(cuenta.NumeroDocumento));

                if (gasto == null || gasto.Estado != "Aprobado")
                {
                    TempData["MensajeError"] = "No se puede registrar el pago porque el gasto asociado aún está pendiente de autorización.";
                    return RedirectToAction("Index", "CuentasPorPagar");
                }
            }

            ViewBag.CuentaResumen = cuenta;

            ViewBag.MetodosPago = new SelectList(new List<string>
    {
        "Efectivo",
        "Transferencia",
        "Depósito",
        "SINPE Móvil",
        "Cheque"
    });
            // cargar select de cuentas disponibles para pago 
            var cuentasResumen = await ObtenerCuentasResumen();

            ViewBag.CuentaPorPagarId = new SelectList(
                _context.CuentasPorPagars
                    .Where(c => c.Activo == true &&
                                c.Estado != "Anulada" &&
                                c.Estado != "Cancelada" &&
                                c.SaldoPendiente > 0),
                "IdCuentaPorPagar",
                "NumeroDocumento",
                cuenta.IdCuentaPorPagar);

            ViewBag.CuentasResumenJson = JsonSerializer.Serialize(cuentasResumen);

            return View(new PagosCuentaPorPagar
            {
                CuentaPorPagarId = cuenta.IdCuentaPorPagar,
                FechaPago = DateOnly.FromDateTime(DateTime.Today),
                Anulado = false
            });
        }

        // POST: Crear pago de cuenta por pagar
        // Registra el pago, actualiza saldo de CxP y sincroniza factura si corresponde.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Create([Bind("CuentaPorPagarId,Monto,MetodoPago,Referencia,Observacion")] PagosCuentaPorPagar pago)
        {
            var cuenta = await _context.CuentasPorPagars
                .Include(c => c.Proveedor)
                .FirstOrDefaultAsync(c => c.IdCuentaPorPagar == pago.CuentaPorPagarId);

            if (cuenta == null)
            {
                ModelState.AddModelError("CuentaPorPagarId", "Debe seleccionar una cuenta por pagar válida.");
            }
            else
            {
                // Valida que la cuenta permita pagos.
                if (cuenta.Activo == false ||
                    cuenta.Estado == "Anulada" ||
                    cuenta.Estado == "Rechazada" ||
                    cuenta.Estado == "Cancelada")
                {
                    ModelState.AddModelError("CuentaPorPagarId", "No se puede registrar un pago para esta cuenta por pagar.");
                }

                if (cuenta.SaldoPendiente <= 0)
                {
                    ModelState.AddModelError("CuentaPorPagarId", "La cuenta seleccionada no tiene saldo pendiente.");
                }

                if ((pago.Monto ?? 0) <= 0)
                {
                    ModelState.AddModelError("Monto", "El monto del pago debe ser mayor a cero.");
                }

                if ((pago.Monto ?? 0) > cuenta.SaldoPendiente)
                {
                    ModelState.AddModelError("Monto", "El monto del pago no puede ser mayor al saldo pendiente.");
                }

                // Si la cuenta nació manualmente en CxP, debe esperar aprobación del gasto.
                // Las CxP nacidas desde factura sí pueden pagarse directamente.
                bool esOrigenManual = cuenta.Origen == "Manual";

                if (esOrigenManual)
                {
                    var gasto = await _context.Gastos
                        .FirstOrDefaultAsync(g =>
    g.NumeroFactura == cuenta.NumeroDocumento ||
    g.Concepto.Contains(cuenta.NumeroDocumento));

                    if (gasto == null || gasto.Estado != "Aprobado")
                    {
                        ModelState.AddModelError("CuentaPorPagarId", "No se puede registrar el pago porque el gasto asociado aún está pendiente de autorización.");
                    }
                }
            }

            pago.FechaPago = DateOnly.FromDateTime(DateTime.Today);
            pago.Anulado = false;

            ModelState.Remove("FechaPago");
            ModelState.Remove("Anulado");
            ModelState.Remove("CuentaPorPagar");

            if (ModelState.IsValid)
            {
                _context.Add(pago);
                await _context.SaveChangesAsync();

                cuenta.MontoPagado = (cuenta.MontoPagado ?? 0) + (pago.Monto ?? 0);

                // Calcula el saldo real de la CxP.
                decimal totalCuenta =
                    (cuenta.SubTotal ?? 0)
                    + (cuenta.Impuesto ?? 0);

                cuenta.SaldoPendiente =
                    totalCuenta
                    - (cuenta.AnticipoAplicado ?? 0)
                    - (cuenta.MontoPagado ?? 0);

                if (cuenta.SaldoPendiente < 0)
                {
                    cuenta.SaldoPendiente = 0;
                }

                ActualizarEstadoCuenta(cuenta);

                _context.Update(cuenta);

                // Si la CxP nació desde factura, sincroniza el saldo y estado de la factura.
                if (cuenta.TipoDocumento == "Factura" && !string.IsNullOrWhiteSpace(cuenta.NumeroFactura))
                {
                    var factura = await _context.Facturas
                        .FirstOrDefaultAsync(f => f.NumeroFactura == cuenta.NumeroFactura);

                    if (factura != null)
                    {
                        factura.MontoPagado = cuenta.MontoPagado ?? 0;
                        factura.SaldoPendiente = cuenta.SaldoPendiente;

                        if (factura.SaldoPendiente <= 0)
                        {
                            factura.SaldoPendiente = 0;
                            factura.Estado = "Pagada";
                        }
                        else if (factura.MontoPagado > 0)
                        {
                            factura.Estado = "Parcial";
                        }
                        else if (factura.FechaVencimiento < DateOnly.FromDateTime(DateTime.Today))
                        {
                            factura.Estado = "Vencida";
                        }
                        else
                        {
                            factura.Estado = "Pendiente";
                        }

                        _context.Update(factura);
                    }
                }

                await _context.SaveChangesAsync();

                // Genera movimiento contable por el pago de la cuenta por pagar.
                // Debe: Cuentas por Pagar
                // Haber: Caja General
                await RegistrarMovimientoContablePagoCxP(pago, cuenta);

                // Auditoría del pago registrado.
                await RegistrarAuditoria(
                    "Crear",
                    pago.IdPagoCuentaPorPagar,
                    "Se registró un pago CxP por ₡" +
                    (pago.Monto ?? 0).ToString("N2") +
                    " aplicado al documento " +
                    cuenta.NumeroDocumento
                );

                TempData["MensajeExito"] = "El pago fue registrado correctamente.";
                return RedirectToAction(nameof(Index));
            }

            var cuentasResumen = await ObtenerCuentasResumen();

            ViewBag.CuentaPorPagarId = new SelectList(
                _context.CuentasPorPagars
                    .Where(c => c.Activo == true &&
                                c.Estado != "Anulada" &&
                                c.Estado != "Cancelada" &&
                                c.SaldoPendiente > 0),
                "IdCuentaPorPagar",
                "NumeroDocumento",
                pago.CuentaPorPagarId);

            ViewBag.CuentasResumenJson = JsonSerializer.Serialize(cuentasResumen);

            ViewBag.MetodosPago = new SelectList(new List<string>
    {
        "Efectivo",
        "Transferencia",
        "Depósito",
        "SINPE Móvil",
        "Cheque"
    }, pago.MetodoPago);



            return View(pago);
        }
        private async Task<List<object>> ObtenerCuentasResumen()
        {
            var hoy = DateOnly.FromDateTime(DateTime.Today);

            return await _context.CuentasPorPagars
                .Include(c => c.Proveedor)
                .Where(c => c.Activo == true &&
                            c.Estado != "Anulada" &&
                            c.Estado != "Cancelada" &&
                            c.SaldoPendiente > 0)
                .Select(c => new
                {
                    id = c.IdCuentaPorPagar,
                    proveedor = c.Proveedor != null ? c.Proveedor.Nombre : "Sin proveedor",
                    documento = c.NumeroDocumento,
                    tipo = c.TipoDocumento,
                    fechaEmision = c.FechaEmision.ToString("dd/MM/yyyy"),
                    fechaVencimiento = c.FechaVencimiento.ToString("dd/MM/yyyy"),
                    diasCredito = c.DiasCredito,
                    diasRestantes = c.FechaVencimiento.DayNumber - hoy.DayNumber,
                    montoOriginal = c.MontoOriginal,
                    montoPagado = c.MontoPagado ?? 0,
                    saldoPendiente = c.SaldoPendiente,
                    estado = c.Estado,
                    pagosRealizados = _context.PagosCuentaPorPagars
                        .Count(p => p.CuentaPorPagarId == c.IdCuentaPorPagar && p.Anulado == false),
                    ultimoPago = _context.PagosCuentaPorPagars
                        .Where(p => p.CuentaPorPagarId == c.IdCuentaPorPagar && p.Anulado == false)
                        .OrderByDescending(p => p.FechaPago)
                        .Select(p => p.FechaPago.ToString("dd/MM/yyyy"))
                        .FirstOrDefault()
                })
                .ToListAsync<object>();
        }

        //GET anular
        public async Task<IActionResult> Delete(int? id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var pago = await _context.PagosCuentaPorPagars
                .Include(p => p.CuentaPorPagar)
                    .ThenInclude(c => c.Proveedor)
                .FirstOrDefaultAsync(p => p.IdPagoCuentaPorPagar == id);

            if (pago == null)
            {
                return NotFound();
            }
            // ============================================================
            // BLOQUEO POR CIERRE CONTABLE DEL PAGO
            // ============================================================
            bool periodoCerrado =
                await PeriodoContableCerradoAsync(pago.FechaPago);

            if (periodoCerrado)
            {
                TempData["MensajeError"] =
                    $"No se puede anular el pago porque el período " +
                    $"{pago.FechaPago:MM/yyyy} se encuentra cerrado. " +
                    $"Para realizar correcciones, primero debe reabrirse el período.";

                return RedirectToAction(
                    nameof(Index),
                    new
                    {
                        anio = pago.FechaPago.Year,
                        mes = pago.FechaPago.Month
                    });
            }
            if (pago.Anulado)
            {
                TempData["MensajeError"] = "Este pago ya se encuentra anulado.";
                return RedirectToAction(nameof(Index));
            }

            return View(pago);
        }

        // POST: Anular pago de cuenta por pagar
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> DeleteConfirmed(int id, string? motivoAnulacion)
        {
            var pago = await _context.PagosCuentaPorPagars
                .Include(p => p.CuentaPorPagar)
                .FirstOrDefaultAsync(p => p.IdPagoCuentaPorPagar == id);

            if (pago == null)
            {
                return NotFound();
            }
            // ============================================================
            // BLOQUEO POR CIERRE CONTABLE DEL PAGO
            // PROTECCIÓN CONTRA POST DIRECTO
            // ============================================================
            bool periodoCerrado =
                await PeriodoContableCerradoAsync(pago.FechaPago);

            if (periodoCerrado)
            {
                TempData["MensajeError"] =
                    $"No se puede anular el pago porque el período " +
                    $"{pago.FechaPago:MM/yyyy} se encuentra cerrado. " +
                    $"Para realizar correcciones, primero debe reabrirse el período.";

                return RedirectToAction(
                    nameof(Index),
                    new
                    {
                        anio = pago.FechaPago.Year,
                        mes = pago.FechaPago.Month
                    });
            }
            if (string.IsNullOrWhiteSpace(motivoAnulacion))
            {
                ModelState.AddModelError("motivoAnulacion", "Debe ingresar el motivo de anulación.");
                pago.CuentaPorPagar = await _context.CuentasPorPagars
       .Include(c => c.Proveedor)
       .FirstOrDefaultAsync(c => c.IdCuentaPorPagar == pago.CuentaPorPagarId);
                return View("Delete", pago);
            }

            if (!System.Text.RegularExpressions.Regex.IsMatch(
                motivoAnulacion,
                @"^[a-zA-ZáéíóúÁÉÍÓÚñÑ0-9\s.,;:/()#-]*$"))
            {
                ModelState.AddModelError("motivoAnulacion", "El motivo de anulación contiene caracteres no permitidos.");
                return View("Delete", pago);
            }

            if (motivoAnulacion.Trim().Length < 5)
            {
                ModelState.AddModelError("motivoAnulacion", "El motivo de anulación debe contener al menos 5 caracteres.");
                return View("Delete", pago);
            }

            if (!motivoAnulacion.Any(char.IsLetter))
            {
                ModelState.AddModelError("motivoAnulacion", "El motivo de anulación debe contener al menos una letra.");
                return View("Delete", pago);
            }

            if (pago.Anulado)
            {
                TempData["MensajeError"] = "Este pago ya se encuentra anulado.";
                return RedirectToAction(nameof(Index));
            }

            var cuenta = await _context.CuentasPorPagars
                .FirstOrDefaultAsync(c => c.IdCuentaPorPagar == pago.CuentaPorPagarId);

            if (cuenta == null)
            {
                return NotFound();
            }

            pago.Anulado = true;
            pago.MotivoAnulacion = motivoAnulacion.Trim();

            cuenta.MontoPagado = (cuenta.MontoPagado ?? 0) - (pago.Monto ?? 0);

            if (cuenta.MontoPagado < 0)
            {
                cuenta.MontoPagado = 0;
            }

            // inclusion de impuesto en calculo 
            decimal totalCuenta =
       (cuenta.SubTotal ?? 0)
       + (cuenta.Impuesto ?? 0);

            cuenta.SaldoPendiente =
                totalCuenta
                - (cuenta.AnticipoAplicado ?? 0)
                - (cuenta.MontoPagado ?? 0);

            if (cuenta.SaldoPendiente < 0)
            {
                cuenta.SaldoPendiente = 0;
            }

            ActualizarEstadoCuenta(cuenta);

            _context.Update(pago);
            _context.Update(cuenta);

            // Sincroniza la factura de compra crédito asociada a la CxP.
            // Si se anula un pago CxP, la factura debe recuperar saldo pendiente.
            var factura = await _context.Facturas
                .FirstOrDefaultAsync(f =>
                    f.NumeroFactura == cuenta.NumeroFactura ||
                    f.NumeroFactura == cuenta.NumeroDocumento);

            if (factura != null)
            {
                factura.MontoPagado = cuenta.MontoPagado ?? 0;
                factura.SaldoPendiente = cuenta.SaldoPendiente;

                if (factura.SaldoPendiente <= 0)
                {
                    factura.SaldoPendiente = 0;
                    factura.Estado = "Pagada";
                }
                else if (factura.MontoPagado > 0)
                {
                    factura.Estado = "Parcial";
                }
                else if (factura.FechaVencimiento < DateOnly.FromDateTime(DateTime.Today))
                {
                    factura.Estado = "Vencida";
                }
                else
                {
                    factura.Estado = "Pendiente";
                }

                _context.Update(factura);
            }

            await _context.SaveChangesAsync();

            // Genera reversión contable por la anulación del pago CxP.
            // Debe: Caja General
            // Haber: Cuentas por Pagar
            await RevertirMovimientoContablePagoCxP(pago, cuenta);

            //AUDITORIA
            await RegistrarAuditoria(
    "Anular",
    pago.IdPagoCuentaPorPagar,
    "Se anuló el pago CxP del documento " +
    cuenta.NumeroDocumento +
    ". Motivo: " + motivoAnulacion.Trim()
);
            TempData["MensajeExito"] = "El pago fue anulado correctamente y la cuenta por pagar fue actualizada.";
            return RedirectToAction(nameof(Index));
        }
        //actualizar estado pcxp
        private void ActualizarEstadoCuenta(CuentasPorPagar cuenta)
        {
            if (cuenta.SaldoPendiente <= 0)
            {
                cuenta.SaldoPendiente = 0;
                cuenta.Estado = "Cancelada";
            }
            else if (cuenta.FechaVencimiento < DateOnly.FromDateTime(DateTime.Today))
            {
                cuenta.Estado = "Vencida";
            }
            else if ((cuenta.MontoPagado ?? 0) > 0)
            {
                cuenta.Estado = "Parcial";
            }
            else
            {
                cuenta.Estado = "Pendiente";
            }
        }
        // Registra el movimiento contable cuando se paga una CxP.
        // Debe: Cuentas por Pagar
        // Haber: Caja General
        private async Task RegistrarMovimientoContablePagoCxP(
            PagosCuentaPorPagar pago,
            CuentasPorPagar cuenta)
        {
            decimal monto = pago.Monto ?? 0;

            if (monto <= 0)
            {
                return;
            }

            string referencia = "PAGO-CXP-" + pago.IdPagoCuentaPorPagar;

            bool yaExiste = await _context.MovimientosContables
                .AnyAsync(m =>
                    m.Referencia == referencia &&
                    m.Estado != "Anulado");

            if (yaExiste)
            {
                return;
            }

            var usuarioSesion = HttpContext.Session.GetString("Usuario");

            var usuario = await _context.Usuarios
                .FirstOrDefaultAsync(u => u.Nombre == usuarioSesion);

            var cuentaCaja = await _context.CuentasContables
                .FirstOrDefaultAsync(c => c.Codigo == "CC001");

            var cuentaCxP = await _context.CuentasContables
                .FirstOrDefaultAsync(c => c.Codigo == "CC002");

            if (cuentaCaja == null || cuentaCxP == null)
            {
                return;
            }

            // Debe: Cuentas por Pagar
            var movimientoDebe = new MovimientosContable
            {
                Fecha = pago.FechaPago,
                CuentaContableId = cuentaCxP.IdCuentaContable,
                TipoMovimiento = "Debe",
                Monto = monto,
                OrigenModulo = "Pagos CxP",
                OrigenId = pago.IdPagoCuentaPorPagar,
                Referencia = referencia,
                Descripcion = "Pago aplicado a cuenta por pagar " + cuenta.NumeroDocumento,
                ProyectoId = cuenta.ProyectoId,
                CentroCostoId = cuenta.CentroCostoId,
                Debe = monto,
                Haber = 0,
                Estado = "Registrado",
                UsuarioId = usuario != null ? usuario.IdUsuario : null
            };

            // Haber: Caja General
            var movimientoHaber = new MovimientosContable
            {
                Fecha = pago.FechaPago,
                CuentaContableId = cuentaCaja.IdCuentaContable,
                TipoMovimiento = "Haber",
                Monto = monto,
                OrigenModulo = "Pagos CxP",
                OrigenId = pago.IdPagoCuentaPorPagar,
                Referencia = referencia,
                Descripcion = "Salida de caja por pago de cuenta por pagar " + cuenta.NumeroDocumento,
                ProyectoId = cuenta.ProyectoId,
                CentroCostoId = cuenta.CentroCostoId,
                Debe = 0,
                Haber = monto,
                Estado = "Registrado",
                UsuarioId = usuario != null ? usuario.IdUsuario : null
            };

            _context.MovimientosContables.Add(movimientoDebe);
            _context.MovimientosContables.Add(movimientoHaber);

            // Actualiza saldos contables.
            // CxP disminuye y Caja disminuye.
            cuentaCxP.SaldoActual = cuentaCxP.SaldoActual - monto;
            cuentaCaja.SaldoActual = cuentaCaja.SaldoActual - monto;

            if (cuentaCxP.SaldoActual < 0)
            {
                cuentaCxP.SaldoActual = 0;
            }

            _context.CuentasContables.Update(cuentaCxP);
            _context.CuentasContables.Update(cuentaCaja);

            await _context.SaveChangesAsync();
        }


        // Registra la reversión contable cuando se anula un pago de CxP.
        // Debe: Caja General
        // Haber: Cuentas por Pagar
        private async Task RevertirMovimientoContablePagoCxP(
            PagosCuentaPorPagar pago,
            CuentasPorPagar cuenta)
        {
            decimal monto = pago.Monto ?? 0;

            if (monto <= 0)
            {
                return;
            }

            string referenciaOriginal = "PAGO-CXP-" + pago.IdPagoCuentaPorPagar;
            string referenciaReversa = "REV-PAGO-CXP-" + pago.IdPagoCuentaPorPagar;

            bool reversaYaExiste = await _context.MovimientosContables
                .AnyAsync(m => m.Referencia == referenciaReversa);

            if (reversaYaExiste)
            {
                return;
            }

            var usuarioSesion = HttpContext.Session.GetString("Usuario");

            var usuario = await _context.Usuarios
                .FirstOrDefaultAsync(u => u.Nombre == usuarioSesion);

            var cuentaCaja = await _context.CuentasContables
                .FirstOrDefaultAsync(c => c.Codigo == "CC001");

            var cuentaCxP = await _context.CuentasContables
                .FirstOrDefaultAsync(c => c.Codigo == "CC002");

            if (cuentaCaja == null || cuentaCxP == null)
            {
                return;
            }

            // Marca los movimientos originales como anulados.
            var movimientosOriginales = await _context.MovimientosContables
                .Where(m =>
                    m.Referencia == referenciaOriginal &&
                    m.Estado != "Anulado")
                .ToListAsync();

            foreach (var movimiento in movimientosOriginales)
            {
                movimiento.Estado = "Anulado";
                movimiento.Anulado = true;
                movimiento.FechaAnulacion = DateTime.Now;
                movimiento.MotivoAnulacion = "Anulación del pago CxP " + cuenta.NumeroDocumento;
                _context.MovimientosContables.Update(movimiento);
            }

            // Debe: Caja General
            var movimientoDebe = new MovimientosContable
            {
                Fecha = DateOnly.FromDateTime(DateTime.Today),
                CuentaContableId = cuentaCaja.IdCuentaContable,
                TipoMovimiento = "Debe",
                Monto = monto,
                OrigenModulo = "Pagos CxP",
                OrigenId = pago.IdPagoCuentaPorPagar,
                Referencia = referenciaReversa,
                Descripcion = "Reversión por anulación de pago CxP " + cuenta.NumeroDocumento,
                ProyectoId = cuenta.ProyectoId,
                CentroCostoId = cuenta.CentroCostoId,
                Debe = monto,
                Haber = 0,
                Estado = "Reversado",
                UsuarioId = usuario != null ? usuario.IdUsuario : null
            };

            // Haber: Cuentas por Pagar
            var movimientoHaber = new MovimientosContable
            {
                Fecha = DateOnly.FromDateTime(DateTime.Today),
                CuentaContableId = cuentaCxP.IdCuentaContable,
                TipoMovimiento = "Haber",
                Monto = monto,
                OrigenModulo = "Pagos CxP",
                OrigenId = pago.IdPagoCuentaPorPagar,
                Referencia = referenciaReversa,
                Descripcion = "Aumento de cuenta por pagar por reversión de pago " + cuenta.NumeroDocumento,
                ProyectoId = cuenta.ProyectoId,
                CentroCostoId = cuenta.CentroCostoId,
                Debe = 0,
                Haber = monto,
                Estado = "Reversado",
                UsuarioId = usuario != null ? usuario.IdUsuario : null
            };
            _context.MovimientosContables.Add(movimientoDebe);
            _context.MovimientosContables.Add(movimientoHaber);

            // Primero guarda las reversas para que tengan IdMovimientoContable.
            await _context.SaveChangesAsync();

            foreach (var movimiento in movimientosOriginales)
            {
                if (movimiento.TipoMovimiento == "Debe")
                {
                    movimiento.MovimientoReversionId = movimientoDebe.IdMovimientoContable;
                }
                else if (movimiento.TipoMovimiento == "Haber")
                {
                    movimiento.MovimientoReversionId = movimientoHaber.IdMovimientoContable;
                }

                _context.MovimientosContables.Update(movimiento);
            }

            // Actualiza saldos contables por reversión.
            // Caja vuelve a aumentar y CxP vuelve a aumentar.
            cuentaCaja.SaldoActual = cuentaCaja.SaldoActual + monto;
            cuentaCxP.SaldoActual = cuentaCxP.SaldoActual + monto;

            _context.CuentasContables.Update(cuentaCaja);
            _context.CuentasContables.Update(cuentaCxP);

            await _context.SaveChangesAsync();
        }
        private void CargarCombosCxP(int? cuentaSeleccionada = null, string? metodoSeleccionado = null)
        {
            ViewData["CuentaPorPagarId"] = new SelectList(
                _context.CuentasPorPagars
                    .Include(c => c.Proveedor)
                    .Where(c => c.Activo == true &&
                                c.Estado != "Anulada" &&
                                c.Estado != "Cancelada" &&
                                c.SaldoPendiente > 0),
                "IdCuentaPorPagar",
                "NumeroDocumento",
                cuentaSeleccionada
            );

            ViewBag.MetodosPago = new SelectList(new List<string>
            {
                "Efectivo",
                "Transferencia",
                "Depósito",
                "SINPE Móvil",
                "Cheque"
            }, metodoSeleccionado);

            ViewBag.CuentasResumen = _context.CuentasPorPagars
                .Include(c => c.Proveedor)
                .Where(c => c.Activo == true &&
                            c.Estado != "Anulada" &&
                            c.Estado != "Cancelada" &&
                            c.SaldoPendiente > 0)
                .Select(c => new
                {
                    id = c.IdCuentaPorPagar,
                    proveedor = c.Proveedor.Nombre,
                    documento = c.NumeroDocumento,
                    tipo = c.TipoDocumento,
                    diasCredito = c.DiasCredito,
                    fechaEmision = c.FechaEmision.ToString("dd/MM/yyyy"),
                    fechaVencimiento = c.FechaVencimiento.ToString("dd/MM/yyyy"),
                    //nota:ultima modificacion
                    montoOriginal = c.MontoOriginal,
                    montoTotal = (c.SubTotal ?? 0) + (c.Impuesto ?? 0),
                    montoPagado = c.MontoPagado ?? 0,
                    saldoPendiente = c.SaldoPendiente,
                    estado = c.Estado,
                    diasRestantes = c.FechaVencimiento.DayNumber - DateOnly.FromDateTime(DateTime.Today).DayNumber,

                    pagosRealizados = c.PagosCuentaPorPagars
                        .Where(p => p.Anulado == false)
                        .Count(),

                    ultimoPago = c.PagosCuentaPorPagars
                        .Where(p => p.Anulado == false)
                        .OrderByDescending(p => p.FechaPago)
                        .Select(p => p.FechaPago.ToString("dd/MM/yyyy"))
                        .FirstOrDefault()
                })
                .ToList();
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
        // Registra las acciones realizadas sobre pagos de cuentas por pagar.
        private async Task RegistrarAuditoria(string accion, int registroId, string descripcion)
    {
        var usuarioSesion = HttpContext.Session.GetString("Usuario");

        var usuario = await _context.Usuarios
            .FirstOrDefaultAsync(u => u.Nombre == usuarioSesion);

        if (usuario != null)
        {
            var auditoria = new Auditorium
            {
                UsuarioId = usuario.IdUsuario,
                Tabla = "PagosCuentaPorPagar",
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
    
}