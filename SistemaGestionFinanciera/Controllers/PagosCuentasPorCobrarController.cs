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
using System.Text.Json;
using ClosedXML.Excel;
using System.IO;
using System.Globalization;

namespace SistemaGestionFinanciera.Controllers
{
    public class PagosCuentasPorCobrarController : BaseController
    {
        private readonly SistemaFinancieroContext _context;

        public PagosCuentasPorCobrarController(SistemaFinancieroContext context)
        {
            _context = context;
        }

        // LISTADO DE PAGOS DE CUENTAS POR COBRAR
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

            // Sin filtros, las cards muestran únicamente el mes actual.
            if (!hayFiltros)
            {
                var hoy = DateOnly.FromDateTime(DateTime.Today);
                var inicioMes = new DateOnly(hoy.Year, hoy.Month, 1);
                var finMes = inicioMes.AddMonths(1).AddDays(-1);

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

            ViewBag.CantidadPagos = pagosParaTotales.Count;

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

            // Conserva filtros.
            ViewBag.Buscar = buscar;
            ViewBag.TipoDocumento = tipoDocumento;
            ViewBag.Estado = estado;
            ViewBag.MetodoPago = metodoPago;
            ViewBag.ProyectoIdValor = proyectoId;
            ViewBag.Anio = anio;
            ViewBag.Mes = mes;
            ViewBag.FechaInicio = fechaInicio?.ToString("yyyy-MM-dd");
            ViewBag.FechaFin = fechaFin?.ToString("yyyy-MM-dd");

            ViewBag.ProyectoId = new SelectList(
                await _context.Proyectos
                    .Where(p => p.Activo == true)
                    .OrderBy(p => p.Nombre)
                    .ToListAsync(),
                "IdProyecto",
                "Nombre",
                proyectoId);

            ViewBag.Anios = await _context.PagosCuentaPorCobrars
                .Select(p => p.FechaPago.Year)
                .Distinct()
                .OrderByDescending(a => a)
                .ToListAsync();

            return View(listaPagos);
        }
        // ===========================================================
        // CONSULTA COMPARTIDA PARA INDEX, REPORTE Y EXCEL
        // ===========================================================
        private async Task<List<PagosCuentaPorCobrar>>
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
            var consulta = _context.PagosCuentaPorCobrars
                .Include(p => p.CuentaPorCobrar)
                    .ThenInclude(c => c.ClienteBeneficiario)
                .Include(p => p.CuentaPorCobrar)
                    .ThenInclude(c => c.Proyecto)
                .AsQueryable();

            // Documento, referencia, observación o cliente.
            if (!string.IsNullOrWhiteSpace(buscar))
            {
                buscar = buscar.Trim();

                consulta = consulta.Where(p =>
                    (p.Referencia != null &&
                     p.Referencia.Contains(buscar)) ||

                    (p.Observacion != null &&
                     p.Observacion.Contains(buscar)) ||

                    (p.CuentaPorCobrar != null &&
                     p.CuentaPorCobrar.NumeroDocumento != null &&
                     p.CuentaPorCobrar.NumeroDocumento.Contains(buscar)) ||

                    (p.CuentaPorCobrar != null &&
                     p.CuentaPorCobrar.ClienteBeneficiario != null &&
                     p.CuentaPorCobrar.ClienteBeneficiario.Nombre
                        .Contains(buscar)));
            }

            if (!string.IsNullOrWhiteSpace(tipoDocumento))
            {
                consulta = consulta.Where(p =>
                    p.CuentaPorCobrar != null &&
                    p.CuentaPorCobrar.TipoDocumento ==
                        tipoDocumento);
            }

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

            if (!string.IsNullOrWhiteSpace(metodoPago))
            {
                consulta = consulta.Where(p =>
                    p.MetodoPago == metodoPago);
            }

            if (proyectoId.HasValue &&
                proyectoId.Value > 0)
            {
                consulta = consulta.Where(p =>
                    p.CuentaPorCobrar != null &&
                    p.CuentaPorCobrar.ProyectoId ==
                        proyectoId.Value);
            }

            if (anio.HasValue)
            {
                consulta = consulta.Where(p =>
                    p.FechaPago.Year == anio.Value);
            }

            if (mes.HasValue)
            {
                consulta = consulta.Where(p =>
                    p.FechaPago.Month == mes.Value);
            }

            if (fechaInicio.HasValue)
            {
                consulta = consulta.Where(p =>
                    p.FechaPago >= fechaInicio.Value);
            }

            if (fechaFin.HasValue)
            {
                consulta = consulta.Where(p =>
                    p.FechaPago <= fechaFin.Value);
            }

            return await consulta
                .OrderByDescending(p => p.FechaPago)
                .ThenByDescending(p =>
                    p.IdPagoCuentaPorCobrar)
                .ToListAsync();
        }
        // ===========================================================
        // VISTA PREVIA DEL HISTORIAL DE PAGOS CXC
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

            ViewBag.TotalRegistros = pagos.Count;

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

            // Resumen por método.
            ViewBag.CantidadEfectivo = pagosActivos.Count(p =>
                p.MetodoPago == "Efectivo");

            ViewBag.CantidadTransferencia = pagosActivos.Count(p =>
                p.MetodoPago == "Transferencia");

            ViewBag.CantidadDeposito = pagosActivos.Count(p =>
                p.MetodoPago == "Depósito");

            ViewBag.CantidadSinpe = pagosActivos.Count(p =>
                p.MetodoPago == "SINPE Móvil");

            ViewBag.CantidadCheque = pagosActivos.Count(p =>
                p.MetodoPago == "Cheque");

            ViewBag.FechaGeneracion = DateTime.Now;

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
        // EXPORTAR HISTORIAL DE PAGOS CXC A EXCEL
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
                    p.IdPagoCuentaPorCobrar)
                .ToList();

            var pagosActivos = pagos
                .Where(p => p.Anulado == false)
                .ToList();

            using var workbook = new XLWorkbook();

            var hoja = workbook.Worksheets.Add(
                "Pagos CxC");

            const int totalColumnas = 11;

            // =======================================================
            // ENCABEZADO
            // =======================================================
            hoja.Cell("A1").Value =
                "Unión Cantonal de Asociaciones de Desarrollo de Santa Ana";

            hoja.Range(1, 1, 1, totalColumnas).Merge();

            hoja.Cell("A2").Value =
                "Historial de Pagos de Cuentas por Cobrar";

            hoja.Range(2, 1, 2, totalColumnas).Merge();

            hoja.Cell("A3").Value =
                $"Generado el {DateTime.Now:dd/MM/yyyy HH:mm}";

            hoja.Range(3, 1, 3, totalColumnas).Merge();

            hoja.Range(1, 1, 1, totalColumnas)
                .Style.Font.Bold = true;

            hoja.Range(1, 1, 1, totalColumnas)
                .Style.Font.FontSize = 16;

            hoja.Range(1, 1, 1, totalColumnas)
                .Style.Font.FontColor =
                    XLColor.FromHtml("#0F5C64");

            hoja.Range(2, 1, 2, totalColumnas)
                .Style.Font.Bold = true;

            hoja.Range(2, 1, 2, totalColumnas)
                .Style.Font.FontSize = 13;

            hoja.Range(1, 1, 3, totalColumnas)
                .Style.Alignment.Horizontal =
                    XLAlignmentHorizontalValues.Center;

            hoja.Range(1, 1, 3, totalColumnas)
                .Style.Alignment.Vertical =
                    XLAlignmentVerticalValues.Center;

            // =======================================================
            // FILTROS
            // =======================================================
            var filtros = new List<string>();

            if (!string.IsNullOrWhiteSpace(buscar))
                filtros.Add($"Búsqueda: {buscar}");

            if (!string.IsNullOrWhiteSpace(tipoDocumento))
                filtros.Add(
                    $"Tipo de documento: {tipoDocumento}");

            if (!string.IsNullOrWhiteSpace(estado))
                filtros.Add($"Estado: {estado}");

            if (!string.IsNullOrWhiteSpace(metodoPago))
                filtros.Add($"Método: {metodoPago}");

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
                filtros.Add($"Año: {anio.Value}");

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
                filtros.Add(
                    $"Desde: {fechaInicio.Value:dd/MM/yyyy}");

            if (fechaFin.HasValue)
                filtros.Add(
                    $"Hasta: {fechaFin.Value:dd/MM/yyyy}");

            hoja.Cell("A5").Value =
                "Filtros aplicados:";

            hoja.Cell("A5").Style.Font.Bold = true;

            hoja.Cell("B5").Value = filtros.Any()
                ? string.Join(" | ", filtros)
                : "Sin filtros. Se muestran todos los pagos de cuentas por cobrar.";

            hoja.Range(5, 2, 5, totalColumnas).Merge();

            hoja.Range(5, 2, 5, totalColumnas)
                .Style.Alignment.WrapText = true;

            // =======================================================
            // ENCABEZADOS
            // =======================================================
            const int filaEncabezado = 7;

            string[] encabezados =
            {
        "Fecha de pago",
        "N.º documento",
        "Tipo de documento",
        "Origen",
        "Cliente / Beneficiario",
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
            int fila = filaEncabezado + 1;

            foreach (var pago in pagos)
            {
                hoja.Cell(fila, 1).Value =
                    pago.FechaPago
                        .ToDateTime(TimeOnly.MinValue);

                hoja.Cell(fila, 2).Value =
                    pago.CuentaPorCobrar?.NumeroDocumento
                    ?? "Sin documento";

                hoja.Cell(fila, 3).Value =
                    pago.CuentaPorCobrar?.TipoDocumento
                    ?? "Sin tipo";

                hoja.Cell(fila, 4).Value =
                    pago.CuentaPorCobrar?.Origen
                    ?? "Sin origen";

                hoja.Cell(fila, 5).Value =
                    pago.CuentaPorCobrar?
                        .ClienteBeneficiario?.Nombre
                    ?? "Sin cliente";

                hoja.Cell(fila, 6).Value =
                    pago.CuentaPorCobrar?.Proyecto?.Nombre
                    ?? "Sin proyecto";

                hoja.Cell(fila, 7).Value =
                    pago.Monto ?? 0m;

                hoja.Cell(fila, 8).Value =
                    pago.MetodoPago ?? "Sin método";

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

            int filaInicioDatos = filaEncabezado + 1;
            int filaFinDatos = fila - 1;

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

                rangoDatos.Style.Alignment.WrapText = false;

                hoja.Range(
                        filaInicioDatos,
                        1,
                        filaFinDatos,
                        1)
                    .Style.DateFormat.Format =
                        "dd/MM/yyyy";

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

                filaFinDatos = filaInicioDatos;
                fila = filaInicioDatos + 1;
            }

            // =======================================================
            // TOTALES
            // =======================================================
            decimal totalPagado = pagosActivos
                .Sum(p => p.Monto ?? 0m);

            int filaTotales = fila + 1;

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
            // RESUMEN
            // =======================================================
            int filaResumen = filaTotales + 3;

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
        pagos.Count(p => p.Anulado),
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
            // COLUMNAS
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

            hoja.SheetView.FreezeRows(filaEncabezado);

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

            using var stream = new MemoryStream();

            workbook.SaveAs(stream);

            string nombreArchivo =
                $"Historial_Pagos_CxC_" +
                $"{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";

            return File(
                stream.ToArray(),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                nombreArchivo);
        }

        // GET: PagosCuentasPorCobrar/Detalles
        public async Task<IActionResult> Details(int? id, string? origen, int? cuentaId)
        {

            ViewBag.Origen = origen;
            ViewBag.CuentaId = cuentaId;

            if (id == null)
            {
                return NotFound();
            }

            var pagosCuentaPorCobrar = await _context.PagosCuentaPorCobrars
                .Include(p => p.CuentaPorCobrar)

                    .ThenInclude(c => c.ClienteBeneficiario)
                .FirstOrDefaultAsync(m => m.IdPagoCuentaPorCobrar == id);

            if (pagosCuentaPorCobrar == null)
            {
                return NotFound();
            }

            return View(pagosCuentaPorCobrar);
        }

        // GET: Crear pago de cuenta por cobrar
        // Este método solo permite abrir el formulario si el cobro viene desde una CxC específica.
        public async Task<IActionResult> Create(int? cuentaPorCobrarId)
        {
            if (cuentaPorCobrarId == null || cuentaPorCobrarId == 0)
            {
                TempData["MensajeError"] = "El cobro debe registrarse desde una cuenta por cobrar específica.";
                return RedirectToAction("Index", "CuentasPorCobrar");
            }

            var cuenta = await _context.CuentasPorCobrars
                .Include(c => c.ClienteBeneficiario)
                .FirstOrDefaultAsync(c => c.IdCuentaPorCobrar == cuentaPorCobrarId);

            if (cuenta == null)
            {
                return NotFound();
            }

            // No se permiten cobros a cuentas anuladas, canceladas o sin saldo.
            if (cuenta.Activo == false ||
                cuenta.Estado == "Anulada" ||
                cuenta.Estado == "Cancelada" ||
                cuenta.SaldoPendiente <= 0)
            {
                TempData["MensajeError"] = "No se puede registrar un cobro para esta cuenta por cobrar.";
                return RedirectToAction("Index", "CuentasPorCobrar");
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

            var cuentasResumen = await ObtenerCuentasResumen();

            ViewBag.CuentaPorCobrarId = new SelectList(
                _context.CuentasPorCobrars
                    .Where(c => c.Activo == true &&
                                c.Estado != "Anulada" &&
                                c.Estado != "Cancelada" &&
                                c.SaldoPendiente > 0),
                "IdCuentaPorCobrar",
                "NumeroDocumento",
                cuenta.IdCuentaPorCobrar);

            ViewBag.CuentasResumenJson = JsonSerializer.Serialize(cuentasResumen);

            return View(new PagosCuentaPorCobrar
            {
                CuentaPorCobrarId = cuenta.IdCuentaPorCobrar,
                FechaPago = DateOnly.FromDateTime(DateTime.Today),
                Anulado = false
            });
        }

        // POST: Crear pago de cuenta por cobrar
        // Registra el cobro, actualiza saldo de CxC, sincroniza factura, genera ingreso y recalcula presupuesto.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Create([Bind("CuentaPorCobrarId,Monto,MetodoPago,Referencia,Observacion")] PagosCuentaPorCobrar pago)
        {

            var cuenta = await _context.CuentasPorCobrars
                .Include(c => c.Factura)
                .Include(c => c.ClienteBeneficiario)
                .FirstOrDefaultAsync(c => c.IdCuentaPorCobrar == pago.CuentaPorCobrarId);
       

            if (cuenta == null)
            {
                ModelState.AddModelError("CuentaPorCobrarId", "Debe seleccionar una cuenta por cobrar válida.");
            }
            else
            {
                // Valida que la cuenta permita cobros.
                if (cuenta.Activo == false ||
                    cuenta.Estado == "Anulada" ||
                    cuenta.Estado == "Cancelada")
                {
                    ModelState.AddModelError("CuentaPorCobrarId", "No se puede registrar un cobro para esta cuenta por cobrar.");
                }

                if (cuenta.SaldoPendiente <= 0)
                {
                    ModelState.AddModelError("CuentaPorCobrarId", "La cuenta seleccionada no tiene saldo pendiente.");
                }

                if ((pago.Monto ?? 0) <= 0)
                {
                    ModelState.AddModelError("Monto", "El monto del cobro debe ser mayor a cero.");
                }

                if ((pago.Monto ?? 0) > cuenta.SaldoPendiente)
                {
                    ModelState.AddModelError("Monto", "El monto del cobro no puede ser mayor al saldo pendiente.");
                }
            }

            // Campos calculados por el sistema.
            pago.FechaPago = DateOnly.FromDateTime(DateTime.Today);
            pago.Anulado = false;

            ModelState.Remove("FechaPago");
            ModelState.Remove("Anulado");
            ModelState.Remove("CuentaPorCobrar");

            if (ModelState.IsValid)
            {
                _context.Add(pago);
                await _context.SaveChangesAsync();

                // Actualiza monto cobrado.
                cuenta.MontoPagado = (cuenta.MontoPagado ?? 0) + (pago.Monto ?? 0);

                // Recalcula saldo pendiente.
                cuenta.SaldoPendiente =
                    (cuenta.MontoOriginal ?? 0)
                    - (cuenta.AnticipoAplicado ?? 0)
                    - (cuenta.MontoPagado ?? 0);

                if (cuenta.SaldoPendiente < 0)
                {
                    cuenta.SaldoPendiente = 0;
                }

                // Actualiza estado de la CxC.
                ActualizarEstadoCuenta(cuenta);

                _context.Update(cuenta);

                // Si la CxC nació desde una factura, sincroniza los montos
                // y el estado de la factura de venta a crédito.
                if (cuenta.Factura != null)
                {
                    cuenta.Factura.MontoPagado = cuenta.MontoPagado ?? 0m;
                    cuenta.Factura.SaldoPendiente = cuenta.SaldoPendiente;

                    // Si ya no queda saldo, la factura está completamente pagada.
                    if (cuenta.Factura.SaldoPendiente <= 0)
                    {
                        cuenta.Factura.SaldoPendiente = 0;
                        cuenta.Factura.Estado = "Pagada";
                    }
                    // Si conserva saldo y ya pasó la fecha de vencimiento,
                    // debe quedar Vencida aunque haya recibido pagos parciales.
                    else if (cuenta.Factura.FechaVencimiento <
                             DateOnly.FromDateTime(DateTime.Today))
                    {
                        cuenta.Factura.Estado = "Vencida";
                    }
                    // Si tiene pagos, conserva saldo y todavía no ha vencido,
                    // queda como pago parcial.
                    else if (cuenta.Factura.MontoPagado > 0)
                    {
                        cuenta.Factura.Estado = "Parcial";
                    }
                    // Sin pagos y sin haber vencido, continúa pendiente.
                    else
                    {
                        cuenta.Factura.Estado = "Pendiente";
                    }

                    _context.Update(cuenta.Factura);
                }

                await _context.SaveChangesAsync();

                //Genera ingreso automático por el cobro recibido.
                bool ingresoPagoYaExiste = await _context.Ingresos
                    .AnyAsync(i => i.Comprobante == "PAGO-CXC-" + pago.IdPagoCuentaPorCobrar);

                if (!ingresoPagoYaExiste)
                {
                    var ingresoAutomatico = new Ingreso
                    {
                        FacturaId = cuenta.FacturaId,
                        Fecha = pago.FechaPago,

                        CategoriaIngresoId = cuenta.Factura != null && cuenta.Factura.CategoriaIngresoId != null
                            ? cuenta.Factura.CategoriaIngresoId.Value
                            : 1,

                        ProyectoId = cuenta.ProyectoId,
                        PresupuestoMensualId =
    cuenta.PresupuestoMensualId,

                        MotivoPendientePresupuestario =
    cuenta.PresupuestoMensualId.HasValue
        ? null
        : "Ingreso generado desde CxC sin período presupuestario.",

                        CentroCostoId = cuenta.Factura != null
                            ? cuenta.Factura.CentroCostoId
                            : null,

                        ClienteBeneficiarioId = cuenta.ClienteBeneficiarioId,
                        TipoIngreso = "Cobro CxC",
                        Origen = "CxC",
                        Fuente = cuenta.NumeroDocumento,

                        MontoEsperado = pago.Monto,
                        MontoReal = pago.Monto,
                        Diferencia = 0,

                        EsAnticipo = false,
                        Descripcion = "Ingreso generado automáticamente por pago de CxC " + cuenta.NumeroDocumento,
                        Estado = "Pagado",
                        Observaciones = pago.Observacion,
                        Comprobante = "PAGO-CXC-" + pago.IdPagoCuentaPorCobrar,
                        Activo = true
                    };

                    _context.Ingresos.Add(ingresoAutomatico);
                    await _context.SaveChangesAsync();

                    // Recalcula presupuesto mensual relacionado.
                    await RecalcularPresupuestoPorPagoCxC(ingresoAutomatico);
                }
                

                // Genera movimiento contable por el cobro de la cuenta por cobrar.
                // Debe: Caja General
                // Haber: Cuentas por Cobrar
                await RegistrarMovimientoContablePagoCxC(pago, cuenta);

                // Auditoría.
                await RegistrarAuditoria(
                    "Crear",
                    pago.IdPagoCuentaPorCobrar,
                    "Se registró un pago CxC por ₡" +
                    (pago.Monto ?? 0).ToString("N2") +
                    " aplicado al documento " +
                    cuenta.NumeroDocumento
                );

                TempData["MensajeExito"] = "El pago fue registrado correctamente.";
                return RedirectToAction(nameof(Index));
            }

            var cuentasResumen = await ObtenerCuentasResumen();

            ViewBag.CuentaPorCobrarId = new SelectList(
                _context.CuentasPorCobrars
                    .Where(c => c.Activo == true &&
                                c.Estado != "Anulada" &&
                                c.Estado != "Cancelada" &&
                                c.SaldoPendiente > 0),
                "IdCuentaPorCobrar",
                "NumeroDocumento",
                pago.CuentaPorCobrarId);

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

        // Obtiene el resumen de cuentas por cobrar disponibles para mostrar en la vista Create.
        private async Task<List<object>> ObtenerCuentasResumen()
        {
            var hoy = DateOnly.FromDateTime(DateTime.Today);

            return await _context.CuentasPorCobrars
                .Include(c => c.ClienteBeneficiario)
                .Where(c => c.Activo == true &&
                            c.Estado != "Anulada" &&
                            c.Estado != "Cancelada" &&
                            c.SaldoPendiente > 0)
                .Select(c => new
                {
                    id = c.IdCuentaPorCobrar,
                    cliente = c.ClienteBeneficiario != null ? c.ClienteBeneficiario.Nombre : "Sin cliente",
                    documento = c.NumeroDocumento,
                    tipo = c.TipoDocumento,
                    fechaEmision = c.FechaEmision.ToString("dd/MM/yyyy"),
                    fechaVencimiento = c.FechaVencimiento.ToString("dd/MM/yyyy"),
                    diasCredito = c.DiasCredito,
                    diasRestantes = c.FechaVencimiento.DayNumber - hoy.DayNumber,
                    montoOriginal = c.MontoOriginal ?? 0,
                    montoPagado = c.MontoPagado ?? 0,
                    saldoPendiente = c.SaldoPendiente,
                    estado = c.Estado,

                    pagosRealizados = _context.PagosCuentaPorCobrars
                        .Count(p => p.CuentaPorCobrarId == c.IdCuentaPorCobrar && p.Anulado == false),

                    ultimoPago = _context.PagosCuentaPorCobrars
                        .Where(p => p.CuentaPorCobrarId == c.IdCuentaPorCobrar && p.Anulado == false)
                        .OrderByDescending(p => p.FechaPago)
                        .Select(p => p.FechaPago.ToString("dd/MM/yyyy"))
                        .FirstOrDefault()
                })
                .ToListAsync<object>();
        }


        // GET: PagosCuentasPorCobrar/eliminar/5
        public async Task<IActionResult> Delete(int? id)
        {
            if (id == null)
            {
                return NotFound();
            }
         
            // Buscar el pago con su cuenta y cliente
            var pago = await _context.PagosCuentaPorCobrars
                .Include(p => p.CuentaPorCobrar)
                    .ThenInclude(c => c.ClienteBeneficiario)
                .FirstOrDefaultAsync(p => p.IdPagoCuentaPorCobrar == id);

            if (pago == null)
            {
                return NotFound();
            }
   // ============================================================
            // BLOQUEO POR CIERRE CONTABLE
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
            // Evitar anular dos veces el mismo pago
            if (pago.Anulado)
            {
                TempData["MensajeError"] = "Este pago ya se encuentra anulado.";
                return RedirectToAction(nameof(Index));
            }

            return View(pago);
        }
  
        // POST: PagosCuentasPorCobrar/Anular/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> DeleteConfirmed(int id, string motivoAnulacion)
        {
            // Buscar el pago que se desea anular
            var pago = await _context.PagosCuentaPorCobrars
                .Include(p => p.CuentaPorCobrar)
                    .ThenInclude(c => c.ClienteBeneficiario)
                .FirstOrDefaultAsync(p => p.IdPagoCuentaPorCobrar == id);

            if (pago == null)
            {
                return NotFound();
            }
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
            // Validar que el motivo no venga vacío
            if (string.IsNullOrWhiteSpace(motivoAnulacion))
            {
                ModelState.AddModelError("motivoAnulacion", "Debe ingresar el motivo de anulación.");
                return View("Delete", pago);
            }

            // Validar caracteres permitidos en el motivo
            if (!System.Text.RegularExpressions.Regex.IsMatch(
                    motivoAnulacion,
                    @"^[a-zA-ZáéíóúÁÉÍÓÚñÑ0-9\s.,;:/()#-]*$"))
            {
                ModelState.AddModelError("motivoAnulacion", "El motivo de anulación contiene caracteres no permitidos.");
                return View("Delete", pago);
            }
        

            if (motivoAnulacion.Trim().Length < 5)
            {
                ModelState.AddModelError("motivoAnulacion",
                    "El motivo de anulación debe contener al menos 5 caracteres.");

                return View("Delete", pago);
            }
            if (!motivoAnulacion.Any(char.IsLetter))
            {
                ModelState.AddModelError("motivoAnulacion",
                    "El motivo de anulación debe contener al menos una letra.");

                return View("Delete", pago);
            }

            // // Todas las validaciones fueron correctas, guardar el motivo de la anulación
            pago.MotivoAnulacion = motivoAnulacion.Trim();
            // Evitar anular dos veces el mismo pago
            if (pago.Anulado)
            {
                TempData["MensajeError"] = "Este pago ya se encuentra anulado.";
                return RedirectToAction(nameof(Index));
            }

            // Buscar la cuenta por cobrar relacionada
            var cuenta = await _context.CuentasPorCobrars
                .FirstOrDefaultAsync(c => c.IdCuentaPorCobrar == pago.CuentaPorCobrarId);

            if (cuenta == null)
            {
                return NotFound();
            }

            decimal montoPago = pago.Monto ?? 0;

           


            // Marcar el pago como anulado sin eliminarlo
            pago.Anulado = true;

            // Restar el pago anulado del monto pagado de la cuenta
            cuenta.MontoPagado = (cuenta.MontoPagado ?? 0) - montoPago;

            if (cuenta.MontoPagado < 0)
            {
                cuenta.MontoPagado = 0;
            }

            // Recalcular el saldo pendiente
            cuenta.SaldoPendiente = (cuenta.MontoOriginal ?? 0)
                                  - (cuenta.AnticipoAplicado ?? 0)
                                  - (cuenta.MontoPagado ?? 0);

            if (cuenta.SaldoPendiente < 0)
            {
                cuenta.SaldoPendiente = 0;
            }

            // Actualizar el estado de la cuenta
            if (cuenta.SaldoPendiente <= 0)
            {
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

            _context.Update(pago);
            _context.Update(cuenta);
            // Sincroniza la factura de venta crédito asociada a la CxC.
            var factura = await _context.Facturas
                .FirstOrDefaultAsync(f => f.IdFactura == cuenta.FacturaId);

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

            // Buscar el ingreso automático generado por este pago de CxC.
            // No se elimina físicamente para conservar trazabilidad. 
            var ingreso = await _context.Ingresos
                .FirstOrDefaultAsync(i =>
                    i.Comprobante == "PAGO-CXC-" + pago.IdPagoCuentaPorCobrar);

            if (ingreso != null)
            {
                // Se marca como anulado para que ya no afecte presupuesto ni reportes.
                ingreso.Estado = "Anulado";
                ingreso.Activo = false;

                // Se conserva la observación del motivo para auditoría funcional.
                ingreso.Observaciones =
                    "Ingreso anulado automáticamente por anulación del pago CxC. Motivo: "
                    + motivoAnulacion.Trim();

                _context.Update(ingreso);
                await _context.SaveChangesAsync();

                // Como el ingreso dejó de estar activo, se recalcula el presupuesto.
                await RecalcularPresupuestoPorPagoCxC(ingreso);
            }


            // Genera reversión contable por la anulación del pago CxC.
            // Debe: Cuentas por Cobrar
            // Haber: Caja General
            await RevertirMovimientoContablePagoCxC(pago, cuenta);

            await RegistrarAuditoria(
    "Anular",
    pago.IdPagoCuentaPorCobrar,
    "Se anuló el pago CxC del documento " + cuenta.NumeroDocumento +
    ". Motivo: " + motivoAnulacion.Trim()
);
            TempData["MensajeExito"] = "El pago fue anulado correctamente y la cuenta por cobrar fue actualizada.";

            return RedirectToAction(nameof(Index));
        }

        // Actualiza el estado de la cuenta por cobrar según saldo, pagos y vencimiento.
        private void ActualizarEstadoCuenta(CuentasPorCobrar cuenta)
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

        // Recalcula el presupuesto mensual cuando un pago de CxC genera un ingreso.
        // El presupuesto se actualiza con base en el proyecto y la fecha del ingreso.
        private async Task RecalcularPresupuestoPorPagoCxC(Ingreso ingreso)
    {
            if (!ingreso.PresupuestoMensualId.HasValue)
            {
                return;
            }

            var presupuesto =
                await _context.PresupuestosMensuales
                    .FirstOrDefaultAsync(p =>
                        p.IdPresupuestoMensual ==
                        ingreso.PresupuestoMensualId.Value);

            if (presupuesto == null)
            {
                return;
            }
            // Recalcula el total real de ingresos del periodo.
            presupuesto.TotalIngresoReal =
    await _context.Ingresos
        .Where(i =>
            i.PresupuestoMensualId ==
                presupuesto.IdPresupuestoMensual &&
            i.Activo == true &&
            i.Estado != "Anulado")
        .SumAsync(i => i.MontoReal ?? 0);

            // Recalcula el total real de gastos aprobados del periodo.
            presupuesto.TotalGastoReal =
     await _context.Gastos
         .Where(g =>
             g.PresupuestoMensualId ==
                 presupuesto.IdPresupuestoMensual &&
             g.Activo == true &&
             g.Estado == "Aprobado")
         .SumAsync(g => g.MontoTotal ?? 0);
            // Diferencia de ingreso: real menos planificado.
            // Si es positiva, se superó la meta de ingresos.
            presupuesto.DiferenciaIngreso =
            (presupuesto.TotalIngresoReal ?? 0) -
            (presupuesto.MontoIngresadoPlanificado ?? 0);

        // Saldo de gasto: planificado menos gasto real.
        // Si es positivo, todavía hay presupuesto disponible para gastar.
        presupuesto.DiferenciaGasto =
            (presupuesto.MontoGastoPlanificado ?? 0) -
            (presupuesto.TotalGastoReal ?? 0);

        // Balance real del periodo: ingresos reales menos gastos reales.
        presupuesto.SaldoDisponible =
            (presupuesto.TotalIngresoReal ?? 0) -
            (presupuesto.TotalGastoReal ?? 0);

        // Estado general del presupuesto.
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

        await _context.SaveChangesAsync();
    }

        // Registra el movimiento contable cuando se recibe un pago de CxC.
        // Debe: Caja General
        // Haber: Cuentas por Cobrar
        private async Task RegistrarMovimientoContablePagoCxC(
            PagosCuentaPorCobrar pago,
            CuentasPorCobrar cuenta)
        {
            decimal monto = pago.Monto ?? 0;

            if (monto <= 0)
            {
                return;
            }

            string referencia = "PAGO-CXC-" + pago.IdPagoCuentaPorCobrar;

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

            var cuentaCxC = await _context.CuentasContables
                .FirstOrDefaultAsync(c => c.Codigo == "CC006");

            if (cuentaCaja == null || cuentaCxC == null)
            {
                return;
            }

            // Debe: Caja General
            var movimientoDebe = new MovimientosContable
            {
                Fecha = pago.FechaPago,
                CuentaContableId = cuentaCaja.IdCuentaContable,
                TipoMovimiento = "Debe",
                Monto = monto,
                OrigenModulo = "Pagos CxC",
                OrigenId = pago.IdPagoCuentaPorCobrar,
                Referencia = referencia,
                Descripcion = "Cobro recibido de cuenta por cobrar " + cuenta.NumeroDocumento,
                ProyectoId = cuenta.ProyectoId,
                CentroCostoId = cuenta.Factura != null ? cuenta.Factura.CentroCostoId : null,
                Debe = monto,
                Haber = 0,
                Estado = "Registrado",
                UsuarioId = usuario != null ? usuario.IdUsuario : null
            };

            // Haber: Cuentas por Cobrar
            var movimientoHaber = new MovimientosContable
            {
                Fecha = pago.FechaPago,
                CuentaContableId = cuentaCxC.IdCuentaContable,
                TipoMovimiento = "Haber",
                Monto = monto,
                OrigenModulo = "Pagos CxC",
                OrigenId = pago.IdPagoCuentaPorCobrar,
                Referencia = referencia,
                Descripcion = "Disminución de cuenta por cobrar por pago recibido " + cuenta.NumeroDocumento,
                ProyectoId = cuenta.ProyectoId,
                CentroCostoId = cuenta.Factura != null ? cuenta.Factura.CentroCostoId : null,
                Debe = 0,
                Haber = monto,
                Estado = "Registrado",
                UsuarioId = usuario != null ? usuario.IdUsuario : null
            };

            _context.MovimientosContables.Add(movimientoDebe);
            _context.MovimientosContables.Add(movimientoHaber);

            // Actualiza saldos contables.
            // Caja aumenta porque entra dinero.
            // CxC disminuye porque baja la deuda pendiente.
            // CAMBIO APLICADO PARA CXC MANUAL:
            // No se fuerza CC006 a cero, porque eso rompe la reversión
            // cuando se anula un pago de CxC manual.
            // Actualiza saldos contables.
            cuentaCaja.SaldoActual = cuentaCaja.SaldoActual + monto;
            cuentaCxC.SaldoActual = cuentaCxC.SaldoActual - monto;

          /*  if (cuentaCxC.SaldoActual < 0)
            {
                cuentaCxC.SaldoActual = 0;
            }  */

            _context.CuentasContables.Update(cuentaCaja);
            _context.CuentasContables.Update(cuentaCxC);

            await _context.SaveChangesAsync();
        }


        // Registra la reversión contable cuando se anula un pago de CxC.
        // Debe: Cuentas por Cobrar
        // Haber: Caja General
        private async Task RevertirMovimientoContablePagoCxC(
            PagosCuentaPorCobrar pago,
            CuentasPorCobrar cuenta)
        {
            decimal monto = pago.Monto ?? 0;

            if (monto <= 0)
            {
                return;
            }

            string referenciaOriginal = "PAGO-CXC-" + pago.IdPagoCuentaPorCobrar;
            string referenciaReversa = "REV-PAGO-CXC-" + pago.IdPagoCuentaPorCobrar;

            bool reversaYaExiste = await _context.MovimientosContables
                     .AnyAsync(m =>
                     m.Referencia == referenciaReversa);

            if (reversaYaExiste)
            {
                return;
            }

            var usuarioSesion = HttpContext.Session.GetString("Usuario");

            var usuario = await _context.Usuarios
                .FirstOrDefaultAsync(u => u.Nombre == usuarioSesion);

            var cuentaCaja = await _context.CuentasContables
                .FirstOrDefaultAsync(c => c.Codigo == "CC001");

            var cuentaCxC = await _context.CuentasContables
                .FirstOrDefaultAsync(c => c.Codigo == "CC006");

            if (cuentaCaja == null || cuentaCxC == null)
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
                movimiento.MotivoAnulacion = "Anulación del pago CxC " + cuenta.NumeroDocumento;
                movimiento.MovimientoReversionId = null;
                _context.MovimientosContables.Update(movimiento);
            }

            // Debe: Cuentas por Cobrar 
            var movimientoDebe = new MovimientosContable
            {
                Fecha = DateOnly.FromDateTime(DateTime.Today),
                CuentaContableId = cuentaCxC.IdCuentaContable,
                TipoMovimiento = "Debe",
                Monto = monto,
                Referencia = referenciaReversa,
                Descripcion = "Reversión por anulación de pago CxC " + cuenta.NumeroDocumento,
                ProyectoId = cuenta.ProyectoId,
                CentroCostoId = cuenta.Factura != null ? cuenta.Factura.CentroCostoId : null,
                Debe = monto,
                Haber = 0,
                Estado = "Reversado",
                UsuarioId = usuario != null ? usuario.IdUsuario : null,
                OrigenModulo = "Pagos CxC",
                OrigenId = pago.IdPagoCuentaPorCobrar
            };

            // Haber: Caja General
            var movimientoHaber = new MovimientosContable
            {
                Fecha = DateOnly.FromDateTime(DateTime.Today),
                CuentaContableId = cuentaCaja.IdCuentaContable,
                TipoMovimiento = "Haber",
                Monto = monto,
                Referencia = referenciaReversa,
                Descripcion = "Salida de caja por reversión de pago CxC " + cuenta.NumeroDocumento,
                ProyectoId = cuenta.ProyectoId,
                CentroCostoId = cuenta.Factura != null ? cuenta.Factura.CentroCostoId : null,
                Debe = 0,
                Haber = monto,
                Estado = "Reversado",
                UsuarioId = usuario != null ? usuario.IdUsuario : null,
                OrigenModulo = "Pagos CxC",
                OrigenId = pago.IdPagoCuentaPorCobrar
            };

            _context.MovimientosContables.Add(movimientoDebe);
            _context.MovimientosContables.Add(movimientoHaber);

            // Actualiza saldos contables por reversión.
            // Caja disminuye y CxC vuelve a aumentar.
            cuentaCaja.SaldoActual = cuentaCaja.SaldoActual - monto;
            cuentaCxC.SaldoActual = cuentaCxC.SaldoActual + monto;

            // Primero guarda las reversas para que tengan IdMovimientoContable.
            await _context.SaveChangesAsync();

            foreach (var movimiento in movimientosOriginales)
            {
                if (movimiento.TipoMovimiento == "Debe")
                {
                    movimiento.MovimientoReversionId = movimientoHaber.IdMovimientoContable;
                }
                else if (movimiento.TipoMovimiento == "Haber")
                {
                    movimiento.MovimientoReversionId = movimientoDebe.IdMovimientoContable;
                }

                _context.MovimientosContables.Update(movimiento);
            }
            // Actualiza saldos contables por reversión.
            // Caja vuelve a aumentar y CxP vuelve a aumentar.
            _context.CuentasContables.Update(cuentaCaja);
            _context.CuentasContables.Update(cuentaCxC);

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
        //AUDITORIA
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
                    Tabla = "PagosCuentaPorCobrar",
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