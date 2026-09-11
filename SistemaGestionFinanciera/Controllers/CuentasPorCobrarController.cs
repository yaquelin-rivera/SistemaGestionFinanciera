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
using ClosedXML.Excel;
using System.IO;

namespace SistemaGestionFinanciera.Controllers
{
    public class CuentasPorCobrarController : BaseController
    {
        private readonly SistemaFinancieroContext _context;

        public CuentasPorCobrarController(SistemaFinancieroContext context)
        {
            _context = context;
        }

        // GET: CuentasPorCobrar
        public async Task<IActionResult> Index(
            string? buscar,
            string? tipoDocumento,
            string? estado,
            string? origen,
            int? proyectoId,
            int? anio,
            int? mes,
            DateOnly? fechaInicio,
            DateOnly? fechaFin,
            bool proximasVencer = false)
        {
            // Actualiza los estados antes de aplicar filtros, alertas y totales.
            await ActualizarEstadosCuentasPorCobrar();

            // Detecta si el usuario aplicó filtros.
            bool hayFiltros =
                !string.IsNullOrWhiteSpace(buscar) ||
                !string.IsNullOrWhiteSpace(tipoDocumento) ||
                !string.IsNullOrWhiteSpace(estado) ||
                !string.IsNullOrWhiteSpace(origen) ||
                (proyectoId.HasValue && proyectoId.Value > 0) ||
                anio.HasValue ||
                mes.HasValue ||
                proximasVencer ||
                fechaInicio.HasValue ||
                fechaFin.HasValue;

            // Obtiene las cuentas por cobrar con los filtros aplicados.
            // La consulta se comparte con el PDF y Excel.
            var listaCuentas = await ConstruirConsultaReporte(
                buscar,
                tipoDocumento,
                estado,
                origen,
                proyectoId,
                anio,
                mes,
                fechaInicio,
                fechaFin,
                proximasVencer);

            // Mantiene los filtros seleccionados en pantalla.
            ViewBag.Buscar = buscar;
            ViewBag.TipoDocumento = tipoDocumento;
            ViewBag.Estado = estado;
            ViewBag.Origen = origen;
            ViewBag.ProyectoSeleccionado = proyectoId;
            ViewBag.Anio = anio;
            ViewBag.Mes = mes;
            ViewBag.ProximasVencerFiltro = proximasVencer;
            ViewBag.FechaInicio = fechaInicio?.ToString("yyyy-MM-dd");
            ViewBag.FechaFin = fechaFin?.ToString("yyyy-MM-dd");

            // ================================================================
            // ALERTAS DE CUENTAS POR COBRAR
            // ================================================================

            var hoyAlertas = DateOnly.FromDateTime(DateTime.Today);
            var fechaLimiteProximas = hoyAlertas.AddDays(7);

            // ---------------------------------------------------------------
            // CUENTAS VENCIDAS
            // Incluye todas las cuentas activas vencidas con saldo pendiente.
            // ---------------------------------------------------------------
            var cuentasVencidas = await _context.CuentasPorCobrars
                .AsNoTracking()
                .Where(c =>
                    c.Activo == true &&
                    c.Estado == "Vencida" &&
                    c.SaldoPendiente > 0)
                .ToListAsync();

            ViewBag.CuentasVencidas =
                cuentasVencidas.Count;

            ViewBag.SaldoTotalVencido =
                cuentasVencidas.Sum(c => c.SaldoPendiente);

            // ---------------------------------------------------------------
            // CUENTAS PRÓXIMAS A VENCER
            // Incluye cuentas pendientes o parciales con saldo,
            // cuyo vencimiento ocurre entre hoy y los próximos 7 días.
            // ---------------------------------------------------------------
            var cuentasProximasVencer = await _context.CuentasPorCobrars
                .AsNoTracking()
                .Where(c =>
                    c.Activo == true &&
                    c.Estado != "Anulada" &&
                    c.Estado != "Cancelada" &&
                    c.Estado != "Vencida" &&
                    c.SaldoPendiente > 0 &&
                    c.FechaVencimiento >= hoyAlertas &&
                    c.FechaVencimiento <= fechaLimiteProximas)
                .ToListAsync();

            ViewBag.CuentasProximasVencer =
                cuentasProximasVencer.Count;

            ViewBag.SaldoProximasVencer =
                cuentasProximasVencer.Sum(c => c.SaldoPendiente);

            var cuentasParaTotales = listaCuentas;

            // Si no existen filtros, las cards muestran solamente el mes actual.
            if (!hayFiltros)
            {
                var hoy = DateOnly.FromDateTime(DateTime.Today);
                var inicioMes = new DateOnly(hoy.Year, hoy.Month, 1);
                var finMes = inicioMes.AddMonths(1).AddDays(-1);

                cuentasParaTotales = listaCuentas
                    .Where(c =>
                        c.FechaEmision >= inicioMes &&
                        c.FechaEmision <= finMes)
                    .ToList();

                ViewBag.TipoResumen = "Mes actual";
            }
            else
            {
                ViewBag.TipoResumen = "Filtro aplicado";
            }

            // Solo suma cuentas activas y no anuladas.
            var cuentasValidas = cuentasParaTotales
                .Where(c =>
                    c.Activo == true &&
                    c.Estado != "Anulada")
                .ToList();

            ViewBag.TotalPorCobrar = cuentasValidas
                .Sum(c => c.MontoOriginal ?? 0m);

            ViewBag.TotalCobrado = cuentasValidas
                .Sum(c => c.MontoPagado ?? 0m);

            ViewBag.TotalSaldoPendiente = cuentasValidas
                .Sum(c => c.SaldoPendiente);

            // Proyectos activos para el filtro.
            ViewBag.ProyectoId = new SelectList(
                await _context.Proyectos
                    .Where(p => p.Activo == true)
                    .OrderBy(p => p.Nombre)
                    .ToListAsync(),
                "IdProyecto",
                "Nombre",
                proyectoId);

            // Años existentes en las cuentas por cobrar.
            ViewBag.Anios = await _context.CuentasPorCobrars
                .Select(c => c.FechaEmision.Year)
                .Distinct()
                .OrderByDescending(a => a)
                .ToListAsync();

            return View(listaCuentas);
        }

        // ===========================================================
        // CONSTRUYE LA CONSULTA PARA INDEX, PDF Y EXCEL
        // ===========================================================
        private async Task<List<CuentasPorCobrar>> ConstruirConsultaReporte(
            string? buscar,
            string? tipoDocumento,
            string? estado,
            string? origen,
            int? proyectoId,
            int? anio,
            int? mes,
            DateOnly? fechaInicio,
            DateOnly? fechaFin,
            bool proximasVencer = false)
        {
            var cuentas = _context.CuentasPorCobrars
                .Include(c => c.ClienteBeneficiario)
                .Include(c => c.Proyecto)
                .Include(c => c.PresupuestoMensualRegistro)
                .AsQueryable();

            if (proximasVencer)
            {
                var hoy = DateOnly.FromDateTime(DateTime.Today);
                var fechaLimite = hoy.AddDays(7);

                cuentas = cuentas.Where(c =>
                    c.Activo == true &&
                    c.Estado != "Anulada" &&
                    c.Estado != "Cancelada" &&
                    c.Estado != "Vencida" &&
                    c.SaldoPendiente > 0 &&
                    c.FechaVencimiento >= hoy &&
                    c.FechaVencimiento <= fechaLimite);
            }
            // Buscar
            if (!string.IsNullOrWhiteSpace(buscar))
            {
                cuentas = cuentas.Where(c =>
                    (c.NumeroDocumento != null && c.NumeroDocumento.Contains(buscar)) ||
                    (c.Concepto != null && c.Concepto.Contains(buscar)) ||
                    (c.ClienteBeneficiario != null &&
                     c.ClienteBeneficiario.Nombre.Contains(buscar)));
            }

            // Tipo de documento
            if (!string.IsNullOrWhiteSpace(tipoDocumento))
            {
                cuentas = cuentas.Where(c =>
                    c.TipoDocumento == tipoDocumento);
            }

            // Estado
            if (!string.IsNullOrWhiteSpace(estado))
            {
                cuentas = cuentas.Where(c =>
                    c.Estado == estado);
            }

            // Origen
            if (!string.IsNullOrWhiteSpace(origen))
            {
                cuentas = cuentas.Where(c =>
                    c.Origen == origen);
            }

            // Proyecto
            if (proyectoId.HasValue && proyectoId.Value > 0)
            {
                cuentas = cuentas.Where(c =>
                    c.ProyectoId == proyectoId.Value);
            }

            // Año
            if (anio.HasValue)
            {
                cuentas = cuentas.Where(c =>
                    c.FechaEmision.Year == anio.Value);
            }

            // Mes
            if (mes.HasValue)
            {
                cuentas = cuentas.Where(c =>
                    c.FechaEmision.Month == mes.Value);
            }

            // Fecha inicial
            if (fechaInicio.HasValue)
            {
                cuentas = cuentas.Where(c =>
                    c.FechaEmision >= fechaInicio.Value);
            }

            // Fecha final
            if (fechaFin.HasValue)
            {
                cuentas = cuentas.Where(c =>
                    c.FechaEmision <= fechaFin.Value);
            }

            return await cuentas
                .OrderByDescending(c => c.IdCuentaPorCobrar)
                .ToListAsync();
        }
        // ===========================================================
        // VISTA PREVIA DEL REPORTE DE CUENTAS POR COBRAR
        // ===========================================================
        public async Task<IActionResult> Reporte(
            string? buscar,
            string? tipoDocumento,
            string? estado,
            string? origen,
            int? proyectoId,
            int? anio,
            int? mes,
            DateOnly? fechaInicio,
            DateOnly? fechaFin,
            bool proximasVencer = false)
        {
            // Actualiza los estados antes de generar la vista previa del reporte.
            await ActualizarEstadosCuentasPorCobrar();

            var cuentas = await ConstruirConsultaReporte(
                buscar,
                tipoDocumento,
                estado,
                origen,
                proyectoId,
                anio,
                mes,
                fechaInicio,
                fechaFin, proximasVencer);

            // Orden definitivo del reporte:
            // fecha más reciente y luego ID más reciente.
            cuentas = cuentas
                .OrderByDescending(c => c.FechaEmision)
                .ThenByDescending(c => c.IdCuentaPorCobrar)
                .ToList();

            // Las anuladas se muestran en la tabla,
            // pero no participan en los totales financieros.
            var cuentasValidas = cuentas
                .Where(c =>
                    c.Activo == true &&
                    c.Estado != "Anulada")
                .ToList();

            ViewBag.TotalRegistros = cuentas.Count;

            ViewBag.TotalPorCobrar = cuentasValidas
                .Sum(c => c.MontoOriginal ?? 0m);

            ViewBag.TotalCobrado = cuentasValidas
                .Sum(c => c.MontoPagado ?? 0m);

            ViewBag.TotalSaldoPendiente = cuentasValidas
                .Sum(c => c.SaldoPendiente);

            // Cantidades por estado.
            ViewBag.CantidadPendientes = cuentasValidas.Count(c =>
                c.Estado == "Pendiente");

            ViewBag.CantidadParciales = cuentasValidas.Count(c =>
                c.Estado == "Parcial");

            ViewBag.CantidadCanceladas = cuentasValidas.Count(c =>
                c.Estado == "Cancelada");

            ViewBag.CantidadVencidas = cuentasValidas.Count(c =>
                c.Estado == "Vencida");


            ViewBag.CantidadAnuladas = cuentas.Count(c =>
                c.Estado == "Anulada");

            ViewBag.FechaGeneracion = DateTime.Now;

            // Conserva los filtros para mostrarlos en el reporte.
            ViewBag.Buscar = buscar;
            ViewBag.TipoDocumento = tipoDocumento;
            ViewBag.Estado = estado;
            ViewBag.Origen = origen;
            ViewBag.ProyectoId = proyectoId;
            ViewBag.Anio = anio;
            ViewBag.Mes = mes;

            ViewBag.ProximasVencerFiltro = proximasVencer;
            ViewBag.FechaInicio = fechaInicio?.ToString("yyyy-MM-dd");
            ViewBag.FechaFin = fechaFin?.ToString("yyyy-MM-dd");

            // Nombre del proyecto para que el reporte no muestre solamente el ID.
            ViewBag.NombreProyecto = proyectoId.HasValue
                ? await _context.Proyectos
                    .Where(p => p.IdProyecto == proyectoId.Value)
                    .Select(p => p.Nombre)
                    .FirstOrDefaultAsync()
                : null;

            // Registra la generación de la vista previa del reporte.
            // RegistroId se establece en 0 porque corresponde a un reporte general
            // y no a una cuenta por cobrar específica.
            await RegistrarAuditoria(
                "Vista previa PDF",
                0,
                "Se generó la vista previa del reporte de cuentas por cobrar. " +
                ConstruirDescripcionFiltrosReporte(
                    buscar,
                    tipoDocumento,
                    estado,
                    origen,
                    proyectoId,
                    anio,
                    mes,
                    fechaInicio,
                    fechaFin)
            );

            return View(cuentas);
        }

        // ===========================================================
        // EXPORTAR REPORTE DE CUENTAS POR COBRAR A EXCEL
        // ===========================================================
        public async Task<IActionResult> ExportarExcel(
            string? buscar,
            string? tipoDocumento,
            string? estado,
            string? origen,
            int? proyectoId,
            int? anio,
            int? mes,
            DateOnly? fechaInicio,
            DateOnly? fechaFin, bool proximasVencer = false)
        {
            // Actualiza los estados antes de construir el archivo Excel.
            await ActualizarEstadosCuentasPorCobrar();

            var cuentas = await ConstruirConsultaReporte(
                buscar,
                tipoDocumento,
                estado,
                origen,
                proyectoId,
                anio,
                mes,
                fechaInicio,
                fechaFin,
                proximasVencer);

            // Orden definitivo del reporte:
            // fecha más reciente y luego ID más reciente.
            cuentas = cuentas
                .OrderByDescending(c => c.FechaEmision)
                .ThenByDescending(c => c.IdCuentaPorCobrar)
                .ToList();

            using var workbook = new XLWorkbook();

            var hoja = workbook.Worksheets.Add("Cuentas por Cobrar");

            const int totalColumnas = 13;

            // ===========================================================
            // ENCABEZADO PRINCIPAL
            // ===========================================================
            hoja.Cell("A1").Value =
                "Unión Cantonal de Asociaciones de Desarrollo de Santa Ana";

            hoja.Range(1, 1, 1, totalColumnas).Merge();

            hoja.Cell("A2").Value =
                "Reporte de Cuentas por Cobrar";

            hoja.Range(2, 1, 2, totalColumnas).Merge();

            hoja.Cell("A3").Value =
                $"Generado el {DateTime.Now:dd/MM/yyyy HH:mm}";

            hoja.Range(3, 1, 3, totalColumnas).Merge();

            hoja.Range(1, 1, 1, totalColumnas)
                .Style.Font.Bold = true;

            hoja.Range(1, 1, 1, totalColumnas)
                .Style.Font.FontSize = 16;

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

            // ===========================================================
            // FILTROS APLICADOS
            // ===========================================================
            var filtros = new List<string>();

            if (!string.IsNullOrWhiteSpace(buscar))
            {
                filtros.Add($"Búsqueda: {buscar}");
            }

            if (!string.IsNullOrWhiteSpace(tipoDocumento))
            {
                filtros.Add($"Tipo de documento: {tipoDocumento}");
            }

            if (!string.IsNullOrWhiteSpace(estado))
            {
                filtros.Add($"Estado: {estado}");
            }

            if (!string.IsNullOrWhiteSpace(origen))
            {
                filtros.Add($"Origen: {origen}");
            }

            if (proyectoId.HasValue && proyectoId.Value > 0)
            {
                string? nombreProyecto = await _context.Proyectos
                    .Where(p => p.IdProyecto == proyectoId.Value)
                    .Select(p => p.Nombre)
                    .FirstOrDefaultAsync();

                filtros.Add(
                    $"Proyecto: {nombreProyecto ?? proyectoId.Value.ToString()}");
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
                        new System.Globalization.CultureInfo("es-CR"));

                nombreMes =
                    char.ToUpper(nombreMes[0]) +
                    nombreMes.Substring(1);

                filtros.Add($"Mes: {nombreMes}");
            }
            if (proximasVencer)
            {
                filtros.Add("Vencimiento: Próximos 7 días");
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

            hoja.Cell("A5").Value = "Filtros aplicados:";
            hoja.Cell("A5").Style.Font.Bold = true;

            hoja.Cell("B5").Value = filtros.Any()
                ? string.Join(" | ", filtros)
                : "Sin filtros. Se muestran todas las cuentas por cobrar.";

            hoja.Range(5, 2, 5, totalColumnas).Merge();

            hoja.Range(5, 2, 5, totalColumnas)
                .Style.Alignment.WrapText = true;

            // ===========================================================
            // ENCABEZADOS DE LA TABLA
            // ===========================================================
            const int filaEncabezado = 7;

            string[] encabezados =
            {
        "N.º documento",
        "Tipo de documento",
        "Origen",
        "Cliente / Beneficiario",
        "Proyecto",
        "Concepto",
        "Fecha de emisión",
        "Fecha de vencimiento",
        "Días de crédito",
        "Monto original",
        "Monto pagado",
        "Saldo pendiente",
        "Estado"
    };

            for (int columna = 0; columna < encabezados.Length; columna++)
            {
                hoja.Cell(filaEncabezado, columna + 1).Value =
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

            rangoEncabezado.Style.Border.TopBorder =
                XLBorderStyleValues.Thin;

            rangoEncabezado.Style.Border.BottomBorder =
                XLBorderStyleValues.Thin;

            rangoEncabezado.Style.Border.LeftBorder =
                XLBorderStyleValues.Thin;

            rangoEncabezado.Style.Border.RightBorder =
                XLBorderStyleValues.Thin;

            hoja.Row(filaEncabezado).Height = 25;

            // ===========================================================
            // DATOS
            // ===========================================================
            int fila = filaEncabezado + 1;

            foreach (var cuenta in cuentas)
            {
                hoja.Cell(fila, 1).Value =
                    cuenta.NumeroDocumento ?? "Sin número";

                hoja.Cell(fila, 2).Value =
                    cuenta.TipoDocumento ?? "Sin tipo";

                hoja.Cell(fila, 3).Value =
                    cuenta.Origen ?? "Sin origen";

                hoja.Cell(fila, 4).Value =
                    cuenta.ClienteBeneficiario?.Nombre
                    ?? "Sin cliente o beneficiario";

                hoja.Cell(fila, 5).Value =
                    cuenta.Proyecto?.Nombre ?? "Sin proyecto";

                hoja.Cell(fila, 6).Value =
                    cuenta.Concepto ?? "Sin concepto";

                hoja.Cell(fila, 7).Value =
                    cuenta.FechaEmision.ToDateTime(TimeOnly.MinValue);

                hoja.Cell(fila, 8).Value =
                    cuenta.FechaVencimiento.ToDateTime(TimeOnly.MinValue);

                hoja.Cell(fila, 9).Value =
                    cuenta.DiasCredito ?? 0;

                hoja.Cell(fila, 10).Value =
                    cuenta.MontoOriginal ?? 0m;

                hoja.Cell(fila, 11).Value =
                    cuenta.MontoPagado ?? 0m;

                hoja.Cell(fila, 12).Value =
                    cuenta.SaldoPendiente;

                hoja.Cell(fila, 13).Value =
                    cuenta.Estado ?? "Sin estado";

                fila++;
            }

            // ===========================================================
            // FORMATO DE LOS DATOS
            // ===========================================================
            int filaInicioDatos = filaEncabezado + 1;
            int filaFinDatos = fila - 1;

            if (cuentas.Any())
            {
                var rangoDatos = hoja.Range(
                    filaInicioDatos,
                    1,
                    filaFinDatos,
                    totalColumnas);

                rangoDatos.Style.Border.TopBorder =
                    XLBorderStyleValues.Thin;

                rangoDatos.Style.Border.BottomBorder =
                    XLBorderStyleValues.Thin;

                rangoDatos.Style.Border.LeftBorder =
                    XLBorderStyleValues.Thin;

                rangoDatos.Style.Border.RightBorder =
                    XLBorderStyleValues.Thin;

                rangoDatos.Style.Border.TopBorderColor =
                    XLColor.LightGray;

                rangoDatos.Style.Border.BottomBorderColor =
                    XLColor.LightGray;

                rangoDatos.Style.Border.LeftBorderColor =
                    XLColor.LightGray;

                rangoDatos.Style.Border.RightBorderColor =
                    XLColor.LightGray;

                rangoDatos.Style.Alignment.Vertical =
                    XLAlignmentVerticalValues.Center;

                // Formato de fechas.
                hoja.Range(
                        filaInicioDatos,
                        7,
                        filaFinDatos,
                        8)
                    .Style.DateFormat.Format = "dd/MM/yyyy";

                // Formato de días de crédito.
                hoja.Range(
                        filaInicioDatos,
                        9,
                        filaFinDatos,
                        9)
                    .Style.NumberFormat.Format = "0";

                // Formato monetario.
                hoja.Range(
                        filaInicioDatos,
                        10,
                        filaFinDatos,
                        12)
                    .Style.NumberFormat.Format =
                        "₡ #,##0.00";

                hoja.Range(
                        filaInicioDatos,
                        9,
                        filaFinDatos,
                        13)
                    .Style.Alignment.Horizontal =
                        XLAlignmentHorizontalValues.Center;

                // Ajuste de texto en columnas extensas.
                hoja.Range(
                        filaInicioDatos,
                        1,
                        filaFinDatos,
                        6)
                    .Style.Alignment.WrapText = true;

                // Filtro automático.
                hoja.Range(
                        filaEncabezado,
                        1,
                        filaFinDatos,
                        totalColumnas)
                    .SetAutoFilter();

                // Resalta las cuentas anuladas.
                foreach (var cuenta in cuentas)
                {
                    int indice = cuentas.IndexOf(cuenta);
                    int filaCuenta = filaInicioDatos + indice;

                    if (cuenta.Estado == "Anulada")
                    {
                        hoja.Range(
                                filaCuenta,
                                1,
                                filaCuenta,
                                totalColumnas)
                            .Style.Fill.BackgroundColor =
                                XLColor.FromHtml("#F8D7DA");

                        hoja.Range(
                                filaCuenta,
                                1,
                                filaCuenta,
                                totalColumnas)
                            .Style.Font.FontColor =
                                XLColor.FromHtml("#842029");
                    }
                    else if (cuenta.Estado == "Vencida")
                    {
                        hoja.Range(
                                filaCuenta,
                                1,
                                filaCuenta,
                                totalColumnas)
                            .Style.Fill.BackgroundColor =
                                XLColor.FromHtml("#FFF3CD");
                    }
                }
            }
            else
            {
                hoja.Cell(filaInicioDatos, 1).Value =
                    "No se encontraron cuentas por cobrar con los filtros aplicados.";

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

            // ===========================================================
            // TOTALES FINANCIEROS
            // ===========================================================
            var cuentasValidas = cuentas
                .Where(c =>
                    c.Activo == true &&
                    c.Estado != "Anulada")
                .ToList();

            decimal totalMontoOriginal = cuentasValidas
                .Sum(c => c.MontoOriginal ?? 0m);

            decimal totalMontoPagado = cuentasValidas
                .Sum(c => c.MontoPagado ?? 0m);

            decimal totalSaldoPendiente = cuentasValidas
                .Sum(c => c.SaldoPendiente);

            int filaTotales = fila + 1;

            hoja.Cell(filaTotales, 8).Value =
                "TOTALES:";

            hoja.Range(
                    filaTotales,
                    8,
                    filaTotales,
                    9)
                .Merge();

            hoja.Cell(filaTotales, 10).Value =
                totalMontoOriginal;

            hoja.Cell(filaTotales, 11).Value =
                totalMontoPagado;

            hoja.Cell(filaTotales, 12).Value =
                totalSaldoPendiente;

            hoja.Cell(filaTotales, 13).Value =
                $"{cuentas.Count} registro(s)";

            var rangoTotales = hoja.Range(
                filaTotales,
                8,
                filaTotales,
                totalColumnas);

            rangoTotales.Style.Font.Bold = true;

            rangoTotales.Style.Fill.BackgroundColor =
                XLColor.FromHtml("#D9EDEF");

            rangoTotales.Style.Border.TopBorder =
                XLBorderStyleValues.Thin;

            rangoTotales.Style.Border.BottomBorder =
                XLBorderStyleValues.Thin;

            rangoTotales.Style.Border.LeftBorder =
                XLBorderStyleValues.Thin;

            rangoTotales.Style.Border.RightBorder =
                XLBorderStyleValues.Thin;

            rangoTotales.Style.Alignment.Vertical =
                XLAlignmentVerticalValues.Center;

            rangoTotales.Style.Alignment.Horizontal =
                XLAlignmentHorizontalValues.Center;

            hoja.Range(
                    filaTotales,
                    10,
                    filaTotales,
                    12)
                .Style.NumberFormat.Format =
                    "₡ #,##0.00";

            // ===========================================================
            // RESUMEN POR ESTADO
            // ===========================================================
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
                .Style.Font.Bold = true;

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
                .Style.Fill.BackgroundColor =
                    XLColor.FromHtml("#0F5C64");

            hoja.Range(
                    filaResumen,
                    1,
                    filaResumen,
                    5)
                .Style.Alignment.Horizontal =
                    XLAlignmentHorizontalValues.Center;

            string[] encabezadosResumen =
            {
        "Pendientes",
        "Parciales",
        "Canceladas",
        "Vencidas",
        "Anuladas"
    };

            for (int columna = 0;
                 columna < encabezadosResumen.Length;
                 columna++)
            {
                hoja.Cell(
                    filaResumen + 1,
                    columna + 1).Value =
                        encabezadosResumen[columna];
            }

            hoja.Cell(filaResumen + 2, 1).Value =
                cuentasValidas.Count(c =>
                    c.Estado == "Pendiente");

            hoja.Cell(filaResumen + 2, 2).Value =
                cuentasValidas.Count(c =>
                    c.Estado == "Parcial");

            hoja.Cell(filaResumen + 2, 3).Value =
                cuentasValidas.Count(c =>
                    c.Estado == "Cancelada");

            hoja.Cell(filaResumen + 2, 4).Value =
                cuentasValidas.Count(c =>
                    c.Estado == "Vencida");

            hoja.Cell(filaResumen + 2, 5).Value =
                cuentas.Count(c =>
                    c.Estado == "Anulada");

            var rangoEncabezadoResumen = hoja.Range(
                filaResumen + 1,
                1,
                filaResumen + 1,
                5);

            rangoEncabezadoResumen.Style.Font.Bold = true;

            rangoEncabezadoResumen.Style.Fill.BackgroundColor =
                XLColor.FromHtml("#D9EDEF");

            rangoEncabezadoResumen.Style.Alignment.Horizontal =
                XLAlignmentHorizontalValues.Center;

            var rangoValoresResumen = hoja.Range(
                filaResumen + 2,
                1,
                filaResumen + 2,
                5);

            rangoValoresResumen.Style.Alignment.Horizontal =
                XLAlignmentHorizontalValues.Center;

            var rangoResumenCompleto = hoja.Range(
                filaResumen + 1,
                1,
                filaResumen + 2,
                5);

            rangoResumenCompleto.Style.Border.TopBorder =
                XLBorderStyleValues.Thin;

            rangoResumenCompleto.Style.Border.BottomBorder =
                XLBorderStyleValues.Thin;

            rangoResumenCompleto.Style.Border.LeftBorder =
                XLBorderStyleValues.Thin;

            rangoResumenCompleto.Style.Border.RightBorder =
                XLBorderStyleValues.Thin;

            // ===========================================================
            // CONFIGURACIÓN DE COLUMNAS
            // ===========================================================
            hoja.Column(1).Width = 18;
            hoja.Column(2).Width = 20;
            hoja.Column(3).Width = 16;
            hoja.Column(4).Width = 30;
            hoja.Column(5).Width = 25;
            hoja.Column(6).Width = 35;
            hoja.Column(7).Width = 17;
            hoja.Column(8).Width = 19;
            hoja.Column(9).Width = 14;
            hoja.Column(10).Width = 18;
            hoja.Column(11).Width = 18;
            hoja.Column(12).Width = 18;
            hoja.Column(13).Width = 15;

            hoja.Rows().AdjustToContents();

            // Mantiene visible el encabezado al desplazarse.
            hoja.SheetView.FreezeRows(filaEncabezado);

            // ===========================================================
            // CONFIGURACIÓN DE IMPRESIÓN
            // ===========================================================
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

            // ===========================================================
            // GENERACIÓN DEL ARCHIVO
            // ===========================================================
            using var stream = new MemoryStream();

            workbook.SaveAs(stream);

            string nombreArchivo =
                $"Reporte_Cuentas_por_Cobrar_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";

            // Registra la exportación únicamente después de haber construido
            // correctamente el archivo Excel.
            await RegistrarAuditoria(
                "Exportar Excel",
                0,
                "Se exportó el reporte de cuentas por cobrar a Excel. " +
                ConstruirDescripcionFiltrosReporte(
                    buscar,
                    tipoDocumento,
                    estado,
                    origen,
                    proyectoId,
                    anio,
                    mes,
                    fechaInicio,
                    fechaFin)
            );

            return File(
                stream.ToArray(),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                nombreArchivo);
        }

        // GET: CuentasPorCobrar/Details/5
        public async Task<IActionResult> Details(int? id, string? origen, int? pagoId)
        {
            ViewBag.PagoId = pagoId;

            if (id == null)
            {
                return NotFound();
            }

            var cuentasPorCobrar = await _context.CuentasPorCobrars
                .Include(c => c.ClienteBeneficiario)
                .Include(c => c.Proyecto)
                .Include(c => c.PresupuestoMensualRegistro)
                .Include(c => c.PagosCuentaPorCobrars)
                .FirstOrDefaultAsync(m => m.IdCuentaPorCobrar == id);

            if (cuentasPorCobrar == null)
            {
                return NotFound();
            }

            ViewBag.Origen = origen;

            return View(cuentasPorCobrar);
        }

        // GET: CuentasPorCobrar/Create
        public async Task<IActionResult> Create()
        {
            CargarCombos();

            await CargarOpcionesProyectoPeriodoCxC();

            ViewBag.FechaEmision =
                DateOnly.FromDateTime(DateTime.Today);

            return View();
        }

        // POST: CuentasPorCobrar/Create
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Create(
    [Bind("ClienteBeneficiarioId,TipoDocumento,Concepto,MontoOriginal,DiasCredito,Observacion")]CuentasPorCobrar cuentasPorCobrar,
    string? SeleccionProyectoPeriodoCxC)
        {
            // Obtiene la fecha actual del sistema.
            var fechaHoy = DateOnly.FromDateTime(DateTime.Today);

            // Valida cliente o beneficiario.
            if (cuentasPorCobrar.ClienteBeneficiarioId == null || cuentasPorCobrar.ClienteBeneficiarioId <= 0)
            {
                ModelState.AddModelError("ClienteBeneficiarioId", "Debe seleccionar el cliente o beneficiario.");
            }

            // Valida días de crédito.
            if (cuentasPorCobrar.DiasCredito == null || cuentasPorCobrar.DiasCredito <= 0)
            {
                ModelState.AddModelError("DiasCredito", "Debe seleccionar los días de crédito.");
            }

            // Valida monto original.
            if (cuentasPorCobrar.MontoOriginal == null || cuentasPorCobrar.MontoOriginal <= 0)
            {
                ModelState.AddModelError("MontoOriginal", "El monto original debe ser mayor a cero.");
            }

            // Campos calculados o controlados por el sistema.
            ModelState.Remove("NumeroDocumento");
            ModelState.Remove("FechaEmision");
            ModelState.Remove("FechaVencimiento");
            ModelState.Remove("MontoPagado");
            ModelState.Remove("AnticipoAplicado");
            ModelState.Remove("SaldoPendiente");
            ModelState.Remove("Estado");
            ModelState.Remove("Activo");
            ModelState.Remove("FacturaId");
            ModelState.Remove("Factura");
            ModelState.Remove("ClienteBeneficiario");
            ModelState.Remove("Proyecto");
            ModelState.Remove("PagosCuentaPorCobrars");
            ModelState.Remove("Origen");
            ModelState.Remove("EstadoAutorizacion");
            ModelState.Remove("MotivoAnulacion");

            await AplicarSeleccionProyectoPeriodoCxC(
    cuentasPorCobrar,
    SeleccionProyectoPeriodoCxC);

            await ValidarRelacionesActivas(cuentasPorCobrar);
            if (ModelState.IsValid)
            {
                // Fecha de emisión automática.
                cuentasPorCobrar.FechaEmision = fechaHoy;

                // Fecha de vencimiento automática según días de crédito.
                cuentasPorCobrar.FechaVencimiento = fechaHoy.AddDays(cuentasPorCobrar.DiasCredito ?? 0);

                // Número automático según tipo de documento.
                cuentasPorCobrar.NumeroDocumento = GenerarNumeroDocumentoCxC(cuentasPorCobrar.TipoDocumento);

                // Las CxC manuales no nacen desde factura.
                cuentasPorCobrar.FacturaId = null;

                // Inicializa montos.
                cuentasPorCobrar.MontoPagado = 0;
                cuentasPorCobrar.AnticipoAplicado = 0;
                cuentasPorCobrar.SaldoPendiente = cuentasPorCobrar.MontoOriginal ?? 0;

                // Estado inicial.
                cuentasPorCobrar.Estado = "Pendiente";
                cuentasPorCobrar.Activo = true;

                // Regla de negocio:
                // Las cuentas creadas directamente desde CxC nacen con origen manual.
                cuentasPorCobrar.Origen = "Manual";

                // Regla de negocio:
                // Las CxC manuales quedan pendientes de autorización.
                cuentasPorCobrar.EstadoAutorizacion = "Pendiente";

                _context.Add(cuentasPorCobrar);
                await _context.SaveChangesAsync();

                // ================================================================
                // CAMBIO APLICADO - ASIENTO DE NACIMIENTO CXC MANUAL
                // Crea el asiento contable inicial para cuentas por cobrar manuales.
                // No toca facturas.
                // No toca pagos.
                // ================================================================
                await RegistrarAsientoNacimientoCxCManual(cuentasPorCobrar);

                // Auditoría.
                await RegistrarAuditoria(
                    "Crear",
                    cuentasPorCobrar.IdCuentaPorCobrar,
                    "Se creó la cuenta por cobrar " + cuentasPorCobrar.NumeroDocumento +
                    " por un monto de ₡" + cuentasPorCobrar.MontoOriginal
                );

                TempData["MensajeExito"] = "La cuenta por cobrar fue registrada correctamente.";
                return RedirectToAction(nameof(Index));
            }

            // Si hay errores, recarga combos.
            CargarCombos(
                cuentasPorCobrar.ClienteBeneficiarioId,
                cuentasPorCobrar.ProyectoId,
                cuentasPorCobrar.TipoDocumento,
                cuentasPorCobrar.DiasCredito);

            await CargarOpcionesProyectoPeriodoCxC(
    SeleccionProyectoPeriodoCxC); 

            ViewBag.FechaEmision = fechaHoy;

            return View(cuentasPorCobrar);
        }

        // GET: CuentasPorCobrar/Edit/5
        public async Task<IActionResult> Edit(int? id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var cuentasPorCobrar = await _context.CuentasPorCobrars.FindAsync(id);

            if (cuentasPorCobrar == null)
            {
                return NotFound();
            }
            bool periodoCerrado =
    await PeriodoContableCerradoAsync(cuentasPorCobrar.FechaEmision);

            if (periodoCerrado)
            {
                TempData["MensajeError"] =
                    $"No se puede modificar la cuenta por cobrar porque el período " +
                    $"{cuentasPorCobrar.FechaEmision:MM/yyyy} se encuentra cerrado. " +
                    $"Para realizar correcciones, primero debe reabrirse el período.";

                return RedirectToAction(
                    nameof(Index),
                    new
                    {
                        anio = cuentasPorCobrar.FechaEmision.Year,
                        mes = cuentasPorCobrar.FechaEmision.Month
                    });
            }
            // No se modifica una cuenta anulada.
            if (cuentasPorCobrar.Estado == "Anulada" ||
     cuentasPorCobrar.Estado == "Cancelada")
            {
                TempData["MensajeError"] =
                    "No se puede modificar una cuenta por cobrar cancelada o anulada.";

                return RedirectToAction(nameof(Index));
            }
            // Si nació desde factura, se debe modificar desde facturación.
            if (cuentasPorCobrar.Origen == "Factura" || cuentasPorCobrar.FacturaId != null)
            {
                TempData["MensajeError"] =
                    "Esta cuenta por cobrar fue generada desde una factura. Debe modificar la factura de origen para sincronizar correctamente el sistema.";

                return RedirectToAction(nameof(Index));
            }

            CargarCombosEdit(
    cuentasPorCobrar.ClienteBeneficiarioId,
    cuentasPorCobrar.ProyectoId,
    cuentasPorCobrar.TipoDocumento,
    cuentasPorCobrar.DiasCredito);

            string? seleccionActual =
    cuentasPorCobrar.PresupuestoMensualId.HasValue
        ? $"PRES-{cuentasPorCobrar.PresupuestoMensualId.Value}"
        : cuentasPorCobrar.ProyectoId.HasValue
            ? $"PROY-{cuentasPorCobrar.ProyectoId.Value}"
            : null;

            await CargarOpcionesProyectoPeriodoCxC(
                seleccionActual);

            return View(cuentasPorCobrar);
        }

        // POST: CuentasPorCobrar/Edit/5
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Edit(int id,[Bind("IdCuentaPorCobrar,ClienteBeneficiarioId,TipoDocumento,Concepto,MontoOriginal,DiasCredito,Observacion")]
    CuentasPorCobrar cuentasPorCobrar, string? SeleccionProyectoPeriodoCxC)
        {
            if (id != cuentasPorCobrar.IdCuentaPorCobrar)
            {
                return NotFound();
            }

            var cuentaActual = await _context.CuentasPorCobrars
                .FirstOrDefaultAsync(c => c.IdCuentaPorCobrar == id);

            if (cuentaActual == null)
            {
                return NotFound();
            }
            bool periodoCerrado =
    await PeriodoContableCerradoAsync(cuentaActual.FechaEmision);

            if (periodoCerrado)
            {
                TempData["MensajeError"] =
                    $"No se puede modificar la cuenta por cobrar porque el período " +
                    $"{cuentaActual.FechaEmision:MM/yyyy} se encuentra cerrado. " +
                    $"Para realizar correcciones, primero debe reabrirse el período.";

                return RedirectToAction(
                    nameof(Index),
                    new
                    {
                        anio = cuentaActual.FechaEmision.Year,
                        mes = cuentaActual.FechaEmision.Month
                    });
            }
            // No se modifica una cuenta anulada.
            if (cuentaActual.Estado == "Anulada")
            {
                TempData["MensajeError"] = "No se puede modificar una cuenta por cobrar anulada.";
                return RedirectToAction(nameof(Index));
            }

            // Si nació desde factura, se debe modificar desde facturación.
            if (cuentaActual.Origen == "Factura" || cuentaActual.FacturaId != null)
            {
                TempData["MensajeError"] =
                    "Esta cuenta por cobrar fue generada desde una factura. Debe modificar la factura de origen para sincronizar correctamente el sistema.";

                return RedirectToAction(nameof(Index));
            }

            // Verifica si tiene pagos activos.
            bool tienePagos = await _context.PagosCuentaPorCobrars
                .AnyAsync(p => p.CuentaPorCobrarId == cuentaActual.IdCuentaPorCobrar
                            && p.Anulado == false);

            // Regla de negocio:
            // Si ya tiene pagos, solo se permite modificar la observación.
            if (tienePagos)
            {
                cuentaActual.Observacion = cuentasPorCobrar.Observacion;

                await _context.SaveChangesAsync();

                await RegistrarAuditoria(
                    "Modificar",
                    cuentaActual.IdCuentaPorCobrar,
                    "Se modificó la observación de la cuenta por cobrar " + cuentaActual.NumeroDocumento
                );

                TempData["MensajeExito"] = "La observación fue modificada correctamente.";
                return RedirectToAction(nameof(Index));
            }

            // Validaciones cuando no tiene pagos.
            if (cuentasPorCobrar.ClienteBeneficiarioId == null || cuentasPorCobrar.ClienteBeneficiarioId <= 0)
            {
                ModelState.AddModelError("ClienteBeneficiarioId", "Debe seleccionar el cliente o beneficiario.");
            }

            if (cuentasPorCobrar.DiasCredito == null || cuentasPorCobrar.DiasCredito <= 0)
            {
                ModelState.AddModelError("DiasCredito", "Debe seleccionar los días de crédito.");
            }

            if (cuentasPorCobrar.MontoOriginal == null || cuentasPorCobrar.MontoOriginal <= 0)
            {
                ModelState.AddModelError("MontoOriginal", "El monto original debe ser mayor a cero.");
            }

            // Campos calculados o controlados por el sistema.
            ModelState.Remove("NumeroDocumento");
            ModelState.Remove("FechaEmision");
            ModelState.Remove("FechaVencimiento");
            ModelState.Remove("MontoPagado");
            ModelState.Remove("AnticipoAplicado");
            ModelState.Remove("SaldoPendiente");
            ModelState.Remove("Estado");
            ModelState.Remove("Activo");
            ModelState.Remove("FacturaId");
            ModelState.Remove("Factura");
            ModelState.Remove("ClienteBeneficiario");
            ModelState.Remove("Proyecto");
            ModelState.Remove("PresupuestoMensualRegistro");
            ModelState.Remove("PagosCuentaPorCobrars");
            ModelState.Remove("Origen");
            ModelState.Remove("EstadoAutorizacion");
            ModelState.Remove("MotivoAnulacion");

            await AplicarSeleccionProyectoPeriodoCxC(
                cuentasPorCobrar,
                SeleccionProyectoPeriodoCxC,
                  cuentaActual.PresupuestoMensualId);

            await ValidarRelacionesActivas(
                cuentasPorCobrar,
                cuentaActual.ClienteBeneficiarioId,
                cuentaActual.ProyectoId);

            if (!ModelState.IsValid)
            {
                CargarCombosEdit(
                    cuentasPorCobrar.ClienteBeneficiarioId,
                    cuentasPorCobrar.ProyectoId,
                    cuentasPorCobrar.TipoDocumento,
                    cuentasPorCobrar.DiasCredito);

                await CargarOpcionesProyectoPeriodoCxC(
                    SeleccionProyectoPeriodoCxC);

                return View(cuentasPorCobrar);
            }
            // ================================================================
            // CAMBIO APLICADO - MODIFICACIÓN DE ASIENTO CXC MANUAL
            // Si la CxC manual cambia datos contables y no tiene pagos activos,
            // se revierte el asiento de nacimiento anterior.
            // Luego, más abajo, se genera nuevamente con los datos modificados.
            // No aplica para CxC nacidas desde factura.
            // ================================================================
            bool requiereActualizarAsientoManual =
                cuentaActual.Origen == "Manual" &&
                cuentaActual.FacturaId == null &&
                (
                    cuentaActual.MontoOriginal != cuentasPorCobrar.MontoOriginal ||
                    cuentaActual.ProyectoId != cuentasPorCobrar.ProyectoId ||
                      cuentaActual.PresupuestoMensualId != cuentasPorCobrar.PresupuestoMensualId ||
                    cuentaActual.DiasCredito != cuentasPorCobrar.DiasCredito
                );

            if (requiereActualizarAsientoManual)
            {
                await RevertirAsientoNacimientoCxCManual(
                    cuentaActual,
                    "Modificación de CxC manual antes de regenerar asiento contable"
                );
            }

            // Actualiza datos permitidos.
            cuentaActual.ClienteBeneficiarioId = cuentasPorCobrar.ClienteBeneficiarioId;
            cuentaActual.ProyectoId = cuentasPorCobrar.ProyectoId;

            cuentaActual.PresupuestoMensualId =
    cuentasPorCobrar.PresupuestoMensualId;
            // No se modifica el tipo de documento para no romper el consecutivo.
            // cuentaActual.TipoDocumento = cuentasPorCobrar.TipoDocumento;

            cuentaActual.Concepto = cuentasPorCobrar.Concepto;
            cuentaActual.MontoOriginal = cuentasPorCobrar.MontoOriginal;
            cuentaActual.DiasCredito = cuentasPorCobrar.DiasCredito;
            cuentaActual.Observacion = cuentasPorCobrar.Observacion;

            // Recalcula vencimiento.
            cuentaActual.FechaVencimiento = cuentaActual.FechaEmision.AddDays(cuentasPorCobrar.DiasCredito ?? 0);

            // Recalcula saldo pendiente.
            cuentaActual.SaldoPendiente =
                (cuentaActual.MontoOriginal ?? 0)
                - (cuentaActual.MontoPagado ?? 0)
                - (cuentaActual.AnticipoAplicado ?? 0);

            if (cuentaActual.SaldoPendiente < 0)
            {
                cuentaActual.SaldoPendiente = 0;
            }

            // Actualiza estado según saldo, pagos y vencimiento.
            ActualizarEstadoCuenta(cuentaActual);

            await _context.SaveChangesAsync();


            // SINCRONIZA INGRESOS GENERADOS POR PAGOS CXC
            // Si la CxC fue modificada, cualquier ingreso generado anteriormente
            // desde sus pagos conserva su monto/fecha/estado, pero actualiza
            // proyecto, período presupuestario y cliente.
            await SincronizarIngresosPagosCxC(cuentaActual);

            // CAMBIO APLICADO - NUEVO ASIENTO MODIFICADO CXC MANUAL
            // Después de guardar los nuevos datos, se registra el asiento nuevo
            // con monto/proyecto/datos actualizados.
            if (requiereActualizarAsientoManual)
            {
                await RegistrarAsientoNacimientoCxCManual(cuentaActual);
            }

            await RegistrarAuditoria(
                "Modificar",
                cuentaActual.IdCuentaPorCobrar,
                "Se modificó la cuenta por cobrar " + cuentaActual.NumeroDocumento
            );

            TempData["MensajeExito"] = "La cuenta por cobrar fue modificada correctamente.";
            return RedirectToAction(nameof(Index));
        }

        // SINCRONIZA INGRESOS GENERADOS DESDE PAGOS DE UNA CXC
        // Mantiene alineados los datos administrativos/presupuestarios
        // de los ingresos generados por pagos CxC cuando la cuenta se edita.
        // IMPORTANTE: No cambia monto, fecha, estado ni condición de anulación del ingreso. Solo sincroniza los datos que pertenecen a la CxC de origen.
        private async Task SincronizarIngresosPagosCxC(
            CuentasPorCobrar cuenta)
        {
            if (string.IsNullOrWhiteSpace(cuenta.NumeroDocumento))
            {
                return;
            }

            var ingresosAsociados = await _context.Ingresos
                .Where(i =>
                    i.Fuente == cuenta.NumeroDocumento &&
                    i.Comprobante != null &&
                    i.Comprobante.StartsWith("PAGO-CXC-"))
                .ToListAsync();

            if (!ingresosAsociados.Any())
            {
                return;
            }

            foreach (var ingreso in ingresosAsociados)
            {
                // Sincroniza proyecto.
                ingreso.ProyectoId =
                    cuenta.ProyectoId;

                // Sincroniza período presupuestario.
                ingreso.PresupuestoMensualId =
                    cuenta.PresupuestoMensualId;

                // Sincroniza cliente / beneficiario.
                ingreso.ClienteBeneficiarioId =
                    cuenta.ClienteBeneficiarioId;
            }

            await _context.SaveChangesAsync();
        }
        private async Task ValidarRelacionesActivas(
    CuentasPorCobrar cuenta,
    int? clienteActualId = null,
    int? proyectoActualId = null)
        {
            bool clienteValido = await _context.ClientesBeneficiarios.AnyAsync(c =>
                c.IdClienteBeneficiario == cuenta.ClienteBeneficiarioId &&
                (c.Activo || c.IdClienteBeneficiario == clienteActualId) &&
                c.Nombre != "Consumidor Final");

            if (!clienteValido)
            {
                ModelState.AddModelError(
                    "ClienteBeneficiarioId",
                    "Debe seleccionar un cliente o beneficiario activo.");
            }

            if (cuenta.ProyectoId.HasValue)
            {
                bool proyectoValido = await _context.Proyectos.AnyAsync(p =>
                    p.IdProyecto == cuenta.ProyectoId &&
                    (p.Activo || p.IdProyecto == proyectoActualId));

                if (!proyectoValido)
                {
                    ModelState.AddModelError(
                        "ProyectoId",
                        "Debe seleccionar un proyecto activo.");
                }
            }
        }
        // GET: CuentasPorCobrar/Delete/5
        public async Task<IActionResult> Delete(int? id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var cuentasPorCobrar = await _context.CuentasPorCobrars
                .Include(c => c.ClienteBeneficiario)
                .Include(c => c.Proyecto)
                .Include(c => c.PresupuestoMensualRegistro)
                .FirstOrDefaultAsync(m => m.IdCuentaPorCobrar == id);

            if (cuentasPorCobrar == null)
            {
                return NotFound();
            }
            // ============================================================
            // BLOQUEO POR CIERRE CONTABLE
            // ============================================================
            bool periodoCerrado =
                await PeriodoContableCerradoAsync(cuentasPorCobrar.FechaEmision);

            if (periodoCerrado)
            {
                TempData["MensajeError"] =
                    $"No se puede anular la cuenta por cobrar porque el período " +
                    $"{cuentasPorCobrar.FechaEmision:MM/yyyy} se encuentra cerrado. " +
                    $"Para realizar correcciones, primero debe reabrirse el período.";

                return RedirectToAction(
                    nameof(Index),
                    new
                    {
                        anio = cuentasPorCobrar.FechaEmision.Year,
                        mes = cuentasPorCobrar.FechaEmision.Month
                    });
            }
            if (cuentasPorCobrar.Estado == "Anulada")
            {
                TempData["MensajeError"] =
                    "La cuenta por cobrar ya se encuentra anulada.";

                return RedirectToAction(nameof(Index));
            }

            if (cuentasPorCobrar.Estado == "Cancelada")
            {
                TempData["MensajeError"] =
                    "No se puede anular una cuenta por cobrar cancelada.";

                return RedirectToAction(nameof(Index));
            }

            bool tienePagosActivos = await _context.PagosCuentaPorCobrars
                .AnyAsync(p =>
                    p.CuentaPorCobrarId == cuentasPorCobrar.IdCuentaPorCobrar &&
                    !p.Anulado);

            if (tienePagosActivos)
            {
                TempData["MensajeError"] =
                    "No se puede anular la cuenta porque tiene pagos activos.";

                return RedirectToAction(nameof(Index));
            }
            return View(cuentasPorCobrar);
        }

        // POST: CuentasPorCobrar/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> DeleteConfirmed(int id, string motivoAnulacion)
        {
            var cuentasPorCobrar = await _context.CuentasPorCobrars
                .FirstOrDefaultAsync(c => c.IdCuentaPorCobrar == id);

            if (cuentasPorCobrar == null)
            {
                return NotFound();
            }
            // ============================================================
            // BLOQUEO POR CIERRE CONTABLE
            // ============================================================
            bool periodoCerrado =
                await PeriodoContableCerradoAsync(cuentasPorCobrar.FechaEmision);

            if (periodoCerrado)
            {
                TempData["MensajeError"] =
                    $"No se puede anular la cuenta por cobrar porque el período " +
                    $"{cuentasPorCobrar.FechaEmision:MM/yyyy} se encuentra cerrado. " +
                    $"Para realizar correcciones, primero debe reabrirse el período.";

                return RedirectToAction(
                    nameof(Index),
                    new
                    {
                        anio = cuentasPorCobrar.FechaEmision.Year,
                        mes = cuentasPorCobrar.FechaEmision.Month
                    });
            }
            // Valida motivo obligatorio.
            if (string.IsNullOrWhiteSpace(motivoAnulacion))
            {
                TempData["MensajeError"] = "Debe ingresar el motivo de anulación.";
                return RedirectToAction(nameof(Delete), new { id });
            }

            // Valida que el motivo tenga al menos tres letras y caracteres permitidos.
            if (!System.Text.RegularExpressions.Regex.IsMatch(
                motivoAnulacion,
                @"^(?=(?:.*[a-zA-ZáéíóúÁÉÍÓÚñÑ]){3,})[a-zA-ZáéíóúÁÉÍÓÚñÑ0-9\s.,;:/()#-]+$"))
            {
                TempData["MensajeError"] = "El motivo de anulación debe contener al menos tres letras y solo caracteres permitidos.";
                return RedirectToAction(nameof(Delete), new { id });
            }

            // Si la cuenta nació desde factura,
            // no se permite anularla desde CxC.
            if (cuentasPorCobrar.Origen == "Factura" || cuentasPorCobrar.FacturaId != null)
            {
                TempData["MensajeError"] =
                    "Esta cuenta por cobrar fue generada desde una factura. Debe anular la factura de origen para sincronizar correctamente el sistema.";

                return RedirectToAction(nameof(Index));
            }

            // Bloquea anulación si tiene pagos activos.
            bool tienePagosActivos = await _context.PagosCuentaPorCobrars
                .AnyAsync(p => p.CuentaPorCobrarId == cuentasPorCobrar.IdCuentaPorCobrar
                            && p.Anulado == false);

            if (tienePagosActivos)
            {
                TempData["MensajeError"] =
                    "No se puede anular esta cuenta por cobrar porque tiene pagos activos registrados.";

                return RedirectToAction(nameof(Index));
            }

            // Anula la cuenta por cobrar.
            cuentasPorCobrar.Estado = "Anulada";
            cuentasPorCobrar.Activo = false;
            cuentasPorCobrar.MontoPagado = 0;
            cuentasPorCobrar.SaldoPendiente = 0;
            cuentasPorCobrar.MotivoAnulacion = motivoAnulacion.Trim();

            // Anula pagos asociados si existieran.
            var pagosAsociados = await _context.PagosCuentaPorCobrars
                .Where(p => p.CuentaPorCobrarId == cuentasPorCobrar.IdCuentaPorCobrar)
                .ToListAsync();

            foreach (var pago in pagosAsociados)
            {
                pago.Anulado = true;
                pago.MotivoAnulacion =
                    "Pago anulado automáticamente por anulación de la cuenta por cobrar "
                    + cuentasPorCobrar.NumeroDocumento;
               
            }

            // Anula ingresos automáticos generados desde pagos CxC.
            var ingresosAsociados = await _context.Ingresos
                .Where(i => i.Fuente == cuentasPorCobrar.NumeroDocumento
                            && i.Comprobante != null
                            && i.Comprobante.StartsWith("PAGO-CXC-"))
                .ToListAsync();

            foreach (var ingreso in ingresosAsociados)
            {
                ingreso.Estado = "Anulado";
                ingreso.Activo = false;
                ingreso.Observaciones =
                    (ingreso.Observaciones ?? "") +
                    " | Anulado automáticamente por anulación de CxC. Motivo: "
                    + motivoAnulacion.Trim();

                _context.Update(ingreso);
            }

            _context.Update(cuentasPorCobrar);
            await _context.SaveChangesAsync();

            // ================================================================
            // CAMBIO APLICADO - REVERSIÓN ASIENTO NACIMIENTO CXC MANUAL
            // Revierte el asiento inicial de la CxC manual cuando se anula.
            // No toca facturas.
            // No toca pagos activos porque ya están bloqueados arriba.
            // ================================================================
            await RevertirAsientoNacimientoCxCManual(
                cuentasPorCobrar,
                motivoAnulacion.Trim()
            );

            await RegistrarAuditoria(
                "Anular",
                cuentasPorCobrar.IdCuentaPorCobrar,
                "Se anuló la cuenta por cobrar " + cuentasPorCobrar.NumeroDocumento +
                ". Motivo: " + motivoAnulacion.Trim()
            );

            TempData["MensajeExito"] = "La cuenta por cobrar fue anulada correctamente.";
            return RedirectToAction(nameof(Index));
        }
        private async Task CargarOpcionesProyectoPeriodoCxC( string? seleccionActual = null)
        {
            int? idPresupuestoActual = null;

            if (!string.IsNullOrWhiteSpace(seleccionActual) &&
                seleccionActual.StartsWith("PRES-"))
            {
                string valorIdActual =
                    seleccionActual.Substring(5);

                if (int.TryParse(
                    valorIdActual,
                    out int idActual))
                {
                    idPresupuestoActual = idActual;
                }
            }
            var cultura =
                new System.Globalization.CultureInfo("es-CR");

            var hoy =
                DateOnly.FromDateTime(DateTime.Today);

            var fechaMinima =
                hoy.AddMonths(-3);

            int anioMinimo = fechaMinima.Year;
            int mesMinimo = fechaMinima.Month;

            // ============================================================
            // 1. PERÍODOS PRESUPUESTARIOS UTILIZABLES
            // Aprobados o Borrador de proyectos activos
            // ============================================================
            var periodos = await _context.PresupuestosMensuales
    .AsNoTracking()
    .Include(pm => pm.Proyecto)
    .Where(pm =>

        (
            pm.ProyectoId != null &&
            pm.Proyecto != null &&
            pm.Proyecto.Activo &&

            (
                pm.EstadoAprobacion == "Aprobado" ||
                pm.EstadoAprobacion == "Borrador" ||
                pm.EstadoAprobacion == "Rechazado"
            ) &&

            // NUEVO:
            // No mostrar períodos contables cerrados.
            !_context.CierresContables.Any(c =>
                c.Anio == pm.Anio &&
                c.Mes == pm.Mes &&
                c.Estado == "Cerrado") &&

            (
                pm.Anio > anioMinimo ||
                (pm.Anio == anioMinimo &&
                 pm.Mes >= mesMinimo)
            )
        )

        // En Edit conserva únicamente
        // el presupuesto histórico actualmente asignado.
        ||
        (
            idPresupuestoActual.HasValue &&
            pm.IdPresupuestoMensual ==
                idPresupuestoActual.Value
        )
    )
    .OrderByDescending(pm => pm.Anio)
    .ThenByDescending(pm => pm.Mes)
    .ThenBy(pm => pm.Proyecto!.Nombre)
    .ToListAsync();


            var opciones =
                new List<SelectListItem>();

            foreach (var periodo in periodos)
            {
                string nombreMes =
                    new DateTime(
                        periodo.Anio,
                        periodo.Mes,
                        1)
                    .ToString("MMMM", cultura);

                nombreMes =
                    char.ToUpper(nombreMes[0]) +
                    nombreMes.Substring(1);

                string valor =
                    $"PRES-{periodo.IdPresupuestoMensual}";

                opciones.Add(
                    new SelectListItem
                    {
                        Value = valor,
                        Text =
                            $"{periodo.Proyecto!.Nombre} — " +
                            $"{nombreMes} {periodo.Anio} " +
                            $"({periodo.EstadoAprobacion})",
                        Selected =
                            valor == seleccionActual
                    });
            }

            // ============================================================
            // 2. PROYECTOS ACTIVOS SIN PERÍODO UTILIZABLE
            // ============================================================

            var proyectosSinPeriodo =
                await _context.Proyectos
                    .AsNoTracking()
                    .Where(p =>
                        p.Activo &&
                        !p.PresupuestosMensuales.Any(pm =>
    (
        pm.EstadoAprobacion == "Aprobado" ||
        pm.EstadoAprobacion == "Borrador" ||
        pm.EstadoAprobacion == "Rechazado"
    ) &&
                // NUEVO:
                // Un período cerrado ya no cuenta
                // como período utilizable.
                !_context.CierresContables.Any(c =>
                    c.Anio == pm.Anio &&
                    c.Mes == pm.Mes &&
                    c.Estado == "Cerrado") &&
                            (
                                pm.Anio > anioMinimo ||
                                (pm.Anio == anioMinimo &&
                                 pm.Mes >= mesMinimo)
                            )))
                    .OrderBy(p => p.Nombre)
                    .ToListAsync();

            foreach (var proyecto in proyectosSinPeriodo)
            {
                string valor =
                    $"PROY-{proyecto.IdProyecto}";

                opciones.Add(
                    new SelectListItem
                    {
                        Value = valor,
                        Text =
                            $"{proyecto.Nombre} — " +
                            "Sin período presupuestario",
                        Selected =
                            valor == seleccionActual
                    });
            }

            ViewBag.OpcionesProyectoPeriodoCxC =
                opciones;

            ViewBag.SeleccionProyectoPeriodoCxC =
                seleccionActual;
        }
        private async Task AplicarSeleccionProyectoPeriodoCxC(
    CuentasPorCobrar cuenta,
    string? seleccionProyectoPeriodo,
    int? presupuestoActualId = null)
        {
            ModelState.Remove(nameof(cuenta.ProyectoId));
            ModelState.Remove(nameof(cuenta.PresupuestoMensualId));

            cuenta.ProyectoId = null;
            cuenta.PresupuestoMensualId = null;

            if (string.IsNullOrWhiteSpace(
                seleccionProyectoPeriodo))
            {
                ModelState.AddModelError(
                    nameof(cuenta.ProyectoId),
                    "Debe seleccionar un proyecto y período presupuestario.");

                return;
            }

            // ============================================================
            // PERÍODO PRESUPUESTARIO
            // ============================================================

            if (seleccionProyectoPeriodo.StartsWith("PRES-"))
            {
                string valorId =
                    seleccionProyectoPeriodo.Substring(5);

                if (!int.TryParse(valorId, out int idPresupuesto))
                {
                    ModelState.AddModelError(
                        nameof(cuenta.ProyectoId),
                        "La selección presupuestaria no es válida.");

                    return;
                }

                var presupuesto =
                    await _context.PresupuestosMensuales
                        .Include(pm => pm.Proyecto)
                        .FirstOrDefaultAsync(pm =>
                            pm.IdPresupuestoMensual ==
                            idPresupuesto);

                if (presupuesto == null ||
                    presupuesto.Proyecto == null ||
                    !presupuesto.Proyecto.Activo ||
                    (
                        presupuesto.EstadoAprobacion != "Aprobado" &&
                        presupuesto.EstadoAprobacion != "Borrador" &&
                        presupuesto.EstadoAprobacion != "Rechazado"
                    ))
                {
                    ModelState.AddModelError(
                        nameof(cuenta.ProyectoId),
                        "El período presupuestario seleccionado ya no está disponible.");

                    return;
                }
                // ============================================================
                // NUEVO:
                // No permitir utilizar un período contable cerrado,
                // excepto si en Edit es exactamente el presupuesto histórico
                // que ya tenía asignado la CxC.
                // ============================================================

                bool esPresupuestoHistoricoActual =
                    presupuestoActualId.HasValue &&
                    presupuesto.IdPresupuestoMensual ==
                        presupuestoActualId.Value;

                if (!esPresupuestoHistoricoActual)
                {
                    bool periodoCerrado =
                        await _context.CierresContables
                            .AsNoTracking()
                            .AnyAsync(c =>
                                c.Anio == presupuesto.Anio &&
                                c.Mes == presupuesto.Mes &&
                                c.Estado == "Cerrado");

                    if (periodoCerrado)
                    {
                        ModelState.AddModelError(
                            nameof(cuenta.ProyectoId),
                            "El período presupuestario seleccionado pertenece a un período contable cerrado.");

                        return;
                    }
                }

                cuenta.ProyectoId =
                    presupuesto.ProyectoId;

                cuenta.PresupuestoMensualId =
                    presupuesto.IdPresupuestoMensual;

                return;
            }

            // ============================================================
            // PROYECTO SIN PERÍODO
            // ============================================================

            if (seleccionProyectoPeriodo.StartsWith("PROY-"))
            {
                string valorId =
                    seleccionProyectoPeriodo.Substring(5);

                if (!int.TryParse(valorId, out int idProyecto))
                {
                    ModelState.AddModelError(
                        nameof(cuenta.ProyectoId),
                        "El proyecto seleccionado no es válido.");

                    return;
                }

                var proyecto =
                    await _context.Proyectos
                        .FirstOrDefaultAsync(p =>
                            p.IdProyecto == idProyecto &&
                            p.Activo);

                if (proyecto == null)
                {
                    ModelState.AddModelError(
                        nameof(cuenta.ProyectoId),
                        "El proyecto seleccionado no está disponible.");

                    return;
                }

                cuenta.ProyectoId =
                    proyecto.IdProyecto;

                cuenta.PresupuestoMensualId =
                    null;

                return;
            }

            ModelState.AddModelError(
                nameof(cuenta.ProyectoId),
                "La selección de proyecto y período no es válida.");
        }

        // Carga listas usadas en Create y Edit.
        private void CargarCombos(
            int? clienteBeneficiarioId = null,
            int? proyectoId = null,
            string? tipoDocumento = null,
            int? diasCredito = null)
        {
            ViewData["ClienteBeneficiarioId"] = new SelectList(
                _context.ClientesBeneficiarios
                    .Where(c => c.Activo == true && c.Nombre != "Consumidor Final")
                .OrderBy(c => c.Nombre),
                "IdClienteBeneficiario",
                "Nombre",
                clienteBeneficiarioId);

            ViewData["ProyectoId"] = new SelectList(
                _context.Proyectos.Where(p => p.Activo == true),
                "IdProyecto",
                "Nombre",
                proyectoId);

            ViewBag.TiposDocumento = new SelectList(new List<string>
            {
                "Donación",
                "Cuota",
                "Convenio",
                "Ajuste"
            }, tipoDocumento);

            ViewBag.DiasCredito = new SelectList(new List<int>
            {
                8, 15, 30, 45, 60, 90, 120, 150, 180
            }, diasCredito);
        }
        private void CargarCombosEdit(
    int? clienteBeneficiarioId,
    int? proyectoId,
    string? tipoDocumento,
    int? diasCredito)
        {
            ViewData["ClienteBeneficiarioId"] = new SelectList(
                _context.ClientesBeneficiarios
                    .Where(c =>
                        (c.Activo ||
                         c.IdClienteBeneficiario == clienteBeneficiarioId)
                        &&
                        c.Nombre != "Consumidor Final")
                    .OrderBy(c => c.Nombre),
                "IdClienteBeneficiario",
                "Nombre",
                clienteBeneficiarioId);

            ViewData["ProyectoId"] = new SelectList(
                _context.Proyectos
                    .Where(p =>
                        p.Activo ||
                        p.IdProyecto == proyectoId)
                    .OrderBy(p => p.Nombre),
                "IdProyecto",
                "Nombre",
                proyectoId);

            ViewBag.TiposDocumento = new SelectList(new List<string>
    {
        "Donación",
        "Cuota",
        "Convenio",
        "Ajuste"
    }, tipoDocumento);

            ViewBag.DiasCredito = new SelectList(new List<int>
    {
        8,15,30,45,60,90,120,150,180
    }, diasCredito);
        }
        // Genera el número automático según el tipo de documento.
        private string GenerarNumeroDocumentoCxC(string? tipoDocumento)
        {
            string prefijo = tipoDocumento switch
            {
                "Donación" => "DON",
                "Cuota" => "CUO",
                "Convenio" => "CON",
                "Ajuste" => "AJU",
                _ => "CXC"
            };

            var ultimoDocumento = _context.CuentasPorCobrars
                .Where(c => c.TipoDocumento == tipoDocumento && c.NumeroDocumento != null)
                .OrderByDescending(c => c.IdCuentaPorCobrar)
                .Select(c => c.NumeroDocumento)
                .FirstOrDefault();

            int consecutivo = 1;

            if (!string.IsNullOrEmpty(ultimoDocumento))
            {
                var partes = ultimoDocumento.Split('-');

                if (partes.Length == 2 && int.TryParse(partes[1], out int numeroActual))
                {
                    consecutivo = numeroActual + 1;
                }
            }

            return $"{prefijo}-{consecutivo.ToString("D5")}";
        }
        // ================================================================
        // CAMBIO APLICADO - ASIENTO DE NACIMIENTO CXC MANUAL
        // Debe: Cuentas por Cobrar CC006
        // Haber: Cuenta de ingreso según tipo de documento
        // ================================================================
        private async Task RegistrarAsientoNacimientoCxCManual(CuentasPorCobrar cuenta)
        {
            decimal monto = cuenta.MontoOriginal ?? 0;

            if (monto <= 0)
            {
                return;
            }

            string referencia = "CXC-MANUAL-" + cuenta.IdCuentaPorCobrar;

            bool yaExiste = await _context.MovimientosContables
                .AnyAsync(m => m.Referencia == referencia && m.Estado != "Anulado");

            if (yaExiste)
            {
                return;
            }

            var cuentaCxC = await _context.CuentasContables
                .FirstOrDefaultAsync(c => c.Codigo == "CC006");

            var cuentaIngreso = await ObtenerCuentaIngresoCxCManual(cuenta.TipoDocumento);

            if (cuentaCxC == null || cuentaIngreso == null)
            {
                return;
            }

            var usuarioId = HttpContext.Session.GetInt32("UsuarioId") ?? 1;

            var movimientoDebe = new MovimientosContable
            {
                Fecha = cuenta.FechaEmision,
                CuentaContableId = cuentaCxC.IdCuentaContable,
                ProyectoId = cuenta.ProyectoId,
                TipoMovimiento = "Debe",
                Monto = monto,
                Debe = monto,
                Haber = 0,
                Referencia = referencia,
                OrigenModulo = "CxC Manual",
                OrigenId = cuenta.IdCuentaPorCobrar,
                Descripcion = "Registro inicial de cuenta por cobrar manual " + cuenta.NumeroDocumento,
                Estado = "Registrado",
                EsAutomatico = true,
                UsuarioId = usuarioId,
                FechaCreacion = DateTime.Now
            };

            var movimientoHaber = new MovimientosContable
            {
                Fecha = cuenta.FechaEmision,
                CuentaContableId = cuentaIngreso.IdCuentaContable,
                ProyectoId = cuenta.ProyectoId,
                TipoMovimiento = "Haber",
                Monto = monto,
                Debe = 0,
                Haber = monto,
                Referencia = referencia,
                OrigenModulo = "CxC Manual",
                OrigenId = cuenta.IdCuentaPorCobrar,
                Descripcion = "Ingreso reconocido por cuenta por cobrar manual " + cuenta.NumeroDocumento,
                Estado = "Registrado",
                EsAutomatico = true,
                UsuarioId = usuarioId,
                FechaCreacion = DateTime.Now
            };

            _context.MovimientosContables.Add(movimientoDebe);
            _context.MovimientosContables.Add(movimientoHaber);

            AplicarSaldoCuenta(cuentaCxC, monto, 0);
            AplicarSaldoCuenta(cuentaIngreso, 0, monto);

            _context.CuentasContables.Update(cuentaCxC);
            _context.CuentasContables.Update(cuentaIngreso);

            await _context.SaveChangesAsync();
        }

        // ================================================================
        // CAMBIO APLICADO - REVERSIÓN ASIENTO NACIMIENTO CXC MANUAL
        // Debe: Cuenta de ingreso
        // Haber: Cuentas por Cobrar CC006
        // ================================================================
        private async Task RevertirAsientoNacimientoCxCManual(
            CuentasPorCobrar cuenta,
            string motivoAnulacion)
        {
            decimal monto = cuenta.MontoOriginal ?? 0;

            if (monto <= 0)
            {
                return;
            }

            string referenciaOriginal = "CXC-MANUAL-" + cuenta.IdCuentaPorCobrar;
            // ================================================================
            // CAMBIO APLICADO - REVERSIÓN CXC MANUAL CON REFERENCIA ÚNICA
            // Permite modificar una CxC manual más de una vez.
            // Antes se bloqueaba porque ya existía REV-CXC-MANUAL-ID.
            // ================================================================
            string referenciaReversa =
                "REV-CXC-MANUAL-" + cuenta.IdCuentaPorCobrar + "-" + DateTime.Now.ToString("yyyyMMddHHmmss");

            var movimientosOriginales = await _context.MovimientosContables
       .Where(m =>
           m.Referencia == referenciaOriginal &&
           m.OrigenModulo == "CxC Manual" &&
           m.Estado == "Registrado" &&
           m.Anulado == false &&
           m.MovimientoReversionId == null)
       .ToListAsync();

            if (!movimientosOriginales.Any())
            {
                return;
            }

            var cuentaCxC = await _context.CuentasContables
                .FirstOrDefaultAsync(c => c.Codigo == "CC006");

            var cuentaIngreso = await ObtenerCuentaIngresoCxCManual(cuenta.TipoDocumento);

            if (cuentaCxC == null || cuentaIngreso == null)
            {
                return;
            }

            var usuarioId = HttpContext.Session.GetInt32("UsuarioId") ?? 1;

            foreach (var movimiento in movimientosOriginales)
            {
                movimiento.Estado = "Anulado";
                movimiento.Anulado = true;
                movimiento.FechaAnulacion = DateTime.Now;
                movimiento.MotivoAnulacion = "Anulación de CxC manual. Motivo: " + motivoAnulacion;

                _context.MovimientosContables.Update(movimiento);
            }

            var movimientoDebeIngreso = new MovimientosContable
            {
                Fecha = DateOnly.FromDateTime(DateTime.Today),
                CuentaContableId = cuentaIngreso.IdCuentaContable,
                ProyectoId = cuenta.ProyectoId,
                TipoMovimiento = "Debe",
                Monto = monto,
                Debe = monto,
                Haber = 0,
                Referencia = referenciaReversa,
                OrigenModulo = "CxC Manual",
                OrigenId = cuenta.IdCuentaPorCobrar,
                Descripcion = "Reversión de ingreso por anulación de CxC manual " + cuenta.NumeroDocumento,
                Estado = "Reversado",
                EsAutomatico = true,
                UsuarioId = usuarioId,
                FechaCreacion = DateTime.Now
            };

            var movimientoHaberCxC = new MovimientosContable
            {
                Fecha = DateOnly.FromDateTime(DateTime.Today),
                CuentaContableId = cuentaCxC.IdCuentaContable,
                ProyectoId = cuenta.ProyectoId,
                TipoMovimiento = "Haber",
                Monto = monto,
                Debe = 0,
                Haber = monto,
                Referencia = referenciaReversa,
                OrigenModulo = "CxC Manual",
                OrigenId = cuenta.IdCuentaPorCobrar,
                Descripcion = "Reversión de cuenta por cobrar manual " + cuenta.NumeroDocumento,
                Estado = "Reversado",
                EsAutomatico = true,
                UsuarioId = usuarioId,
                FechaCreacion = DateTime.Now
            };

            _context.MovimientosContables.Add(movimientoDebeIngreso);
            _context.MovimientosContables.Add(movimientoHaberCxC);

            // ================================================================
            // CAMBIO APLICADO - ENLACE DE REVERSIÓN CXC MANUAL
            // Relaciona cada movimiento original con su movimiento de reversión.
            // Esto permite mostrar en Details el botón celeste:
            // "Reversión del movimiento #..."
            // ================================================================
            await _context.SaveChangesAsync();

            foreach (var movimientoOriginal in movimientosOriginales)
            {
                if (movimientoOriginal.TipoMovimiento == "Debe")
                {
                    movimientoOriginal.MovimientoReversionId = movimientoHaberCxC.IdMovimientoContable;
                }
                else if (movimientoOriginal.TipoMovimiento == "Haber")
                {
                    movimientoOriginal.MovimientoReversionId = movimientoDebeIngreso.IdMovimientoContable;
                }

                _context.MovimientosContables.Update(movimientoOriginal);
            }

            AplicarSaldoCuenta(cuentaIngreso, monto, 0);
            AplicarSaldoCuenta(cuentaCxC, 0, monto);

            _context.CuentasContables.Update(cuentaIngreso);
            _context.CuentasContables.Update(cuentaCxC);

            await _context.SaveChangesAsync();
        }

        // ================================================================
        // CAMBIO APLICADO - CUENTA INGRESO PARA CXC MANUAL
        // Ajustá los códigos si en tu catálogo tienen otro número.
        // ================================================================
        private async Task<CuentasContable?> ObtenerCuentaIngresoCxCManual(string? tipoDocumento)
        {
            string codigoIngreso = tipoDocumento switch
            {
                "Cuota" => "CC003",
                "Donación" => "CC007",
                "Convenio" => "CC009",
                "Ajuste" => "CC014",
                _ => "CC007"
            };

            return await _context.CuentasContables
                .FirstOrDefaultAsync(c => c.Codigo == codigoIngreso);
        }

        // ================================================================
        // CAMBIO APLICADO - ACTUALIZACIÓN DE SALDOS CONTABLES
        // Respeta naturaleza de la cuenta.
        // Deudora: aumenta con Debe.
        // Acreedora: aumenta con Haber.
        // ================================================================
        private void AplicarSaldoCuenta(CuentasContable cuenta, decimal debe, decimal haber)
        {
            if (cuenta.Naturaleza == "Deudora")
            {
                cuenta.SaldoActual = cuenta.SaldoActual + debe - haber;
            }
            else
            {
                cuenta.SaldoActual = cuenta.SaldoActual + haber - debe;
            }
        }
        // ================================================================
        // CONSTRUYE LA DESCRIPCIÓN DE LOS FILTROS USADOS EN UN REPORTE
        // Se utiliza en Auditoría para conocer qué información final,  fue consultada o exportada por el usuario.
        private string ConstruirDescripcionFiltrosReporte(
            string? buscar,
            string? tipoDocumento,
            string? estado,
            string? origen,
            int? proyectoId,
            int? anio,
            int? mes,
            DateOnly? fechaInicio,
            DateOnly? fechaFin)
        {
            var filtros = new List<string>();

            if (!string.IsNullOrWhiteSpace(buscar))
            {
                filtros.Add($"Búsqueda={buscar.Trim()}");
            }

            if (!string.IsNullOrWhiteSpace(tipoDocumento))
            {
                filtros.Add($"Tipo={tipoDocumento}");
            }

            if (!string.IsNullOrWhiteSpace(estado))
            {
                filtros.Add($"Estado={estado}");
            }

            if (!string.IsNullOrWhiteSpace(origen))
            {
                filtros.Add($"Origen={origen}");
            }

            if (proyectoId.HasValue && proyectoId.Value > 0)
            {
                filtros.Add($"ProyectoId={proyectoId.Value}");
            }

            if (anio.HasValue)
            {
                filtros.Add($"Año={anio.Value}");
            }

            if (mes.HasValue)
            {
                filtros.Add($"Mes={mes.Value}");
            }

            if (fechaInicio.HasValue)
            {
                filtros.Add($"Desde={fechaInicio.Value:dd/MM/yyyy}");
            }

            if (fechaFin.HasValue)
            {
                filtros.Add($"Hasta={fechaFin.Value:dd/MM/yyyy}");
            }

            return filtros.Any()
                ? "Filtros aplicados: " + string.Join(", ", filtros) + "."
                : "Sin filtros aplicados.";
        }
        // Registra auditoría de acciones realizadas sobre CxC.
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
                    Tabla = "CuentasPorCobrar",
                    RegistroId = registroId,
                    Accion = accion,
                    Descripcion = descripcion,
                    Fecha = DateTime.Now
                };

                _context.Auditoria.Add(auditoria);
                await _context.SaveChangesAsync();
            }
        }

        // ================================================================
        // ACTUALIZA LOS ESTADOS DE TODAS LAS CUENTAS POR COBRAR
        // ================================================================
        // Esta actualización no crea ingresos, pagos ni movimientos contables.
        // Únicamente mantiene correcto el estado administrativo de la CxC
        // según su saldo pendiente, pagos realizados y fecha de vencimiento.
        private async Task ActualizarEstadosCuentasPorCobrar()
        {
            // Se trabaja únicamente con la fecha, sin considerar la hora.
            DateOnly hoy = DateOnly.FromDateTime(DateTime.Today);

            // Obtiene las cuentas activas.
            // Las cuentas anuladas no deben cambiar nuevamente de estado.
            var cuentas = await _context.CuentasPorCobrars
                .Where(c =>
                    c.Activo == true &&
                    c.Estado != "Anulada")
                .ToListAsync();

            bool huboCambios = false;

            foreach (var cuenta in cuentas)
            {
                string nuevoEstado;

                // Si ya no existe saldo pendiente, la cuenta está cancelada.
                if (cuenta.SaldoPendiente <= 0)
                {
                    nuevoEstado = "Cancelada";
                }
                // Si conserva saldo y la fecha ya pasó, está vencida.
                // Vencida tiene prioridad sobre Parcial.
                else if (cuenta.FechaVencimiento < hoy)
                {
                    nuevoEstado = "Vencida";
                }
                // Si recibió algún pago, todavía tiene saldo y no ha vencido,
                // se mantiene como pago parcial.
                else if ((cuenta.MontoPagado ?? 0m) > 0)
                {
                    nuevoEstado = "Parcial";
                }
                // Si no tiene pagos y todavía no vence, continúa pendiente.
                else
                {
                    nuevoEstado = "Pendiente";
                }

                // Solo modifica el registro cuando el estado realmente cambió.
                if (cuenta.Estado != nuevoEstado)
                {
                    cuenta.Estado = nuevoEstado;
                    huboCambios = true;
                }
            }

            // Realiza un único guardado para evitar múltiples accesos a la BD.
            if (huboCambios)
            {
                await _context.SaveChangesAsync();
            }
        }
        // ================================================================
        // ACTUALIZA EL ESTADO DE UNA CUENTA POR COBRAR
        // ================================================================
        // Se utiliza cuando una única cuenta cambia por edición,
        // creación o registro de pagos.
        private void ActualizarEstadoCuenta(CuentasPorCobrar cuenta)
        {
            // Una cuenta anulada conserva siempre su estado.
            if (cuenta.Estado == "Anulada")
            {
                return;
            }

            if (cuenta.SaldoPendiente <= 0)
            {
                cuenta.SaldoPendiente = 0;
                cuenta.Estado = "Cancelada";
            }
            else if (cuenta.FechaVencimiento < DateOnly.FromDateTime(DateTime.Today))
            {
                cuenta.Estado = "Vencida";
            }
            else if ((cuenta.MontoPagado ?? 0m) > 0)
            {
                cuenta.Estado = "Parcial";
            }
            else
            {
                cuenta.Estado = "Pendiente";
            }
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
        private bool CuentasPorCobrarExists(int id)
        {
            return _context.CuentasPorCobrars.Any(e => e.IdCuentaPorCobrar == id);
        }
    }
}