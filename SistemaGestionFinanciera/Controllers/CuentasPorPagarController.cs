using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.EntityFrameworkCore;
using SistemaGestionFinanciera.Data;
using SistemaGestionFinanciera.Models;
using ClosedXML.Excel;
using System.IO;
using System.Globalization;

namespace SistemaGestionFinanciera.Controllers
{
    public class CuentasPorPagarController : BaseController
    {
        private readonly SistemaFinancieroContext _context;

        public CuentasPorPagarController(SistemaFinancieroContext context)
        {
            _context = context;
        }
        // LISTADO DE CUENTAS POR PAGAR
        // GET cuentas por pagar 
        public async Task<IActionResult> Index(
            string? buscar,
            string? tipoDocumento,
            string? estado,
            DateOnly? fechaInicio,
            DateOnly? fechaFin,
            int? categoriaGastoId,
            int? centroCostoId,
            int? proyectoId,
            int? anio,
            int? mes,
            string? origen,
            string? estadoAutorizacion,
            bool proximasVencer = false)
        {
            // Actualiza vencimientos antes de aplicar filtros, alerta y totales.
            await ActualizarEstadosCuentasPorPagar();

            // Detecta si el usuario aplicó algún filtro.
            bool hayFiltros =
                !string.IsNullOrWhiteSpace(buscar) ||
                !string.IsNullOrWhiteSpace(tipoDocumento) ||
                !string.IsNullOrWhiteSpace(estado) ||
                !string.IsNullOrWhiteSpace(origen) ||
                !string.IsNullOrWhiteSpace(estadoAutorizacion) ||
                proximasVencer ||
                fechaInicio.HasValue ||
                fechaFin.HasValue ||
               
                (categoriaGastoId.HasValue && categoriaGastoId.Value > 0) ||
                (centroCostoId.HasValue && centroCostoId.Value > 0) ||
                (proyectoId.HasValue && proyectoId.Value > 0) ||
                anio.HasValue ||
                mes.HasValue;

            // Obtiene las cuentas utilizando la misma consulta
            // compartida por Index, Reporte y Excel.
            var listaCuentas = await ConstruirConsultaReporte(
                buscar,
                tipoDocumento,
                estado,
                fechaInicio,
                fechaFin,
                categoriaGastoId,
                centroCostoId,
                proyectoId,
                anio,
                mes,
                origen,
                estadoAutorizacion,
                proximasVencer);

            // Conserva los filtros seleccionados.
            ViewBag.Buscar = buscar;
            ViewBag.TipoDocumento = tipoDocumento;
            ViewBag.Estado = estado;
            ViewBag.Origen = origen;

            ViewBag.EstadoAutorizacion = estadoAutorizacion;
            ViewBag.ProximasVencerFiltro = proximasVencer;

            ViewBag.CategoriaGastoSeleccionada = categoriaGastoId;
            ViewBag.CentroCostoSeleccionado = centroCostoId;
            ViewBag.ProyectoSeleccionado = proyectoId;

            ViewBag.Anio = anio;
            ViewBag.Mes = mes;

            ViewBag.FechaInicio = fechaInicio?.ToString("yyyy-MM-dd");
            ViewBag.FechaFin = fechaFin?.ToString("yyyy-MM-dd");

            // ================================================================
            // ALERTAS DE CUENTAS VENCIDAS
            // ================================================================

            var cuentasVencidas =
                await _context.CuentasPorPagars
                    .AsNoTracking()
                    .Where(c =>
                        c.Activo == true &&
                        c.Estado == "Vencida" &&
                        c.Estado != "Anulada" &&
                        c.SaldoPendiente > 0)
                    .ToListAsync();

            // Total de cuentas vencidas.
            ViewBag.CuentasVencidas =
                cuentasVencidas.Count;

            // Saldo total de todas las cuentas vencidas.
            ViewBag.SaldoTotalVencido =
                cuentasVencidas.Sum(c =>
                    c.SaldoPendiente);

            // Cuentas vencidas que todavía no han sido autorizadas.
            var vencidasPendientesAutorizacion =
                cuentasVencidas
                    .Where(c =>
                        c.EstadoAutorizacion == "Pendiente")
                    .ToList();

            ViewBag.VencidasPendientesAutorizacion =
                vencidasPendientesAutorizacion.Count;

            ViewBag.SaldoVencidoPendienteAutorizacion =
                vencidasPendientesAutorizacion.Sum(c =>
                    c.SaldoPendiente);

            var cuentasParaTotales = listaCuentas;
            // ================================================================
            // ALERTA DE CUENTAS PRÓXIMAS A VENCER
            // Se consideran cuentas con saldo pendiente que vencen desde hoy
            // hasta los siguientes siete días.
            // ================================================================

            var hoyProximas =
                DateOnly.FromDateTime(DateTime.Today);

            var fechaLimiteProximas =
                hoyProximas.AddDays(7);

            var cuentasProximasVencer =
                await _context.CuentasPorPagars
                    .AsNoTracking()
                    .Where(c =>
                        c.Activo == true &&
                        c.Estado != "Anulada" &&
                        c.Estado != "Cancelada" &&
                        c.Estado != "Vencida" &&
                        c.SaldoPendiente > 0 &&
                        c.FechaVencimiento >= hoyProximas &&
                        c.FechaVencimiento <= fechaLimiteProximas)
                    .ToListAsync();

            ViewBag.ProximasVencer =
                cuentasProximasVencer.Count;

            ViewBag.SaldoProximasVencer =
                cuentasProximasVencer.Sum(c =>
                    c.SaldoPendiente);
            // Sin filtros, las cards muestran solamente el mes actual.
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

            // Las anuladas permanecen visibles, pero no se suman.
            var cuentasValidas = cuentasParaTotales
                .Where(c =>
                    c.Activo == true &&
                    c.Estado != "Anulada")
                .ToList();

            ViewBag.TotalObligaciones = cuentasValidas
                .Sum(c =>
                    (c.SubTotal ?? 0m) +
                    (c.Impuesto ?? 0m));

            ViewBag.TotalPagado = cuentasValidas
                .Sum(c => c.MontoPagado ?? 0m);

            ViewBag.TotalSaldoPendiente = cuentasValidas
                .Sum(c => c.SaldoPendiente);

            // Categorías activas.
            ViewBag.CategoriaGastoId = new SelectList(
                await _context.CategoriasGastos
                    .Where(c => c.Activo == true)
                    .OrderBy(c => c.Nombre)
                    .ToListAsync(),
                "IdCategoriaGasto",
                "Nombre",
                categoriaGastoId);

            // Centros de costo activos.
            ViewBag.CentroCostoId = new SelectList(
                await _context.CentrosCostos
                    .Where(c => c.Activo == true)
                    .OrderBy(c => c.Nombre)
                    .ToListAsync(),
                "IdCentroCosto",
                "Nombre",
                centroCostoId);

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
            ViewBag.Anios = await _context.CuentasPorPagars
                .Select(c => c.FechaEmision.Year)
                .Distinct()
                .OrderByDescending(a => a)
                .ToListAsync();

            return View(listaCuentas);
        }
        // ===========================================================
        // CONSULTA COMPARTIDA PARA INDEX, REPORTE Y EXCEL
        // ===========================================================
        private async Task<List<CuentasPorPagar>> ConstruirConsultaReporte(
            string? buscar,
            string? tipoDocumento,
            string? estado,
            DateOnly? fechaInicio,
            DateOnly? fechaFin,
            int? categoriaGastoId,
            int? centroCostoId,
            int? proyectoId,
            int? anio,
            int? mes,
            string? origen,
            string? estadoAutorizacion = null,
            bool proximasVencer = false)
        {
            var cuentas = _context.CuentasPorPagars
                .Include(c => c.Proveedor)
                .Include(c => c.Proyecto)
                .Include(c => c.PresupuestoMensualRegistro)
                .Include(c => c.CategoriaGasto)
                .Include(c => c.CentroCosto)
                .AsQueryable();

            // Búsqueda por documento, factura o proveedor.
            if (!string.IsNullOrWhiteSpace(buscar))
            {
                buscar = buscar.Trim();

                cuentas = cuentas.Where(c =>
                    (c.NumeroDocumento != null &&
                     c.NumeroDocumento.Contains(buscar)) ||

                    (c.NumeroFactura != null &&
                     c.NumeroFactura.Contains(buscar)) ||

                    (c.Concepto != null &&
                     c.Concepto.Contains(buscar)) ||

                    (c.Proveedor != null &&
                     c.Proveedor.Nombre.Contains(buscar)));
            }

            // Tipo de documento.
            if (!string.IsNullOrWhiteSpace(tipoDocumento))
            {
                cuentas = cuentas.Where(c =>
                    c.TipoDocumento == tipoDocumento);
            }

            // Estado.
            if (!string.IsNullOrWhiteSpace(estado))
            {
                cuentas = cuentas.Where(c =>
                    c.Estado == estado);
            }

            // Origen.
            if (!string.IsNullOrWhiteSpace(origen))
            {
                cuentas = cuentas.Where(c =>
                    c.Origen == origen);
            }
            // Estado de autorización.
            // Solo aplica a CxP manuales que continúan vigentes.
            if (!string.IsNullOrWhiteSpace(estadoAutorizacion))
            {
                cuentas = cuentas.Where(c =>
                    c.Origen == "Manual" &&
                    c.Estado != "Anulada" &&
                    c.EstadoAutorizacion == estadoAutorizacion);
            }
            // Categoría de gasto.
            if (categoriaGastoId.HasValue &&
                categoriaGastoId.Value > 0)
            {
                cuentas = cuentas.Where(c =>
                    c.CategoriaGastoId == categoriaGastoId.Value);
            }

            // Centro de costo.
            if (centroCostoId.HasValue &&
                centroCostoId.Value > 0)
            {
                cuentas = cuentas.Where(c =>
                    c.CentroCostoId == centroCostoId.Value);
            }

            // Proyecto.
            if (proyectoId.HasValue &&
                proyectoId.Value > 0)
            {
                cuentas = cuentas.Where(c =>
                    c.ProyectoId == proyectoId.Value);
            }

            // Año.
            if (anio.HasValue)
            {
                cuentas = cuentas.Where(c =>
                    c.FechaEmision.Year == anio.Value);
            }

            // Mes.
            if (mes.HasValue)
            {
                cuentas = cuentas.Where(c =>
                    c.FechaEmision.Month == mes.Value);
            }

            // Fecha inicial.
            if (fechaInicio.HasValue)
            {
                cuentas = cuentas.Where(c =>
                    c.FechaEmision >= fechaInicio.Value);
            }

            // Fecha final.
            if (fechaFin.HasValue)
            {
                cuentas = cuentas.Where(c =>
                    c.FechaEmision <= fechaFin.Value);
            }
            // Filtro especial: cuentas próximas a vencer durante
            // los siguientes siete días.
            if (proximasVencer)
            {
                var hoyProximas =
                    DateOnly.FromDateTime(DateTime.Today);

                var fechaLimiteProximas =
                    hoyProximas.AddDays(7);

                cuentas = cuentas.Where(c =>
                    c.Activo == true &&
                    c.Estado != "Anulada" &&
                    c.Estado != "Cancelada" &&
                    c.Estado != "Vencida" &&
                    c.SaldoPendiente > 0 &&
                    c.FechaVencimiento >= hoyProximas &&
                    c.FechaVencimiento <= fechaLimiteProximas);
            }
            // Registros más recientes primero.
            return await cuentas
                .OrderByDescending(c => c.FechaEmision)
                .ThenByDescending(c => c.IdCuentaPorPagar)
                .ToListAsync();
        }
        // ===========================================================
        // VISTA PREVIA DEL REPORTE DE CUENTAS POR PAGAR
        // ===========================================================
        public async Task<IActionResult> Reporte(
            string? buscar,
            string? tipoDocumento,
            string? estado,
            DateOnly? fechaInicio,
            DateOnly? fechaFin,
            int? categoriaGastoId,
            int? centroCostoId,
            int? proyectoId,
            int? anio,
            int? mes,
            string? origen,
            string? estadoAutorizacion,
            bool proximasVencer = false)
        {
            // Actualiza los estados antes de generar la vista previa.
            await ActualizarEstadosCuentasPorPagar();

            var cuentas = await ConstruirConsultaReporte(
                buscar,
                tipoDocumento,
                estado,
                fechaInicio,
                fechaFin,
                categoriaGastoId,
                centroCostoId,
                proyectoId,
                anio,
                mes,
                origen,
                estadoAutorizacion,
                proximasVencer);

            // Las anuladas aparecen en el reporte,
            // pero no participan en los totales.
            var cuentasValidas = cuentas
                .Where(c =>
                    c.Activo == true &&
                    c.Estado != "Anulada")
                .ToList();

            ViewBag.TotalRegistros = cuentas.Count;

            ViewBag.TotalObligaciones = cuentasValidas
                .Sum(c =>
                    (c.SubTotal ?? 0m) +
                    (c.Impuesto ?? 0m));

            ViewBag.TotalPagado = cuentasValidas
                .Sum(c => c.MontoPagado ?? 0m);

            ViewBag.TotalSaldoPendiente = cuentasValidas
                .Sum(c => c.SaldoPendiente);

            // Resumen por estado.
            ViewBag.CantidadPendientes = cuentasValidas
                .Count(c => c.Estado == "Pendiente");

            ViewBag.CantidadParciales = cuentasValidas
                .Count(c => c.Estado == "Parcial");

            ViewBag.CantidadCanceladas = cuentasValidas
                .Count(c => c.Estado == "Cancelada");

            ViewBag.CantidadVencidas = cuentasValidas
                .Count(c => c.Estado == "Vencida");

            ViewBag.CantidadAnuladas = cuentas
                .Count(c => c.Estado == "Anulada");

            ViewBag.FechaGeneracion = DateTime.Now;

            // Conserva los filtros.
            ViewBag.Buscar = buscar;
            ViewBag.TipoDocumento = tipoDocumento;
            ViewBag.Estado = estado;
            ViewBag.EstadoAutorizacion = estadoAutorizacion;
            ViewBag.ProximasVencerFiltro = proximasVencer;
            ViewBag.Origen = origen;

            ViewBag.CategoriaGastoId = categoriaGastoId;
            ViewBag.CentroCostoId = centroCostoId;
            ViewBag.ProyectoId = proyectoId;

            ViewBag.Anio = anio;
            ViewBag.Mes = mes;

            ViewBag.FechaInicio = fechaInicio?.ToString("yyyy-MM-dd");
            ViewBag.FechaFin = fechaFin?.ToString("yyyy-MM-dd");

            // Nombres para mostrar en lugar de los ID.
            ViewBag.NombreCategoria = categoriaGastoId.HasValue
                ? await _context.CategoriasGastos
                    .Where(c =>
                        c.IdCategoriaGasto == categoriaGastoId.Value)
                    .Select(c => c.Nombre)
                    .FirstOrDefaultAsync()
                : null;

            ViewBag.NombreCentroCosto = centroCostoId.HasValue
                ? await _context.CentrosCostos
                    .Where(c =>
                        c.IdCentroCosto == centroCostoId.Value)
                    .Select(c => c.Nombre)
                    .FirstOrDefaultAsync()
                : null;

            ViewBag.NombreProyecto = proyectoId.HasValue
                ? await _context.Proyectos
                    .Where(p =>
                        p.IdProyecto == proyectoId.Value)
                    .Select(p => p.Nombre)
                    .FirstOrDefaultAsync()
                : null;
            // Registra la vista previa del reporte financiero.
            // RegistroId = 0 porque corresponde a un conjunto de CxP.
            await RegistrarAuditoria(
                "Vista previa PDF",
                0,
                "Se generó la vista previa del reporte de cuentas por pagar. " +
                ConstruirDescripcionFiltrosReporte(
                    buscar,
                    tipoDocumento,
                    estado,
                    fechaInicio,
                    fechaFin,
                    categoriaGastoId,
                    centroCostoId,
                    proyectoId,
                    anio,
                    mes,
                    origen)
            );

            return View(cuentas);
           
        }
        // ===========================================================
        // EXPORTAR REPORTE DE CUENTAS POR PAGAR A EXCEL
        // ===========================================================
        public async Task<IActionResult> ExportarExcel(
            string? buscar,
            string? tipoDocumento,
            string? estado,
            DateOnly? fechaInicio,
            DateOnly? fechaFin,
            int? categoriaGastoId,
            int? centroCostoId,
            int? proyectoId,
            int? anio,
            int? mes,
            string? origen,
            string? estadoAutorizacion, 
            bool proximasVencer = false)
        {
            // Actualiza los estados antes de construir el archivo Excel.
            await ActualizarEstadosCuentasPorPagar();

            var cuentas = await ConstruirConsultaReporte(
                buscar,
                tipoDocumento,
                estado,
                fechaInicio,
                fechaFin,
                categoriaGastoId,
                centroCostoId,
                proyectoId,
                anio,
                mes,
                origen,
                estadoAutorizacion,
                proximasVencer);

            // Orden definitivo:
            // fecha más reciente y luego ID más reciente.
            cuentas = cuentas
                .OrderByDescending(c => c.FechaEmision)
                .ThenByDescending(c => c.IdCuentaPorPagar)
                .ToList();

            using var workbook = new XLWorkbook();

            var hoja = workbook.Worksheets.Add("Cuentas por Pagar");

            const int totalColumnas = 15;

            // ===========================================================
            // ENCABEZADO PRINCIPAL
            // ===========================================================
            hoja.Cell("A1").Value =
                "Unión Cantonal de Asociaciones de Desarrollo de Santa Ana";

            hoja.Range(1, 1, 1, totalColumnas).Merge();

            hoja.Cell("A2").Value =
                "Reporte de Cuentas por Pagar";

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
            if (!string.IsNullOrWhiteSpace(estadoAutorizacion))
            {
                filtros.Add($"Autorización: {estadoAutorizacion}");
            }
            if (proximasVencer)
            {
                filtros.Add(
                    "Vencimiento: Próximos 7 días");
            }
            if (categoriaGastoId.HasValue &&
                categoriaGastoId.Value > 0)
            {
                string? nombreCategoria =
                    await _context.CategoriasGastos
                        .Where(c =>
                            c.IdCategoriaGasto ==
                            categoriaGastoId.Value)
                        .Select(c => c.Nombre)
                        .FirstOrDefaultAsync();

                filtros.Add(
                    $"Categoría: " +
                    $"{nombreCategoria ?? categoriaGastoId.Value.ToString()}");
            }

            if (centroCostoId.HasValue &&
                centroCostoId.Value > 0)
            {
                string? nombreCentroCosto =
                    await _context.CentrosCostos
                        .Where(c =>
                            c.IdCentroCosto ==
                            centroCostoId.Value)
                        .Select(c => c.Nombre)
                        .FirstOrDefaultAsync();

                filtros.Add(
                    $"Centro de costo: " +
                    $"{nombreCentroCosto ?? centroCostoId.Value.ToString()}");
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
                        new System.Globalization.CultureInfo("es-CR"));

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

            hoja.Cell("A5").Value = "Filtros aplicados:";
            hoja.Cell("A5").Style.Font.Bold = true;

            hoja.Cell("B5").Value = filtros.Any()
                ? string.Join(" | ", filtros)
                : "Sin filtros. Se muestran todas las cuentas por pagar.";

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
        "Proveedor",
        "Categoría",
        "Centro de costo",
        "Proyecto",
        "Concepto",
        "Fecha de emisión",
        "Fecha de vencimiento",
        "Días de crédito",
        "Monto total",
        "Monto pagado",
        "Saldo pendiente",
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
                decimal montoTotal =
                    (cuenta.SubTotal ?? 0m) +
                    (cuenta.Impuesto ?? 0m);

                hoja.Cell(fila, 1).Value =
                    cuenta.NumeroDocumento ?? "Sin número";

                hoja.Cell(fila, 2).Value =
                    cuenta.TipoDocumento ?? "Sin tipo";

                hoja.Cell(fila, 3).Value =
                    cuenta.Origen ?? "Sin origen";

                hoja.Cell(fila, 4).Value =
                    cuenta.Proveedor?.Nombre ?? "Sin proveedor";

                hoja.Cell(fila, 5).Value =
                    cuenta.CategoriaGasto?.Nombre ?? "Sin categoría";

                hoja.Cell(fila, 6).Value =
                    cuenta.CentroCosto?.Nombre ?? "Sin centro de costo";

                hoja.Cell(fila, 7).Value =
                    cuenta.Proyecto?.Nombre ?? "Sin proyecto";

                hoja.Cell(fila, 8).Value =
                    cuenta.Concepto ?? "Sin concepto";

                hoja.Cell(fila, 9).Value =
                    cuenta.FechaEmision
                        .ToDateTime(TimeOnly.MinValue);

                hoja.Cell(fila, 10).Value =
                    cuenta.FechaVencimiento
                        .ToDateTime(TimeOnly.MinValue);

                hoja.Cell(fila, 11).Value =
                    cuenta.DiasCredito ?? 0;

                hoja.Cell(fila, 12).Value =
                    montoTotal;

                hoja.Cell(fila, 13).Value =
                    cuenta.MontoPagado ?? 0m;

                hoja.Cell(fila, 14).Value =
                    cuenta.SaldoPendiente;

                hoja.Cell(fila, 15).Value =
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

                // No obliga a todas las filas a crecer.
                rangoDatos.Style.Alignment.WrapText = false;

                // Fechas.
                hoja.Range(
                        filaInicioDatos,
                        9,
                        filaFinDatos,
                        10)
                    .Style.DateFormat.Format =
                        "dd/MM/yyyy";

                // Días de crédito.
                hoja.Range(
                        filaInicioDatos,
                        11,
                        filaFinDatos,
                        11)
                    .Style.NumberFormat.Format =
                        "0";

                // Montos.
                hoja.Range(
                        filaInicioDatos,
                        12,
                        filaFinDatos,
                        14)
                    .Style.NumberFormat.Format =
                        "₡ #,##0.00";

                // Centrado de fechas, días y estado.
                hoja.Range(
                        filaInicioDatos,
                        9,
                        filaFinDatos,
                        11)
                    .Style.Alignment.Horizontal =
                        XLAlignmentHorizontalValues.Center;

                hoja.Range(
                        filaInicioDatos,
                        15,
                        filaFinDatos,
                        15)
                    .Style.Alignment.Horizontal =
                        XLAlignmentHorizontalValues.Center;

                // Montos alineados a la derecha.
                hoja.Range(
                        filaInicioDatos,
                        12,
                        filaFinDatos,
                        14)
                    .Style.Alignment.Horizontal =
                        XLAlignmentHorizontalValues.Right;

                // Autofiltro en toda la tabla.
                hoja.Range(
                        filaEncabezado,
                        1,
                        filaFinDatos,
                        totalColumnas)
                    .SetAutoFilter();

                // Colores por estado.
                for (int indice = 0;
                     indice < cuentas.Count;
                     indice++)
                {
                    var cuenta = cuentas[indice];

                    int filaCuenta =
                        filaInicioDatos + indice;

                    if (cuenta.Estado == "Anulada")
                    {
                        // Rojo tenue, igual al patrón de CxC.
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
                        // Amarillo tenue.
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
                    "No se encontraron cuentas por pagar " +
                    "con los filtros aplicados.";

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

            decimal totalObligaciones =
                cuentasValidas.Sum(c =>
                    (c.SubTotal ?? 0m) +
                    (c.Impuesto ?? 0m));

            decimal totalPagado =
                cuentasValidas.Sum(c =>
                    c.MontoPagado ?? 0m);

            decimal totalSaldoPendiente =
                cuentasValidas.Sum(c =>
                    c.SaldoPendiente);

            int filaTotales = fila + 1;

            hoja.Cell(filaTotales, 10).Value =
                "TOTALES:";

            hoja.Range(
                    filaTotales,
                    10,
                    filaTotales,
                    11)
                .Merge();

            hoja.Cell(filaTotales, 12).Value =
                totalObligaciones;

            hoja.Cell(filaTotales, 13).Value =
                totalPagado;

            hoja.Cell(filaTotales, 14).Value =
                totalSaldoPendiente;

            hoja.Cell(filaTotales, 15).Value =
                cuentas.Count == 1
                    ? "1 registro"
                    : $"{cuentas.Count} registros";

            var rangoTotales = hoja.Range(
                filaTotales,
                10,
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
                    12,
                    filaTotales,
                    14)
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
            hoja.Column(3).Width = 14;
            hoja.Column(4).Width = 28;
            hoja.Column(5).Width = 24;
            hoja.Column(6).Width = 25;
            hoja.Column(7).Width = 28;
            hoja.Column(8).Width = 34;
            hoja.Column(9).Width = 17;
            hoja.Column(10).Width = 19;
            hoja.Column(11).Width = 14;
            hoja.Column(12).Width = 18;
            hoja.Column(13).Width = 18;
            hoja.Column(14).Width = 18;
            hoja.Column(15).Width = 15;

            // Alto uniforme para que las filas no se deformen.
            if (cuentas.Any())
            {
                for (int numeroFila = filaInicioDatos;
                     numeroFila <= filaFinDatos;
                     numeroFila++)
                {
                    hoja.Row(numeroFila).Height = 20;
                }
            }

            // Mantiene visible el encabezado.
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

            // Fuente uniforme.
            hoja.Range(
                    1,
                    1,
                    filaResumen + 2,
                    totalColumnas)
                .Style.Font.FontName = "Arial";

            // ===========================================================
            // GENERACIÓN DEL ARCHIVO
            // ===========================================================
            using var stream = new MemoryStream();

            workbook.SaveAs(stream);

            string nombreArchivo =
                $"Reporte_Cuentas_por_Pagar_" +
                $"{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";
            // Registra la exportación únicamente después de haber construido
            // correctamente el archivo Excel.
            await RegistrarAuditoria(
                "Exportar Excel",
                0,
                "Se exportó el reporte de cuentas por pagar a Excel. " +
                ConstruirDescripcionFiltrosReporte(
                    buscar,
                    tipoDocumento,
                    estado,
                    fechaInicio,
                    fechaFin,
                    categoriaGastoId,
                    centroCostoId,
                    proyectoId,
                    anio,
                    mes,
                    origen)
            );
            return File(
                stream.ToArray(),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                nombreArchivo);
        }

        // GET: CuentasPorPagar/Details/5
        public async Task<IActionResult> Details(int? id, string? origen, int? pagoId)
        {
            ViewBag.PagoId = pagoId;

            if (id == null)
            {
                return NotFound();
            }

            var cuentasPorPagar = await _context.CuentasPorPagars
                .Include(c => c.Proveedor)
                .Include(c => c.Proyecto)
                .Include(c => c.PresupuestoMensualRegistro)
                .Include(c => c.CategoriaGasto)
                .Include(c => c.CentroCosto)
                .Include(c => c.PagosCuentaPorPagars)
                .FirstOrDefaultAsync(m => m.IdCuentaPorPagar == id);

            if (cuentasPorPagar == null)
            {
                return NotFound();
            }

            ViewBag.Origen = origen;

            return View(cuentasPorPagar);
        }

        // GET: CuentasPorPagar/Crear
        public IActionResult Create()
        {
            CargarCombos();

            CargarPeriodosPresupuestarios();

            ViewBag.FechaEmision =
                DateOnly.FromDateTime(DateTime.Today);

            return View();
        }

        // POST: CuentasPorPagar/Create
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Create(
    [Bind("ProveedorId,PresupuestoMensualId,CategoriaGastoId,CentroCostoId,TipoDocumento,Concepto,NumeroFactura,MontoOriginal,Descuento,TipoImpuesto,DiasCredito,Observacion")]CuentasPorPagar cuentasPorPagar)
        {
                // Obtiene la fecha actual del sistema.
                var fechaHoy = DateOnly.FromDateTime(DateTime.Today);
            // ================================================================
            // VALIDAR PROYECTO / PERÍODO PRESUPUESTARIO
            // ================================================================

            PresupuestosMensuale? presupuestoSeleccionado = null;

            if (!cuentasPorPagar.PresupuestoMensualId.HasValue)
            {
                ModelState.AddModelError(
                    "PresupuestoMensualId",
                    "Debe seleccionar un proyecto y período presupuestario.");
            }
            else
            {
                var fechaMinima = fechaHoy.AddMonths(-3);

                int anioMinimo = fechaMinima.Year;
                int mesMinimo = fechaMinima.Month;

                presupuestoSeleccionado =
                    await _context.PresupuestosMensuales
                        .Include(pm => pm.Proyecto)
                        .FirstOrDefaultAsync(pm =>
                            pm.IdPresupuestoMensual ==
                                cuentasPorPagar.PresupuestoMensualId.Value &&
                            pm.EstadoAprobacion == "Aprobado" &&
                            pm.ProyectoId != null &&
                            pm.Proyecto != null &&
                            pm.Proyecto.Activo &&
            // NUEVO
            !_context.CierresContables.Any(c =>
                c.Anio == pm.Anio &&
                c.Mes == pm.Mes &&
                c.Estado == "Cerrado") &&
                            (
                                pm.Anio > anioMinimo ||
                                (pm.Anio == anioMinimo &&
                                 pm.Mes >= mesMinimo)
                            ));

                if (presupuestoSeleccionado == null)
                {
                    ModelState.AddModelError(
                        "PresupuestoMensualId",
                        "El período presupuestario seleccionado no está disponible.");
                }
                else
                {
                    // El proyecto se obtiene del presupuesto mensual seleccionado.
                    cuentasPorPagar.ProyectoId =
                        presupuestoSeleccionado.ProyectoId;
                }
            }

            // Valida que exista un proveedor seleccionado.
            if (cuentasPorPagar.ProveedorId <= 0)
                {
                    ModelState.AddModelError("ProveedorId", "Debe seleccionar el proveedor.");
                }


                // Valida que exista una categoría de gasto.
                if (cuentasPorPagar.CategoriaGastoId == null || cuentasPorPagar.CategoriaGastoId <= 0)
                {
                    ModelState.AddModelError("CategoriaGastoId", "Debe seleccionar la categoría de gasto.");
                }
                // valida centro costo 
            if (cuentasPorPagar.CentroCostoId == null || cuentasPorPagar.CentroCostoId <= 0)
            {
                ModelState.AddModelError("CentroCostoId", "Debe seleccionar el centro de costo.");
            }
            // Valida los días de crédito.
            if (cuentasPorPagar.DiasCredito == null || cuentasPorPagar.DiasCredito <= 0)
                {
                    ModelState.AddModelError("DiasCredito", "Debe seleccionar los días de crédito.");
                }

                // El descuento nunca puede ser mayor que el monto original.
                if ((cuentasPorPagar.Descuento ?? 0) > cuentasPorPagar.MontoOriginal)
                {
                    ModelState.AddModelError("Descuento", "El descuento no puede ser mayor al monto original.");
                }

                // Estos datos los calcula el sistema automáticamente.
                ModelState.Remove("NumeroDocumento");
                ModelState.Remove("FechaEmision");
                ModelState.Remove("FechaVencimiento");
                ModelState.Remove("SubTotal");
                ModelState.Remove("Impuesto");
                ModelState.Remove("MontoPagado");
                ModelState.Remove("AnticipoAplicado");
                ModelState.Remove("SaldoPendiente");
                ModelState.Remove("Estado");
                ModelState.Remove("Activo");
                ModelState.Remove("Proveedor");
                ModelState.Remove("Proyecto");
            ModelState.Remove("PresupuestoMensualRegistro");
            ModelState.Remove("CategoriaGasto");
                ModelState.Remove("CentroCosto");
                ModelState.Remove("PagosCuentaPorPagars");
                ModelState.Remove("MotivoAnulacion");

                if (ModelState.IsValid)
                {
                    // Asigna automáticamente la fecha de emisión.
                    cuentasPorPagar.FechaEmision = fechaHoy;

                    // Calcula la fecha de vencimiento según los días de crédito.
                    cuentasPorPagar.FechaVencimiento = fechaHoy.AddDays(cuentasPorPagar.DiasCredito ?? 0);

                    // Genera el número consecutivo del documento.
                    cuentasPorPagar.NumeroDocumento = GenerarNumeroDocumentoCxP(cuentasPorPagar.TipoDocumento);

                    // Inicializa los montos de pago.
                    cuentasPorPagar.MontoPagado = 0;
                    cuentasPorPagar.AnticipoAplicado = 0;

                    // Calcula el subtotal descontando el descuento aplicado.
                    cuentasPorPagar.SubTotal =
                        cuentasPorPagar.MontoOriginal - (cuentasPorPagar.Descuento ?? 0);

                    // Obtiene el porcentaje del impuesto seleccionado.
                    decimal porcentajeImpuesto = cuentasPorPagar.TipoImpuesto switch
                    {
                        "IVA 13%" => 0.13m,
                        "IVA 4%" => 0.04m,
                        "IVA 2%" => 0.02m,
                        "IVA 1%" => 0.01m,
                        _ => 0m
                    };

                    // Calcula automáticamente el impuesto.
                    cuentasPorPagar.Impuesto =
                        (cuentasPorPagar.SubTotal ?? 0) * porcentajeImpuesto;

                    // Calcula el saldo pendiente de pago.
                    cuentasPorPagar.SaldoPendiente =
                        (cuentasPorPagar.SubTotal ?? 0)
                        + (cuentasPorPagar.Impuesto ?? 0)
                        - (cuentasPorPagar.MontoPagado ?? 0)
                        - (cuentasPorPagar.AnticipoAplicado ?? 0);

                    // Estado inicial de la cuenta por pagar.
                    cuentasPorPagar.Estado = "Pendiente";
                    cuentasPorPagar.Activo = true;

                // Las cuentas creadas directamente desde CxP nacen como origen manual.
                // Estas requieren aprobación del gasto antes de permitir pagos.
                cuentasPorPagar.Origen = "Manual";

                // Las CxP manuales deben esperar la aprobación del gasto automático.
                cuentasPorPagar.EstadoAutorizacion = "Pendiente";

                // Guarda el registro en la base de datos.
                _context.Add(cuentasPorPagar);
                    await _context.SaveChangesAsync(); 

                // Integración: cuando la cuenta por pagar se registra manualmente,
                // también se crea un gasto pendiente de autorización.
                bool gastoYaExiste = await _context.Gastos
                    .AnyAsync(g => g.NumeroFactura == cuentasPorPagar.NumeroDocumento);

                if (!gastoYaExiste)
                {
                    var gastoAutomatico = new Gasto
                    {
                        Fecha = cuentasPorPagar.FechaEmision,
                        CategoriaGastoId = cuentasPorPagar.CategoriaGastoId.Value,
                        ProyectoId = cuentasPorPagar.ProyectoId,
                        PresupuestoMensualId = cuentasPorPagar.PresupuestoMensualId,
                        ProveedorId = cuentasPorPagar.ProveedorId,
                        CentroCostoId = cuentasPorPagar.CentroCostoId,
                        NumeroFactura = cuentasPorPagar.NumeroDocumento,
                        Concepto = "Gasto generado automáticamente desde la cuenta por pagar " + cuentasPorPagar.NumeroDocumento,
                        Subtotal = cuentasPorPagar.SubTotal,
                        Impuesto = cuentasPorPagar.Impuesto,
                        Descuento = cuentasPorPagar.Descuento ?? 0,
                        TipoImpuesto = cuentasPorPagar.TipoImpuesto,
                        MontoTotal = (cuentasPorPagar.SubTotal ?? 0) + (cuentasPorPagar.Impuesto ?? 0),
                        Observacion = cuentasPorPagar.Observacion,
                        Estado = "Pendiente de autorización",
                        Activo = true
                    };

                    _context.Gastos.Add(gastoAutomatico);
                    await _context.SaveChangesAsync();
                }
                await RegistrarAuditoria(
    "Crear",
    cuentasPorPagar.IdCuentaPorPagar,
    "Se registró la cuenta por pagar " + cuentasPorPagar.NumeroDocumento +
    " por un monto de ₡" + cuentasPorPagar.SaldoPendiente
);
                TempData["MensajeExito"] = "La cuenta por pagar fue registrada correctamente.";

                    return RedirectToAction(nameof(Index));
                }

                // Si ocurre un error vuelve a cargar los combos.
                CargarCombos(
                    cuentasPorPagar.ProveedorId,
                    cuentasPorPagar.ProyectoId,
                    cuentasPorPagar.TipoDocumento,
                    cuentasPorPagar.DiasCredito,
                    cuentasPorPagar.CategoriaGastoId,
                    cuentasPorPagar.CentroCostoId,
                    cuentasPorPagar.TipoImpuesto);

            CargarPeriodosPresupuestarios(
    cuentasPorPagar.PresupuestoMensualId);
            // Conserva la fecha mostrada en el formulario.
            ViewBag.FechaEmision = fechaHoy;

                return View(cuentasPorPagar);
            }


        /* GET: CuentasPorPagar/Edit/5
        public async Task<IActionResult> Edit(int? id)
    {
        if (id == null)
        {
            return NotFound();
        }

        var cuentasPorPagar = await _context.CuentasPorPagars.FindAsync(id);

        if (cuentasPorPagar == null)
        {
            return NotFound();
        }

        if (cuentasPorPagar.Estado == "Anulada")
        {
            TempData["MensajeError"] = "No se puede modificar una cuenta por pagar anulada.";
            return RedirectToAction(nameof(Index));
        }

        CargarCombosEdit(
 cuentasPorPagar.ProveedorId,
 cuentasPorPagar.ProyectoId,
 cuentasPorPagar.TipoDocumento,
 cuentasPorPagar.DiasCredito,
 cuentasPorPagar.CategoriaGastoId,
 cuentasPorPagar.CentroCostoId,
 cuentasPorPagar.TipoImpuesto);

        return View(cuentasPorPagar);
    }*/
        // ================================================================
        // GET: CuentasPorPagar/Edit/5
        // ================================================================
        public async Task<IActionResult> Edit(int? id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var cuenta = await _context.CuentasPorPagars
                .FirstOrDefaultAsync(c =>
                    c.IdCuentaPorPagar == id.Value);

            if (cuenta == null)
            {
                return NotFound();
            }
            // ============================================================
            // BLOQUEO POR CIERRE CONTABLE
            // ============================================================
            bool periodoCerrado =
                await PeriodoContableCerradoAsync(cuenta.FechaEmision);

            if (periodoCerrado)
            {
                TempData["MensajeError"] =
                    $"No se puede modificar la cuenta por pagar porque el período " +
                    $"{cuenta.FechaEmision:MM/yyyy} se encuentra cerrado. " +
                    $"Para realizar correcciones, primero debe reabrirse el período.";

                return RedirectToAction(
                    nameof(Index),
                    new
                    {
                        anio = cuenta.FechaEmision.Year,
                        mes = cuenta.FechaEmision.Month
                    });
            }
            
            // Las CxP provenientes de factura se gestionan desde Facturación.
            if (string.Equals(
                cuenta.Origen,
                "Factura",
                StringComparison.OrdinalIgnoreCase))
            {
                TempData["MensajeError"] =
                    "Esta cuenta por pagar fue generada desde una factura. " +
                    "Debe modificarse desde el módulo de Facturación.";

                return RedirectToAction(nameof(Index));
            }
// ============================================================
            // CXP MANUAL APROBADA: SOLO CONSULTA Y PAGOS
            // ============================================================
            bool gastoAprobado = await _context.Gastos
                .AnyAsync(g =>
                    g.NumeroFactura == cuenta.NumeroDocumento &&
                    g.Estado == "Aprobado");

            if (gastoAprobado ||
                cuenta.EstadoAutorizacion == "Aprobado")
            {
                TempData["MensajeError"] =
                    "Esta cuenta por pagar ya fue aprobada. " +
                    "Solo puede consultar sus detalles y registrar pagos.";

                return RedirectToAction(nameof(Index));
            }
            // Las cuentas anuladas o canceladas son únicamente de consulta.
            if (cuenta.Estado == "Anulada" ||
                cuenta.Estado == "Cancelada")
            {
                TempData["MensajeError"] =
                    $"No se puede modificar una cuenta por pagar {cuenta.Estado.ToLower()}. " +
                    "El registro se conserva únicamente para consulta.";

                return RedirectToAction(nameof(Index));
            }

            CargarCombosEdit(
                cuenta.ProveedorId,
                cuenta.ProyectoId,
                cuenta.TipoDocumento,
                cuenta.DiasCredito,
                cuenta.CategoriaGastoId,
                cuenta.CentroCostoId,
                cuenta.TipoImpuesto);

            CargarPeriodosPresupuestariosEdit(
    cuenta.PresupuestoMensualId);
            // ================================================================
            // MOTIVO DE RECHAZO DEL GASTO ASOCIADO
            // ================================================================
            if (cuenta.EstadoAutorizacion == "Rechazado")
            {
                var gastoRechazado = await _context.Gastos
                    .AsNoTracking()
                    .FirstOrDefaultAsync(g =>
                        g.NumeroFactura == cuenta.NumeroDocumento ||
                        (
                            g.Concepto != null &&
                            g.Concepto.Contains(cuenta.NumeroDocumento)
                        ));

                ViewBag.MotivoRechazo =
                    gastoRechazado?.MotivoRechazo;
            }
            return View(cuenta);
        }

        // POST: CuentasPorPagar/Edit/5
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Edit(int id, [Bind("IdCuentaPorPagar,ProveedorId,PresupuestoMensualId,CategoriaGastoId,CentroCostoId,TipoDocumento,Concepto,NumeroFactura,MontoOriginal,Descuento,TipoImpuesto,DiasCredito,Observacion")] CuentasPorPagar cuentasPorPagar)
        {
            if (id != cuentasPorPagar.IdCuentaPorPagar)
            {
                return NotFound();
            }

            var cuentaActual = await _context.CuentasPorPagars
                .FirstOrDefaultAsync(c => c.IdCuentaPorPagar == id);

            if (cuentaActual == null)
            {
                return NotFound();
            }
            // ============================================================
            // BLOQUEO POR CIERRE CONTABLE
            // ============================================================
            bool periodoCerrado =
                await PeriodoContableCerradoAsync(cuentaActual.FechaEmision);

            if (periodoCerrado)
            {
                TempData["MensajeError"] =
                    $"No se puede modificar la cuenta por pagar porque el período " +
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

            bool estabaRechazada =
    cuentaActual.EstadoAutorizacion == "Rechazado";
            // Protección contra solicitudes POST manipuladas:
            // una CxP de factura solo se modifica desde Facturación.
            if (string.Equals(
                cuentaActual.Origen,
                "Factura",
                StringComparison.OrdinalIgnoreCase))
            {
                TempData["MensajeError"] =
                    "Esta cuenta por pagar fue generada desde una factura. " +
                    "Debe modificarse desde el módulo de Facturación.";

                return RedirectToAction(nameof(Index));
            }

            // Las cuentas finalizadas son únicamente de consulta.
            if (cuentaActual.Estado == "Anulada" ||
                cuentaActual.Estado == "Cancelada")
            {
                TempData["MensajeError"] =
                    $"No se puede modificar una cuenta por pagar {cuentaActual.Estado.ToLower()}.";

                return RedirectToAction(nameof(Index));
            }

            bool tienePagos = await _context.PagosCuentaPorPagars
                .AnyAsync(p => p.CuentaPorPagarId == cuentaActual.IdCuentaPorPagar
                            && p.Anulado == false);

            // Si tiene pagos, solo se permite modificar observación.
            if (tienePagos)
            {
                cuentaActual.Observacion = cuentasPorPagar.Observacion;

                if (estabaRechazada)
                {
                    cuentaActual.EstadoAutorizacion = "Pendiente";
                }

                await _context.SaveChangesAsync();

                await RegistrarAuditoria(
                    "Modificar",
                    cuentaActual.IdCuentaPorPagar,
                    "Se modificó la observación de la cuenta por pagar " + cuentaActual.NumeroDocumento
                );

                TempData["MensajeExito"] = "La observación fue modificada correctamente.";
                return RedirectToAction(nameof(Index));
            }
           
            // CXP MANUAL APROBADA: NO ADMITE MODIFICACIONES
            
            var gastoAprobado = await _context.Gastos
                .FirstOrDefaultAsync(g =>
                    g.NumeroFactura == cuentaActual.NumeroDocumento &&
                    g.Estado == "Aprobado");

            if (gastoAprobado != null ||
                cuentaActual.EstadoAutorizacion == "Aprobado")
            {
                TempData["MensajeError"] =
                    "Esta cuenta por pagar ya fue aprobada. " +
                    "Solo puede consultar sus detalles y registrar pagos.";

                return RedirectToAction(nameof(Index));
            }

            // Validaciones solo cuando NO tiene pagos.
            if (cuentasPorPagar.ProveedorId <= 0)
                ModelState.AddModelError("ProveedorId", "Debe seleccionar el proveedor.");

            if (cuentasPorPagar.CategoriaGastoId == null || cuentasPorPagar.CategoriaGastoId <= 0)
                ModelState.AddModelError("CategoriaGastoId", "Debe seleccionar la categoría de gasto.");

            if (cuentasPorPagar.CentroCostoId == null || cuentasPorPagar.CentroCostoId <= 0)
                ModelState.AddModelError("CentroCostoId", "Debe seleccionar el centro de costo.");

            if (cuentasPorPagar.DiasCredito == null || cuentasPorPagar.DiasCredito <= 0)
                ModelState.AddModelError("DiasCredito", "Debe seleccionar los días de crédito.");

            if (cuentasPorPagar.MontoOriginal <= 0)
                ModelState.AddModelError("MontoOriginal", "El monto original debe ser mayor a cero.");

            if ((cuentasPorPagar.Descuento ?? 0) > cuentasPorPagar.MontoOriginal)
                ModelState.AddModelError("Descuento", "El descuento no puede ser mayor al monto original.");

            ModelState.Remove("NumeroDocumento");
            ModelState.Remove("FechaEmision");
            ModelState.Remove("FechaVencimiento");
            ModelState.Remove("SubTotal");
            ModelState.Remove("Impuesto");
            ModelState.Remove("MontoPagado");
            ModelState.Remove("AnticipoAplicado");
            ModelState.Remove("SaldoPendiente");
            ModelState.Remove("Estado");
            ModelState.Remove("Activo");
            ModelState.Remove("Proveedor");
            ModelState.Remove("Proyecto");
            ModelState.Remove(
    "PresupuestoMensualRegistro");
            ModelState.Remove("CategoriaGasto");
            ModelState.Remove("CentroCosto");
            ModelState.Remove("PagosCuentaPorPagars");
            ModelState.Remove("MotivoAnulacion");

            CargarPeriodosPresupuestariosEdit(
    cuentasPorPagar.PresupuestoMensualId);

            PresupuestosMensuale? presupuestoSeleccionado = null;

            if (!cuentasPorPagar.PresupuestoMensualId.HasValue)
            {
                ModelState.AddModelError(
                    "PresupuestoMensualId",
                    "Debe seleccionar un proyecto y período presupuestario.");
            }
            else
            {
                presupuestoSeleccionado =
                    await _context.PresupuestosMensuales
                        .Include(pm => pm.Proyecto)
                        .FirstOrDefaultAsync(pm =>
                            pm.IdPresupuestoMensual ==
                                cuentasPorPagar.PresupuestoMensualId.Value &&

                            (
                                // Puede conservar el presupuesto histórico actual
                                pm.IdPresupuestoMensual ==
                                    cuentaActual.PresupuestoMensualId

                                ||

                                // O cambiar a un presupuesto actualmente utilizable
                                (
                                    pm.EstadoAprobacion == "Aprobado" &&
                                    pm.ProyectoId != null &&
                                    pm.Proyecto != null &&
                                    pm.Proyecto.Activo &&

                                    // NUEVO:
                                    // No permitir mover la CxP
                                    // hacia un período contable cerrado.
                                    !_context.CierresContables.Any(c =>
                                        c.Anio == pm.Anio &&
                                        c.Mes == pm.Mes &&
                                        c.Estado == "Cerrado")
                                )
                            ));

                if (presupuestoSeleccionado == null)
                {
                    ModelState.AddModelError(
                        "PresupuestoMensualId",
                        "El período presupuestario seleccionado no está disponible.");
                }
                else
                {
                    cuentasPorPagar.ProyectoId =
                        presupuestoSeleccionado.ProyectoId;
                }
            
        }
            if (!ModelState.IsValid)
            {
                CargarCombosEdit(
     cuentasPorPagar.ProveedorId,
     cuentasPorPagar.ProyectoId,
     cuentasPorPagar.TipoDocumento,
     cuentasPorPagar.DiasCredito,
     cuentasPorPagar.CategoriaGastoId,
     cuentasPorPagar.CentroCostoId,
     cuentasPorPagar.TipoImpuesto);

                return View(cuentasPorPagar);
            }

            cuentaActual.ProveedorId = cuentasPorPagar.ProveedorId;
            cuentaActual.ProyectoId =
      cuentasPorPagar.ProyectoId;

            cuentaActual.PresupuestoMensualId =
                cuentasPorPagar.PresupuestoMensualId;
            cuentaActual.CategoriaGastoId = cuentasPorPagar.CategoriaGastoId;
            cuentaActual.CentroCostoId = cuentasPorPagar.CentroCostoId;
            // No se modifica el tipo de documento para no romper el correlativo.
            // cuentaActual.TipoDocumento = cuentasPorPagar.TipoDocumento;
            cuentaActual.Concepto = cuentasPorPagar.Concepto; 
            cuentaActual.NumeroFactura = cuentasPorPagar.NumeroFactura;
            cuentaActual.MontoOriginal = cuentasPorPagar.MontoOriginal;
            cuentaActual.Descuento = cuentasPorPagar.Descuento ?? 0;
            cuentaActual.TipoImpuesto = cuentasPorPagar.TipoImpuesto;
            cuentaActual.DiasCredito = cuentasPorPagar.DiasCredito;
            cuentaActual.Observacion = cuentasPorPagar.Observacion;

            cuentaActual.FechaVencimiento = cuentaActual.FechaEmision.AddDays(cuentasPorPagar.DiasCredito ?? 0);

            cuentaActual.SubTotal = cuentaActual.MontoOriginal - (cuentaActual.Descuento ?? 0); 

            decimal porcentajeImpuesto = cuentaActual.TipoImpuesto switch
            {
                "IVA 13%" => 0.13m,
                "IVA 4%" => 0.04m,
                "IVA 2%" => 0.02m,
                "IVA 1%" => 0.01m,
                _ => 0m
            };

            cuentaActual.Impuesto = (cuentaActual.SubTotal ?? 0) * porcentajeImpuesto;

            cuentaActual.SaldoPendiente =
                (cuentaActual.SubTotal ?? 0)
                + (cuentaActual.Impuesto ?? 0)
                - (cuentaActual.MontoPagado ?? 0)
                - (cuentaActual.AnticipoAplicado ?? 0);

            if (cuentaActual.SaldoPendiente < 0)
            {
                cuentaActual.SaldoPendiente = 0;
            }

            ActualizarEstadoCuenta(cuentaActual);

            // Si esta CxP manual generó un gasto automático,
            // se actualiza el gasto mientras siga pendiente de autorización. 
            var gastoAutomatico = await _context.Gastos
                .FirstOrDefaultAsync(g => g.NumeroFactura == cuentaActual.NumeroDocumento);

            if (gastoAutomatico != null &&
            (
                gastoAutomatico.Estado == "Pendiente de autorización" ||
                gastoAutomatico.Estado == "Rechazado"
            ))
            {
                gastoAutomatico.Fecha = cuentaActual.FechaEmision;
                gastoAutomatico.CategoriaGastoId = cuentaActual.CategoriaGastoId.Value;
                gastoAutomatico.ProyectoId = cuentaActual.ProyectoId;
                gastoAutomatico.PresupuestoMensualId =
    cuentaActual.PresupuestoMensualId;
                gastoAutomatico.ProveedorId = cuentaActual.ProveedorId;
                gastoAutomatico.CentroCostoId = cuentaActual.CentroCostoId;
                gastoAutomatico.Concepto = "Gasto generado automáticamente desde la cuenta por pagar " + cuentaActual.NumeroDocumento;
                gastoAutomatico.Subtotal = cuentaActual.SubTotal;
                gastoAutomatico.Impuesto = cuentaActual.Impuesto;
                gastoAutomatico.Descuento = cuentaActual.Descuento ?? 0;
                gastoAutomatico.TipoImpuesto = cuentaActual.TipoImpuesto;
                gastoAutomatico.MontoTotal = (cuentaActual.SubTotal ?? 0) + (cuentaActual.Impuesto ?? 0);
                gastoAutomatico.Observacion = cuentaActual.Observacion;
                if (estabaRechazada)
                {
                    gastoAutomatico.Estado = "Pendiente de autorización";
                    gastoAutomatico.MotivoRechazo = null;
                    gastoAutomatico.FechaRechazo = null;
                    gastoAutomatico.UsuarioRechazoId = null;
                }
                _context.Update(gastoAutomatico);
            }
            if (estabaRechazada)
            {
                cuentaActual.EstadoAutorizacion = "Pendiente";
            }
            await _context.SaveChangesAsync();
      
            await RegistrarAuditoria(
                "Modificar",
                cuentaActual.IdCuentaPorPagar,
                "Se modificó la cuenta por pagar " + cuentaActual.NumeroDocumento
            );

            TempData["MensajeExito"] = "La cuenta por pagar fue modificada correctamente.";
            return RedirectToAction(nameof(Index));
        }
        private void CargarPeriodosPresupuestariosEdit(
    int? presupuestoMensualId)
        {
            var hoy = DateOnly.FromDateTime(DateTime.Today);
            var fechaMinima = hoy.AddMonths(-3);

            int anioMinimo = fechaMinima.Year;
            int mesMinimo = fechaMinima.Month;

            var periodos =
                _context.PresupuestosMensuales
                    .AsNoTracking()
                    .Include(pm => pm.Proyecto)
                    .Where(pm =>
                        (
                            pm.EstadoAprobacion == "Aprobado" &&
                            pm.ProyectoId != null &&
                            pm.Proyecto != null &&
                            pm.Proyecto.Activo &&
                                // NUEVO:
                    // No permitir seleccionar otros períodos cerrados.
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
                        ||
                        pm.IdPresupuestoMensual ==
                            presupuestoMensualId
                    )
                    .OrderByDescending(pm => pm.Anio)
.ThenByDescending(pm => pm.Mes)
.ThenBy(pm => pm.Proyecto!.Nombre)
                    .Select(pm => new
                    {
                        pm.IdPresupuestoMensual,

                        Nombre =
                            pm.Proyecto!.Nombre + " — " +
                            new DateTime(
                                pm.Anio,
                                pm.Mes,
                                1)
                            .ToString(
                                "MMMM yyyy",
                                new CultureInfo("es-CR"))
                    })
                    .ToList();

            ViewBag.PresupuestosMensualId =
                new SelectList(
                    periodos,
                    "IdPresupuestoMensual",
                    "Nombre",
                    presupuestoMensualId);
        }
        // ================================================================
        // GET: CuentasPorPagar/Delete/5
        // ================================================================
        public async Task<IActionResult> Delete(int? id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var cuenta = await _context.CuentasPorPagars
                .Include(c => c.Proveedor)
                .Include(c => c.Proyecto)
                  .Include(c => c.PresupuestoMensualRegistro)
                .FirstOrDefaultAsync(c =>
                    c.IdCuentaPorPagar == id.Value);

            if (cuenta == null)
            {
                return NotFound();
            }
            // ============================================================
            // BLOQUEO POR CIERRE CONTABLE
            // ============================================================
            bool periodoCerrado =
                await PeriodoContableCerradoAsync(cuenta.FechaEmision);

            if (periodoCerrado)
            {
                TempData["MensajeError"] =
                    $"No se puede anular la cuenta por pagar porque el período " +
                    $"{cuenta.FechaEmision:MM/yyyy} se encuentra cerrado. " +
                    $"Para realizar correcciones, primero debe reabrirse el período.";

                return RedirectToAction(
                    nameof(Index),
                    new
                    {
                        anio = cuenta.FechaEmision.Year,
                        mes = cuenta.FechaEmision.Month
                    });
            }
            // Las CxP creadas desde factura deben anularse desde Facturación.
            if (string.Equals(
                cuenta.Origen,
                "Factura",
                StringComparison.OrdinalIgnoreCase))
            {
                TempData["MensajeError"] =
                    "Esta cuenta por pagar fue generada desde una factura. " +
                    "Debe anular la factura de origen desde el módulo de Facturación.";

                return RedirectToAction(nameof(Index));
            }

            if (cuenta.Estado == "Anulada")
            {
                TempData["MensajeError"] =
                    "La cuenta por pagar ya se encuentra anulada.";

                return RedirectToAction(nameof(Index));
            }

            if (cuenta.Estado == "Cancelada")
            {
                TempData["MensajeError"] =
                    "No se puede anular una cuenta por pagar cancelada.";

                return RedirectToAction(nameof(Index));
            }

            bool tienePagosActivos =
                await _context.PagosCuentaPorPagars
                    .AnyAsync(p =>
                        p.CuentaPorPagarId ==
                            cuenta.IdCuentaPorPagar &&
                        p.Anulado == false);

            if (tienePagosActivos)
            {
                TempData["MensajeError"] =
                    "No se puede anular esta cuenta por pagar porque posee pagos activos. " +
                    "Primero debe anular los pagos relacionados.";

                return RedirectToAction(nameof(Index));
            }

            return View(cuenta);
        }


        // ================================================================
        // POST: CuentasPorPagar/Delete/5
        // ================================================================
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> DeleteConfirmed(
            int id,
            string motivoAnulacion)
        {
            var cuentasPorPagar =
                await _context.CuentasPorPagars.FindAsync(id);

            if (cuentasPorPagar == null)
            {
                return NotFound();
            }
            // ============================================================
            // BLOQUEO POR CIERRE CONTABLE
            // PROTECCIÓN CONTRA POST DIRECTO
            // ============================================================
            bool periodoCerrado =
                await PeriodoContableCerradoAsync(cuentasPorPagar.FechaEmision);

            if (periodoCerrado)
            {
                TempData["MensajeError"] =
                    $"No se puede anular la cuenta por pagar porque el período " +
                    $"{cuentasPorPagar.FechaEmision:MM/yyyy} se encuentra cerrado. " +
                    $"Para realizar correcciones, primero debe reabrirse el período.";

                return RedirectToAction(
                    nameof(Index),
                    new
                    {
                        anio = cuentasPorPagar.FechaEmision.Year,
                        mes = cuentasPorPagar.FechaEmision.Month
                    });
            }
            // Protección contra POST directo o manipulado.
            if (cuentasPorPagar.Estado == "Anulada")
            {
                TempData["MensajeError"] =
                    "La cuenta por pagar ya se encuentra anulada.";

                return RedirectToAction(nameof(Index));
            }

            if (cuentasPorPagar.Estado == "Cancelada")
            {
                TempData["MensajeError"] =
                    "No se puede anular una cuenta por pagar cancelada.";

                return RedirectToAction(nameof(Index));
            }

            if (string.Equals(
                cuentasPorPagar.Origen,
                "Factura",
                StringComparison.OrdinalIgnoreCase))
            {
                TempData["MensajeError"] =
                    "Esta cuenta por pagar fue generada desde una factura. " +
                    "Debe anular la factura de origen desde el módulo de Facturación.";

                return RedirectToAction(nameof(Index));
            }

            if (string.IsNullOrWhiteSpace(motivoAnulacion))
            {
                TempData["MensajeError"] =
                    "Debe ingresar el motivo de anulación.";

                return RedirectToAction(
                    nameof(Delete),
                    new { id });
            }

            motivoAnulacion = motivoAnulacion.Trim();

            if (!System.Text.RegularExpressions.Regex.IsMatch(
                motivoAnulacion,
                @"^(?=(?:.*[a-zA-ZáéíóúÁÉÍÓÚñÑ]){3,})[a-zA-ZáéíóúÁÉÍÓÚñÑ0-9\s.,;:/()#-]+$"))
            {
                TempData["MensajeError"] =
                    "El motivo de anulación debe contener al menos tres letras y solo caracteres permitidos.";

                return RedirectToAction(
                    nameof(Delete),
                    new { id });
            }

            bool tienePagosActivos =
                await _context.PagosCuentaPorPagars
                    .AnyAsync(p =>
                        p.CuentaPorPagarId ==
                            cuentasPorPagar.IdCuentaPorPagar &&
                        p.Anulado == false);

            if (tienePagosActivos)
            {
                TempData["MensajeError"] =
                    "No se puede anular esta cuenta por pagar porque posee pagos activos. " +
                    "Primero debe anular los pagos relacionados.";

                return RedirectToAction(nameof(Index));
            }

            // Anular la CxP.
            cuentasPorPagar.Estado = "Anulada";
            cuentasPorPagar.EstadoAutorizacion = "Anulada";
            cuentasPorPagar.Activo = false;
            cuentasPorPagar.MontoPagado = 0;
            cuentasPorPagar.SaldoPendiente = 0;
            cuentasPorPagar.MotivoAnulacion =
                motivoAnulacion;

            // Buscar el gasto generado por esta CxP manual.
            var gastoAsociado =
                await _context.Gastos
                    .FirstOrDefaultAsync(g =>
                        g.NumeroFactura ==
                            cuentasPorPagar.NumeroDocumento ||
                        (
                            g.Concepto != null &&
                            cuentasPorPagar.NumeroDocumento != null &&
                            g.Concepto.Contains(
                                cuentasPorPagar.NumeroDocumento)
                        ));

            if (gastoAsociado != null)
            {
                gastoAsociado.Estado = "Anulado";
                gastoAsociado.Activo = false;

                gastoAsociado.Observacion =
                    (gastoAsociado.Observacion ?? "") +
                    " | Anulado automáticamente por anulación de CxP manual. Motivo: " +
                    motivoAnulacion;

                await RevertirAsientoNacimientoCxPManual(
                    gastoAsociado,
                    cuentasPorPagar,
                    motivoAnulacion);
            }

            await _context.SaveChangesAsync();

            await RegistrarAuditoria(
                "Anular",
                cuentasPorPagar.IdCuentaPorPagar,
                "Se anuló la cuenta por pagar " +
                cuentasPorPagar.NumeroDocumento +
                ". Motivo: " +
                motivoAnulacion);

            TempData["MensajeExito"] =
                "La cuenta por pagar fue anulada correctamente.";

            return RedirectToAction(nameof(Index));
        }
        private void CargarCombos(
                int? proveedorId = null,
                int? proyectoId = null,
                string? tipoDocumento = null,
                int? diasCredito = null,
                int? categoriaGastoId = null,
                int? centroCostoId = null,
                string? tipoImpuesto = null)
        {
            ViewData["ProveedorId"] = new SelectList(
                _context.Proveedores.Where(p => p.Activo == true),
                "IdProveedor",
                "Nombre",
                proveedorId);

            ViewData["ProyectoId"] = new SelectList(
                _context.Proyectos.Where(p => p.Activo == true),
                "IdProyecto",
                "Nombre",
                proyectoId);

            ViewData["CategoriaGastoId"] = new SelectList(
                _context.CategoriasGastos.Where(c => c.Activo == true),
                "IdCategoriaGasto",
                "Nombre",
                categoriaGastoId);

            ViewData["CentroCostoId"] = new SelectList(
                _context.CentrosCostos.Where(c => c.Activo == true),
                "IdCentroCosto",
                "Nombre",
                centroCostoId);

            ViewBag.TiposDocumento = new SelectList(new List<string>
            {
                "Obligación",
                "Convenio",
                "Honorarios",
                "Préstamo",
                "Ajuste",
                "Servicio pendiente",
                "Otro"
            }, tipoDocumento);

            ViewBag.DiasCredito = new SelectList(new List<int>
            {
                8, 15, 30, 45, 60, 90, 120, 150, 180
            }, diasCredito);

            ViewBag.TiposImpuesto = new SelectList(new List<string>
            {
                "Sin impuesto",
                "IVA 13%",
                "IVA 4%",
                "IVA 2%",
                "IVA 1%"
            }, tipoImpuesto);
        }
        private void CargarPeriodosPresupuestarios(
    int? presupuestoMensualId = null)
        {
            var hoy = DateOnly.FromDateTime(DateTime.Today);
            var fechaMinima = hoy.AddMonths(-3);

            int anioMinimo = fechaMinima.Year;
            int mesMinimo = fechaMinima.Month;

            var periodos = _context.PresupuestosMensuales
                .AsNoTracking()
                .Include(pm => pm.Proyecto)
                .Where(pm =>
                    pm.EstadoAprobacion == "Aprobado" &&
                    pm.ProyectoId != null &&
                    pm.Proyecto != null &&
                    pm.Proyecto.Activo &&
        // NUEVO:
        // No permitir usar presupuestos
        // de períodos contables cerrados.
        !_context.CierresContables.Any(c =>
            c.Anio == pm.Anio &&
            c.Mes == pm.Mes &&
            c.Estado == "Cerrado") &&
                    (
                        pm.Anio > anioMinimo ||
                        (pm.Anio == anioMinimo &&
                         pm.Mes >= mesMinimo)
                    ))
               .OrderByDescending(pm => pm.Anio)
.ThenByDescending(pm => pm.Mes)
.ThenBy(pm => pm.Proyecto!.Nombre)
                .Select(pm => new
                {
                    pm.IdPresupuestoMensual,

                    Nombre =
                        pm.Proyecto!.Nombre + " — " +
                        new DateTime(pm.Anio, pm.Mes, 1)
                            .ToString(
                                "MMMM yyyy",
                                new CultureInfo("es-CR"))
                })
                .ToList();

            ViewBag.PresupuestosMensualId =
                new SelectList(
                    periodos,
                    "IdPresupuestoMensual",
                    "Nombre",
                    presupuestoMensualId);
        }
        private void CargarCombosEdit(
    int? proveedorId,
    int? proyectoId,
    string? tipoDocumento,
    int? diasCredito,
    int? categoriaGastoId,
    int? centroCostoId,
    string? tipoImpuesto)
        {
            ViewData["ProveedorId"] = new SelectList(
                _context.Proveedores
                    .Where(p =>
                        p.Activo == true ||
                        p.IdProveedor == proveedorId)
                    .OrderBy(p => p.Nombre),
                "IdProveedor",
                "Nombre",
                proveedorId);

            ViewData["ProyectoId"] = new SelectList(
                _context.Proyectos
                    .Where(p =>
                        p.Activo == true ||
                        p.IdProyecto == proyectoId)
                    .OrderBy(p => p.Nombre),
                "IdProyecto",
                "Nombre",
                proyectoId);

            ViewData["CategoriaGastoId"] = new SelectList(
                _context.CategoriasGastos
                    .Where(c =>
                        c.Activo == true ||
                        c.IdCategoriaGasto == categoriaGastoId)
                    .OrderBy(c => c.Nombre),
                "IdCategoriaGasto",
                "Nombre",
                categoriaGastoId);

            ViewData["CentroCostoId"] = new SelectList(
                _context.CentrosCostos
                    .Where(c =>
                        c.Activo == true ||
                        c.IdCentroCosto == centroCostoId)
                    .OrderBy(c => c.Nombre),
                "IdCentroCosto",
                "Nombre",
                centroCostoId);

            ViewBag.TiposDocumento = new SelectList(
                new List<string>
                {
            "Obligación",
            "Convenio",
            "Honorarios",
            "Préstamo",
            "Ajuste",
            "Servicio pendiente",
            "Otro"
                },
                tipoDocumento);

            ViewBag.DiasCredito = new SelectList(
                new List<int>
                {
            8, 15, 30, 45, 60, 90, 120, 150, 180
                },
                diasCredito);

            ViewBag.TiposImpuesto = new SelectList(
                new List<string>
                {
            "Sin impuesto",
            "IVA 13%",
            "IVA 4%",
            "IVA 2%",
            "IVA 1%"
                },
                tipoImpuesto);
        }

        private string GenerarNumeroDocumentoCxP(string? tipoDocumento)
        {
            string prefijo = tipoDocumento switch
            {
                "Obligación" => "OBL",
                "Convenio" => "CON",
                "Honorarios" => "HON",
                "Préstamo" => "PRE",
                "Ajuste" => "AJP",
                "Servicio pendiente" => "SER",
                "Otro" => "OTR",
                _ => "CXP"
            };

            var ultimoDocumento = _context.CuentasPorPagars
                .Where(c => c.TipoDocumento == tipoDocumento && c.NumeroDocumento != null)
                .OrderByDescending(c => c.IdCuentaPorPagar)
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
        // DESCRIBE LOS FILTROS UTILIZADOS EN REPORTES DE CXP
        // ================================================================
        // Permite registrar en Auditoría qué información fue consultada
        // o exportada por el usuario.
        private string ConstruirDescripcionFiltrosReporte(
            string? buscar,
            string? tipoDocumento,
            string? estado,
            DateOnly? fechaInicio,
            DateOnly? fechaFin,
            int? categoriaGastoId,
            int? centroCostoId,
            int? proyectoId,
            int? anio,
            int? mes,
            string? origen)
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

            if (categoriaGastoId.HasValue &&
                categoriaGastoId.Value > 0)
            {
                filtros.Add($"CategoriaGastoId={categoriaGastoId.Value}");
            }

            if (centroCostoId.HasValue &&
                centroCostoId.Value > 0)
            {
                filtros.Add($"CentroCostoId={centroCostoId.Value}");
            }

            if (proyectoId.HasValue &&
                proyectoId.Value > 0)
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
        // Registra las acciones realizadas sobre cuentas por pagar. AUDITORIA
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
                    Tabla = "CuentasPorPagar",
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
        // ACTUALIZA EL ESTADO DE UNA CUENTA POR PAGAR ESPECÍFICA
        // ================================================================
        // Se utiliza después de crear, modificar o aplicar pagos a una CxP.
        // No guarda cambios; el método que la llama realiza SaveChangesAsync().
        private void ActualizarEstadoCuenta(CuentasPorPagar cuenta)
        {
            // Una cuenta anulada conserva siempre su estado.
            if (cuenta.Estado == "Anulada")
            {
                return;
            }

            // Si ya no existe saldo, la obligación está cancelada.
            if (cuenta.SaldoPendiente <= 0)
            {
                cuenta.SaldoPendiente = 0;
                cuenta.Estado = "Cancelada";
            }
            // Si conserva saldo y pasó la fecha, está vencida.
            // Vencida tiene prioridad sobre Parcial.
            else if (cuenta.FechaVencimiento <
                     DateOnly.FromDateTime(DateTime.Today))
            {
                cuenta.Estado = "Vencida";
            }
            // Si tiene pagos, conserva saldo y todavía no vence, está parcial.
            else if ((cuenta.MontoPagado ?? 0m) > 0)
            {
                cuenta.Estado = "Parcial";
            }
            // Sin pagos y sin vencer, continúa pendiente.
            else
            {
                cuenta.Estado = "Pendiente";
            }
        }
        // ================================================================
        // ACTUALIZA POR VENCIMIENTO TODAS LAS CUENTAS POR PAGAR
        // ================================================================
        // No crea gastos, pagos ni movimientos contables.
        // Únicamente corrige el estado administrativo por el paso del tiempo.
        private async Task ActualizarEstadosCuentasPorPagar()
        {
            var cuentas = await _context.CuentasPorPagars
                .Where(c =>
                    c.Activo == true &&
                    c.Estado != "Anulada")
                .ToListAsync();

            bool huboCambios = false;

            foreach (var cuenta in cuentas)
            {
                string estadoAnterior = cuenta.Estado;

                // Reutiliza exactamente la regla individual.
                ActualizarEstadoCuenta(cuenta);

                if (cuenta.Estado != estadoAnterior)
                {
                    huboCambios = true;
                }
            }

            // Un único guardado para todas las cuentas que cambiaron.
            if (huboCambios)
            {
                await _context.SaveChangesAsync();
            }
        }
        // ================================================================
        // CAMBIO APLICADO - REVERSIÓN ASIENTO NACIMIENTO CXP MANUAL
        // Revierte el asiento contable generado cuando se aprobó el gasto
        // asociado a una CxP manual.
        //
        // Asiento original esperado:
        // Debe: cuenta de gasto
        // Haber: CC002 Cuentas por Pagar
        //
        // Reversión:
        // Debe: CC002 Cuentas por Pagar
        // Haber: cuenta de gasto
        // ================================================================
        private async Task RevertirAsientoNacimientoCxPManual(
            Gasto gasto,
            CuentasPorPagar cuentaPorPagar,
            string motivoAnulacion)
        {
            if (gasto == null)
            {
                return;
            }

            string referenciaOriginal = gasto.NumeroFactura ?? cuentaPorPagar.NumeroDocumento ?? "";
            string referenciaReversa =
                "REV-CXP-MANUAL-" + cuentaPorPagar.IdCuentaPorPagar + "-" + DateTime.Now.ToString("yyyyMMddHHmmss");

            if (string.IsNullOrWhiteSpace(referenciaOriginal))
            {
                return;
            }

            var movimientosOriginales = await _context.MovimientosContables
                .Where(m =>
                    m.OrigenModulo == "Gastos" &&
                    m.OrigenId == gasto.IdGasto &&
                    m.Estado == "Registrado" &&
                    m.Anulado == false &&
                    m.MovimientoReversionId == null)
                .ToListAsync();

            if (!movimientosOriginales.Any())
            {
                return;
            }

            var usuarioId = HttpContext.Session.GetInt32("UsuarioId") ?? 1;

            foreach (var movimiento in movimientosOriginales)
            {
                movimiento.Estado = "Anulado";
                movimiento.Anulado = true;
                movimiento.FechaAnulacion = DateTime.Now;
                movimiento.MotivoAnulacion =
                    "Anulación de CxP manual " + cuentaPorPagar.NumeroDocumento +
                    ". Motivo: " + motivoAnulacion;

                _context.MovimientosContables.Update(movimiento);
            }

            foreach (var movimiento in movimientosOriginales)
            {
                var movimientoReversa = new MovimientosContable
                {
                    Fecha = DateOnly.FromDateTime(DateTime.Today),
                    CuentaContableId = movimiento.CuentaContableId,
                    ProyectoId = movimiento.ProyectoId,
                    CentroCostoId = movimiento.CentroCostoId,

                    TipoMovimiento = movimiento.TipoMovimiento == "Debe" ? "Haber" : "Debe",
                    Monto = movimiento.Monto,

                    Debe = movimiento.Haber,
                    Haber = movimiento.Debe,

                    Referencia = referenciaReversa,
                    OrigenModulo = "Gastos",
                    OrigenId = gasto.IdGasto,
                    Descripcion = "Reversión de asiento por anulación de CxP manual " + cuentaPorPagar.NumeroDocumento,

                    Estado = "Reversado",
                    EsAutomatico = true,
                    UsuarioId = usuarioId,
                    FechaCreacion = DateTime.Now
                };

                _context.MovimientosContables.Add(movimientoReversa);

                await _context.SaveChangesAsync();

                movimiento.MovimientoReversionId = movimientoReversa.IdMovimientoContable;
                _context.MovimientosContables.Update(movimiento);

                var cuentaContable = await _context.CuentasContables
                    .FirstOrDefaultAsync(c => c.IdCuentaContable == movimiento.CuentaContableId);

                if (cuentaContable != null)
                {
                    AplicarSaldoCuenta(cuentaContable, movimientoReversa.Debe, movimientoReversa.Haber);
                    _context.CuentasContables.Update(cuentaContable);
                }
            }

            await _context.SaveChangesAsync();
        }

        // ================================================================
        // CAMBIO APLICADO - ACTUALIZACIÓN DE SALDOS CONTABLES CXP MANUAL
        // Respeta la naturaleza contable.
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
        private async Task<bool> PeriodoContableCerradoAsync(DateOnly fecha)
        {
            return await _context.CierresContables
                .AsNoTracking()
                .AnyAsync(c =>
                    c.Anio == fecha.Year &&
                    c.Mes == fecha.Month &&
                    c.Estado == "Cerrado");
        }
        private bool CuentasPorPagarExists(int id)
        {
            return _context.CuentasPorPagars.Any(e => e.IdCuentaPorPagar == id);
        }
        
    }
}