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
using SistemaGestionFinanciera.Services;
using ClosedXML.Excel;
using System.IO;
using System.Globalization;
using DocumentFormat.OpenXml.InkML;

namespace SistemaGestionFinanciera.Controllers
{
    public class FacturasController : BaseController
    {
        private readonly SistemaFinancieroContext _context;
        private readonly ContabilidadService _contabilidadService;

        public FacturasController(
            SistemaFinancieroContext context,
            ContabilidadService contabilidadService)
        {
            _context = context;
            _contabilidadService = contabilidadService;
        }

        // GET: Facturas
        public async Task<IActionResult> Index(
            string? tipoFactura,
            string? estado,
            string? condicionPago,
            int? clienteId,
            int? proveedorId,
            int? proyectoId,
            int? anio,
            int? mes,
            DateOnly? fechaDesde,
            DateOnly? fechaHasta,
            bool proximasVencer = false)
        {
            // Detecta si el usuario aplicó algún filtro.
            bool hayFiltros =
                !string.IsNullOrWhiteSpace(tipoFactura) ||
                !string.IsNullOrWhiteSpace(estado) ||
                !string.IsNullOrWhiteSpace(condicionPago) ||
                clienteId.HasValue ||
                proveedorId.HasValue ||
                proyectoId.HasValue ||
                anio.HasValue ||
                mes.HasValue ||
                fechaDesde.HasValue ||
                fechaHasta.HasValue ||
                proximasVencer;

            // Obtiene las facturas con los filtros aplicados.
            var facturas = await ConstruirConsultaReporte(
                tipoFactura,
                estado,
                condicionPago,
                clienteId,
                proveedorId,
                proyectoId,
                anio,
                mes,
                fechaDesde,
                fechaHasta,
                proximasVencer);

            var facturasParaTotales = facturas;
            // ================================================================
            // ALERTAS DE FACTURACIÓN
            // ================================================================

            var hoyAlertas =
                DateOnly.FromDateTime(DateTime.Today);

            var fechaLimiteProximas =
                hoyAlertas.AddDays(7);

            // ---------------------------------------------------------------
            // FACTURAS PRÓXIMAS A VENCER
            // Incluye ventas y compras a crédito con saldo pendiente.
            // ---------------------------------------------------------------
            var facturasProximasVencer =
                await _context.Facturas
                    .AsNoTracking()
                    .Where(f =>
                        f.Estado != "Anulada" &&
                        f.Estado != "Pagada" &&
                        f.Estado != "Vencida" &&
                        f.SaldoPendiente > 0 &&
                        f.FechaVencimiento.HasValue &&
                        f.FechaVencimiento.Value >= hoyAlertas &&
                        f.FechaVencimiento.Value <= fechaLimiteProximas &&
                        (
                            f.CondicionPago == "Credito" ||
                            f.CondicionPago == "Crédito"
                        ))
                    .ToListAsync();

            ViewBag.TotalFacturasProximas =
                facturasProximasVencer.Count;

            ViewBag.MontoFacturasProximas =
                facturasProximasVencer.Sum(f =>
                    f.SaldoPendiente);

            ViewBag.VentasProximas =
                facturasProximasVencer.Count(f =>
                    f.TipoFactura == "Venta");

            ViewBag.ComprasProximas =
                facturasProximasVencer.Count(f =>
                    f.TipoFactura == "Compra");

            // ---------------------------------------------------------------
            // FACTURAS VENCIDAS
            // Incluye todas las ventas y compras vencidas con saldo.
            // ---------------------------------------------------------------
            var facturasVencidas =
                await _context.Facturas
                    .AsNoTracking()
                    .Where(f =>
                        f.Estado != "Anulada" &&
                        f.Estado == "Vencida" &&
                        f.SaldoPendiente > 0)
                    .ToListAsync();

            ViewBag.TotalFacturasVencidas =
                facturasVencidas.Count;

            ViewBag.MontoFacturasVencidas =
                facturasVencidas.Sum(f =>
                    f.SaldoPendiente);

            ViewBag.VentasVencidas =
                facturasVencidas.Count(f =>
                    f.TipoFactura == "Venta");

            ViewBag.ComprasVencidas =
                facturasVencidas.Count(f =>
                    f.TipoFactura == "Compra");

            // Si no hay filtros, las tarjetas muestran solamente el mes actual.
            if (!hayFiltros)
            {
                var hoy = DateOnly.FromDateTime(DateTime.Today);
                var inicioMes = new DateOnly(hoy.Year, hoy.Month, 1);
                var finMes = inicioMes.AddMonths(1).AddDays(-1);

                facturasParaTotales = facturas
                    .Where(f =>
                        f.FechaEmision >= inicioMes &&
                        f.FechaEmision <= finMes)
                    .ToList();

                ViewBag.TipoResumen = "Mes actual";
            }
            else
            {
                ViewBag.TipoResumen = "Filtro aplicado";
            }

            // Las facturas anuladas no participan en los totales financieros.
            var facturasValidas = facturasParaTotales
                .Where(f => f.Estado != "Anulada")
                .ToList();

            // Totales de facturas de venta.
            ViewBag.TotalVentas = facturasValidas
                .Where(f => f.TipoFactura == "Venta")
                .Sum(f => f.Total);

            ViewBag.TotalCobrado = facturasValidas
                .Where(f => f.TipoFactura == "Venta")
                .Sum(f => f.MontoPagado);

            ViewBag.SaldoPorCobrar = facturasValidas
                .Where(f => f.TipoFactura == "Venta")
                .Sum(f => f.SaldoPendiente);

            // Totales de facturas de compra.
            ViewBag.TotalCompras = facturasValidas
                .Where(f => f.TipoFactura == "Compra")
                .Sum(f => f.Total);

            ViewBag.TotalPagadoCompras = facturasValidas
                .Where(f => f.TipoFactura == "Compra")
                .Sum(f => f.MontoPagado);

            ViewBag.SaldoPorPagar = facturasValidas
                .Where(f => f.TipoFactura == "Compra")
                .Sum(f => f.SaldoPendiente);

            // Cantidades por estado.
            ViewBag.CantidadPendientes = facturasValidas.Count(f =>
                f.Estado == "Pendiente");

            ViewBag.CantidadVencidas = facturasValidas.Count(f =>
                f.Estado == "Vencida");

            ViewBag.CantidadPagadas = facturasValidas.Count(f =>
                f.Estado == "Pagada");

            ViewBag.CantidadAnuladas = facturasParaTotales.Count(f =>
                f.Estado == "Anulada");

            // Conserva los filtros seleccionados en la pantalla.
            ViewData["TipoFactura"] = tipoFactura;
            ViewData["Estado"] = estado;
            ViewData["CondicionPago"] = condicionPago;
            ViewData["ClienteId"] = clienteId;
            ViewData["ProveedorId"] = proveedorId;
            ViewData["ProyectoId"] = proyectoId;
            ViewData["Anio"] = anio;
            ViewData["Mes"] = mes;
            ViewData["FechaDesde"] = fechaDesde?.ToString("yyyy-MM-dd");
            ViewData["FechaHasta"] = fechaHasta?.ToString("yyyy-MM-dd");
            ViewData["ProximasVencer"] = proximasVencer;

            // Carga solamente clientes activos.
            ViewBag.Clientes = new SelectList(
                await _context.ClientesBeneficiarios
                    .Where(c => c.Activo)
                    .OrderBy(c => c.Nombre)
                    .ToListAsync(),
                "IdClienteBeneficiario",
                "Nombre",
                clienteId);

            // Carga solamente proveedores activos.
            ViewBag.Proveedores = new SelectList(
                await _context.Proveedores
                    .Where(p => p.Activo)
                    .OrderBy(p => p.Nombre)
                    .ToListAsync(),
                "IdProveedor",
                "Nombre",
                proveedorId);

            // Carga solamente proyectos activos.
            ViewBag.Proyectos = new SelectList(
                await _context.Proyectos
                    .Where(p => p.Activo)
                    .OrderBy(p => p.Nombre)
                    .ToListAsync(),
                "IdProyecto",
                "Nombre",
                proyectoId);

            // Obtiene los años existentes en las facturas.
            ViewBag.Anios = await _context.Facturas
                .Select(f => f.FechaEmision.Year)
                .Distinct()
                .OrderByDescending(a => a)
                .ToListAsync();

            return View(facturas);
        }
        // ===========================================================
        // CONSTRUYE LA CONSULTA PARA INDEX, PDF Y EXCEL
        // ===========================================================
        private async Task<List<Factura>> ConstruirConsultaReporte(
            string? tipoFactura,
            string? estado,
            string? condicionPago,
            int? clienteId,
            int? proveedorId,
            int? proyectoId,
            int? anio,
            int? mes,
            DateOnly? fechaDesde,
            DateOnly? fechaHasta,
            bool proximasVencer = false)
        {
            var consulta = _context.Facturas
                .Include(f => f.ClienteBeneficiario)
                .Include(f => f.Proveedor)
                .Include(f => f.Proyecto)
                .Include(f => f.PresupuestoMensualRegistro)
                .Include(f => f.CategoriaIngreso)
                .Include(f => f.CategoriaGasto)
                .Include(f => f.CentroCosto)
                .AsQueryable();
            // Filtro especial activado desde la alerta:
            // facturas que vencen entre hoy y los próximos siete días.
            if (proximasVencer)
            {
                var hoyFiltro=
                    DateOnly.FromDateTime(DateTime.Today);

                var fechaLimite =
                    hoyFiltro.AddDays(7);

                consulta = consulta.Where(f =>
                    f.Estado != "Anulada" &&
                    f.Estado != "Pagada" &&
                    f.Estado != "Vencida" &&
                    f.SaldoPendiente > 0 &&
                    f.FechaVencimiento.HasValue &&
                    f.FechaVencimiento.Value >= hoyFiltro &&
                    f.FechaVencimiento.Value <= fechaLimite &&
                    (
                        f.CondicionPago == "Credito" ||
                        f.CondicionPago == "Crédito"
                    ));
            }
            // Filtro por tipo: Venta o Compra.
            if (!string.IsNullOrWhiteSpace(tipoFactura))
            {
                consulta = consulta.Where(f =>
                    f.TipoFactura == tipoFactura);
            }

            // Filtro por estado.
            if (!string.IsNullOrWhiteSpace(estado))
            {
                consulta = consulta.Where(f =>
                    f.Estado == estado);
            }
            // Filtro por condición de pago: Contado o Crédito.
            if (!string.IsNullOrWhiteSpace(condicionPago))
            {
                consulta = consulta.Where(f =>
                    f.CondicionPago == condicionPago);
            }

            // Filtro por cliente.
            if (clienteId.HasValue)
            {
                consulta = consulta.Where(f =>
                    f.ClienteBeneficiarioId == clienteId.Value);
            }

            // Filtro por proveedor.
            if (proveedorId.HasValue)
            {
                consulta = consulta.Where(f =>
                    f.ProveedorId == proveedorId.Value);
            }

            // Filtro por proyecto.
            if (proyectoId.HasValue)
            {
                consulta = consulta.Where(f =>
                    f.ProyectoId == proyectoId.Value);
            }

            // Filtro por año.
            if (anio.HasValue)
            {
                consulta = consulta.Where(f =>
                    f.FechaEmision.Year == anio.Value);
            }

            // Filtro por mes.
            if (mes.HasValue)
            {
                consulta = consulta.Where(f =>
                    f.FechaEmision.Month == mes.Value);
            }

            // Filtro por fecha inicial.
            if (fechaDesde.HasValue)
            {
                consulta = consulta.Where(f =>
                    f.FechaEmision >= fechaDesde.Value);
            }

            // Filtro por fecha final.
            if (fechaHasta.HasValue)
            {
                consulta = consulta.Where(f =>
                    f.FechaEmision <= fechaHasta.Value);
            }

            // Obtiene los registros más recientes primero.
            var facturas = await consulta
                .OrderByDescending(f => f.IdFactura)
                .ToListAsync();

            // Actualiza automáticamente las facturas vencidas.
            bool huboCambios = false;
            var hoy = DateOnly.FromDateTime(DateTime.Today);

            foreach (var factura in facturas)
            {
                bool debeMarcarseVencida =
      factura.FechaVencimiento < hoy &&
      factura.SaldoPendiente > 0 &&
      factura.Estado != "Pagada" &&
      factura.Estado != "Vencida" &&
      factura.Estado != "Anulada";

                if (debeMarcarseVencida)
                {
                    factura.Estado = "Vencida";
                    huboCambios = true;
                }
            }

            if (huboCambios)
            {
                await _context.SaveChangesAsync();
            }

            /*
             * Si el usuario filtró específicamente por estado, es necesario
             * volver a comprobarlo después de actualizar las vencidas.
             * Ejemplo:
             * una factura estaba "Pendiente", pero durante esta consulta pasó
             * a "Vencida"; ya no debe mostrarse si el filtro era Pendiente.
             */
            if (!string.IsNullOrWhiteSpace(estado))
            {
                facturas = facturas
                    .Where(f => f.Estado == estado)
                    .ToList();
            }

            return facturas;
        }
        // ===========================================================
        // VISTA PREVIA DEL REPORTE DE FACTURAS
        // ===========================================================
        public async Task<IActionResult> Reporte(
            string? tipoFactura,
            string? estado,
            string? condicionPago,
            int? clienteId,
            int? proveedorId,
            int? proyectoId,
            int? anio,
            int? mes,
            DateOnly? fechaDesde,
            DateOnly? fechaHasta,
            bool proximasVencer = false)
        {
            var facturas = await ConstruirConsultaReporte(
                tipoFactura,
                estado,
                condicionPago,
                clienteId,
                proveedorId,
                proyectoId,
                anio,
                mes,
                fechaDesde,
                fechaHasta,
                proximasVencer);

            // Para PDF se ordena por fecha y, en caso de empate, por ID.
            facturas = facturas
                .OrderByDescending(f => f.FechaEmision)
                .ThenByDescending(f => f.IdFactura)
                .ToList();

            var facturasValidas = facturas
                .Where(f => f.Estado != "Anulada")
                .ToList();

            var facturasVenta = facturasValidas
                .Where(f => f.TipoFactura == "Venta")
                .ToList();

            var facturasCompra = facturasValidas
                .Where(f => f.TipoFactura == "Compra")
                .ToList();

            ViewBag.TotalRegistros = facturas.Count;

            // Totales de ventas.
            ViewBag.TotalVentas = facturasVenta.Sum(f => f.Total);
            ViewBag.TotalCobrado = facturasVenta.Sum(f => f.MontoPagado);
            ViewBag.SaldoPorCobrar = facturasVenta.Sum(f => f.SaldoPendiente);

            // Totales de compras.
            ViewBag.TotalCompras = facturasCompra.Sum(f => f.Total);
            ViewBag.TotalPagadoCompras = facturasCompra.Sum(f => f.MontoPagado);
            ViewBag.SaldoPorPagar = facturasCompra.Sum(f => f.SaldoPendiente);

            // Cantidades por tipo y estado.
            ViewBag.CantidadVentas = facturas.Count(f =>
                f.TipoFactura == "Venta");

            ViewBag.CantidadCompras = facturas.Count(f =>
                f.TipoFactura == "Compra");

            ViewBag.CantidadPendientes = facturasValidas.Count(f =>
                f.Estado == "Pendiente");

            ViewBag.CantidadVencidas = facturasValidas.Count(f =>
                f.Estado == "Vencida");

            ViewBag.CantidadPagadas = facturasValidas.Count(f =>
                f.Estado == "Pagada");

            ViewBag.CantidadAnuladas = facturas.Count(f =>
                f.Estado == "Anulada");

            ViewBag.FechaGeneracion = DateTime.Now;

            // Conserva los filtros para mostrarlos en el reporte.
            ViewBag.TipoFactura = tipoFactura;
            ViewBag.Estado = estado;
            ViewBag.CondicionPago = condicionPago;
            ViewBag.ClienteId = clienteId;
            ViewBag.ProveedorId = proveedorId;
            ViewBag.ProyectoId = proyectoId;
            ViewBag.Anio = anio;
            ViewBag.Mes = mes;
            ViewBag.FechaDesde = fechaDesde?.ToString("yyyy-MM-dd");
            ViewBag.FechaHasta = fechaHasta?.ToString("yyyy-MM-dd");
            ViewBag.ProximasVencer = proximasVencer;

            // Nombres para que el PDF no muestre solamente el ID.
            ViewBag.NombreCliente = clienteId.HasValue
                ? await _context.ClientesBeneficiarios
                    .Where(c => c.IdClienteBeneficiario == clienteId.Value)
                    .Select(c => c.Nombre)
                    .FirstOrDefaultAsync()
                : null;

            ViewBag.NombreProveedor = proveedorId.HasValue
                ? await _context.Proveedores
                    .Where(p => p.IdProveedor == proveedorId.Value)
                    .Select(p => p.Nombre)
                    .FirstOrDefaultAsync()
                : null;

            ViewBag.NombreProyecto = proyectoId.HasValue
                ? await _context.Proyectos
                    .Where(p => p.IdProyecto == proyectoId.Value)
                    .Select(p => p.Nombre)
                    .FirstOrDefaultAsync()
                : null;
            // Registra la generación de la vista previa del reporte.
            // RegistroId = 0 porque el reporte corresponde a un conjunto
            // de facturas y no a una factura individual.
            await RegistrarAuditoriaFactura(
                "Vista previa PDF",
                0,
                "Se generó la vista previa del reporte de facturas. " +
                ConstruirDescripcionFiltrosReporte(
                    tipoFactura,
                    estado,
                    condicionPago,
                    clienteId,
                    proveedorId,
                    proyectoId,
                    anio,
                    mes,
                    fechaDesde,
                    fechaHasta)
            );

         
            return View(facturas);
        }
        // ===========================================================
        // EXPORTAR REPORTE DE FACTURAS A EXCEL
        // ===========================================================
        public async Task<IActionResult> ExportarExcel(
            string? tipoFactura,
            string? estado,
            string? condicionPago,
            int? clienteId,
            int? proveedorId,
            int? proyectoId,
            int? anio,
            int? mes,
            DateOnly? fechaDesde,
            DateOnly? fechaHasta,
            bool proximasVencer = false)
        {
            var facturas = await ConstruirConsultaReporte(
                tipoFactura,
                estado,
                condicionPago,
                clienteId,
                proveedorId,
                proyectoId,
                anio,
                mes,
                fechaDesde,
                fechaHasta,
                proximasVencer);

            // Orden definitivo del reporte:
            // primero la fecha más reciente y luego el ID más reciente.
            facturas = facturas
                .OrderByDescending(f => f.FechaEmision)
                .ThenByDescending(f => f.IdFactura)
                .ToList();

            using var workbook = new XLWorkbook();

            var hoja = workbook.Worksheets.Add("Facturas");

            const int totalColumnas = 14;

            // ===========================================================
            // ENCABEZADO PRINCIPAL
            // ===========================================================
            hoja.Cell("A1").Value =
                "Unión Cantonal de Asociaciones de Desarrollo de Santa Ana";

            hoja.Range(1, 1, 1, totalColumnas).Merge();

            hoja.Cell("A2").Value = "Reporte de Facturas";

            hoja.Range(2, 1, 2, totalColumnas).Merge();

            hoja.Cell("A3").Value =
                $"Generado el {DateTime.Now:dd/MM/yyyy HH:mm}";

            hoja.Range(3, 1, 3, totalColumnas).Merge();

            hoja.Range(1, 1, 1, totalColumnas).Style.Font.Bold = true;
            hoja.Range(1, 1, 1, totalColumnas).Style.Font.FontSize = 16;

            hoja.Range(2, 1, 2, totalColumnas).Style.Font.Bold = true;
            hoja.Range(2, 1, 2, totalColumnas).Style.Font.FontSize = 13;

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

            if (!string.IsNullOrWhiteSpace(tipoFactura))
            {
                filtros.Add($"Tipo: {tipoFactura}");
            }

            if (!string.IsNullOrWhiteSpace(estado))
            {
                filtros.Add($"Estado: {estado}");
            }
            if (proximasVencer)
            {
                filtros.Add(
                    "Vencimiento: Próximos 7 días");
            }
            if (!string.IsNullOrWhiteSpace(condicionPago))
            {
                filtros.Add($"Condición de pago: {condicionPago}");
            }

            if (clienteId.HasValue)
            {
                string? nombreCliente = await _context.ClientesBeneficiarios
                    .Where(c =>
                        c.IdClienteBeneficiario == clienteId.Value)
                    .Select(c => c.Nombre)
                    .FirstOrDefaultAsync();

                filtros.Add(
                    $"Cliente: {nombreCliente ?? clienteId.Value.ToString()}");
            }

            if (proveedorId.HasValue)
            {
                string? nombreProveedor = await _context.Proveedores
                    .Where(p =>
                        p.IdProveedor == proveedorId.Value)
                    .Select(p => p.Nombre)
                    .FirstOrDefaultAsync();

                filtros.Add(
                    $"Proveedor: {nombreProveedor ?? proveedorId.Value.ToString()}");
            }

            if (proyectoId.HasValue)
            {
                string? nombreProyecto = await _context.Proyectos
                    .Where(p =>
                        p.IdProyecto == proyectoId.Value)
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

            if (fechaDesde.HasValue)
            {
                filtros.Add(
                    $"Desde: {fechaDesde.Value:dd/MM/yyyy}");
            }

            if (fechaHasta.HasValue)
            {
                filtros.Add(
                    $"Hasta: {fechaHasta.Value:dd/MM/yyyy}");
            }

            hoja.Cell("A5").Value = "Filtros aplicados:";
            hoja.Cell("A5").Style.Font.Bold = true;

            hoja.Cell("B5").Value = filtros.Any()
                ? string.Join(" | ", filtros)
                : "Sin filtros. Se muestran todas las facturas.";

            hoja.Range(5, 2, 5, totalColumnas).Merge();
            hoja.Range(5, 2, 5, totalColumnas).Style.Alignment.WrapText = true;

            // ===========================================================
            // ENCABEZADOS DE LA TABLA
            // ===========================================================
            const int filaEncabezado = 7;

            string[] encabezados =
            {
        "Número",
        "Tipo",
        "Condición de pago",
        "Fecha de emisión",
        "Fecha de vencimiento",
        "Cliente / Proveedor",
        "Proyecto",
        "Estado",
        "Subtotal",
        "Descuento",
        "Impuesto",
        "Total",
        "Monto pagado",
        "Saldo pendiente"
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
            rangoEncabezado.Style.Font.FontColor = XLColor.White;

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

            foreach (var factura in facturas)
            {
                string tercero = factura.TipoFactura == "Venta"
                    ? factura.ClienteBeneficiario?.Nombre ?? "Sin cliente"
                    : factura.Proveedor?.Nombre ?? "Sin proveedor";

                hoja.Cell(fila, 1).Value =
                    factura.NumeroFactura ?? "Sin número";

                hoja.Cell(fila, 2).Value =
                    factura.TipoFactura ?? "Sin tipo";

                hoja.Cell(fila, 3).Value =
                    factura.CondicionPago ?? "Sin condición";

                hoja.Cell(fila, 4).Value =
                    factura.FechaEmision.ToDateTime(TimeOnly.MinValue);

                if (factura.FechaVencimiento.HasValue)
                {
                    hoja.Cell(fila, 5).Value =
                        factura.FechaVencimiento.Value
                            .ToDateTime(TimeOnly.MinValue);
                }
                else
                {
                    hoja.Cell(fila, 5).Value = "Sin vencimiento";
                }

                hoja.Cell(fila, 6).Value = tercero;

                hoja.Cell(fila, 7).Value =
                    factura.Proyecto?.Nombre ?? "Sin proyecto";

                hoja.Cell(fila, 8).Value =
                    factura.Estado ?? "Sin estado";

                hoja.Cell(fila, 9).Value = factura.SubTotal;
                hoja.Cell(fila, 10).Value = factura.DescuentoTotal ?? 0m;
                hoja.Cell(fila, 11).Value = factura.ImpuestoTotal;
                hoja.Cell(fila, 12).Value = factura.Total;
                hoja.Cell(fila, 13).Value = factura.MontoPagado;
                hoja.Cell(fila, 14).Value = factura.SaldoPendiente;

                fila++;
            }

            // ===========================================================
            // FORMATO DE DATOS
            // ===========================================================
            if (facturas.Any())
            {
                var rangoDatos = hoja.Range(
                    filaEncabezado + 1,
                    1,
                    fila - 1,
                    totalColumnas);

                rangoDatos.Style.Border.TopBorder =
                    XLBorderStyleValues.Thin;

                rangoDatos.Style.Border.BottomBorder =
                    XLBorderStyleValues.Thin;

                rangoDatos.Style.Border.LeftBorder =
                    XLBorderStyleValues.Thin;

                rangoDatos.Style.Border.RightBorder =
                    XLBorderStyleValues.Thin;

                rangoDatos.Style.Alignment.Vertical =
                    XLAlignmentVerticalValues.Center;

                rangoDatos.Style.Alignment.WrapText = true;

                // Formato de fechas.
                hoja.Range(
                    filaEncabezado + 1,
                    4,
                    fila - 1,
                    5)
                    .Style.DateFormat.Format = "dd/MM/yyyy";

                // Formato monetario.
                hoja.Range(
                    filaEncabezado + 1,
                    9,
                    fila - 1,
                    14)
                    .Style.NumberFormat.Format = "₡#,##0.00";

                // Agrega filtros automáticos al encabezado.
                hoja.Range(
                    filaEncabezado,
                    1,
                    fila - 1,
                    totalColumnas)
                    .SetAutoFilter();
                // ===========================================================
                // COLORES SEGÚN EL ESTADO DE LA FACTURA
                // ===========================================================
                for (int indice = 0; indice < facturas.Count; indice++)
                {
                    var factura = facturas[indice];

                    int filaFactura =
                        filaEncabezado + 1 + indice;

                    var rangoFila = hoja.Range(
                        filaFactura,
                        1,
                        filaFactura,
                        totalColumnas);

                    if (factura.Estado == "Anulada")
                    {
                        // Fondo rojo claro.
                        rangoFila.Style.Fill.BackgroundColor =
                            XLColor.FromHtml("#F8D7DA");

                        // Texto rojo oscuro.
                        rangoFila.Style.Font.FontColor =
                            XLColor.FromHtml("#842029");
                    }
                    else if (factura.Estado == "Vencida")
                    {
                        // Fondo amarillo claro.
                        rangoFila.Style.Fill.BackgroundColor =
                            XLColor.FromHtml("#FFF3CD");

                        // Texto amarillo oscuro para mejorar la lectura.
                        rangoFila.Style.Font.FontColor =
                            XLColor.FromHtml("#664D03");
                    }
                }
            }

            // ===========================================================
            // TOTALES FINANCIEROS
            // ===========================================================
            var facturasValidas = facturas
                .Where(f => f.Estado != "Anulada")
                .ToList();

            var ventas = facturasValidas
                .Where(f => f.TipoFactura == "Venta")
                .ToList();

            var compras = facturasValidas
                .Where(f => f.TipoFactura == "Compra")
                .ToList();

            decimal totalVentas = ventas.Sum(f => f.Total);
            decimal totalCobrado = ventas.Sum(f => f.MontoPagado);
            decimal saldoPorCobrar = ventas.Sum(f => f.SaldoPendiente);

            decimal totalCompras = compras.Sum(f => f.Total);
            decimal totalPagadoCompras = compras.Sum(f => f.MontoPagado);
            decimal saldoPorPagar = compras.Sum(f => f.SaldoPendiente);

            int filaTotales = fila + 1;

            // ===========================================================
            // FILA TOTAL DE VENTAS
            // ===========================================================
            hoja.Cell(filaTotales, 8).Value = "TOTALES DE VENTAS";

            hoja.Range(
                    filaTotales,
                    8,
                    filaTotales,
                    11)
                .Merge();

            hoja.Cell(filaTotales, 12).Value = totalVentas;
            hoja.Cell(filaTotales, 13).Value = totalCobrado;
            hoja.Cell(filaTotales, 14).Value = saldoPorCobrar;

            var rangoTotalVentas = hoja.Range(
                filaTotales,
                8,
                filaTotales,
                totalColumnas);

            rangoTotalVentas.Style.Fill.BackgroundColor =
                XLColor.FromHtml("#CFE5E8");

            rangoTotalVentas.Style.Font.Bold = true;

            rangoTotalVentas.Style.Font.FontColor =
                XLColor.FromHtml("#123F46");

            rangoTotalVentas.Style.Border.OutsideBorder =
                XLBorderStyleValues.Thin;

            rangoTotalVentas.Style.Border.InsideBorder =
                XLBorderStyleValues.Thin;

            rangoTotalVentas.Style.Alignment.Vertical =
                XLAlignmentVerticalValues.Center;

            rangoTotalVentas.Style.Alignment.Horizontal =
                XLAlignmentHorizontalValues.Right;

            // ===========================================================
            // FILA TOTAL DE COMPRAS
            // ===========================================================
            hoja.Cell(filaTotales + 1, 8).Value = "TOTALES DE COMPRAS";

            hoja.Range(
                    filaTotales + 1,
                    8,
                    filaTotales + 1,
                    11)
                .Merge();

            hoja.Cell(filaTotales + 1, 12).Value = totalCompras;
            hoja.Cell(filaTotales + 1, 13).Value = totalPagadoCompras;
            hoja.Cell(filaTotales + 1, 14).Value = saldoPorPagar;

            var rangoTotalCompras = hoja.Range(
                filaTotales + 1,
                8,
                filaTotales + 1,
                totalColumnas);

            rangoTotalCompras.Style.Fill.BackgroundColor =
                XLColor.FromHtml("#DCE9F8");

            rangoTotalCompras.Style.Font.Bold = true;

            rangoTotalCompras.Style.Font.FontColor =
                XLColor.FromHtml("#123F46");

            rangoTotalCompras.Style.Border.OutsideBorder =
                XLBorderStyleValues.Thin;

            rangoTotalCompras.Style.Border.InsideBorder =
                XLBorderStyleValues.Thin;

            rangoTotalCompras.Style.Alignment.Vertical =
                XLAlignmentVerticalValues.Center;

            rangoTotalCompras.Style.Alignment.Horizontal =
                XLAlignmentHorizontalValues.Right;

            // Formato monetario de los totales.
            hoja.Range(
                    filaTotales,
                    12,
                    filaTotales + 1,
                    14)
                .Style.NumberFormat.Format =
                    "₡ #,##0.00";

            // ===========================================================
            // RESUMEN DEL ESTADO DE LAS FACTURAS
            // ===========================================================
            int filaResumen = filaTotales + 4;

            hoja.Cell(filaResumen, 1).Value =
                "RESUMEN DEL ESTADO DE LAS FACTURAS";

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

            tituloResumen.Style.Font.Bold = true;
            tituloResumen.Style.Font.FontColor = XLColor.White;

            tituloResumen.Style.Fill.BackgroundColor =
                XLColor.FromHtml("#0F5C64");

            tituloResumen.Style.Alignment.Horizontal =
                XLAlignmentHorizontalValues.Center;

            tituloResumen.Style.Alignment.Vertical =
                XLAlignmentVerticalValues.Center;

            string[] encabezadosResumen =
            {
    "Ventas",
    "Compras",
    "Pagadas",
    "Pendientes",
    "Vencidas",
    "Anuladas"
};

            for (int columna = 0; columna < encabezadosResumen.Length; columna++)
            {
                hoja.Cell(
                    filaResumen + 1,
                    columna + 1).Value =
                        encabezadosResumen[columna];
            }

            hoja.Cell(filaResumen + 2, 1).Value =
                facturas.Count(f => f.TipoFactura == "Venta");

            hoja.Cell(filaResumen + 2, 2).Value =
                facturas.Count(f => f.TipoFactura == "Compra");

            hoja.Cell(filaResumen + 2, 3).Value =
                facturasValidas.Count(f => f.Estado == "Pagada");

            hoja.Cell(filaResumen + 2, 4).Value =
                facturasValidas.Count(f => f.Estado == "Pendiente");

            hoja.Cell(filaResumen + 2, 5).Value =
                facturasValidas.Count(f => f.Estado == "Vencida");

            hoja.Cell(filaResumen + 2, 6).Value =
                facturas.Count(f => f.Estado == "Anulada");

            var encabezadoResumen = hoja.Range(
                filaResumen + 1,
                1,
                filaResumen + 1,
                6);

            encabezadoResumen.Style.Font.Bold = true;

            encabezadoResumen.Style.Fill.BackgroundColor =
                XLColor.FromHtml("#D9EDEF");

            encabezadoResumen.Style.Alignment.Horizontal =
                XLAlignmentHorizontalValues.Center;

            var valoresResumen = hoja.Range(
                filaResumen + 2,
                1,
                filaResumen + 2,
                6);

            valoresResumen.Style.Alignment.Horizontal =
                XLAlignmentHorizontalValues.Center;

            valoresResumen.Style.Font.Bold = true;

            var rangoResumenCompleto = hoja.Range(
                filaResumen + 1,
                1,
                filaResumen + 2,
                6);

            rangoResumenCompleto.Style.Border.OutsideBorder =
                XLBorderStyleValues.Thin;

            rangoResumenCompleto.Style.Border.InsideBorder =
                XLBorderStyleValues.Thin;

            // ===========================================================
            // AJUSTE FINAL
            // ===========================================================
            hoja.Columns().AdjustToContents();

            // Evita columnas exageradamente anchas.
            foreach (var columna in hoja.ColumnsUsed())
            {
                if (columna.Width > 35)
                {
                    columna.Width = 35;
                }
            }

            hoja.SheetView.FreezeRows(filaEncabezado);

            using var stream = new MemoryStream();

            workbook.SaveAs(stream);

            string nombreArchivo =
                $"Reporte_Facturas_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";

            // Registra la exportación solamente después de haber construido
            // correctamente el archivo Excel.
            await RegistrarAuditoriaFactura(
                "Exportar Excel",
                0,
                "Se exportó el reporte de facturas a Excel. " +
                ConstruirDescripcionFiltrosReporte(
                    tipoFactura,
                    estado,
                    condicionPago,
                    clienteId,
                    proveedorId,
                    proyectoId,
                    anio,
                    mes,
                    fechaDesde,
                    fechaHasta)
            );

            return File(
                stream.ToArray(),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                nombreArchivo);
        }
        // GET: Facturas/DetaLle
        public async Task<IActionResult> Details(int? id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var factura = await _context.Facturas
     .Include(f => f.ClienteBeneficiario)
     .Include(f => f.Proveedor)
     .Include(f => f.CategoriaIngreso)
     .Include(f => f.CategoriaGasto)
     .Include(f => f.Proyecto)
     .Include(f => f.PresupuestoMensualRegistro)
     .Include(f => f.CentroCosto)
     .Include(f => f.FacturaDetalles)
     .FirstOrDefaultAsync(m => m.IdFactura == id);

            if (factura == null)
            {
                return NotFound();
            }

            return View(factura);
        }
        //cargara presupesutos aprobados 
        private void CargarPeriodosPresupuestariosFactura(
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
        // NUEVO
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

        private void CargarOpcionesProyectoPeriodoVenta(
    string? seleccionActual = null)
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

            var cultura = new CultureInfo("es-CR");

            var hoy = DateOnly.FromDateTime(DateTime.Today);
            var fechaMinima = hoy.AddMonths(-3);

            int anioMinimo = fechaMinima.Year;
            int mesMinimo = fechaMinima.Month;

            var opciones = new List<SelectListItem>();

            // ============================================================
            // 1. PERIODOS DISPONIBLES PARA VENTAS
            // Aprobados o Borrador, desde 3 meses atrás + todos los futuros
            // ============================================================

            var periodos = _context.PresupuestosMensuales
                .AsNoTracking()
                .Include(pm => pm.Proyecto)
          .Where(pm =>
    (
        pm.ProyectoId != null &&
        pm.Proyecto != null &&
        pm.Proyecto.Activo &&

        (
            pm.EstadoAprobacion == "Aprobado" ||
            pm.EstadoAprobacion == "Borrador"
        ) &&

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

    (
        idPresupuestoActual.HasValue &&
        pm.IdPresupuestoMensual ==
            idPresupuestoActual.Value
    ))
                .OrderByDescending(pm => pm.Anio)
                .ThenByDescending(pm => pm.Mes)
                .ThenBy(pm => pm.Proyecto!.Nombre)
                .ToList();

            foreach (var periodo in periodos)
            {
                string nombreMes = new DateTime(
                        periodo.Anio,
                        periodo.Mes,
                        1)
                    .ToString("MMMM", cultura);

                nombreMes =
                    char.ToUpper(nombreMes[0]) +
                    nombreMes.Substring(1);

                string valor =
                    $"PRES-{periodo.IdPresupuestoMensual}";

                opciones.Add(new SelectListItem
                {
                    Value = valor,

                    Text =
                        $"{periodo.Proyecto!.Nombre} — " +
                        $"{nombreMes} {periodo.Anio} " +
                        $"({periodo.EstadoAprobacion})",

                    Selected = valor == seleccionActual
                });
            }

            // ============================================================
            // 2. PROYECTOS SIN PERIODO UTILIZABLE
            // ============================================================

            var proyectosSinPeriodo = _context.Proyectos
                .AsNoTracking()
                .Where(p =>
                    p.Activo &&
                    !_context.PresupuestosMensuales.Any(pm =>
                        pm.ProyectoId == p.IdProyecto &&
                        (
                            pm.EstadoAprobacion == "Aprobado" ||
                            pm.EstadoAprobacion == "Borrador"
                        ) &&
// NUEVO
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
                .ToList();

            foreach (var proyecto in proyectosSinPeriodo)
            {
                string valor =
                    $"PROY-{proyecto.IdProyecto}";

                opciones.Add(new SelectListItem
                {
                    Value = valor,
                    Text =
                        $"{proyecto.Nombre} — Sin período presupuestario",
                    Selected = valor == seleccionActual
                });
            }

            ViewBag.OpcionesProyectoPeriodoVenta = opciones;
        }

        // GET: Facturas/Creaar
        public IActionResult Create()
        {
            ViewData["ClienteBeneficiarioId"] = new SelectList(
    _context.ClientesBeneficiarios.Where(c => c.Activo == true),
    "IdClienteBeneficiario",
    "Nombre"
);

            ViewData["ProveedorId"] = new SelectList(
                _context.Proveedores.Where(p => p.Activo == true),
                "IdProveedor",
                "Nombre"
            );


            var factura = new Factura
            {
                FechaEmision = DateOnly.FromDateTime(DateTime.Now),
                DescuentoTotal = 0,
                ImpuestoTotal = 0,
                Total = 0,
                SaldoPendiente = 0,
                Estado = "Pendiente",
                NumeroFactura = "Se generará automáticamente"
            };
            ViewData["CategoriaIngresoId"] = new SelectList(
          _context.CategoriasIngresos
              .Where(c => c.Activo == true)
              .OrderBy(c => c.Nombre),
          "IdCategoriaIngreso",
          "Nombre"
      );

            ViewData["CategoriaGastoId"] = new SelectList(
                _context.CategoriasGastos
                    .Where(c => c.Activo == true)
                    .OrderBy(c => c.Nombre),
                "IdCategoriaGasto",
                "Nombre"
            );

            ViewData["ProyectoId"] = new SelectList(
                _context.Proyectos
                    .Where(p => p.Activo == true)
                    .OrderBy(p => p.Nombre),
                "IdProyecto",
                "Nombre"
            );

            ViewData["CentroCostoId"] = new SelectList(
                _context.CentrosCostos
                    .Where(c => c.Activo == true)
                    .OrderBy(c => c.Nombre),
                "IdCentroCosto",
                "Nombre"
            );
            ViewData["ServicioCatalogoId"] = new SelectList(
                _context.FacturaDetalles
                .Where(s => s.FacturaId == null && s.Activo == true),
                   "IdFacturaDetalle",
                   "Nombre"
            );

            ViewBag.ServiciosCatalogo = _context.FacturaDetalles
                   .Where(s => s.FacturaId == null && s.Activo == true)
                   .OrderBy(s => s.Nombre)
                   .ToList();

            CargarPeriodosPresupuestariosFactura();
            CargarOpcionesProyectoPeriodoVenta();
            return View(factura);
        }

        // POST: Facturas/Crear
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Create(
    [Bind("IdFactura,TipoFactura,CondicionPago,DiasCredito,TipoImpuesto,ClienteBeneficiarioId,ProveedorId,FechaEmision,CategoriaIngresoId,CategoriaGastoId,ProyectoId,PresupuestoMensualId,CentroCostoId,ServicioCatalogoId,TipoIngreso,SubTotal,DescuentoTotal,Observacion")]
    Factura factura,
    string detalleFacturaJson,
    string? SeleccionProyectoPeriodoVenta)
        {
            List<DetalleFacturaJson> detallesFactura = new List<DetalleFacturaJson>();

if (!string.IsNullOrWhiteSpace(detalleFacturaJson))
{
    detallesFactura = JsonSerializer.Deserialize<List<DetalleFacturaJson>>(
        detalleFacturaJson,
        new JsonSerializerOptions
        {
            PropertyNameCaseInsensitive = true
        }
    ) ?? new List<DetalleFacturaJson>();
}

if (detallesFactura.Count == 0)
{
    ModelState.AddModelError("", "Debe agregar al menos una línea al detalle de la factura.");
}

// Validar tipo de factura
if (factura.TipoFactura != "Venta" && factura.TipoFactura != "Compra")
            {
                ModelState.AddModelError("TipoFactura", "Debe seleccionar Venta o Compra.");
            }

            // Venta exige cliente
            if (factura.TipoFactura == "Venta" && factura.ClienteBeneficiarioId == null)
            {
                ModelState.AddModelError("ClienteBeneficiarioId", "Debe seleccionar un cliente para la factura de venta.");
            }


            if (factura.TipoFactura == "Compra" && factura.CategoriaGastoId == null)
            {
                ModelState.AddModelError("CategoriaGastoId", "Debe seleccionar una categoría de gasto para la factura de compra.");
            }
            // Compra exige proveedor
            if (factura.TipoFactura == "Compra" &&
    factura.ProveedorId == null)
            {
                ModelState.AddModelError(
                    "ProveedorId",
                    "Debe seleccionar un proveedor para la factura de compra.");
            }
            // ============================================================
            // BLOQUEO POR CIERRE CONTABLE
            // APLICA A FACTURA DE VENTA Y FACTURA DE COMPRA
            // ============================================================
            bool periodoCerrado =
                await PeriodoContableCerradoAsync(factura.FechaEmision);

            if (periodoCerrado)
            {
                ModelState.AddModelError(
                    "FechaEmision",
                    factura.TipoFactura == "Venta"
                        ? "No se puede registrar la factura de venta porque el período contable seleccionado se encuentra cerrado."
                        : "No se puede registrar la factura de compra porque el período contable seleccionado se encuentra cerrado.");
            }

            // VALIDAR PERÍODO PRESUPUESTARIO EN FACTURA DE COMPRA
            PresupuestosMensuale? presupuestoSeleccionado = null;

            if (factura.TipoFactura == "Compra")
            {
                if (!factura.PresupuestoMensualId.HasValue)
                {
                    ModelState.AddModelError(
                        "PresupuestoMensualId",
                        "Debe seleccionar un proyecto y período presupuestario.");
                }
                else
                {
                    var hoy = DateOnly.FromDateTime(DateTime.Today);
                    var fechaMinima = hoy.AddMonths(-3);

                    int anioMinimo = fechaMinima.Year;
                    int mesMinimo = fechaMinima.Month;

                    presupuestoSeleccionado =
                        await _context.PresupuestosMensuales
                            .Include(pm => pm.Proyecto)
                            .FirstOrDefaultAsync(pm =>
                                pm.IdPresupuestoMensual ==
                                    factura.PresupuestoMensualId.Value &&
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
                        // Nunca confiamos en ProyectoId enviado por navegador.
                        factura.ProyectoId =
                            presupuestoSeleccionado.ProyectoId;
                    }
                }
            }
            else if (factura.TipoFactura == "Venta")
            {
                ModelState.Remove(nameof(factura.ProyectoId));
                ModelState.Remove(nameof(factura.PresupuestoMensualId));

                factura.ProyectoId = null;
                factura.PresupuestoMensualId = null;

                if (string.IsNullOrWhiteSpace(
                    SeleccionProyectoPeriodoVenta))
                {
                    ModelState.AddModelError(
                        "ProyectoId",
                        "Debe seleccionar un proyecto y período presupuestario.");

                }
            
            // ============================================================
            // VENTA ASOCIADA A UN PERIODO
            // ============================================================

            if (SeleccionProyectoPeriodoVenta.StartsWith("PRES-"))
                {
                    string valorId =
                        SeleccionProyectoPeriodoVenta.Substring(5);

                    if (!int.TryParse(
                        valorId,
                        out int idPresupuestoMensual))
                    {
                        ModelState.AddModelError(
                            "ProyectoId",
                            "La selección presupuestaria no es válida.");
                    }
                    else
                    {
                        var hoy = DateOnly.FromDateTime(DateTime.Today);
                        var fechaMinima = hoy.AddMonths(-3);

                        int anioMinimo = fechaMinima.Year;
                        int mesMinimo = fechaMinima.Month;

                        var presupuestoVenta =
                            await _context.PresupuestosMensuales
                                .AsNoTracking()
                                .Include(pm => pm.Proyecto)
                                .FirstOrDefaultAsync(pm =>
                                    pm.IdPresupuestoMensual ==
                                        idPresupuestoMensual &&
                                    pm.ProyectoId != null &&
                                    pm.Proyecto != null &&
                                    pm.Proyecto.Activo &&

                                    (
                                        pm.EstadoAprobacion == "Aprobado" ||
                                        pm.EstadoAprobacion == "Borrador"
                                    ) &&
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

                        if (presupuestoVenta == null)
                        {
                            ModelState.AddModelError(
                                "ProyectoId",
                                "El período presupuestario seleccionado ya no está disponible.");
                        }
                        else
                        {
                            factura.ProyectoId =
                                presupuestoVenta.ProyectoId;

                            factura.PresupuestoMensualId =
                                presupuestoVenta.IdPresupuestoMensual;
                        }
                    }
                }

                // ============================================================
                // VENTA PARA PROYECTO SIN PERIODO
                // ============================================================

                else if (
                    SeleccionProyectoPeriodoVenta.StartsWith("PROY-"))
                {
                    string valorId =
                        SeleccionProyectoPeriodoVenta.Substring(5);

                    if (!int.TryParse(
                        valorId,
                        out int idProyecto))
                    {
                        ModelState.AddModelError(
                            "ProyectoId",
                            "El proyecto seleccionado no es válido.");
                    }
                    else
                    {
                        bool proyectoActivo =
                            await _context.Proyectos.AnyAsync(p =>
                                p.IdProyecto == idProyecto &&
                                p.Activo);

                        if (!proyectoActivo)
                        {
                            ModelState.AddModelError(
                                "ProyectoId",
                                "El proyecto seleccionado está inactivo o no existe.");
                        }
                        else
                        {
                            factura.ProyectoId = idProyecto;
                            factura.PresupuestoMensualId = null;
                        }
                    }
                }
                else
                {
                    ModelState.AddModelError(
                        "ProyectoId",
                        "La selección de proyecto y período presupuestario no es válida.");
                }
            }
            // Fecha vencimiento no puede ser menor a emisión
            if (factura.FechaVencimiento != null && factura.FechaVencimiento < factura.FechaEmision)
            {
                ModelState.AddModelError("FechaVencimiento", "La fecha de vencimiento no puede ser menor que la fecha de emisión.");
            }

            // Descuento no puede ser mayor al subtotal
            if (factura.DescuentoTotal > factura.SubTotal)
            {
                ModelState.AddModelError("DescuentoTotal", "El descuento no puede ser mayor que el subtotal.");
            }
            // valida si es credito obligatorio poner dias de credito
            if (factura.CondicionPago == "Credito" && factura.DiasCredito == null)
            {
                ModelState.AddModelError("DiasCredito", "Debe seleccionar los días de crédito.");
            }

            factura.SubTotal = detallesFactura.Sum(d => d.SubtotalLinea);
            factura.DescuentoTotal = detallesFactura.Sum(d => d.Descuento);
            factura.ImpuestoTotal = detallesFactura.Sum(d => d.ImpuestoLinea);
            factura.Total = detallesFactura.Sum(d => d.TotalLinea);

            if (ModelState.IsValid)
            {
                // Limpiar relación que no corresponde según tipo
                if (factura.TipoFactura == "Venta")
                {
                    factura.ProveedorId = null;
                }
                else if (factura.TipoFactura == "Compra")
                {
                    factura.ClienteBeneficiarioId = null;
                }

                // Número automático

                string prefijo = factura.TipoFactura == "Venta" ? "FV" : "FC";

                int ultimoNumero = await _context.Facturas
                    .Where(f => f.TipoFactura == factura.TipoFactura)
                    .CountAsync();

                factura.NumeroFactura = $"{prefijo}-{(ultimoNumero + 1).ToString("D6")}";

            

                // Lógica de contado o crédito
                if (factura.CondicionPago == "Contado")
                {
                    factura.DiasCredito = null;
                    factura.FechaVencimiento = factura.FechaEmision;
                    factura.MontoPagado = factura.Total;
                    factura.SaldoPendiente = 0;
                    factura.Estado = "Pagada";
                }
                else if (factura.CondicionPago == "Credito")
                {
                    factura.MontoPagado = 0;
                    factura.SaldoPendiente = factura.Total;
                    factura.Estado = "Pendiente";

                    if (factura.DiasCredito.HasValue)
                    {
                        factura.FechaVencimiento = factura.FechaEmision.AddDays(factura.DiasCredito.Value);
                    }
                }
                // guardar datos creados en detalles factura
                _context.Add(factura);
                await _context.SaveChangesAsync();

                foreach (var linea in detallesFactura)
                {
                    var detalle = new FacturaDetalle
                    {
                        FacturaId = factura.IdFactura,
                        Descripcion = linea.ServicioNombre,
                        Nombre = linea.ServicioNombre,
                        Cantidad = linea.Cantidad,
                        PrecioUnitario = linea.PrecioUnitario,
                        Descuento = linea.Descuento,
                        Impuesto = linea.ImpuestoLinea,
                        TotalLinea = linea.TotalLinea,
                        Activo = true
                    };

                    _context.FacturaDetalles.Add(detalle);
                }

                await _context.SaveChangesAsync();
  
 // ==================== REGISTRO CONTABLE AUTOMÁTICO DE LA FACTURA===================================
                
                int usuarioId = HttpContext.Session.GetInt32("UsuarioId") ?? 1;

                // Venta contado
                if (factura.TipoFactura == "Venta" &&
                    factura.CondicionPago == "Contado")
                {
                    await _contabilidadService.RegistrarVentaContadoAsync(
                               factura,
                               usuarioId);
                }

                // Venta crédito
                if (factura.TipoFactura == "Venta" &&
                   (factura.CondicionPago == "Credito" ||
                    factura.CondicionPago == "Crédito"))
                {
                    await _contabilidadService.RegistrarVentaCreditoAsync(
                            factura,
                            usuarioId);
                }

                // Compra contado
                if (factura.TipoFactura == "Compra" &&
                    factura.CondicionPago == "Contado")
                {
                    await _contabilidadService.RegistrarCompraContadoAsync(
                        factura,
                        usuarioId);
                }

                // Compra crédito
                if (factura.TipoFactura == "Compra" &&
                   (factura.CondicionPago == "Credito" ||
                    factura.CondicionPago == "Crédito"))
                {
                    await _contabilidadService.RegistrarCompraCreditoAsync(
                              factura,
                              usuarioId);
                }

                // Integración: si la factura es venta de contado,
                // se genera automáticamente un ingreso relacionado con la factura.
                if (factura.TipoFactura == "Venta" && factura.CondicionPago == "Contado")
                {
                    bool ingresoYaExiste = await _context.Ingresos
                        .AnyAsync(i => i.FacturaId == factura.IdFactura);

                    if (!ingresoYaExiste)
                    {
                        var ingresoAutomatico = new Ingreso
                        {
                            FacturaId = factura.IdFactura,
                            Fecha = factura.FechaEmision,

                            CategoriaIngresoId = factura.CategoriaIngresoId.Value,

                            ProyectoId = factura.ProyectoId,

                            PresupuestoMensualId =
             factura.PresupuestoMensualId,

                            MotivoPendientePresupuestario =
             factura.PresupuestoMensualId.HasValue
                 ? null
                 : "Ingreso generado desde factura sin período presupuestario.",

                            CentroCostoId = factura.CentroCostoId,
                            ClienteBeneficiarioId = factura.ClienteBeneficiarioId,

                            TipoIngreso = factura.TipoIngreso ?? "Facturación",
                            Origen = "Factura",

                            Fuente = factura.NumeroFactura ?? "Factura de venta",
                            MontoEsperado = factura.Total,
                            MontoReal = factura.Total,
                            Diferencia = 0,
                            EsAnticipo = false,

                            Descripcion =
             "Ingreso generado automáticamente desde la factura " +
             factura.NumeroFactura,

                            Estado = "Pagado",
                            Observaciones = factura.Observacion,
                            Comprobante = factura.NumeroFactura,
                            Activo = true
                        };

                        _context.Ingresos.Add(ingresoAutomatico);
                        await _context.SaveChangesAsync();
                        // Después de crear el ingreso automático,
                        // se recalcula el presupuesto mensual y sus detalles.
                        await RecalcularPresupuestoPorFactura(factura);
                    }
                }
                // Integración: si la factura es venta a crédito,
                // se genera automáticamente una cuenta por cobrar.
                // Importante: NO se genera ingreso todavía.
                // El ingreso se generará cuando se registre el pago de la CxC.
                if (factura.TipoFactura == "Venta" && factura.CondicionPago == "Credito")
                {
                    bool cuentaYaExiste = await _context.CuentasPorCobrars
                        .AnyAsync(c => c.FacturaId == factura.IdFactura);

                    if (!cuentaYaExiste)
                    {
                        var cuentaPorCobrar = new CuentasPorCobrar
                        {
                            ClienteBeneficiarioId = factura.ClienteBeneficiarioId,
                            ProyectoId = factura.ProyectoId,
                            PresupuestoMensualId = factura.PresupuestoMensualId,
                            FacturaId = factura.IdFactura,
                            TipoDocumento = "Factura",
                            Origen = "Factura",
                            EstadoAutorizacion = "No aplica",
                            Concepto = "Cuenta por cobrar generada desde la factura " + factura.NumeroFactura,
                            NumeroDocumento = factura.NumeroFactura,
                            FechaEmision = factura.FechaEmision,
                            FechaVencimiento = factura.FechaVencimiento ?? factura.FechaEmision,
                            MontoOriginal = factura.Total,
                            MontoPagado = 0,
                            AnticipoAplicado = 0,
                            DiasCredito = factura.DiasCredito,
                            SaldoPendiente = factura.Total,
                            Estado = "Pendiente",
                            Observacion = factura.Observacion,
                            Activo = true
                        };

                        _context.CuentasPorCobrars.Add(cuentaPorCobrar);
                        await _context.SaveChangesAsync();
                    }
                }

                // Integración: si la factura es compra,
                // se genera automáticamente un gasto.
                // Si además es crédito, también se genera una cuenta por pagar. 
               

                if (factura.TipoFactura == "Compra")
                {
                    bool gastoYaExiste = await _context.Gastos
                        .AnyAsync(g => g.NumeroFactura == factura.NumeroFactura);

                    if (!gastoYaExiste)
                    {

                        // Calcula la foto histórica del presupuesto para el gasto automático.
                        // Esto permite que el Index de Gastos muestre presupuesto, gastado y saldo correctamente.
                        var totalGastadoMes =
      await _context.Gastos
          .Where(g =>
              g.CategoriaGastoId ==
                  factura.CategoriaGastoId &&
              g.PresupuestoMensualId ==
                  factura.PresupuestoMensualId &&
              g.Activo &&
              g.Estado == "Aprobado")
          .SumAsync(g =>
              g.MontoTotal ?? 0);

                        totalGastadoMes += factura.Total;

                        var categoria = await _context.CategoriasGastos
                            .FirstOrDefaultAsync(c => c.IdCategoriaGasto == factura.CategoriaGastoId);

                        decimal presupuestoMensual = categoria?.LimiteMensual ?? 0;
                        decimal saldoDisponible = presupuestoMensual - totalGastadoMes;

                        var gastoAutomatico = new Gasto
                        {
                            Fecha = factura.FechaEmision,
                            CategoriaGastoId = factura.CategoriaGastoId.Value,
                            ProyectoId = factura.ProyectoId,
                            PresupuestoMensualId = factura.PresupuestoMensualId,
                            ProveedorId = factura.ProveedorId,
                            CentroCostoId = factura.CentroCostoId,
                            NumeroFactura = factura.NumeroFactura,
                            Concepto = "Gasto generado automáticamente desde la factura " + factura.NumeroFactura,
                            Subtotal = factura.SubTotal,
                            Impuesto = factura.ImpuestoTotal,
                            Descuento = factura.DescuentoTotal ?? 0,
                            TipoImpuesto = factura.TipoImpuesto,
                            MontoTotal = factura.Total,
                            Observacion = factura.Observacion,

                            PresupuestoMensual = presupuestoMensual,
                            GastadoAcumulado = totalGastadoMes,
                            SaldoDisponible = saldoDisponible,
                            SuperaPresupuesto = presupuestoMensual > 0 && totalGastadoMes > presupuestoMensual,

                            Estado = "Aprobado",
                            Activo = true
                        };

                        _context.Gastos.Add(gastoAutomatico);
                        await _context.SaveChangesAsync();

                        await RecalcularPresupuestoPorFactura(factura);
                    }
                }

                if (factura.TipoFactura == "Compra" && factura.CondicionPago == "Credito")
                {
                    bool cuentaYaExiste = await _context.CuentasPorPagars
                        .AnyAsync(c => c.NumeroFactura == factura.NumeroFactura);

                    if (!cuentaYaExiste)
                    {
                        var cuentaPorPagar = new CuentasPorPagar
                        {
                            ProveedorId = factura.ProveedorId.Value,
                            ProyectoId = factura.ProyectoId,
                            PresupuestoMensualId = factura.PresupuestoMensualId,
                            CategoriaGastoId = factura.CategoriaGastoId,
                            CentroCostoId = factura.CentroCostoId,
                            Origen = "Factura",
                            EstadoAutorizacion = "No aplica",
                            TipoDocumento = "Factura",
                            NumeroDocumento = factura.NumeroFactura,
                            NumeroFactura = factura.NumeroFactura,
                            
                            Concepto = "Cuenta por pagar generada desde la factura " + factura.NumeroFactura,
                            FechaEmision = factura.FechaEmision,
                            FechaVencimiento = factura.FechaVencimiento ?? factura.FechaEmision,
                            DiasCredito = factura.DiasCredito,
                            SubTotal = factura.SubTotal,
                            Impuesto = factura.ImpuestoTotal,
                            Descuento = factura.DescuentoTotal ?? 0,
                            TipoImpuesto = factura.TipoImpuesto,
                            MontoOriginal = factura.Total,
                            MontoPagado = 0,
                            AnticipoAplicado = 0,
                            SaldoPendiente = factura.Total,
                            Estado = "Pendiente",
                            Observacion = factura.Observacion,
                            Activo = true
                        };

                        _context.CuentasPorPagars.Add(cuentaPorPagar);
                        await _context.SaveChangesAsync();
                    }
                }

                // Auditoría: registra la creación de la factura.
                await RegistrarAuditoriaFactura(
                    "Crear",
                    factura.IdFactura,
                    "Se registró la factura " + factura.NumeroFactura +
                    " de tipo " + factura.TipoFactura +
                    " con condición de pago " + factura.CondicionPago +
                    " por un total de ₡" + factura.Total.ToString("N2") + ".");

                TempData["MensajeExito"] = "Factura registrada correctamente.";
                return RedirectToAction(nameof(Index));
            }

            ViewData["ClienteBeneficiarioId"] = new SelectList(
                _context.ClientesBeneficiarios.Where(c => c.Activo == true),
                "IdClienteBeneficiario",
                "Nombre",
                factura.ClienteBeneficiarioId
            );

            ViewData["ProveedorId"] = new SelectList(
                _context.Proveedores.Where(p => p.Activo == true),
                "IdProveedor",
                "Nombre",
                factura.ProveedorId
            );
            ViewData["ServicioCatalogoId"] = new SelectList(
                _context.FacturaDetalles.Where(s => s.FacturaId == null && s.Activo == true),
                "IdFacturaDetalle",
                "Nombre",
                 factura.ServicioCatalogoId
             );
            ViewBag.ServiciosCatalogo = _context.FacturaDetalles
                    .Where(s => s.FacturaId == null && s.Activo == true)
                    .OrderBy(s => s.Nombre)
                    .ToList();

            ViewData["CategoriaIngresoId"] = new SelectList(
      _context.CategoriasIngresos
          .Where(c => c.Activo == true)
          .OrderBy(c => c.Nombre),
      "IdCategoriaIngreso",
      "Nombre",
      factura.CategoriaIngresoId
  );

            ViewData["CategoriaGastoId"] = new SelectList(
                _context.CategoriasGastos
                    .Where(c => c.Activo == true)
                    .OrderBy(c => c.Nombre),
                "IdCategoriaGasto",
                "Nombre",
                factura.CategoriaGastoId
            );

            ViewData["ProyectoId"] = new SelectList(
                _context.Proyectos
                    .Where(p => p.Activo == true)
                    .OrderBy(p => p.Nombre),
                "IdProyecto",
                "Nombre",
                factura.ProyectoId
            );

            ViewData["CentroCostoId"] = new SelectList(
                _context.CentrosCostos
                    .Where(c => c.Activo == true)
                    .OrderBy(c => c.Nombre),
                "IdCentroCosto",
                "Nombre",
                factura.CentroCostoId
            );

            CargarPeriodosPresupuestariosFactura(
    factura.PresupuestoMensualId);


            CargarOpcionesProyectoPeriodoVenta(
                SeleccionProyectoPeriodoVenta);
            ViewBag.DetalleFacturaJson =
                JsonSerializer.Serialize(
                    detallesFactura,
                    new JsonSerializerOptions
                    {
                        PropertyNamingPolicy = JsonNamingPolicy.CamelCase
                    });

            return View(factura);
        }
        // Recalcula el presupuesto mensual relacionado con la factura.
        // Se usa cuando una factura genera un ingreso o gasto automático.
        private async Task RecalcularPresupuestoPorFactura(Factura factura)
        {
            if (factura.ProyectoId == null)
            {
                return;
            }
            PresupuestosMensuale? presupuesto;

            // Las compras se imputan al período presupuestario
            // seleccionado explícitamente.
            if (factura.TipoFactura == "Compra" &&
                factura.PresupuestoMensualId.HasValue)
            {
                presupuesto =
                    await _context.PresupuestosMensuales
                        .Include(p => p.PresupuestoDetalles)
                        .FirstOrDefaultAsync(p =>
                            p.IdPresupuestoMensual ==
                                factura.PresupuestoMensualId.Value);
            }
            else
            {
                // Los ingresos siguen correspondiendo a la fecha
                // real en que se producen.
                presupuesto =
                    await _context.PresupuestosMensuales
                        .Include(p => p.PresupuestoDetalles)
                        .FirstOrDefaultAsync(p =>
                            p.ProyectoId == factura.ProyectoId &&
                            p.Anio == factura.FechaEmision.Year &&
                            p.Mes == factura.FechaEmision.Month);
            }

            if (presupuesto == null)
            {
                return;
            }

            var fechaInicio = new DateOnly(presupuesto.Anio, presupuesto.Mes, 1);
            var fechaFin = fechaInicio.AddMonths(1).AddDays(-1);

            // Recalcula cada detalle del presupuesto.
            foreach (var detalle in presupuesto.PresupuestoDetalles)
            {
                if (detalle.Tipo == "Ingreso")
                {
                    detalle.MontoEjecutado = await _context.Ingresos
                        .Where(i =>
                            i.ProyectoId == presupuesto.ProyectoId &&
                            i.CategoriaIngresoId == detalle.CategoriaIngresoId &&
                            i.Fecha >= fechaInicio &&
                            i.Fecha <= fechaFin &&
                            i.Activo == true &&
                            i.Estado != "Anulado")
                        .SumAsync(i => i.MontoReal ?? 0);
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
             g.MontoTotal ?? 0); ;
                }

                detalle.Diferencia =
                    detalle.MontoPlanificado - (detalle.MontoEjecutado ?? 0);
            }

            // Recalcula totales generales del presupuesto.
            presupuesto.TotalIngresoReal = await _context.Ingresos
                .Where(i =>
                    i.ProyectoId == presupuesto.ProyectoId &&
                    i.Fecha >= fechaInicio &&
                    i.Fecha <= fechaFin &&
                    i.Activo == true &&
                    i.Estado != "Anulado")
                .SumAsync(i => i.MontoReal ?? 0);

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

            presupuesto.DiferenciaIngreso =
                (presupuesto.MontoIngresadoPlanificado ?? 0) -
                (presupuesto.TotalIngresoReal ?? 0);

            presupuesto.DiferenciaGasto =
                (presupuesto.MontoGastoPlanificado ?? 0) -
                (presupuesto.TotalGastoReal ?? 0);

            presupuesto.SaldoDisponible =
                (presupuesto.TotalIngresoReal ?? 0) -
                (presupuesto.TotalGastoReal ?? 0);

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
        private void CargarPeriodosPresupuestariosFacturaEdit(
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

                            // El período debe estar abierto.
                            !_context.CierresContables.Any(c =>
                                c.Anio == pm.Anio &&
                                c.Mes == pm.Mes &&
                                c.Estado == "Cerrado") &&

                            // Mantiene la regla actual de 3 meses.
                            (
                                pm.Anio > anioMinimo ||
                                (
                                    pm.Anio == anioMinimo &&
                                    pm.Mes >= mesMinimo
                                )
                            )
                        )

                        // Conserva solamente el presupuesto histórico
                        // que ya tenía esta factura.
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
        // GET: Facturas/Edit/5
        public async Task<IActionResult> Edit(int? id, string? retorno = null)
        {
            if (id == null)
            {
                return NotFound();
            }

            var factura = await _context.Facturas
                .Include(f => f.ClienteBeneficiario)
                .Include(f => f.Proveedor)
                .Include(f => f.CategoriaIngreso)
                .Include(f => f.CategoriaGasto)
                .Include(f => f.Proyecto)
            
                .Include(f => f.PresupuestoMensualRegistro)
                .Include(f => f.CentroCosto)
                .Include(f => f.FacturaDetalles)
                .FirstOrDefaultAsync(f => f.IdFactura == id);

            if (factura == null)
            {
                return NotFound();
            }
            // ============================================================
            // BLOQUEO POR CIERRE CONTABLE
            // ============================================================
            bool periodoCerrado =
                await PeriodoContableCerradoAsync(factura.FechaEmision);

            if (periodoCerrado)
            {
                string tipo =
                    factura.TipoFactura == "Venta"
                        ? "venta"
                        : "compra";

                TempData["MensajeError"] =
                    $"No se puede modificar la factura de {tipo} porque el período " +
                    $"{factura.FechaEmision:MM/yyyy} se encuentra cerrado. " +
                    $"Para realizar correcciones, primero debe reabrirse el período.";

                return RedirectToAction(
                    nameof(Index),
                    new
                    {
                        anio = factura.FechaEmision.Year,
                        mes = factura.FechaEmision.Month
                    });
            }
            if (factura.Estado == "Anulada")
            {
                TempData["MensajeError"] = "No se puede modificar una factura anulada.";
                return RedirectToAction(nameof(Index));
            }
            // si se modifica si es contado o compra o vneta contado 
            if (factura.CondicionPago == "Credito" || factura.CondicionPago == "Crédito")
            {
                bool tienePagosActivos = false;

                if (factura.TipoFactura == "Compra")
                {
                    tienePagosActivos = await _context.PagosCuentaPorPagars
                        .AnyAsync(p =>
                            p.CuentaPorPagar.NumeroFactura == factura.NumeroFactura &&
                            p.Anulado == false);
                }

                if (factura.TipoFactura == "Venta")
                {
                    tienePagosActivos = await _context.PagosCuentaPorCobrars
                        .AnyAsync(p =>
                            p.CuentaPorCobrar.FacturaId == factura.IdFactura &&
                            p.Anulado == false);
                }

                if (tienePagosActivos)
                {
                    TempData["MensajeError"] =
                        "No se puede modificar esta factura porque tiene pagos activos. Primero debe anular los pagos relacionados.";

                    return RedirectToAction(nameof(Details), new { id = factura.IdFactura });
                }
            }

            RecargarCombosFactura(factura);
            if (factura.TipoFactura == "Venta")
            {
                string? seleccionActual = factura.PresupuestoMensualId.HasValue
                    ? $"PRES-{factura.PresupuestoMensualId.Value}"
                    : factura.ProyectoId.HasValue
                        ? $"PROY-{factura.ProyectoId.Value}"
                        : null;

                CargarOpcionesProyectoPeriodoVenta(seleccionActual);
            }
            ViewBag.Retorno = retorno;
            return View(factura);
        }


        // POST: Facturas/Edit/5
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Edit(
    int id,
    [Bind("IdFactura,CondicionPago,DiasCredito,Observacion,CategoriaIngresoId,CategoriaGastoId,ProyectoId,PresupuestoMensualId,CentroCostoId,TipoIngreso")]
    Factura factura,
    string? SeleccionProyectoPeriodoVenta,
    string? retorno)
        {
            if (id != factura.IdFactura)
            {
                return NotFound();
            }

            var facturaActual = await _context.Facturas
                .FirstOrDefaultAsync(f => f.IdFactura == id);

            if (facturaActual == null)
            {
                return NotFound();
            }
            // ============================================================
            // BLOQUEO POR CIERRE CONTABLE
            // PROTECCIÓN TAMBIÉN CONTRA POST DIRECTO
            // ============================================================
            bool periodoCerrado =
                await PeriodoContableCerradoAsync(facturaActual.FechaEmision);

            if (periodoCerrado)
            {
                string tipo =
                    facturaActual.TipoFactura == "Venta"
                        ? "venta"
                        : "compra";

                TempData["MensajeError"] =
                    $"No se puede modificar la factura de {tipo} porque el período " +
                    $"{facturaActual.FechaEmision:MM/yyyy} se encuentra cerrado. " +
                    $"Para realizar correcciones, primero debe reabrirse el período.";

                return RedirectToAction(
                    nameof(Index),
                    new
                    {
                        anio = facturaActual.FechaEmision.Year,
                        mes = facturaActual.FechaEmision.Month
                    });
            }
            if (facturaActual.Estado == "Anulada")
            {
                TempData["MensajeError"] = "No se puede modificar una factura anulada.";
                return RedirectToAction(nameof(Index));
            }

            PresupuestosMensuale? presupuestoSeleccionado = null;

            if (facturaActual.TipoFactura == "Compra")
            {
                if (!factura.PresupuestoMensualId.HasValue)
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
                   factura.PresupuestoMensualId.Value &&

               (
                   // Puede conservar el presupuesto histórico actual
                   pm.IdPresupuestoMensual ==
                       facturaActual.PresupuestoMensualId

                   ||

                   // O cambiar a un presupuesto actualmente válido
                   (
                       pm.EstadoAprobacion == "Aprobado" &&
                       pm.ProyectoId != null &&
                       pm.Proyecto != null &&
                       pm.Proyecto.Activo &&

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
                        // Proyecto derivado del período seleccionado.
                        factura.ProyectoId =
                            presupuestoSeleccionado.ProyectoId;
                    }
                }
            }
            else if (facturaActual.TipoFactura == "Venta")
            {
                ModelState.Remove(nameof(factura.ProyectoId));
                ModelState.Remove(nameof(factura.PresupuestoMensualId));

                factura.ProyectoId = null;
                factura.PresupuestoMensualId = null;

                if (string.IsNullOrWhiteSpace(SeleccionProyectoPeriodoVenta))
                {
                    ModelState.AddModelError(
                        "ProyectoId",
                        "Debe seleccionar un proyecto y período presupuestario.");
                }
                else if (SeleccionProyectoPeriodoVenta.StartsWith("PRES-"))
                {
                    string valorId =
                        SeleccionProyectoPeriodoVenta.Substring(5);

                    if (!int.TryParse(valorId, out int idPresupuestoMensual))
                    {
                        ModelState.AddModelError(
                            "ProyectoId",
                            "La selección presupuestaria no es válida.");
                    }
                    else
                    {
                        var hoy = DateOnly.FromDateTime(DateTime.Today);
                        var fechaMinima = hoy.AddMonths(-3);

                        int anioMinimo = fechaMinima.Year;
                        int mesMinimo = fechaMinima.Month;

                        var presupuestoVenta =
     await _context.PresupuestosMensuales
         .AsNoTracking()
         .Include(pm => pm.Proyecto)
         .FirstOrDefaultAsync(pm =>
             pm.IdPresupuestoMensual ==
                 idPresupuestoMensual &&

             (
                 // Puede mantener el presupuesto actual
                 pm.IdPresupuestoMensual ==
                     facturaActual.PresupuestoMensualId

                 ||

                 // O cambiar a uno actualmente disponible
                 (
                     pm.ProyectoId != null &&
                     pm.Proyecto != null &&
                     pm.Proyecto.Activo &&

                     (
                         pm.EstadoAprobacion == "Aprobado" ||
                         pm.EstadoAprobacion == "Borrador"
                     ) &&

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
             ));

                        if (presupuestoVenta == null)
                        {
                            ModelState.AddModelError(
                                "ProyectoId",
                                "El período presupuestario seleccionado ya no está disponible.");
                        }
                        else
                        {
                            factura.ProyectoId =
                                presupuestoVenta.ProyectoId;

                            factura.PresupuestoMensualId =
                                presupuestoVenta.IdPresupuestoMensual;
                        }
                    }
                }
                else if (SeleccionProyectoPeriodoVenta.StartsWith("PROY-"))
                {
                    string valorId =
                        SeleccionProyectoPeriodoVenta.Substring(5);

                    if (!int.TryParse(valorId, out int idProyecto))
                    {
                        ModelState.AddModelError(
                            "ProyectoId",
                            "El proyecto seleccionado no es válido.");
                    }
                    else
                    {
                        bool proyectoActivo =
                            await _context.Proyectos.AnyAsync(p =>
                                p.IdProyecto == idProyecto &&
                                p.Activo);

                        if (!proyectoActivo)
                        {
                            ModelState.AddModelError(
                                "ProyectoId",
                                "El proyecto seleccionado está inactivo o no existe.");
                        }
                        else
                        {
                            factura.ProyectoId = idProyecto;
                            factura.PresupuestoMensualId = null;
                        }
                    }
                }
                else
                {
                    ModelState.AddModelError(
                        "ProyectoId",
                        "La selección de proyecto y período presupuestario no es válida.");
                }
            }
            // si se modifica compra o venta contado 
            if (factura.CondicionPago == "Credito" || factura.CondicionPago == "Crédito")
            {
                bool tienePagosActivos = false;

                if (facturaActual.TipoFactura == "Compra")
                {
                    tienePagosActivos = await _context.PagosCuentaPorPagars
                        .AnyAsync(p =>
                            p.CuentaPorPagar.NumeroFactura ==
                                facturaActual.NumeroFactura &&
                            p.Anulado == false);
                }

                if (facturaActual.TipoFactura == "Venta")
                {
                    tienePagosActivos = await _context.PagosCuentaPorCobrars
                        .AnyAsync(p =>
                            p.CuentaPorCobrar.FacturaId ==
                                facturaActual.IdFactura &&
                            p.Anulado == false);
                }

                if (tienePagosActivos)
                {
                    TempData["MensajeError"] =
                        "No se puede modificar esta factura porque tiene pagos activos. Primero debe anular los pagos relacionados.";

                    return RedirectToAction(nameof(Details), new { id = factura.IdFactura });
                }
            }
            if (factura.CondicionPago == "Credito" && factura.DiasCredito == null)
            {
                ModelState.AddModelError("DiasCredito", "Debe seleccionar los días de crédito.");
            }
            
            ModelState.Remove("TipoFactura");
            ModelState.Remove("NumeroFactura");
            ModelState.Remove("ClienteBeneficiarioId");
            ModelState.Remove("PresupuestoMensualRegistro");
            ModelState.Remove("ProveedorId");
            ModelState.Remove("FechaEmision");
            ModelState.Remove("FechaVencimiento");
            ModelState.Remove("SubTotal");
            ModelState.Remove("DescuentoTotal");
            ModelState.Remove("TipoImpuesto");
            ModelState.Remove("ImpuestoTotal");
            ModelState.Remove("Total");
            ModelState.Remove("MontoPagado");
            ModelState.Remove("SaldoPendiente");
            ModelState.Remove("Estado");

            if (!ModelState.IsValid)
            {
                facturaActual.CondicionPago =
                    factura.CondicionPago;

                facturaActual.DiasCredito =
                    factura.DiasCredito;

                facturaActual.Observacion =
                    factura.Observacion;

                facturaActual.CategoriaIngresoId =
                    factura.CategoriaIngresoId;

                facturaActual.CategoriaGastoId =
                    factura.CategoriaGastoId;

                facturaActual.ProyectoId =
                    factura.ProyectoId;

                facturaActual.PresupuestoMensualId =
                    factura.PresupuestoMensualId;

                facturaActual.CentroCostoId =
                    factura.CentroCostoId;

                facturaActual.TipoIngreso =
                    factura.TipoIngreso;

                RecargarCombosFactura(
                    facturaActual);
                if (facturaActual.TipoFactura == "Venta")
                {
                    CargarOpcionesProyectoPeriodoVenta(
                        SeleccionProyectoPeriodoVenta);
                }
                return View(
                    facturaActual);
            }
            // Datos administrativos permitidos
            facturaActual.CondicionPago = factura.CondicionPago;
            facturaActual.DiasCredito = factura.CondicionPago == "Contado" ? null : factura.DiasCredito;
            facturaActual.Observacion = factura.Observacion;

            facturaActual.ProyectoId =
        factura.ProyectoId;

            facturaActual.PresupuestoMensualId =
    factura.PresupuestoMensualId;

            facturaActual.CentroCostoId =
                factura.CentroCostoId;

            if (facturaActual.TipoFactura == "Venta")
            {
                facturaActual.CategoriaIngresoId = factura.CategoriaIngresoId;
                facturaActual.TipoIngreso = factura.TipoIngreso;
                facturaActual.CategoriaGastoId = null;
            }

            if (facturaActual.TipoFactura == "Compra")
            {
                facturaActual.CategoriaGastoId = factura.CategoriaGastoId;
                facturaActual.CategoriaIngresoId = null;
                facturaActual.TipoIngreso = null;
            }

            // Recalcular fechas y estado según condición de pago
            if (facturaActual.CondicionPago == "Contado")
            {
                facturaActual.FechaVencimiento = facturaActual.FechaEmision;
                facturaActual.MontoPagado = facturaActual.Total;
                facturaActual.SaldoPendiente = 0;
                facturaActual.Estado = "Pagada";
            }
            else if (facturaActual.CondicionPago == "Credito" || facturaActual.CondicionPago == "Crédito")
            {
                if (facturaActual.DiasCredito.HasValue)
                {
                    facturaActual.FechaVencimiento = facturaActual.FechaEmision.AddDays(facturaActual.DiasCredito.Value);
                }

                // Si se cambia a crédito, la factura vuelve a quedar pendiente de pago.
                facturaActual.MontoPagado = 0;
                facturaActual.SaldoPendiente = facturaActual.Total;

                if (facturaActual.FechaVencimiento < DateOnly.FromDateTime(DateTime.Today)
                    && facturaActual.SaldoPendiente > 0)
                {
                    facturaActual.Estado = "Vencida";
                }
                else
                {
                    facturaActual.Estado = "Pendiente";
                }
            }

            await _context.SaveChangesAsync();
            // Sincronizar ingreso generado automáticamente
            if (facturaActual.TipoFactura == "Venta" &&
          facturaActual.CondicionPago.Trim() == "Contado")
            {
                var ingreso = await _context.Ingresos
                    .FirstOrDefaultAsync(i => i.FacturaId == facturaActual.IdFactura);

                if (ingreso != null)
                {
                    ingreso.Fecha = facturaActual.FechaEmision;
                    ingreso.CategoriaIngresoId = facturaActual.CategoriaIngresoId.Value;
                    ingreso.ProyectoId = facturaActual.ProyectoId;
                    ingreso.PresupuestoMensualId =
    facturaActual.PresupuestoMensualId;

                    ingreso.MotivoPendientePresupuestario =
                        facturaActual.PresupuestoMensualId.HasValue
                            ? null
                            : "Ingreso generado desde factura sin período presupuestario.";
                    ingreso.CentroCostoId = facturaActual.CentroCostoId;
                    ingreso.ClienteBeneficiarioId = facturaActual.ClienteBeneficiarioId;
                    ingreso.TipoIngreso = facturaActual.TipoIngreso ?? "Facturación";
                    ingreso.Fuente = facturaActual.NumeroFactura;
                    ingreso.MontoEsperado = facturaActual.Total;
                    ingreso.MontoReal = facturaActual.Total;
                    ingreso.Diferencia = 0;
                    ingreso.Descripcion = "Ingreso generado automáticamente desde la factura " + facturaActual.NumeroFactura;
                    ingreso.Comprobante = facturaActual.NumeroFactura;
                    ingreso.Estado = "Pagado";
                    ingreso.Activo = true;

                    await _context.SaveChangesAsync();

                    // Después de actualizar el ingreso relacionado,
                    // se recalcula el presupuesto y el detalle.
                    await RecalcularPresupuestoPorFactura(facturaActual);
                }
            }
            // Sincronizar gasto generado automáticamente desde factura de compra.
            if (facturaActual.TipoFactura == "Compra")
            {
                var gasto = await _context.Gastos
                    .FirstOrDefaultAsync(g => g.NumeroFactura == facturaActual.NumeroFactura);

                if (gasto != null)
                {
                    gasto.Fecha = facturaActual.FechaEmision;
                    gasto.CategoriaGastoId = facturaActual.CategoriaGastoId.Value;
                    gasto.ProyectoId = facturaActual.ProyectoId;
                    gasto.PresupuestoMensualId =
    facturaActual.PresupuestoMensualId;
                    gasto.CentroCostoId = facturaActual.CentroCostoId;
                    gasto.ProveedorId = facturaActual.ProveedorId;
                    gasto.Subtotal = facturaActual.SubTotal;
                    gasto.Impuesto = facturaActual.ImpuestoTotal;
                    gasto.Descuento = facturaActual.DescuentoTotal ?? 0;
                    gasto.TipoImpuesto = facturaActual.TipoImpuesto;
                    gasto.MontoTotal = facturaActual.Total;
                    gasto.Observacion = facturaActual.Observacion;
                    gasto.Activo = true;

                    // Recalcula la foto presupuestaria de la categoría del gasto.
                    var totalGastadoMes =
       await _context.Gastos
           .Where(g =>
               g.CategoriaGastoId ==
                   gasto.CategoriaGastoId &&
               g.PresupuestoMensualId ==
                   gasto.PresupuestoMensualId &&
               g.Activo &&
               g.Estado == "Aprobado" &&
               g.IdGasto != gasto.IdGasto)
           .SumAsync(g =>
               g.MontoTotal ?? 0);


                    var categoria = await _context.CategoriasGastos
                        .FirstOrDefaultAsync(c => c.IdCategoriaGasto == gasto.CategoriaGastoId);

                    decimal presupuestoMensual = categoria?.LimiteMensual ?? 0;
                    decimal saldoDisponible = presupuestoMensual - totalGastadoMes;

                    gasto.PresupuestoMensual = presupuestoMensual;
                    gasto.GastadoAcumulado = totalGastadoMes;
                    gasto.SaldoDisponible = saldoDisponible;
                    gasto.SuperaPresupuesto = presupuestoMensual > 0 && totalGastadoMes > presupuestoMensual;

                    await _context.SaveChangesAsync();

                    // Después de actualizar el gasto, se recalcula el presupuesto mensual.
                    await RecalcularPresupuestoPorFactura(facturaActual);
                }
            }
            // Sincronizar cuenta por cobrar generada desde factura de venta a crédito.
            if (facturaActual.TipoFactura == "Venta" &&
                (facturaActual.CondicionPago == "Credito" || facturaActual.CondicionPago == "Crédito"))
            {
                var cuentaPorCobrar = await _context.CuentasPorCobrars
                    .FirstOrDefaultAsync(c => c.FacturaId == facturaActual.IdFactura);

                if (cuentaPorCobrar != null)
                {
                    cuentaPorCobrar.ClienteBeneficiarioId = facturaActual.ClienteBeneficiarioId;
                    cuentaPorCobrar.ProyectoId = facturaActual.ProyectoId;
                    cuentaPorCobrar.PresupuestoMensualId =
    facturaActual.PresupuestoMensualId;
                    cuentaPorCobrar.TipoDocumento = "Factura";
                    cuentaPorCobrar.NumeroDocumento = facturaActual.NumeroFactura;
                    cuentaPorCobrar.Concepto = "Cuenta por cobrar generada desde la factura " + facturaActual.NumeroFactura;
                    cuentaPorCobrar.FechaEmision = facturaActual.FechaEmision;
                    cuentaPorCobrar.FechaVencimiento = facturaActual.FechaVencimiento ?? facturaActual.FechaEmision;
                    cuentaPorCobrar.DiasCredito = facturaActual.DiasCredito;
                    cuentaPorCobrar.MontoOriginal = facturaActual.Total;
                    cuentaPorCobrar.SaldoPendiente =
                        facturaActual.Total -
                        (cuentaPorCobrar.MontoPagado ?? 0) -
                        (cuentaPorCobrar.AnticipoAplicado ?? 0);

                    if (cuentaPorCobrar.SaldoPendiente <= 0)
                    {
                        cuentaPorCobrar.SaldoPendiente = 0;
                        cuentaPorCobrar.Estado = "Cancelada";
                    }
                    else if ((cuentaPorCobrar.MontoPagado ?? 0) > 0)
                    {
                        cuentaPorCobrar.Estado = "Parcialmente pagada";
                    }
                    else if (cuentaPorCobrar.FechaVencimiento < DateOnly.FromDateTime(DateTime.Today))
                    {
                        cuentaPorCobrar.Estado = "Vencida";
                    }
                    else
                    {
                        cuentaPorCobrar.Estado = "Pendiente";
                    }

                    cuentaPorCobrar.Observacion = facturaActual.Observacion;
                    cuentaPorCobrar.Activo = true;

                    await _context.SaveChangesAsync();
                }
            }

            // Sincronizar cuenta por pagar generada desde factura de compra a crédito.
            if (facturaActual.TipoFactura == "Compra" &&
                (facturaActual.CondicionPago == "Credito" || facturaActual.CondicionPago == "Crédito"))
            {
                var cuentaPorPagar = await _context.CuentasPorPagars
                    .FirstOrDefaultAsync(c => c.NumeroFactura == facturaActual.NumeroFactura);

                if (cuentaPorPagar != null)
                {
                    cuentaPorPagar.ProveedorId = facturaActual.ProveedorId.Value;
                    cuentaPorPagar.ProyectoId = facturaActual.ProyectoId;
                    cuentaPorPagar.PresupuestoMensualId =
    facturaActual.PresupuestoMensualId;
                    cuentaPorPagar.CategoriaGastoId = facturaActual.CategoriaGastoId;
                    cuentaPorPagar.CentroCostoId = facturaActual.CentroCostoId;
                    cuentaPorPagar.TipoDocumento = "Factura";
                    cuentaPorPagar.NumeroDocumento = facturaActual.NumeroFactura;
                    cuentaPorPagar.NumeroFactura = facturaActual.NumeroFactura;
                    cuentaPorPagar.Concepto = "Cuenta por pagar generada desde la factura " + facturaActual.NumeroFactura;
                    cuentaPorPagar.FechaEmision = facturaActual.FechaEmision;
                    cuentaPorPagar.FechaVencimiento = facturaActual.FechaVencimiento ?? facturaActual.FechaEmision;
                    cuentaPorPagar.DiasCredito = facturaActual.DiasCredito;
                    cuentaPorPagar.SubTotal = facturaActual.SubTotal;
                    cuentaPorPagar.Impuesto = facturaActual.ImpuestoTotal;
                    cuentaPorPagar.Descuento = facturaActual.DescuentoTotal ?? 0;
                    cuentaPorPagar.TipoImpuesto = facturaActual.TipoImpuesto;
                    cuentaPorPagar.MontoOriginal = facturaActual.Total;
                    cuentaPorPagar.SaldoPendiente =
                        facturaActual.Total -
                        (cuentaPorPagar.MontoPagado ?? 0) -
                        (cuentaPorPagar.AnticipoAplicado ?? 0);

                    if (cuentaPorPagar.SaldoPendiente <= 0)
                    {
                        cuentaPorPagar.SaldoPendiente = 0;
                        cuentaPorPagar.Estado = "Cancelada";
                    }
                    else if ((cuentaPorPagar.MontoPagado ?? 0) > 0)
                    {
                        cuentaPorPagar.Estado = "Parcial";
                    }
                    else if (cuentaPorPagar.FechaVencimiento < DateOnly.FromDateTime(DateTime.Today))
                    {
                        cuentaPorPagar.Estado = "Vencida";
                    }
                    else
                    {
                        cuentaPorPagar.Estado = "Pendiente";
                    }

                    cuentaPorPagar.Observacion = facturaActual.Observacion;
                    cuentaPorPagar.Activo = true;

                    await _context.SaveChangesAsync();
                }
            }
            // Sincronizar ingresos generados por pagos de CxC
            // cuando la factura de venta a crédito cambia de proyecto, categoría o centro de costo.
            if (facturaActual.TipoFactura == "Venta" &&
                (facturaActual.CondicionPago == "Credito" || facturaActual.CondicionPago == "Crédito"))
            {
                var ingresosCxC = await _context.Ingresos
                    .Where(i => i.FacturaId == facturaActual.IdFactura
                        && i.Activo == true
                        && i.Comprobante != null
                        && i.Comprobante.StartsWith("PAGO-CXC-"))
                    .ToListAsync();

                foreach (var ingreso in ingresosCxC)
                {
                    ingreso.CategoriaIngresoId = facturaActual.CategoriaIngresoId.Value;
                    ingreso.ProyectoId = facturaActual.ProyectoId;
                    ingreso.PresupuestoMensualId =
    facturaActual.PresupuestoMensualId;

                    ingreso.MotivoPendientePresupuestario =
                        facturaActual.PresupuestoMensualId.HasValue
                            ? null
                            : "Ingreso generado desde CxC sin período presupuestario.";
                    ingreso.CentroCostoId = facturaActual.CentroCostoId;
                    ingreso.ClienteBeneficiarioId = facturaActual.ClienteBeneficiarioId;
                    ingreso.TipoIngreso = facturaActual.TipoIngreso ?? "Cobro CxC";
                    ingreso.Fuente = facturaActual.NumeroFactura;
                    ingreso.Descripcion = "Ingreso generado automáticamente por pago de CxC " + facturaActual.NumeroFactura;
                    ingreso.Observaciones = facturaActual.Observacion;
                }

                await _context.SaveChangesAsync();    
            }

                // Después de sincronizar los ingresos de la venta crédito,
                // se recalcula el presupuesto mensual relacionado.
                await RecalcularPresupuestoPorFactura(facturaActual);

                // Reversa el asiento contable anterior de la factura.
                int usuarioId = HttpContext.Session.GetInt32("UsuarioId") ?? 1;   
                await _contabilidadService.ReversarMovimientosAsync(
                    "Facturas",
                    facturaActual.IdFactura,
                    usuarioId,
                    "Reversión por modificación de factura " + facturaActual.NumeroFactura);

             // generar asiento nuevo
            if (facturaActual.TipoFactura == "Venta" &&
    facturaActual.CondicionPago == "Contado")
            {
                await _contabilidadService.RegistrarVentaContadoAsync(facturaActual, usuarioId);
            }

            if (facturaActual.TipoFactura == "Venta" &&
               (facturaActual.CondicionPago == "Credito" ||
                facturaActual.CondicionPago == "Crédito"))
            {
                await _contabilidadService.RegistrarVentaCreditoAsync(facturaActual, usuarioId);
            }

            if (facturaActual.TipoFactura == "Compra" &&
                facturaActual.CondicionPago == "Contado")
            {
                await _contabilidadService.RegistrarCompraContadoAsync(facturaActual, usuarioId);
            }

            if (facturaActual.TipoFactura == "Compra" &&
               (facturaActual.CondicionPago == "Credito" ||
                facturaActual.CondicionPago == "Crédito"))
            {
                await _contabilidadService.RegistrarCompraCreditoAsync(facturaActual, usuarioId);
            }


            // Auditoría: registra la modificación de la factura.
            await RegistrarAuditoriaFactura(
                "Modificar",
                facturaActual.IdFactura,
                "Se modificó la factura " + facturaActual.NumeroFactura +
                " de tipo " + facturaActual.TipoFactura +
                " con condición de pago " + facturaActual.CondicionPago +
                ".");
            TempData["MensajeExito"] = "Factura modificada correctamente.";

            if (retorno == "Ingresos")
            {
                return RedirectToAction(
                    "Index",
                    "Ingresos",
                    new { pendientePresupuesto = true });
            }

            return RedirectToAction(
                nameof(Details),
                new { id = facturaActual.IdFactura });
        }

        private void RecargarCombosFactura(Factura factura)
        {
            ViewData["ClienteBeneficiarioId"] = new SelectList(
                _context.ClientesBeneficiarios
                    .Where(c =>
                        c.Activo == true ||
                        c.IdClienteBeneficiario == factura.ClienteBeneficiarioId)
                    .OrderBy(c => c.Nombre),
                "IdClienteBeneficiario",
                "Nombre",
                factura.ClienteBeneficiarioId
            );

            ViewData["ProveedorId"] = new SelectList(
                _context.Proveedores
                    .Where(p =>
                        p.Activo == true ||
                        p.IdProveedor == factura.ProveedorId)
                    .OrderBy(p => p.Nombre),
                "IdProveedor",
                "Nombre",
                factura.ProveedorId
            );

            ViewData["CategoriaIngresoId"] = new SelectList(
                _context.CategoriasIngresos
                    .Where(c =>
                        c.Activo == true ||
                        c.IdCategoriaIngreso == factura.CategoriaIngresoId)
                    .OrderBy(c => c.Nombre),
                "IdCategoriaIngreso",
                "Nombre",
                factura.CategoriaIngresoId
            );

            ViewData["CategoriaGastoId"] = new SelectList(
                _context.CategoriasGastos
                    .Where(c =>
                        c.Activo == true ||
                        c.IdCategoriaGasto == factura.CategoriaGastoId)
                    .OrderBy(c => c.Nombre),
                "IdCategoriaGasto",
                "Nombre",
                factura.CategoriaGastoId
            );

            ViewData["ProyectoId"] = new SelectList(
                _context.Proyectos
                    .Where(p =>
                        p.Activo == true ||
                        p.IdProyecto == factura.ProyectoId)
                    .OrderBy(p => p.Nombre),
                "IdProyecto",
                "Nombre",
                factura.ProyectoId
            );

            ViewData["CentroCostoId"] = new SelectList(
                _context.CentrosCostos
                    .Where(c =>
                        c.Activo == true ||
                        c.IdCentroCosto == factura.CentroCostoId)
                    .OrderBy(c => c.Nombre),
                "IdCentroCosto",
                "Nombre",
                factura.CentroCostoId
            );
            if (factura.TipoFactura == "Compra")
            {
                CargarPeriodosPresupuestariosFacturaEdit(
                    factura.PresupuestoMensualId);
            }
        }
        // GET: Facturas/anular/5
        public async Task<IActionResult> Delete(int? id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var factura = await _context.Facturas
                .Include(f => f.ClienteBeneficiario)
                .Include(f => f.Proveedor)
                .Include(f => f.CategoriaIngreso)
                .Include(f => f.CategoriaGasto)
                .Include(f => f.Proyecto)
                .Include(f => f.PresupuestoMensualRegistro)
                .Include(f => f.CentroCosto)
                .Include(f => f.FacturaDetalles)
                .FirstOrDefaultAsync(m => m.IdFactura == id);

            if (factura == null)
            {
                return NotFound();
            }
            if (factura.Estado == "Anulada")
            {
                TempData["MensajeError"] = "La factura ya se encuentra anulada.";
                return RedirectToAction(nameof(Index));
            }

            // ============================================================
            // BLOQUEO POR CIERRE CONTABLE
            // ============================================================
            bool periodoCerrado =
                await PeriodoContableCerradoAsync(factura.FechaEmision);

            if (periodoCerrado)
            {
                string tipo =
                    factura.TipoFactura == "Venta"
                        ? "venta"
                        : "compra";

                TempData["MensajeError"] =
                    $"No se puede anular la factura de {tipo} porque el período " +
                    $"{factura.FechaEmision:MM/yyyy} se encuentra cerrado. " +
                    $"Para realizar correcciones, primero debe reabrirse el período.";

                return RedirectToAction(
                    nameof(Index),
                    new
                    {
                        anio = factura.FechaEmision.Year,
                        mes = factura.FechaEmision.Month
                    });
            }

            // si se anula si es venbta contado o compra contado 
            if (factura.CondicionPago == "Credito" || factura.CondicionPago == "Crédito")
            {
                bool tienePagosActivos = false;

                if (factura.TipoFactura == "Compra")
                {
                    tienePagosActivos = await _context.PagosCuentaPorPagars
                        .AnyAsync(p =>
                            p.CuentaPorPagar.NumeroFactura == factura.NumeroFactura &&
                            p.Anulado == false);
                }

                if (factura.TipoFactura == "Venta")
                {
                    tienePagosActivos = await _context.PagosCuentaPorCobrars
                        .AnyAsync(p =>
                            p.CuentaPorCobrar.FacturaId == factura.IdFactura &&
                            p.Anulado == false);
                }

                if (tienePagosActivos)
                {
                    TempData["MensajeError"] =
                        "No se puede anular esta factura porque tiene pagos activos. Primero debe anular los pagos relacionados.";

                    return RedirectToAction(nameof(Details), new { id = factura.IdFactura });
                }
            }

            return View(factura);
        }
        // POST: Facturas/anular
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> DeleteConfirmed(int id)
        {
            var factura = await _context.Facturas
                .FirstOrDefaultAsync(f => f.IdFactura == id);

            if (factura == null)
            {
                return NotFound();
            }
            // ============================================================
            // BLOQUEO POR CIERRE CONTABLE
            // PROTECCIÓN CONTRA POST DIRECTO
            // ============================================================
            bool periodoCerrado =
                await PeriodoContableCerradoAsync(factura.FechaEmision);

            if (periodoCerrado)
            {
                string tipo =
                    factura.TipoFactura == "Venta"
                        ? "venta"
                        : "compra";

                TempData["MensajeError"] =
                    $"No se puede anular la factura de {tipo} porque el período " +
                    $"{factura.FechaEmision:MM/yyyy} se encuentra cerrado. " +
                    $"Para realizar correcciones, primero debe reabrirse el período.";

                return RedirectToAction(
                    nameof(Index),
                    new
                    {
                        anio = factura.FechaEmision.Year,
                        mes = factura.FechaEmision.Month
                    });
            }
            if (factura.Estado == "Anulada")
            {
                TempData["MensajeError"] = "La factura ya se encuentra anulada.";
                return RedirectToAction(nameof(Index));
            }
            

            // Anula la factura principal.
            factura.Estado = "Anulada";
            factura.SaldoPendiente = 0;
            factura.MontoPagado = 0;

            await _context.SaveChangesAsync();

            // =====================================================
            // 1. ANULAR VENTA CONTADO
            // Factura venta contado genera ingreso automático.
            // Al anular la factura, también se anula el ingreso.
           
            if (factura.TipoFactura == "Venta" && factura.CondicionPago == "Contado")
            {
                var ingreso = await _context.Ingresos
                    .FirstOrDefaultAsync(i => i.FacturaId == factura.IdFactura);

                if (ingreso != null)
                {
                    ingreso.Estado = "Anulado";
                    ingreso.Activo = false;
                    ingreso.Observaciones = "Ingreso anulado automáticamente por anulación de la factura " + factura.NumeroFactura;

                    await _context.SaveChangesAsync();
                }

                await RecalcularPresupuestoPorFactura(factura);
            }

            // =====================================================
            // 2. ANULAR VENTA CRÉDITO
            // Factura venta crédito genera CxC.
            // Los ingresos reales vienen de los pagos de CxC.
            // Por eso se anula la CxC y los ingresos generados por pagos.

            if (factura.TipoFactura == "Compra" &&
    (factura.CondicionPago == "Credito" || factura.CondicionPago == "Crédito"))
            {
                bool tienePagosCxP = await _context.PagosCuentaPorPagars
                    .AnyAsync(p =>
                        p.CuentaPorPagar.NumeroFactura == factura.NumeroFactura &&
                        p.Anulado == false);

                if (tienePagosCxP)
                {
                    TempData["MensajeError"] =
                        "No se puede anular esta factura porque la cuenta por pagar tiene pagos activos. Primero debe anular los pagos.";

                    return RedirectToAction(nameof(Index));
                }
            }

            if (factura.TipoFactura == "Venta" &&
                (factura.CondicionPago == "Credito" || factura.CondicionPago == "Crédito"))
            {
                var cuentaPorCobrar = await _context.CuentasPorCobrars
                    .FirstOrDefaultAsync(c => c.FacturaId == factura.IdFactura);

                if (cuentaPorCobrar != null)
                {
                    // Anula la cuenta por cobrar generada por la factura.
                    cuentaPorCobrar.Estado = "Anulada";
                    cuentaPorCobrar.Activo = false;
                    cuentaPorCobrar.MontoPagado = 0;
                    cuentaPorCobrar.SaldoPendiente = 0;
                    cuentaPorCobrar.Observacion = "Cuenta por cobrar anulada automáticamente por anulación de la factura " + factura.NumeroFactura;

                    // Anula todos los pagos registrados sobre esa cuenta.
                    var pagosCxC = await _context.PagosCuentaPorCobrars
                        .Where(p => p.CuentaPorCobrarId == cuentaPorCobrar.IdCuentaPorCobrar
                                    && p.Anulado == false)
                        .ToListAsync();

                    foreach (var pago in pagosCxC)
                    {
                        pago.Anulado = true;
                        pago.MotivoAnulacion = "Pago anulado automáticamente por anulación de la factura " + factura.NumeroFactura;
                    }

                    await _context.SaveChangesAsync();
                }

                // Anula ingresos generados por pagos de esa CxC.
                var ingresosCxC = await _context.Ingresos
      .Where(i => i.FacturaId == factura.IdFactura
          && i.Comprobante != null
          && i.Comprobante.StartsWith("PAGO-CXC-"))
      .ToListAsync();

                foreach (var ingreso in ingresosCxC)
                {
                    ingreso.Estado = "Anulado";
                    ingreso.Activo = false;
                    ingreso.Observaciones = "Ingreso anulado automáticamente por anulación de la factura " + factura.NumeroFactura;
                }

                await _context.SaveChangesAsync();

                await RecalcularPresupuestoPorFactura(factura);
            }

            // =====================================================
            // 3. ANULAR COMPRA CONTADO / COMPRA CRÉDITO
            // Toda factura de compra genera gasto automático.
            // Al anular la factura, se anula el gasto.
            if (factura.TipoFactura == "Compra")
            {
                var gasto = await _context.Gastos
                    .FirstOrDefaultAsync(g => g.NumeroFactura == factura.NumeroFactura);

                if (gasto != null)
                {
                    gasto.Estado = "Anulado";
                    gasto.Activo = false;
                    gasto.Observacion = "Gasto anulado automáticamente por anulación de la factura " + factura.NumeroFactura;

                    await _context.SaveChangesAsync();
                }

                await RecalcularPresupuestoPorFactura(factura);
            }

            // =====================================================
            // 4. ANULAR COMPRA CRÉDITO
            // Factura compra crédito genera CxP.
            // Al anular la factura, también se anula la cuenta por pagar.
            if (factura.TipoFactura == "Compra" &&
                (factura.CondicionPago == "Credito" || factura.CondicionPago == "Crédito"))
            {
                var cuentaPorPagar = await _context.CuentasPorPagars
    .FirstOrDefaultAsync(c =>
        c.NumeroFactura == factura.NumeroFactura ||
        c.NumeroDocumento == factura.NumeroFactura);

                if (cuentaPorPagar != null)
                {
                    // Anula la cuenta por pagar generada por la factura.
                    cuentaPorPagar.Estado = "Anulada";
                    cuentaPorPagar.Activo = false;
                    cuentaPorPagar.MontoPagado = 0;
                    cuentaPorPagar.SaldoPendiente = 0;
                    cuentaPorPagar.Observacion = "Cuenta por pagar anulada automáticamente por anulación de la factura " + factura.NumeroFactura;

                    // Anula todos los pagos registrados sobre esa cuenta.
                    var pagosCxP = await _context.PagosCuentaPorPagars
                        .Where(p => p.CuentaPorPagarId == cuentaPorPagar.IdCuentaPorPagar
                                    && p.Anulado == false)
                        .ToListAsync();

                    foreach (var pago in pagosCxP)
                    {
                        pago.Anulado = true;
                        pago.MotivoAnulacion = "Pago anulado automáticamente por anulación de la factura " + factura.NumeroFactura;
                    }

                    await _context.SaveChangesAsync();
                }

                await RecalcularPresupuestoPorFactura(factura);
            }

            // =====================================================
            // 5. AUDITORÍA
            // Registra la anulación de la factura.
            // Ajusta los nombres de campos si tu modelo Auditoria usa otros.

            int usuarioId = HttpContext.Session.GetInt32("UsuarioId") ?? 1;

            // Reversa los movimientos contables generados por la factura.
            // La reversión se hace desde el módulo de origen, no desde Movimientos Contables.
            await _contabilidadService.ReversarMovimientosAsync(
                "Facturas",
                factura.IdFactura,
                usuarioId,
                "Reversión por anulación de factura " + factura.NumeroFactura);


            var auditoria = new Auditorium
            {
                UsuarioId = usuarioId,
                Tabla = "Facturas",
                RegistroId = factura.IdFactura,
                Accion = "Anulación de factura",
                Descripcion = "Se anuló la factura " + factura.NumeroFactura +
                              " de tipo " + factura.TipoFactura +
                              " con condición de pago " + factura.CondicionPago,
                Fecha = DateTime.Now
            };

            _context.Auditoria.Add(auditoria);
            await _context.SaveChangesAsync();

            TempData["MensajeExito"] = "Factura anulada correctamente y registros relacionados sincronizados.";

            return RedirectToAction(nameof(Index));
        }

        // GET: Facturas/Imprimir/5
        public async Task<IActionResult> Imprimir(int? id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var factura = await _context.Facturas
                .Include(f => f.ClienteBeneficiario)
                .Include(f => f.Proveedor)
                .Include(f => f.CategoriaIngreso)
                .Include(f => f.CategoriaGasto)
                .Include(f => f.Proyecto)

                .Include(f => f.CentroCosto)
                .Include(f => f.FacturaDetalles)
                .FirstOrDefaultAsync(f => f.IdFactura == id);

            if (factura == null)
            {
                return NotFound();
            }

            return View(factura);
        }
        // ================================================================
        // DESCRIBE LOS FILTROS UTILIZADOS EN REPORTES DE FACTURAS
        // ================================================================
        // Permite registrar en Auditoría qué información fue consultada
        // o exportada por el usuario.
        private string ConstruirDescripcionFiltrosReporte(
            string? tipoFactura,
            string? estado,
            string? condicionPago,
            int? clienteId,
            int? proveedorId,
            int? proyectoId,
            int? anio,
            int? mes,
            DateOnly? fechaDesde,
            DateOnly? fechaHasta)
        {
            var filtros = new List<string>();

            if (!string.IsNullOrWhiteSpace(tipoFactura))
            {
                filtros.Add($"Tipo={tipoFactura}");
            }

            if (!string.IsNullOrWhiteSpace(estado))
            {
                filtros.Add($"Estado={estado}");
            }

            if (!string.IsNullOrWhiteSpace(condicionPago))
            {
                filtros.Add($"Condición={condicionPago}");
            }

            if (clienteId.HasValue && clienteId.Value > 0)
            {
                filtros.Add($"ClienteId={clienteId.Value}");
            }

            if (proveedorId.HasValue && proveedorId.Value > 0)
            {
                filtros.Add($"ProveedorId={proveedorId.Value}");
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

            if (fechaDesde.HasValue)
            {
                filtros.Add($"Desde={fechaDesde.Value:dd/MM/yyyy}");
            }

            if (fechaHasta.HasValue)
            {
                filtros.Add($"Hasta={fechaHasta.Value:dd/MM/yyyy}");
            }

            return filtros.Any()
                ? "Filtros aplicados: " + string.Join(", ", filtros) + "."
                : "Sin filtros aplicados.";
        }
        // Registra auditoría de las acciones realizadas en el módulo de facturación.
        private async Task RegistrarAuditoriaFactura(string accion, int registroId, string descripcion)
        {
            int usuarioId = HttpContext.Session.GetInt32("UsuarioId") ?? 1;

            var auditoria = new Auditorium
            {
                UsuarioId = usuarioId,
                Tabla = "Facturas",
                RegistroId = registroId,
                Accion = accion,
                Descripcion = descripcion,
                Fecha = DateTime.Now
            };

            _context.Auditoria.Add(auditoria);
            await _context.SaveChangesAsync();
        }
        public class DetalleFacturaJson
        {
            public int? ServicioId { get; set; }
            public string ServicioNombre { get; set; } = string.Empty;
            public int Cantidad { get; set; }
            public decimal PrecioUnitario { get; set; }
            public decimal Descuento { get; set; }
            public string TipoImpuesto { get; set; } = string.Empty;
            public decimal SubtotalLinea { get; set; }
            public decimal ImpuestoLinea { get; set; }
            public decimal TotalLinea { get; set; }
        }
   // ================================================================
    // VALIDAR SI UN PERÍODO CONTABLE ESTÁ CERRADO
    // ================================================================
    private async Task<bool> PeriodoContableCerradoAsync(DateOnly fecha)
    {
        return await _context.CierresContables
            .AsNoTracking()
            .AnyAsync(c =>
                c.Anio == fecha.Year &&
                c.Mes == fecha.Month &&
                c.Estado == "Cerrado");
    }

    }
 
}
