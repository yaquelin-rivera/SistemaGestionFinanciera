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
using SistemaGestionFinanciera.Services;
using ClosedXML.Excel;
using System.Globalization;

namespace SistemaGestionFinanciera.Controllers
{
    public class GastosController : BaseController
    {
        private readonly SistemaFinancieroContext _context;
        private readonly ContabilidadService _contabilidadService;
        public GastosController(
      SistemaFinancieroContext context,
      ContabilidadService contabilidadService)
        {
            _context = context;
            _contabilidadService = contabilidadService;
        }

        // GET: Gastos
        public async Task<IActionResult> Index(
            string? buscar,
            string? estado,
            string? origen,
            int? proyectoId,
            int? anio,
            int? mes,
            DateOnly? fechaInicio,
            DateOnly? fechaFin,
            bool? excedido)
        {
            // Detecta si el usuario aplicó filtros.
            bool hayFiltros =
                !string.IsNullOrWhiteSpace(buscar) ||
                !string.IsNullOrWhiteSpace(estado) ||
                !string.IsNullOrWhiteSpace(origen) ||
                proyectoId.HasValue ||
                anio.HasValue ||
                mes.HasValue ||
                fechaInicio.HasValue ||
                fechaFin.HasValue ||
                 excedido == true;
            // OBTENER GASTOS SEGÚN LOS FILTROS APLICADOS
            // La consulta, los filtros, el origen y el control de categoría se centralizan en ConstruirConsultaReporte() para reutilizar exactamente la misma información en Index, PDF y Excel.
          
            var gastos = await ConstruirConsultaReporte(
                buscar,
                estado,
                origen,
                proyectoId,
                anio,
                mes,
                fechaInicio,
                fechaFin,
                excedido);

            var gastosParaTotales = gastos;
            // ALERTAS GLOBALES DEL MÓDULO DE GASTOS No dependen de los filtros ni de las tarjetas del mes actual.

            var gastosParaAlertas = await _context.Gastos
                .Include(g => g.CategoriaGasto)
                .Include(g => g.Proveedor)
                .Include(g => g.Proyecto)
                .Include(g => g.CentroCosto)
                .Where(g =>
                    g.Activo &&
                    g.Estado != "Anulado" &&
                    g.Estado != "Rechazado")
                .ToListAsync();

            // Calcula el control presupuestario de todos los gastos
            // antes de contar cuáles exceden su categoría.
            foreach (var gastoAlerta in gastosParaAlertas)
            {
                await CalcularControlCategoriaGastoAsync(gastoAlerta);
            }

            // ---------------------------------------------------------------
            // GASTOS PENDIENTES DE AUTORIZACIÓN

            var gastosPendientesAutorizacion = gastosParaAlertas
                .Where(g =>
                    g.Estado == "Pendiente de autorización")
                .ToList();

            ViewBag.CantidadPendientesAutorizacion =
                gastosPendientesAutorizacion.Count;

            ViewBag.MontoPendientesAutorizacion =
                gastosPendientesAutorizacion.Sum(g =>
                    g.MontoTotal ?? 0m);

            // ---------------------------------------------------------------
            // GASTOS QUE SUPERAN EL LÍMITE DE SU CATEGORÍA

            var gastosExcedidos = gastosParaAlertas
                .Where(g => g.SuperaPresupuesto)
                .ToList();

            ViewBag.CantidadExcedidos =
                gastosExcedidos.Count;

            ViewBag.MontoExcedidos =
                gastosExcedidos.Sum(g =>
                    g.MontoTotal ?? 0m);
        
            // Si no hay filtros, los totales se calculan solo con el mes actual.
            if (!hayFiltros)
            {
                var hoy = DateOnly.FromDateTime(DateTime.Today);
                var inicioMes = new DateOnly(hoy.Year, hoy.Month, 1);
                var finMes = inicioMes.AddMonths(1).AddDays(-1);

                gastosParaTotales = gastos
                    .Where(g => g.Fecha >= inicioMes && g.Fecha <= finMes)
                    .ToList();

                ViewBag.TipoResumen = "Mes actual";
            }
            else
            {
                ViewBag.TipoResumen = "Filtro aplicado";
            }

          
            // Totales financieros.
            ViewBag.TotalRegistrado = gastosParaTotales
                .Where(g => g.Activo && g.Estado != "Anulado" && g.Estado != "Rechazado")
                .Sum(g => g.MontoTotal ?? 0);

            ViewBag.TotalAprobado = gastosParaTotales
                .Where(g => g.Activo && g.Estado == "Aprobado")
                .Sum(g => g.MontoTotal ?? 0);

            ViewBag.TotalPendiente = gastosParaTotales
                .Where(g => g.Activo && g.Estado == "Pendiente de autorización")
                .Sum(g => g.MontoTotal ?? 0);

            ViewBag.TotalNoValido = gastosParaTotales
                .Where(g => g.Estado == "Rechazado" || g.Estado == "Anulado")
                .Sum(g => g.MontoTotal ?? 0);

            // Mantiene valores en pantalla.
            ViewBag.Buscar = buscar;
            ViewBag.Estado = estado;
            ViewBag.Origen = origen;
            ViewBag.ProyectoId = proyectoId;
            ViewBag.Anio = anio;
            ViewBag.Mes = mes;
            ViewBag.FechaInicio = fechaInicio?.ToString("yyyy-MM-dd");
            ViewBag.FechaFin = fechaFin?.ToString("yyyy-MM-dd");
            ViewBag.Excedido = excedido;

            // Carga proyectos para filtro.
            ViewBag.Proyectos = new SelectList(
      await _context.Proyectos
          .Where(p => p.Activo)
          .OrderBy(p => p.Nombre)
          .ToListAsync(),
      "IdProyecto",
      "Nombre",
      proyectoId);

            // Carga automáticamente los años existentes en gastos.
            ViewBag.Anios = await _context.Gastos
                .Select(g => g.Fecha.Year)
                .Distinct()
                .OrderByDescending(a => a)
                .ToListAsync();

            return View(gastos);
        }
        // GET: Gastos/Detalles
        public async Task<IActionResult> Details(int? id)
        {
            var gasto = await _context.Gastos
    .Include(g => g.CategoriaGasto)
    .Include(g => g.Proveedor)
    .Include(g => g.Proyecto)
    .Include(g => g.PresupuestoMensualRegistro)
    .Include(g => g.CentroCosto)
    .FirstOrDefaultAsync(m => m.IdGasto == id);

            if (gasto == null)
            {
                return NotFound();
            }
            await CalcularControlCategoriaGastoAsync(gasto);

            ViewBag.PresupuestoMensual =
                gasto.PresupuestoMensual ?? 0m;

            ViewBag.GastadoMes =
                gasto.GastadoAcumulado ?? 0m;

            ViewBag.SaldoDisponible =
                gasto.SaldoDisponible ?? 0m;

            ViewBag.EstadoPresupuesto =
                gasto.SuperaPresupuesto
                    ? "Excedido"
                    : "Dentro del límite";

            return View(gasto);
        }

        // CONSTRUYE LA CONSULTA BASE PARA REPORTES Y CONSULTAS
        // ===========================================================
        // Este método centraliza todos los filtros utilizados por:
      
        // • Index()
        // • Reporte()
        // • ExportarExcel()
     
        // De esta forma toda la información mostrada será exactamente
        // la misma en pantalla, PDF y Excel.
        private async Task<List<Gasto>> ConstruirConsultaReporte(
            string? buscar,
            string? estado,
            string? origen,
            int? proyectoId,
            int? anio,
            int? mes,
            DateOnly? fechaInicio,
            DateOnly? fechaFin,
            bool? excedido)
        {
            // Detecta si existen filtros distintos al filtro rápido
            // de "Límite excedido".
            bool hayFiltros =
                !string.IsNullOrWhiteSpace(buscar) ||
                !string.IsNullOrWhiteSpace(estado) ||
                !string.IsNullOrWhiteSpace(origen) ||
                proyectoId.HasValue ||
                anio.HasValue ||
                mes.HasValue ||
                fechaInicio.HasValue ||
                fechaFin.HasValue;

            // ===========================================================
            // CONSULTA BASE
            // ===========================================================

            var consulta = _context.Gastos
                .Include(g => g.CategoriaGasto)
                .Include(g => g.Proveedor)
                .Include(g => g.Proyecto)
                .Include(g => g.PresupuestoMensualRegistro)
                .Include(g => g.CentroCosto)

                .AsQueryable();

            // ===========================================================
            // FILTROS
            // ===========================================================

            if (fechaInicio.HasValue)
                consulta = consulta.Where(g => g.Fecha >= fechaInicio.Value);

            if (fechaFin.HasValue)
                consulta = consulta.Where(g => g.Fecha <= fechaFin.Value);

            if (proyectoId.HasValue)
                consulta = consulta.Where(g => g.ProyectoId == proyectoId.Value);

            if (anio.HasValue)
                consulta = consulta.Where(g => g.Fecha.Year == anio.Value);

            if (mes.HasValue)
                consulta = consulta.Where(g => g.Fecha.Month == mes.Value);

            if (!string.IsNullOrWhiteSpace(buscar))
            {
                consulta = consulta.Where(g =>
                    g.Concepto.Contains(buscar) ||
                    (g.NumeroFactura != null && g.NumeroFactura.Contains(buscar)) ||
                    (g.Proveedor != null && g.Proveedor.Nombre.Contains(buscar)) ||
                    (g.CategoriaGasto != null && g.CategoriaGasto.Nombre.Contains(buscar)) ||
                    (g.Proyecto != null && g.Proyecto.Nombre.Contains(buscar)));
            }

            if (!string.IsNullOrWhiteSpace(estado))
                consulta = consulta.Where(g => g.Estado == estado);

            // ===========================================================
            // OBTENER LISTA
            // ===========================================================

            var gastos = await consulta
                .OrderByDescending(g => g.IdGasto)
                .ToListAsync();

            // ===========================================================
            // CALCULAR CONTROL PRESUPUESTARIO
            // ===========================================================

            foreach (var gasto in gastos)
            {
                await CalcularControlCategoriaGastoAsync(gasto);
            }
            // ===========================================================
            // FILTRO RÁPIDO
            // ===========================================================

            if (excedido == true)
            {
                gastos = gastos
                    .Where(g =>
                        g.SuperaPresupuesto &&
                        g.Activo &&
                        g.Estado != "Anulado" &&
                        g.Estado != "Rechazado")
                    .ToList();
            }

            // ===========================================================
            // FILTRO POR ORIGEN
            // ===========================================================

            if (!string.IsNullOrWhiteSpace(origen))
            {
                gastos = gastos.Where(g =>
                {
                    bool esCxP =
                        !string.IsNullOrWhiteSpace(g.NumeroFactura) &&
                        g.Concepto != null &&
                        g.Concepto.ToLower().Contains("cuenta por pagar");

                    bool esFactura =
                        !string.IsNullOrWhiteSpace(g.NumeroFactura) &&
                        g.Concepto != null &&
                        g.Concepto.ToLower().Contains("factura");

                    bool esManual = !esCxP && !esFactura;

                    return
                        (origen == "Manual" && esManual) ||
                        (origen == "CxP" && esCxP) ||
                        (origen == "Factura" && esFactura);

                }).ToList();
            }

            return gastos;
        }

        // VISTA PREVIA DEL REPORTE DE GASTOS
        // Genera una vista imprimible con los mismos filtros aplicados en el Index.
        // Esta vista permitirá: Revisar el reporte en pantalla.  Imprimirlo.
        public async Task<IActionResult> Reporte(
            string? buscar,
            string? estado,
            string? origen,
            int? proyectoId,
            int? anio,
            int? mes,
            DateOnly? fechaInicio,
            DateOnly? fechaFin,
            bool? excedido)
        {
            // Obtiene exactamente los mismos gastos que muestra el Index.
            var gastos = await ConstruirConsultaReporte(
                buscar,
                estado,
                origen,
                proyectoId,
                anio,
                mes,
                fechaInicio,
                fechaFin,
                excedido);

            // ===========================================================
            // TOTALES DEL REPORTE

            // Total de registros encontrados.
            ViewBag.TotalRegistros = gastos.Count;

            // Total registrado:
            // Excluye gastos anulados y rechazados.
            ViewBag.TotalRegistrado = gastos
                .Where(g =>
                    g.Activo &&
                    g.Estado != "Anulado" &&
                    g.Estado != "Rechazado")
                .Sum(g => g.MontoTotal ?? 0m);

            // Total aprobado.
            ViewBag.TotalAprobado = gastos
                .Where(g =>
                    g.Activo &&
                    g.Estado == "Aprobado")
                .Sum(g => g.MontoTotal ?? 0m);

            // Total pendiente de autorización.
            ViewBag.TotalPendiente = gastos
                .Where(g =>
                    g.Activo &&
                    g.Estado == "Pendiente de autorización")
                .Sum(g => g.MontoTotal ?? 0m);

            // Total no válido:
            // Incluye rechazados y anulados.
            ViewBag.TotalNoValido = gastos
                .Where(g =>
                    g.Estado == "Rechazado" ||
                    g.Estado == "Anulado")
                .Sum(g => g.MontoTotal ?? 0m);

            // Cantidad de gastos que exceden el límite mensual
            // de su categoría.
            ViewBag.CantidadExcedidos = gastos.Count(g =>
                g.Activo &&
                g.SuperaPresupuesto &&
                g.Estado != "Anulado" &&
                g.Estado != "Rechazado");

            // Fecha y hora de generación del reporte.
            ViewBag.FechaGeneracion = DateTime.Now;

            // ===========================================================
            // CONSERVAR LOS FILTROS PARA MOSTRARLOS EN EL REPORTE

            ViewBag.Buscar = buscar;
            ViewBag.Estado = estado;
            ViewBag.Origen = origen;
            ViewBag.ProyectoId = proyectoId;
            ViewBag.Anio = anio;
            ViewBag.Mes = mes;
            ViewBag.FechaInicio = fechaInicio?.ToString("yyyy-MM-dd");
            ViewBag.FechaFin = fechaFin?.ToString("yyyy-MM-dd");
            ViewBag.Excedido = excedido;

            // Ordena desde el gasto más reciente.
            gastos = gastos
                .OrderByDescending(g => g.Fecha)
                .ThenByDescending(g => g.IdGasto)
                .ToList();

            string descripcionFiltros =
    await ConstruirDescripcionFiltrosReporte(
        buscar,
        estado,
        origen,
        proyectoId,
        anio,
        mes,
        fechaInicio,
        fechaFin,
        excedido);

            await RegistrarAuditoriaGasto(
                "Vista previa PDF",
                0,
                "Se generó la vista previa del reporte de gastos. " +
                descripcionFiltros);

            return View(gastos);
            
        }
        // EXPORTAR REPORTE DE GASTOS A EXCEL
        // Genera un archivo Excel utilizando exactamente los mismos
        // filtros aplicados en el Index y en la vista previa del reporte.
        public async Task<IActionResult> ExportarExcel(
            string? buscar,
            string? estado,
            string? origen,
            int? proyectoId,
            int? anio,
            int? mes,
            DateOnly? fechaInicio,
            DateOnly? fechaFin,
            bool? excedido)
        {
            // ===========================================================
            // OBTENER LOS GASTOS SEGÚN LOS FILTROS
            // ===========================================================

            var gastos = await ConstruirConsultaReporte(
                buscar,
                estado,
                origen,
                proyectoId,
                anio,
                mes,
                fechaInicio,
                fechaFin,
                excedido);

            // Ordenar desde el gasto más reciente.
            gastos = gastos
                .OrderByDescending(g => g.Fecha)
                .ThenByDescending(g => g.IdGasto)
                .ToList();

            // Crear el libro y la hoja de Excel.
            using var workbook = new XLWorkbook();

            var hoja = workbook.Worksheets.Add("Gastos");

            // Cantidad total de columnas del reporte: A hasta M.
            const int totalColumnas = 13;

            // ===========================================================
            // ENCABEZADO PRINCIPAL
            // ===========================================================

            hoja.Cell("A1").Value = "UCAD SANTA ANA";
            hoja.Range(1, 1, 1, totalColumnas).Merge();

            hoja.Cell("A2").Value = "Reporte de Gastos";
            hoja.Range(2, 1, 2, totalColumnas).Merge();

            hoja.Cell("A3").Value =
                $"Generado el {DateTime.Now:dd/MM/yyyy HH:mm}";

            hoja.Range(3, 1, 3, totalColumnas).Merge();

            // Formato del nombre de la institución.
            hoja.Range(1, 1, 1, totalColumnas)
                .Style.Font.Bold = true;

            hoja.Range(1, 1, 1, totalColumnas)
                .Style.Font.FontSize = 16;

            hoja.Range(1, 1, 1, totalColumnas)
                .Style.Alignment.Horizontal =
                    XLAlignmentHorizontalValues.Center;

            // Formato del título del reporte.
            hoja.Range(2, 1, 2, totalColumnas)
                .Style.Font.Bold = true;

            hoja.Range(2, 1, 2, totalColumnas)
                .Style.Font.FontSize = 13;

            hoja.Range(2, 1, 2, totalColumnas)
                .Style.Alignment.Horizontal =
                    XLAlignmentHorizontalValues.Center;

            // Formato de la fecha de generación.
            hoja.Range(3, 1, 3, totalColumnas)
                .Style.Alignment.Horizontal =
                    XLAlignmentHorizontalValues.Center;

            // ===========================================================
            // FILTROS APLICADOS
            // ===========================================================

            var filtros = new List<string>();

            if (!string.IsNullOrWhiteSpace(buscar))
            {
                filtros.Add($"Búsqueda: {buscar}");
            }

            if (!string.IsNullOrWhiteSpace(estado))
            {
                filtros.Add($"Estado: {estado}");
            }

            if (!string.IsNullOrWhiteSpace(origen))
            {
                filtros.Add($"Origen: {origen}");
            }

            if (proyectoId.HasValue)
            {
                string proyecto = await _context.Proyectos
                    .AsNoTracking()
                    .Where(p =>
                        p.IdProyecto == proyectoId.Value)
                    .Select(p => p.Nombre)
                    .FirstOrDefaultAsync()
                    ?? "No encontrado";

                filtros.Add($"Proyecto: {proyecto}");
            }

            if (anio.HasValue)
            {
                filtros.Add($"Año: {anio.Value}");
            }

            if (mes.HasValue)
            {
                string nombreMes = new DateTime(2000, mes.Value, 1)
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

            if (excedido.HasValue)
            {
                filtros.Add(
                    excedido.Value
                        ? "Control presupuestario: Límite de categoría excedido"
                        : "Control presupuestario: Sin límite excedido");
            }

            hoja.Cell("A5").Value = "Filtros aplicados:";
            hoja.Cell("A5").Style.Font.Bold = true;

            hoja.Cell("B5").Value =
                filtros.Any()
                    ? string.Join(" | ", filtros)
                    : "Sin filtros";

            hoja.Range(5, 2, 5, totalColumnas).Merge();

            hoja.Range(5, 2, 5, totalColumnas)
                .Style.Alignment.WrapText = true;

            // ===========================================================
            // ENCABEZADOS DE LA TABLA
            // ===========================================================

            const int filaEncabezado = 7;

            string[] encabezados =
            {
        "Fecha",
        "Concepto",
        "Factura",
        "Categoría",
        "Proveedor",
        "Proyecto",
        "Centro de costo",
        "Origen",
        "Estado",
        "Subtotal",
        "Impuesto",
        "Descuento",
        "Monto total"
    };

            for (int columna = 0;
                 columna < encabezados.Length;
                 columna++)
            {
                hoja.Cell(
                    filaEncabezado,
                    columna + 1).Value = encabezados[columna];
            }

            var rangoEncabezado = hoja.Range(
                filaEncabezado,
                1,
                filaEncabezado,
                encabezados.Length);

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

            // ===========================================================
            // DATOS DEL REPORTE
            // ===========================================================

            int fila = filaEncabezado + 1;

            foreach (var gasto in gastos)
            {
                // Detectar el origen funcional del gasto.
                bool esGastoDesdeCxP =
                    !string.IsNullOrWhiteSpace(gasto.NumeroFactura) &&
                    gasto.Concepto != null &&
                    gasto.Concepto
                        .ToLower()
                        .Contains("cuenta por pagar");

                bool esGastoDesdeFactura =
                    !string.IsNullOrWhiteSpace(gasto.NumeroFactura) &&
                    gasto.Concepto != null &&
                    gasto.Concepto
                        .ToLower()
                        .Contains("factura");

                string origenGasto;

                if (esGastoDesdeCxP)
                {
                    origenGasto = "CxP";
                }
                else if (esGastoDesdeFactura)
                {
                    origenGasto = "Factura";
                }
                else
                {
                    origenGasto = "Manual";
                }

                // Fecha.
                hoja.Cell(fila, 1).Value =
                    gasto.Fecha.ToDateTime(TimeOnly.MinValue);

                // Concepto.
                hoja.Cell(fila, 2).Value =
                    gasto.Concepto ?? "Sin concepto";

                // Número de factura.
                hoja.Cell(fila, 3).Value =
                    string.IsNullOrWhiteSpace(gasto.NumeroFactura)
                        ? "Sin factura"
                        : gasto.NumeroFactura;

                // Categoría.
                hoja.Cell(fila, 4).Value =
                    gasto.CategoriaGasto?.Nombre ??
                    "Sin categoría";

                // Proveedor.
                hoja.Cell(fila, 5).Value =
                    gasto.Proveedor?.Nombre ??
                    "Sin proveedor";

                // Proyecto.
                hoja.Cell(fila, 6).Value =
                    gasto.Proyecto?.Nombre ??
                    "Sin proyecto";

                // Centro de costo.
                hoja.Cell(fila, 7).Value =
                    gasto.CentroCosto?.Nombre ??
                    "Sin centro";

                // Origen.
                hoja.Cell(fila, 8).Value =
                    origenGasto;

                // Estado y alerta presupuestaria.
                string estadoGasto =
                    gasto.Estado ?? "Sin estado";

                if (gasto.SuperaPresupuesto &&
                    gasto.Estado != "Anulado" &&
                    gasto.Estado != "Rechazado")
                {
                    estadoGasto += " - Límite excedido";
                }

                hoja.Cell(fila, 9).Value =
                    estadoGasto;

                // Montos.
                hoja.Cell(fila, 10).Value =
                    gasto.Subtotal ?? 0m;

                hoja.Cell(fila, 11).Value =
                    gasto.Impuesto ?? 0m;

                hoja.Cell(fila, 12).Value =
                    gasto.Descuento ?? 0m;

                hoja.Cell(fila, 13).Value =
                    gasto.MontoTotal ?? 0m;

                fila++;
            }

            // ===========================================================
            // FORMATO DE LA TABLA
            // ===========================================================

            if (gastos.Any())
            {
                // Crear una tabla formal de Excel.
                hoja.Range(
                        filaEncabezado,
                        1,
                        fila - 1,
                        encabezados.Length)
                    .CreateTable();

                // Formato de fecha.
                hoja.Column(1)
                    .Style.DateFormat.Format =
                        "dd/MM/yyyy";

                // Formato de moneda.
                hoja.Columns(10, 13)
                    .Style.NumberFormat.Format =
                        "₡#,##0.00";

                // Mantener visibles los encabezados.
                hoja.SheetView.FreezeRows(filaEncabezado);

                // Alineación vertical de los datos.
                hoja.Range(
                        filaEncabezado + 1,
                        1,
                        fila - 1,
                        encabezados.Length)
                    .Style.Alignment.Vertical =
                        XLAlignmentVerticalValues.Center;

                // Bordes inferiores.
                hoja.Range(
                        filaEncabezado + 1,
                        1,
                        fila - 1,
                        encabezados.Length)
                    .Style.Border.BottomBorder =
                        XLBorderStyleValues.Thin;

                hoja.Range(
                        filaEncabezado + 1,
                        1,
                        fila - 1,
                        encabezados.Length)
                    .Style.Border.BottomBorderColor =
                        XLColor.LightGray;
                // ===========================================================
                // COLORES SEGÚN EL ESTADO DEL GASTO
                // ===========================================================

                for (int indice = 0; indice < gastos.Count; indice++)
                {
                    var gasto = gastos[indice];

                    int filaGasto =
                        filaEncabezado + 1 + indice;

                    var rangoFila = hoja.Range(
                        filaGasto,
                        1,
                        filaGasto,
                        totalColumnas);

                    if (gasto.Estado == "Anulado")
                    {
                        // Fondo rojo claro.
                        rangoFila.Style.Fill.BackgroundColor =
                            XLColor.FromHtml("#F8D7DA");

                        // Texto rojo oscuro.
                        rangoFila.Style.Font.FontColor =
                            XLColor.FromHtml("#842029");
                    }
                    else if (gasto.Estado == "Pendiente de autorización")
                    {
                        // Fondo amarillo claro.
                        rangoFila.Style.Fill.BackgroundColor =
                            XLColor.FromHtml("#FFF3CD");

                        // Texto amarillo oscuro.
                        rangoFila.Style.Font.FontColor =
                            XLColor.FromHtml("#664D03");
                    }
                }
            }

            // ===========================================================
            // TOTALES
            // ===========================================================

            // Total registrado:
            // excluye rechazados, anulados e inactivos.
            decimal totalRegistrado = gastos
                .Where(g =>
                    g.Activo &&
                    g.Estado != "Anulado" &&
                    g.Estado != "Rechazado")
                .Sum(g => g.MontoTotal ?? 0m);

            // Total aprobado.
            decimal totalAprobado = gastos
                .Where(g =>
                    g.Activo &&
                    g.Estado == "Aprobado")
                .Sum(g => g.MontoTotal ?? 0m);

            // Total pendiente.
            decimal totalPendiente = gastos
                .Where(g =>
                    g.Activo &&
                    g.Estado == "Pendiente de autorización")
                .Sum(g => g.MontoTotal ?? 0m);

            // Total no válido.
            decimal totalNoValido = gastos
                .Where(g =>
                    g.Estado == "Rechazado" ||
                    g.Estado == "Anulado")
                .Sum(g => g.MontoTotal ?? 0m);

            int cantidadExcedidos = gastos.Count(g =>
                g.Activo &&
                g.SuperaPresupuesto &&
                g.Estado != "Anulado" &&
                g.Estado != "Rechazado");

            // Fila donde comienzan los totales.
            int filaTotales = fila + 1;

            hoja.Cell(filaTotales, 9).Value =
                "Resumen del reporte";

            hoja.Cell(filaTotales, 9)
                .Style.Font.Bold = true;

            hoja.Cell(filaTotales, 9)
                .Style.Fill.BackgroundColor =
                    XLColor.FromHtml("#D8F0F3");

            // Total registrado.
            hoja.Cell(filaTotales + 1, 9).Value =
                "Total registrado:";

            hoja.Cell(filaTotales + 1, 13).Value =
                totalRegistrado;

            // Total aprobado.
            hoja.Cell(filaTotales + 2, 9).Value =
                "Total aprobado:";

            hoja.Cell(filaTotales + 2, 13).Value =
                totalAprobado;

            // Total pendiente.
            hoja.Cell(filaTotales + 3, 9).Value =
                "Total pendiente:";

            hoja.Cell(filaTotales + 3, 13).Value =
                totalPendiente;

            // Total no válido.
            hoja.Cell(filaTotales + 4, 9).Value =
                "Total no válido:";

            hoja.Cell(filaTotales + 4, 13).Value =
                totalNoValido;

            // Cantidad de gastos excedidos.
            hoja.Cell(filaTotales + 5, 9).Value =
                "Límite de categoría excedido:";

            // Colocar la cantidad.
            hoja.Cell(filaTotales + 5, 13).Value =
                cantidadExcedidos;

            // Mostrar como número entero y no como dinero.
            hoja.Cell(filaTotales + 5, 13)
                .Style.NumberFormat.Format = "0";

            // Negrita en etiquetas.
            hoja.Range(
                    filaTotales + 1,
                    9,
                    filaTotales + 5,
                    9)
                .Style.Font.Bold = true;

            // Formato monetario de los totales.
            hoja.Range(
                    filaTotales + 1,
                    13,
                    filaTotales + 4,
                    13)
                .Style.NumberFormat.Format =
                    "₡#,##0.00";

            // Alinear montos a la derecha.
            hoja.Range(
                    filaTotales + 1,
                    13,
                    filaTotales + 5,
                    13)
                .Style.Alignment.Horizontal =
                    XLAlignmentHorizontalValues.Right;

            // Línea superior del resumen.
            hoja.Range(
                    filaTotales,
                    9,
                    filaTotales,
                    13)
                .Style.Border.TopBorder =
                    XLBorderStyleValues.Medium;

            hoja.Range(
                    filaTotales,
                    9,
                    filaTotales,
                    13)
                .Style.Border.TopBorderColor =
                    XLColor.FromHtml("#0F5C64");

            // ===========================================================
            // ANCHOS Y AJUSTES
            // ===========================================================

            hoja.Columns().AdjustToContents();

            hoja.Column(1).Width = 13;
            hoja.Column(2).Width = 45;
            hoja.Column(3).Width = 16;
            hoja.Column(4).Width = 24;
            hoja.Column(5).Width = 25;
            hoja.Column(6).Width = 30;
            hoja.Column(7).Width = 25;
            hoja.Column(8).Width = 12;
            hoja.Column(9).Width = 27;
            hoja.Column(10).Width = 16;
            hoja.Column(11).Width = 16;
            hoja.Column(12).Width = 16;
            hoja.Column(13).Width = 18;

            // Permitir varias líneas en columnas con texto largo.
            hoja.Column(2)
                .Style.Alignment.WrapText = true;

            hoja.Column(4)
                .Style.Alignment.WrapText = true;

            hoja.Column(5)
                .Style.Alignment.WrapText = true;

            hoja.Column(6)
                .Style.Alignment.WrapText = true;

            hoja.Column(7)
                .Style.Alignment.WrapText = true;

            hoja.Column(9)
                .Style.Alignment.WrapText = true;

            // Alinear montos a la derecha.
            hoja.Columns(10, 13)
                .Style.Alignment.Horizontal =
                    XLAlignmentHorizontalValues.Right;

            // Orientación horizontal al imprimir desde Excel.
            hoja.PageSetup.PageOrientation =
                XLPageOrientation.Landscape;

            hoja.PageSetup.FitToPages(1, 0);

            // Repetir encabezados si el reporte ocupa varias páginas.
            hoja.PageSetup.SetRowsToRepeatAtTop(
                filaEncabezado,
                filaEncabezado);

            // Centrar horizontalmente en la impresión.
            hoja.PageSetup.CenterHorizontally = true;

            // ===========================================================
            // DEVOLVER EL ARCHIVO
            // ===========================================================

            using var stream = new MemoryStream();

            workbook.SaveAs(stream);

            string nombreArchivo =
     $"ReporteGastos_{DateTime.Now:yyyyMMdd_HHmm}.xlsx";

            string descripcionFiltros =
                await ConstruirDescripcionFiltrosReporte(
                    buscar,
                    estado,
                    origen,
                    proyectoId,
                    anio,
                    mes,
                    fechaInicio,
                    fechaFin,
                    excedido);

            await RegistrarAuditoriaGasto(
                "Exportar Excel",
                0,
                "Se exportó el reporte de gastos a Excel. " +
                descripcionFiltros);

            return File(
                stream.ToArray(),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                nombreArchivo);
        }

        // GET: Gastos/Crear
        public async Task<IActionResult> Create()
        {
            await CargarCatalogosCreateAsync();

            return View();
        }

        // POST: Gastos/Crear
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Create(
    [Bind("IdGasto,Fecha,CategoriaGastoId,PresupuestoMensualId,ProveedorId,CentroCostoId,NumeroFactura,Concepto,Subtotal,Impuesto,Descuento,TipoImpuesto,MontoTotal,Observacion")] Gasto gasto)
        {
            //calculando impuesto, descuento, total
            gasto.Subtotal = Math.Round(gasto.Subtotal ?? 0, 2);
            gasto.Impuesto = Math.Round(gasto.Impuesto ?? 0, 2);
            gasto.Descuento = Math.Round(gasto.Descuento ?? 0, 2);
            gasto.MontoTotal = Math.Round(gasto.MontoTotal ?? 0, 2);

            if ((gasto.Descuento ?? 0) > (gasto.MontoTotal ?? 0))
            {
                ModelState.AddModelError("Descuento", "El descuento no puede ser mayor que el monto total.");
            }
            
            if (gasto.Fecha > DateOnly.FromDateTime(DateTime.Today))
            {
                ModelState.AddModelError("Fecha", "La fecha del gasto no puede ser mayor a la fecha actual.");
            }

            bool periodoCerrado =
    await PeriodoContableCerradoAsync(gasto.Fecha);

            if (periodoCerrado)
            {
                ModelState.AddModelError(
                    "Fecha",
                    "No se puede registrar el gasto en este mes porque el período contable seleccionado se encuentra cerrado.");
            }
            // ============================================================
            // VALIDAR PERÍODO PRESUPUESTARIO SELECCIONADO
            // ============================================================

            PresupuestosMensuale? presupuestoSeleccionado = null;

            if (!gasto.PresupuestoMensualId.HasValue)
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
                                gasto.PresupuestoMensualId.Value &&
                            pm.EstadoAprobacion == "Aprobado" &&
                            pm.ProyectoId != null &&
                            pm.Proyecto != null &&
                            pm.Proyecto.Activo &&

            // NUEVO:
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
                    // El proyecto NO viene del navegador.
                    // Se obtiene directamente del presupuesto seleccionado.
                    gasto.ProyectoId =
                        presupuestoSeleccionado.ProyectoId;
                }
            }
            // Validar categoría de gasto activa.
            bool categoriaActiva = await _context.CategoriasGastos
                .AnyAsync(c =>
                    c.IdCategoriaGasto == gasto.CategoriaGastoId &&
                    c.Activo);

            if (!categoriaActiva)
            {
                ModelState.AddModelError(
                    "CategoriaGastoId",
                    "La categoría de gasto seleccionada no está activa.");
            }


            // Validar proveedor activo.
            if (gasto.ProveedorId.HasValue)
            {
                bool proveedorActivo = await _context.Proveedores
                    .AnyAsync(p =>
                        p.IdProveedor == gasto.ProveedorId.Value &&
                        p.Activo);

                if (!proveedorActivo)
                {
                    ModelState.AddModelError(
                        "ProveedorId",
                        "El proveedor seleccionado no está activo.");
                }
            }


            // Validar centro de costo activo.
            if (gasto.CentroCostoId.HasValue)
            {
                bool centroCostoActivo = await _context.CentrosCostos
                    .AnyAsync(c =>
                        c.IdCentroCosto == gasto.CentroCostoId.Value &&
                        c.Activo);

                if (!centroCostoActivo)
                {
                    ModelState.AddModelError(
                        "CentroCostoId",
                        "El centro de costo seleccionado no está activo.");
                }
            }

            if (ModelState.IsValid)
            {         
                //proceso
                gasto.Activo = true;
                gasto.Estado = "Pendiente de autorización";
                _context.Add(gasto);
                //calculo total de gasto delm mes por categoria 
                var totalGastadoMes = _context.Gastos
                      .Where(g => g.CategoriaGastoId == gasto.CategoriaGastoId
                       && g.Activo
                       && g.Estado != "Anulado"
                       && g.Estado != "Rechazado"
                       && g.Fecha.Month == gasto.Fecha.Month
                       && g.Fecha.Year == gasto.Fecha.Year)
                     .Sum(g => g.MontoTotal ?? 0);
                    totalGastadoMes += gasto.MontoTotal ?? 0;

                var categoria = await _context.CategoriasGastos
                    .FirstOrDefaultAsync(c => c.IdCategoriaGasto == gasto.CategoriaGastoId);
               // calculando presupuesto disponible
                decimal presupuestoMensual = categoria?.LimiteMensual ?? 0;
                decimal saldoDisponible = presupuestoMensual - totalGastadoMes;
                // calculo deteccion de presupeusto 
                gasto.PresupuestoMensual = presupuestoMensual;
                gasto.GastadoAcumulado = totalGastadoMes;
                //calculo de saldo disponible
                gasto.SaldoDisponible = saldoDisponible;
                //verifica si el gasto supera el presupeusto mensual
                if (categoria != null)
                {
                    gasto.SuperaPresupuesto = totalGastadoMes > presupuestoMensual;
                }
                await _context.SaveChangesAsync();

                // Auditoría: registra la creación del gasto.
                await RegistrarAuditoriaGasto(
                    "Crear",
                    gasto.IdGasto,
                    "Se registró el gasto " + gasto.Concepto +
                    " por un monto de ₡" + (gasto.MontoTotal ?? 0).ToString("N2") + ".");

                if (gasto.SuperaPresupuesto)
                {
                    TempData["MensajeAdvertencia"] = "El gasto fue registrado correctamente, pero supera el límite mensual de la categoría.";
                }
                else
                {
                    TempData["MensajeExito"] = "El gasto fue registrado correctamente.";
                }
              return RedirectToAction(nameof(Index));
            }

            await CargarCatalogosCreateAsync(gasto);
            return View(gasto);
        }

        // GET: Gastos/Editar
        public async Task<IActionResult> Edit(int? id)
        {
            if (id == null)
            {
                return NotFound();
            }

           
            var gasto = await _context.Gastos.FindAsync(id);
   if (gasto == null)
            {
                return NotFound();
            }
            // ============================================================
            // BLOQUEO POR CIERRE CONTABLE
            // ============================================================
            bool periodoCerrado =
                await PeriodoContableCerradoAsync(gasto.Fecha);

            if (periodoCerrado)
            {
                TempData["MensajeError"] =
                    $"No se puede modificar el gasto porque el período " +
                    $"{gasto.Fecha:MM/yyyy} se encuentra cerrado. " +
                    $"Para realizar correcciones, primero debe reabrirse el período.";

                return RedirectToAction(
                    nameof(Index),
                    new
                    {
                        anio = gasto.Fecha.Year,
                        mes = gasto.Fecha.Month
                    });
            }
            // BLOQUEO DE EDICIÓN DE GASTOS AUTOMÁTICOS
            // Los gastos originados desde CxP o Factura deben modificarse
            // únicamente desde su módulo de origen.
            bool esGastoDesdeCxP =
                !string.IsNullOrWhiteSpace(gasto.NumeroFactura) &&
                gasto.Concepto != null &&
                gasto.Concepto
                    .ToLower()
                    .Contains("cuenta por pagar");

            bool esGastoDesdeFactura =
                !string.IsNullOrWhiteSpace(gasto.NumeroFactura) &&
                gasto.Concepto != null &&
                gasto.Concepto
                    .ToLower()
                    .Contains("factura");

            if (esGastoDesdeCxP)
            {
                TempData["MensajeError"] =
                    "Este gasto fue generado automáticamente desde una cuenta por pagar. " +
                    "Debe modificarse desde el módulo de Cuentas por Pagar.";

                return RedirectToAction(nameof(Index));
            }

            if (esGastoDesdeFactura)
            {
                TempData["MensajeError"] =
                    "Este gasto fue generado automáticamente desde una factura. " +
                    "Debe modificarse desde el módulo de Facturación.";

                return RedirectToAction(nameof(Index));
            }
            if (gasto.Estado == "Aprobado"
            //    || gasto.Estado == "Rechazado"
                || gasto.Estado == "Anulado")
            {
                TempData["MensajeError"] = "Este gasto no puede modificarse porque ya fue aprobado o anulado.";
                return RedirectToAction(nameof(Index));
            }

            await CargarCatalogosEditAsync(gasto);

            return View(gasto);
        }

        // POST: Gastos/Editar
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Edit(int id, [Bind("IdGasto,Fecha,CategoriaGastoId,PresupuestoMensualId,CentroCostoId,NumeroFactura,ProveedorId,Concepto,Subtotal,Impuesto,Descuento,TipoImpuesto,MontoTotal,Observacion")] Gasto gasto)
        {
            if (id != gasto.IdGasto)
            {
                return NotFound();
            }
            var gastoActual = await _context.Gastos
    .AsNoTracking()
    .FirstOrDefaultAsync(g => g.IdGasto == id);

            if (gastoActual == null)
            {
                return NotFound();
            }
            // ============================================================
            // BLOQUEO DEL PERÍODO ORIGINAL
            // ============================================================
            bool periodoOriginalCerrado =
                await PeriodoContableCerradoAsync(gastoActual.Fecha);

            if (periodoOriginalCerrado)
            {
                TempData["MensajeError"] =
                    $"No se puede modificar el gasto porque el período " +
                    $"{gastoActual.Fecha:MM/yyyy} se encuentra cerrado. " +
                    $"Para realizar correcciones, primero debe reabrirse el período.";

                return RedirectToAction(
                    nameof(Index),
                    new
                    {
                        anio = gastoActual.Fecha.Year,
                        mes = gastoActual.Fecha.Month
                    });
            }
            // ================================================================
            // PROTECCIÓN CONTRA POST DIRECTO DE GASTOS AUTOMÁTICOS
            // ================================================================
            bool esGastoDesdeCxP =
                !string.IsNullOrWhiteSpace(gastoActual.NumeroFactura) &&
                gastoActual.Concepto != null &&
                gastoActual.Concepto
                    .ToLower()
                    .Contains("cuenta por pagar");

            bool esGastoDesdeFactura =
                !string.IsNullOrWhiteSpace(gastoActual.NumeroFactura) &&
                gastoActual.Concepto != null &&
                gastoActual.Concepto
                    .ToLower()
                    .Contains("factura");

            if (esGastoDesdeCxP)
            {
                TempData["MensajeError"] =
                    "Este gasto fue generado automáticamente desde una cuenta por pagar. " +
                    "Debe modificarse desde el módulo de Cuentas por Pagar.";

                return RedirectToAction(nameof(Index));
            }

            if (esGastoDesdeFactura)
            {
                TempData["MensajeError"] =
                    "Este gasto fue generado automáticamente desde una factura. " +
                    "Debe modificarse desde el módulo de Facturación.";

                return RedirectToAction(nameof(Index));
            }
            if (gastoActual.Estado == "Aprobado"
    //  || gastoActual.Estado == "Rechazado"
              || gastoActual.Estado == "Anulado")
            {
                TempData["MensajeError"] = "Este gasto no puede modificarse porque ya fue aprobado o anulado.";
                return RedirectToAction(nameof(Index));
            }

            gasto.Subtotal = Math.Round(gasto.Subtotal ?? 0, 2);
            gasto.Impuesto = Math.Round(gasto.Impuesto ?? 0, 2);
            gasto.Descuento = Math.Round(gasto.Descuento ?? 0, 2);
            gasto.MontoTotal = Math.Round(gasto.MontoTotal ?? 0, 2);

            if ((gasto.Descuento ?? 0) > (gasto.MontoTotal ?? 0))
            {
                ModelState.AddModelError("Descuento", "El descuento no puede ser mayor que el monto total.");
            }

            if (gasto.Fecha > DateOnly.FromDateTime(DateTime.Today))
            {
                ModelState.AddModelError("Fecha",
                    "La fecha del gasto no puede ser mayor a la fecha actual.");
            }
            // ============================================================
            // IMPEDIR MOVER EL GASTO HACIA UN PERÍODO CERRADO
            // ============================================================
            bool nuevoPeriodoCerrado =
                await PeriodoContableCerradoAsync(gasto.Fecha);

            if (nuevoPeriodoCerrado)
            {
                ModelState.AddModelError(
                    "Fecha",
                    "La nueva fecha corresponde a un período contable cerrado.");
            }
            // La categoría debe estar activa o ser la categoría histórica del gasto.
            bool categoriaPermitida = await _context.CategoriasGastos
                .AnyAsync(c =>
                    c.IdCategoriaGasto == gasto.CategoriaGastoId &&
                    (c.Activo ||
                     c.IdCategoriaGasto == gastoActual.CategoriaGastoId));

            if (!categoriaPermitida)
            {
                ModelState.AddModelError(
                    "CategoriaGastoId",
                    "La categoría seleccionada no está activa.");
            }


            // El proveedor debe estar activo o ser el proveedor histórico del gasto.
            if (gasto.ProveedorId.HasValue)
            {
                bool proveedorPermitido = await _context.Proveedores
                    .AnyAsync(p =>
                        p.IdProveedor == gasto.ProveedorId.Value &&
                        (p.Activo ||
                         p.IdProveedor == gastoActual.ProveedorId));

                if (!proveedorPermitido)
                {
                    ModelState.AddModelError(
                        "ProveedorId",
                        "El proveedor seleccionado no está activo.");
                }
            }


            // El centro debe estar activo o ser el centro histórico del gasto.
            if (gasto.CentroCostoId.HasValue)
            {
                bool centroPermitido = await _context.CentrosCostos
                    .AnyAsync(c =>
                        c.IdCentroCosto == gasto.CentroCostoId.Value &&
                        (c.Activo ||
                         c.IdCentroCosto == gastoActual.CentroCostoId));

                if (!centroPermitido)
                {
                    ModelState.AddModelError(
                        "CentroCostoId",
                        "El centro de costo seleccionado no está activo.");
                }
            }

            PresupuestosMensuale? presupuestoSeleccionado = null;

            if (!gasto.PresupuestoMensualId.HasValue)
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
                gasto.PresupuestoMensualId.Value &&

            (
                // Puede conservar el presupuesto histórico actual.
                pm.IdPresupuestoMensual ==
                    gastoActual.PresupuestoMensualId

                ||

                // O cambiar a uno actualmente permitido.
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
                    gasto.ProyectoId =
                        presupuestoSeleccionado.ProyectoId;
                }
            }
            if (ModelState.IsValid)
            {
                try
                {
                    // CALCULOS INTERNOS DE PRESUPUESTO, GASTADO AL MES , SALDO DISPONIBLE, ESTADO SI ESTA DENTRO DEL LIMITE O NO Suma gastos activos del mes excluye el gasto actualsuma el nuevo monto editado calcula presupuesto calcula saldo guarda la foto histórica
                    var totalGastadoMes = _context.Gastos
    .Where(g => g.CategoriaGastoId == gasto.CategoriaGastoId
    && g.Activo
    && g.Estado != "Anulado"
    && g.Estado != "Rechazado"
    && g.IdGasto != gasto.IdGasto
    && g.Fecha.Month == gasto.Fecha.Month
    && g.Fecha.Year == gasto.Fecha.Year)
      .Sum(g => g.MontoTotal ?? 0);

                    totalGastadoMes += gasto.MontoTotal ?? 0;

                    var categoria = await _context.CategoriasGastos
                        .FirstOrDefaultAsync(c => c.IdCategoriaGasto == gasto.CategoriaGastoId);

                    decimal presupuestoMensual = categoria?.LimiteMensual ?? 0;

                    decimal saldoDisponible = presupuestoMensual - totalGastadoMes;

                    // Guardar el estado histórico del presupuesto al momento de registrar el gasto
                    gasto.PresupuestoMensual = presupuestoMensual;
                    gasto.GastadoAcumulado = totalGastadoMes;
                    gasto.SaldoDisponible = saldoDisponible;


                    if (categoria != null)
                    {
                        gasto.SuperaPresupuesto = totalGastadoMes > presupuestoMensual;
                    }
                    /*  gasto.Activo = true;
                      gasto.Estado = gastoActual.Estado;
                      _context.Update(gasto);*/

                    gasto.Activo = true;

                    if (gastoActual.Estado == "Rechazado")
                    {
                        gasto.Estado = "Pendiente de autorización";

                        gasto.MotivoRechazo = null;
                        gasto.FechaRechazo = null;
                        gasto.UsuarioRechazoId = null;
                    }
                    else
                    {
                        gasto.Estado = gastoActual.Estado;
                    }

                    _context.Update(gasto);

                    await _context.SaveChangesAsync();
                    //auditoria
                    var cambios = new List<string>();

                    if (gastoActual.CategoriaGastoId != gasto.CategoriaGastoId)
                    {
                        cambios.Add("se modificó la categoría de gasto");
                    }

                    if (gastoActual.ProveedorId != gasto.ProveedorId)
                    {
                        cambios.Add("se modificó el proveedor");
                    }

                    if (gastoActual.CentroCostoId != gasto.CentroCostoId)
                    {
                        cambios.Add("se modificó el centro de costo");
                    }

                    if (gastoActual.PresupuestoMensualId != gasto.PresupuestoMensualId)
                    {
                        cambios.Add("se modificó el proyecto/período presupuestario");
                    }

                    if (gastoActual.Fecha != gasto.Fecha)
                    {
                        cambios.Add(
                            $"fecha: {gastoActual.Fecha:dd/MM/yyyy} → {gasto.Fecha:dd/MM/yyyy}");
                    }

                    if (gastoActual.Concepto != gasto.Concepto)
                    {
                        cambios.Add("se modificó el concepto");
                    }

                    if (gastoActual.MontoTotal != gasto.MontoTotal)
                    {
                        cambios.Add(
                            $"monto total: ₡{(gastoActual.MontoTotal ?? 0):N2} → ₡{(gasto.MontoTotal ?? 0):N2}");
                    }

                    if (gastoActual.Observacion != gasto.Observacion)
                    {
                        cambios.Add("se modificó la observación");
                    }
                    // Auditoría: registra la modificación del gasto.
                    string accionAuditoria =
       gastoActual.Estado == "Rechazado"
           ? "Corregir"
           : "Modificar";

                    string descripcionAuditoria;

                    if (gastoActual.Estado == "Rechazado")
                    {
                        descripcionAuditoria =
                            "Se corrigió el gasto rechazado " +
                            gasto.Concepto +
                            " y fue enviado nuevamente a autorización.";

                        if (cambios.Any())
                        {
                            descripcionAuditoria +=
                                " Cambios: " +
                                string.Join("; ", cambios) +
                                ".";
                        }
                    }
                    else
                    {
                        descripcionAuditoria =
                            "Se modificó el gasto " +
                            gasto.Concepto + ".";

                        if (cambios.Any())
                        {
                            descripcionAuditoria +=
                                " Cambios: " +
                                string.Join("; ", cambios) +
                                ".";
                        }
                    }

                    await RegistrarAuditoriaGasto(
                        accionAuditoria,
                        gasto.IdGasto,
                        descripcionAuditoria);

                    if (gastoActual.Estado == "Rechazado")
                    {
                        TempData["MensajeExito"] =
                            "El gasto fue corregido y enviado nuevamente a autorización.";
                    }
                    else if (gasto.SuperaPresupuesto)
                    {
                        TempData["MensajeAdvertencia"] =
                            "El gasto fue modificado correctamnete, pero supera el límite mensual de la categoría.";
                    }
                    else
                    {
                        TempData["MensajeExito"] =
                            "El gasto fue modificado correctamente.";
                    }
                }
                catch (DbUpdateConcurrencyException)
                {
                    if (!GastoExists(gasto.IdGasto))
                    {
                        return NotFound();
                    }
                    else
                    {
                        throw;
                    }
                }
                return RedirectToAction(nameof(Index));
            }
            await CargarCatalogosEditAsync(
     gasto,
     gastoActual);

            return View(gasto);
        }
        // POST: Gastos/Aprobar/5
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Aprobar(int id)
        {
            var rolUsuario = HttpContext.Session.GetString("Rol");

            if (rolUsuario != "Administrador" && rolUsuario != "Contador")
            {
                TempData["MensajeError"] = "No tiene permisos para aprobar o rechazar gastos.";
                return RedirectToAction(nameof(Index));
            }
            var gasto = await _context.Gastos.FindAsync(id);

            if (gasto == null)
            {
                return NotFound();
            }
            // ============================================================
            // BLOQUEO POR CIERRE CONTABLE
            // ============================================================
            bool periodoCerrado =
                await PeriodoContableCerradoAsync(gasto.Fecha);

            if (periodoCerrado)
            {
                TempData["MensajeError"] =
                    $"No se puede aprobar el gasto porque el período " +
                    $"{gasto.Fecha:MM/yyyy} se encuentra cerrado. " +
                    $"Para aprobarlo, primero debe reabrirse el período.";

                return RedirectToAction(
                    nameof(Index),
                    new
                    {
                        anio = gasto.Fecha.Year,
                        mes = gasto.Fecha.Month
                    });
            }

            if (gasto.Estado != "Pendiente de autorización")
            {
                TempData["MensajeError"] = "Solo se pueden aprobar gastos pendientes de autorización.";
                return RedirectToAction(nameof(Index));
            }

await CalcularControlCategoriaGastoAsync(gasto);

            gasto.Estado = "Aprobado";

            // Si el gasto nació automáticamente desde una CxP manual,
            // también se autoriza la cuenta por pagar para permitir registrar pagos.
            var cuentaPorPagar = await _context.CuentasPorPagars
                .FirstOrDefaultAsync(c =>
                    c.Origen == "Manual" &&
                    c.NumeroDocumento == gasto.NumeroFactura);

            if (cuentaPorPagar != null)
            {
                cuentaPorPagar.EstadoAutorizacion = "Aprobado";
                _context.Update(cuentaPorPagar);
            }

            _context.Update(gasto);
            await _context.SaveChangesAsync();

            // CAMBIO APLICADO - ASIENTO GASTO CXP MANUAL
            // Si el gasto nació desde una CxP manual, NO debe afectar Caja.
            // Debe crear:
            // Debe: cuenta de gasto
            // Haber: CC002 Cuentas por Pagar
            //
            // Si NO es CxP manual, mantiene la lógica normal existente.
            // ================================================================
            if (cuentaPorPagar != null && cuentaPorPagar.Origen == "Manual")
            {
                await RegistrarAsientoGastoCxPManual(
                    gasto,
                    cuentaPorPagar,
                    HttpContext.Session.GetInt32("UsuarioId") ?? 1);
            }
            else
            {
                await _contabilidadService.RegistrarGastoAsync(
                    gasto,
                    HttpContext.Session.GetInt32("UsuarioId") ?? 1);
            }

            // Auditoría: registra la aprobación del gasto.
            await RegistrarAuditoriaGasto(
                "Aprobar",
                gasto.IdGasto,
                "Se aprobó el gasto " + (gasto.NumeroFactura ?? gasto.Concepto) + ".");

            TempData["MensajeExito"] = "El gasto fue aprobado correctamente.";
            return RedirectToAction(nameof(Index));
        }

        // POST: Gastos/Rechazar/5
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Rechazar(int id, string motivoRechazo)
        {
            var rolUsuario = HttpContext.Session.GetString("Rol");

            if (rolUsuario != "Administrador" && rolUsuario != "Contador")
            {
                TempData["MensajeError"] =
                    "No tiene permisos para aprobar o rechazar gastos.";

                return RedirectToAction(nameof(Index));
            }

            if (string.IsNullOrWhiteSpace(motivoRechazo))
            {
                TempData["MensajeError"] =
                    "Debe indicar el motivo del rechazo.";

                return RedirectToAction(nameof(Index));
            }

            motivoRechazo = motivoRechazo.Trim();

            if (motivoRechazo.Length < 5)
            {
                TempData["MensajeError"] =
                    "El motivo del rechazo debe contener al menos 5 caracteres.";

                return RedirectToAction(nameof(Index));
            }

            if (motivoRechazo.Length > 500)
            {
                TempData["MensajeError"] =
                    "El motivo del rechazo no puede superar los 500 caracteres.";

                return RedirectToAction(nameof(Index));
            }

            var gasto = await _context.Gastos.FindAsync(id);

            if (gasto == null)
            {
                return NotFound();
            }
            // ============================================================
            // BLOQUEO POR CIERRE CONTABLE
            // ============================================================
            bool periodoCerrado =
                await PeriodoContableCerradoAsync(gasto.Fecha);

            if (periodoCerrado)
            {
                TempData["MensajeError"] =
                    $"No se puede rechazar el gasto porque el período " +
                    $"{gasto.Fecha:MM/yyyy} se encuentra cerrado. " +
                    $"Para realizar cambios, primero debe reabrirse el período.";

                return RedirectToAction(
                    nameof(Index),
                    new
                    {
                        anio = gasto.Fecha.Year,
                        mes = gasto.Fecha.Month
                    });
            }
            if (gasto.Estado != "Pendiente de autorización")
            {
                TempData["MensajeError"] =
                    "Solo se pueden rechazar gastos pendientes de autorización.";

                return RedirectToAction(nameof(Index));
            }

            gasto.Estado = "Rechazado";
            gasto.MotivoRechazo = motivoRechazo;
            gasto.FechaRechazo = DateTime.Now;
            gasto.UsuarioRechazoId =
                HttpContext.Session.GetInt32("UsuarioId") ?? 1;

            // Si el gasto pertenece a una CxP manual,
            // también se rechaza la autorización de la cuenta origen.
            var cuentaPorPagar = await _context.CuentasPorPagars
                .FirstOrDefaultAsync(c =>
                    c.Origen == "Manual" &&
                    c.NumeroDocumento == gasto.NumeroFactura);

            if (cuentaPorPagar != null)
            {
                cuentaPorPagar.EstadoAutorizacion = "Rechazado";
                _context.Update(cuentaPorPagar);
            }

            _context.Update(gasto);
            await _context.SaveChangesAsync();

            await RegistrarAuditoriaGasto(
                "Rechazar",
                gasto.IdGasto,
                "Se rechazó el gasto " +
                (gasto.NumeroFactura ?? gasto.Concepto) +
                ". Motivo: " + motivoRechazo);

            TempData["MensajeExito"] =
                "El gasto fue rechazado correctamente.";

            return RedirectToAction(nameof(Index));
        }

        // GET: Gastos/eliminar
        public async Task<IActionResult> Delete(int? id)
        {
            if (id == null)
            {
                return NotFound();
            }
            var gasto = await _context.Gastos
                .Include(g => g.CategoriaGasto)
                .Include(g => g.Proveedor)
                .Include(g => g.Proyecto)
                .Include(g => g.PresupuestoMensualRegistro)
                .FirstOrDefaultAsync(m => m.IdGasto == id);
            if (gasto == null)
            {
                return NotFound();
            }
            // EVITAR ANULAR NUEVAMENTE UN GASTO YA ANULADO
            if (!gasto.Activo ||
    gasto.Estado == "Anulado")
            {
                TempData["MensajeError"] =
                    "Este gasto ya se encuentra anulado y no puede anularse nuevamente.";

                return RedirectToAction(
                    nameof(Index),
                    new
                    {
                        anio = gasto.Fecha.Year,
                        mes = gasto.Fecha.Month
                    });
            }
            // BLOQUEO POR CIERRE CONTABLE
            bool periodoCerrado =
    await PeriodoContableCerradoAsync(gasto.Fecha);

            if (periodoCerrado)
            {
                TempData["MensajeError"] =
                    $"No se puede anular el gasto porque el período " +
                    $"{gasto.Fecha:MM/yyyy} se encuentra cerrado. " +
                    $"Para realizar correcciones, primero debe reabrirse el período.";

                return RedirectToAction(
                    nameof(Index),
                    new
                    {
                        anio = gasto.Fecha.Year,
                        mes = gasto.Fecha.Month
                    });
            }
           

            return View(gasto);
        }

        // POST: Gastos/eliminar
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> DeleteConfirmed(int id, string motivoAnulacion)
        {
            if (string.IsNullOrWhiteSpace(motivoAnulacion))
            {
                TempData["MensajeError"] =
                    "Debe indicar el motivo de anulación.";

                return RedirectToAction(nameof(Delete), new { id });
            }

            motivoAnulacion = motivoAnulacion.Trim();

            if (motivoAnulacion.Length < 5)
            {
                TempData["MensajeError"] =
                    "El motivo de anulación debe contener al menos 5 caracteres.";

                return RedirectToAction(nameof(Delete), new { id });
            }

            if (motivoAnulacion.Length > 500)
            {
                TempData["MensajeError"] =
                    "El motivo de anulación no puede superar los 500 caracteres.";

                return RedirectToAction(nameof(Delete), new { id });
            }
            var gasto = await _context.Gastos.FindAsync(id);

            if (gasto == null)
            {
         
            // ============================================================
            // EVITAR ANULAR NUEVAMENTE UN GASTO YA ANULADO
            // ============================================================
            if (!gasto.Activo ||
                    gasto.Estado == "Anulado")
                {
                    TempData["MensajeError"] =
                        "Este gasto ya se encuentra anulado y no puede anularse nuevamente.";

                    return RedirectToAction(
                        nameof(Index),
                        new
                        {
                            anio = gasto.Fecha.Year,
                            mes = gasto.Fecha.Month
                        });
                }
                // ============================================================
                // BLOQUEO POR CIERRE CONTABLE
                // ============================================================
                bool periodoCerrado =
                    await PeriodoContableCerradoAsync(gasto.Fecha);

                if (periodoCerrado)
                {
                    TempData["MensajeError"] =
                        $"No se puede anular el gasto porque el período " +
                        $"{gasto.Fecha:MM/yyyy} se encuentra cerrado. " +
                        $"Para realizar correcciones, primero debe reabrirse el período.";

                    return RedirectToAction(
                        nameof(Index),
                        new
                        {
                            anio = gasto.Fecha.Year,
                            mes = gasto.Fecha.Month
                        });
                }

            

                bool esGastoDesdeCxP =
    !string.IsNullOrWhiteSpace(gasto.NumeroFactura) &&
    gasto.Concepto != null &&
    gasto.Concepto.ToLower().Contains("cuenta por pagar");

                bool esGastoDesdeFactura =
                    !string.IsNullOrWhiteSpace(gasto.NumeroFactura) &&
                    gasto.Concepto != null &&
                    gasto.Concepto.ToLower().Contains("factura");

                bool esGastoManual =
                    !esGastoDesdeCxP &&
                    !esGastoDesdeFactura;

                if (!esGastoManual)
                {
                    TempData["MensajeError"] = "Solo se pueden anular gastos registrados manualmente desde el módulo de Gastos.";
                    return RedirectToAction(nameof(Index));
                }

                // Reversa los movimientos contables originales antes de anular el gasto.
                await _contabilidadService.ReversarMovimientosAsync(
                    "Gastos",
                    gasto.IdGasto,
                    HttpContext.Session.GetInt32("UsuarioId") ?? 1,
                    "Anulación de gasto");


                gasto.Activo = false;
                gasto.Estado = "Anulado";
                gasto.MotivoAnulacion = motivoAnulacion;
                gasto.FechaAnulacion = DateTime.Now;
                gasto.UsuarioAnulacionId =
                    HttpContext.Session.GetInt32("UsuarioId") ?? 1;

                _context.Update(gasto);
                // Auditoría: registra la anulación del gasto.
                await RegistrarAuditoriaGasto(
     "Anular",
     gasto.IdGasto,
     "Se anuló el gasto " + gasto.Concepto +
     " por un monto de ₡" +
     (gasto.MontoTotal ?? 0).ToString("N2") +
     ". Motivo: " + motivoAnulacion);

                TempData["MensajeExito"] = "El gasto fue anulado correctamente.";
            }

            await _context.SaveChangesAsync();
            return RedirectToAction(nameof(Index));
        }
        // presupuesto de ctaegorias gasto 
        private async Task CalcularControlCategoriaGastoAsync(Gasto gasto)
        {
            decimal limiteCategoriaGasto = await _context.CategoriasGastos
                .AsNoTracking()
                .Where(c => c.IdCategoriaGasto == gasto.CategoriaGastoId)
                .Select(c => (decimal?)c.LimiteMensual)
                .FirstOrDefaultAsync() ?? 0m;

            decimal totalAprobadoCategoria = await _context.Gastos
                .AsNoTracking()
                .Where(g =>
                    g.CategoriaGastoId == gasto.CategoriaGastoId &&
                    g.IdGasto != gasto.IdGasto &&
                    g.Activo &&
                    g.Estado == "Aprobado" &&
                    g.Fecha.Month == gasto.Fecha.Month &&
                    g.Fecha.Year == gasto.Fecha.Year)
                .SumAsync(g => g.MontoTotal ?? 0m);

            decimal montoGastoEvaluado =
                gasto.Estado == "Pendiente de autorización" ||
                gasto.Estado == "Aprobado"
                    ? gasto.MontoTotal ?? 0m
                    : 0m;

            decimal totalProyectadoCategoria =
                totalAprobadoCategoria + montoGastoEvaluado;

            decimal saldoCategoria =
                limiteCategoriaGasto - totalProyectadoCategoria;

            gasto.PresupuestoMensual = limiteCategoriaGasto;
            gasto.GastadoAcumulado = totalAprobadoCategoria;
            gasto.SaldoDisponible = saldoCategoria;

            gasto.SuperaPresupuesto =
                limiteCategoriaGasto > 0 &&
                totalProyectadoCategoria > limiteCategoriaGasto;
        }
        // ================================================================
        // CONSTRUYE LA DESCRIPCIÓN DE LOS FILTROS USADOS EN LOS REPORTES
        // ================================================================
        private async Task<string> ConstruirDescripcionFiltrosReporte(
            string? buscar,
            string? estado,
            string? origen,
            int? proyectoId,
            int? anio,
            int? mes,
            DateOnly? fechaInicio,
            DateOnly? fechaFin,
            bool? excedido)
        {
            var filtros = new List<string>();

            if (!string.IsNullOrWhiteSpace(buscar))
            {
                filtros.Add($"Búsqueda={buscar.Trim()}");
            }

            if (!string.IsNullOrWhiteSpace(estado))
            {
                filtros.Add($"Estado={estado}");
            }

            if (!string.IsNullOrWhiteSpace(origen))
            {
                filtros.Add($"Origen={origen}");
            }

            if (proyectoId.HasValue)
            {
                string proyecto = await _context.Proyectos
                    .AsNoTracking()
                    .Where(p =>
                        p.IdProyecto == proyectoId.Value)
                    .Select(p => p.Nombre)
                    .FirstOrDefaultAsync()
                    ?? proyectoId.Value.ToString();

                filtros.Add($"Proyecto={proyecto}");
            }

            if (anio.HasValue)
            {
                filtros.Add($"Año={anio.Value}");
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

                filtros.Add($"Mes={nombreMes}");
            }

            if (fechaInicio.HasValue)
            {
                filtros.Add(
                    $"Desde={fechaInicio.Value:dd/MM/yyyy}");
            }

            if (fechaFin.HasValue)
            {
                filtros.Add(
                    $"Hasta={fechaFin.Value:dd/MM/yyyy}");
            }

            if (excedido == true)
            {
                filtros.Add(
                    "Límite de categoría excedido=Sí");
            }

            return filtros.Any()
                ? "Filtros aplicados: " +
                  string.Join(", ", filtros) +
                  "."
                : "Sin filtros aplicados.";
        }
        // Registra auditoría de las acciones realizadas en el módulo de gastos.
        private async Task RegistrarAuditoriaGasto(string accion, int registroId, string descripcion)
        {
            int usuarioId = HttpContext.Session.GetInt32("UsuarioId") ?? 1;

            var auditoria = new Auditorium
            {
                UsuarioId = usuarioId,
                Tabla = "Gastos",
                RegistroId = registroId,
                Accion = accion,
                Descripcion = descripcion,
                Fecha = DateTime.Now
            };

            _context.Auditoria.Add(auditoria);
            await _context.SaveChangesAsync();
        }

        // Carga los catálogos permitidos para registrar un gasto nuevo.
        private async Task CargarCatalogosCreateAsync(Gasto? gasto = null)
        {
            ViewData["CategoriaGastoId"] = new SelectList(
                await _context.CategoriasGastos
                    .AsNoTracking()
                    .Where(c => c.Activo)
                    .OrderBy(c => c.Nombre)
                    .ToListAsync(),
                "IdCategoriaGasto",
                "Nombre",
                gasto?.CategoriaGastoId
            );

            ViewData["ProveedorId"] = new SelectList(
                await _context.Proveedores
                    .AsNoTracking()
                    .Where(p => p.Activo)
                    .OrderBy(p => p.Nombre)
                    .ToListAsync(),
                "IdProveedor",
                "Nombre",
                gasto?.ProveedorId
            );

            ViewData["CentroCostoId"] = new SelectList(
                await _context.CentrosCostos
                    .AsNoTracking()
                    .Where(c => c.Activo)
                    .OrderBy(c => c.Nombre)
                    .ToListAsync(),
                "IdCentroCosto",
                "Nombre",
                gasto?.CentroCostoId
            );
            // cargar solo proy con pres aprobado 
            var hoy = DateOnly.FromDateTime(DateTime.Today);
            var fechaMinima = hoy.AddMonths(-3);

            int anioMinimo = fechaMinima.Year;
            int mesMinimo = fechaMinima.Month;

            var presupuestosPermitidos =
                await _context.PresupuestosMensuales
                    .AsNoTracking()
                    .Include(pm => pm.Proyecto)
                    .Where(pm =>
                        pm.EstadoAprobacion == "Aprobado" &&
                        pm.ProyectoId != null &&
                        pm.Proyecto != null &&
                        pm.Proyecto.Activo &&

            // No mostrar períodos presupuestarios
            // pertenecientes a meses contablemente cerrados.
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
                    .ToListAsync();

            ViewData["PresupuestoMensualId"] =
                new SelectList(
                    presupuestosPermitidos,
                    "IdPresupuestoMensual",
                    "Nombre",
                    gasto?.PresupuestoMensualId
                );
        }
        // Carga los catálogos para modificar un gasto.
        // Conserva el registro que ya estaba seleccionado aunque después fuera desactivado.
        private async Task CargarCatalogosEditAsync(
     Gasto gasto,
     Gasto? gastoHistorico = null)
        {
            int categoriaHistorica =
                gastoHistorico?.CategoriaGastoId
                ?? gasto.CategoriaGastoId;

            int? proveedorHistorico =
                gastoHistorico?.ProveedorId
                ?? gasto.ProveedorId;

            int? centroHistorico =
                gastoHistorico?.CentroCostoId
                ?? gasto.CentroCostoId;

       

            ViewData["CategoriaGastoId"] = new SelectList(
                await _context.CategoriasGastos
                    .AsNoTracking()
                    .Where(c =>
                        c.Activo ||
                        c.IdCategoriaGasto == categoriaHistorica)
                    .OrderBy(c => c.Nombre)
                    .ToListAsync(),
                "IdCategoriaGasto",
                "Nombre",
                gasto.CategoriaGastoId);

            ViewData["ProveedorId"] = new SelectList(
                await _context.Proveedores
                    .AsNoTracking()
                    .Where(p =>
                        p.Activo ||
                        p.IdProveedor == proveedorHistorico)
                    .OrderBy(p => p.Nombre)
                    .ToListAsync(),
                "IdProveedor",
                "Nombre",
                gasto.ProveedorId);

            ViewData["CentroCostoId"] = new SelectList(
                await _context.CentrosCostos
                    .AsNoTracking()
                    .Where(c =>
                        c.Activo ||
                        c.IdCentroCosto == centroHistorico)
                    .OrderBy(c => c.Nombre)
                    .ToListAsync(),
                "IdCentroCosto",
                "Nombre",
                gasto.CentroCostoId);

            int? presupuestoHistorico =
    gastoHistorico?.PresupuestoMensualId
    ?? gasto.PresupuestoMensualId;

            var hoy = DateOnly.FromDateTime(DateTime.Today);
            var fechaMinima = hoy.AddMonths(-3);

            int anioMinimo = fechaMinima.Year;
            int mesMinimo = fechaMinima.Month;

            var presupuestosPermitidos =
      await _context.PresupuestosMensuales
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

              // Conserva únicamente el presupuesto histórico actual.
              ||
              pm.IdPresupuestoMensual == presupuestoHistorico
          )
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
                    .ToListAsync();

            ViewData["PresupuestoMensualId"] =
                new SelectList(
                    presupuestosPermitidos,
                    "IdPresupuestoMensual",
                    "Nombre",
                    gasto.PresupuestoMensualId
                );
        }


        private bool GastoExists(int id)
        {
            return _context.Gastos.Any(e => e.IdGasto == id);
        }
        // ================================================================
        // CAMBIO APLICADO - REGISTRO CONTABLE CXP MANUAL APROBADA
        // Este método solo aplica para gastos nacidos desde CxP manual.
        // No toca facturas.
        // No toca gastos normales.
        // No toca pagos.
        // ================================================================
        private async Task RegistrarAsientoGastoCxPManual(
            Gasto gasto,
            CuentasPorPagar cuentaPorPagar,
            int usuarioId)
        {
            decimal monto = gasto.MontoTotal ?? 0;

            if (monto <= 0)
            {
                return;
            }

            string referencia = cuentaPorPagar.NumeroDocumento ?? gasto.NumeroFactura ?? "";

            if (string.IsNullOrWhiteSpace(referencia))
            {
                return;
            }

            bool yaExiste = await _context.MovimientosContables
                .AnyAsync(m =>
                    m.Referencia == referencia &&
                    m.OrigenModulo == "Gastos" &&
                    m.OrigenId == gasto.IdGasto &&
                    m.Estado == "Registrado" &&
                    m.Anulado == false);

            if (yaExiste)
            {
                return;
            }

            var categoriaGasto = await _context.CategoriasGastos
      .Include(c => c.CuentaContable)
      .FirstOrDefaultAsync(c => c.IdCategoriaGasto == gasto.CategoriaGastoId);

            var cuentaGasto = categoriaGasto?.CuentaContable;

            var cuentaCxP = await _context.CuentasContables
                .FirstOrDefaultAsync(c => c.Codigo == "CC002");

            if (cuentaGasto == null || cuentaCxP == null)
            {
                return;
            }

            var movimientoDebe = new MovimientosContable
            {
                Fecha = gasto.Fecha,
                CuentaContableId = cuentaGasto.IdCuentaContable,
                ProyectoId = gasto.ProyectoId,
                CentroCostoId = gasto.CentroCostoId,
                TipoMovimiento = "Debe",
                Monto = monto,
                Debe = monto,
                Haber = 0,
                Referencia = referencia,
                OrigenModulo = "Gastos",
                OrigenId = gasto.IdGasto,
                Descripcion = "Registro contable de gasto por CxP manual " + referencia,
                Estado = "Registrado",
                EsAutomatico = true,
                UsuarioId = usuarioId,
                FechaCreacion = DateTime.Now
            };

            var movimientoHaber = new MovimientosContable
            {
                Fecha = gasto.Fecha,
                CuentaContableId = cuentaCxP.IdCuentaContable,
                ProyectoId = gasto.ProyectoId,
                CentroCostoId = gasto.CentroCostoId,
                TipoMovimiento = "Haber",
                Monto = monto,
                Debe = 0,
                Haber = monto,
                Referencia = referencia,
                OrigenModulo = "Gastos",
                OrigenId = gasto.IdGasto,
                Descripcion = "Registro de cuenta por pagar por CxP manual " + referencia,
                Estado = "Registrado",
                EsAutomatico = true,
                UsuarioId = usuarioId,
                FechaCreacion = DateTime.Now
            };

            _context.MovimientosContables.Add(movimientoDebe);
            _context.MovimientosContables.Add(movimientoHaber);

            AplicarSaldoCuenta(cuentaGasto, monto, 0);
            AplicarSaldoCuenta(cuentaCxP, 0, monto);

            _context.CuentasContables.Update(cuentaGasto);
            _context.CuentasContables.Update(cuentaCxP);

            await _context.SaveChangesAsync();
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
    }
}
