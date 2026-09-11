using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using SistemaGestionFinanciera.Data;
using SistemaGestionFinanciera.Models;
using ClosedXML.Excel;
using System.Globalization;

namespace SistemaGestionFinanciera.Controllers
{
    public class MovimientosContablesController : BaseController
    {
        private readonly SistemaFinancieroContext _context;

        // Constructor del controlador.
        // Inicializa el acceso a la base de datos.
        public MovimientosContablesController(SistemaFinancieroContext context)
        {
            _context = context;
        }
        // CONSULTA GENERAL DE MOVIMIENTOS CONTABLES
        // Este método muestra el libro diario del sistema.
        // Permite consultar los movimientos contables generadosautomáticamente desde los diferentes módulos:
        //  Ingresos, Gastos, Facturación, Cuentas por Cobrar, Cuentas por Pagar
        // tambien calcula automáticamente los indicadores financieros mostrados en la parte superior de la vista.
        public async Task<IActionResult> Index(
            string buscar,
            int? cuentaContableId,
            string tipoMovimiento,
            string origenModulo,
            int? proyectoId,
            int? centroCostoId,
            string estado,
            int? anio,
            int? mes,
            DateOnly? desde,
            DateOnly? hasta,
              string? origen)
        {
            // Consulta principal de movimientos contables.
            var movimientos = _context.MovimientosContables
                .Include(m => m.CuentaContable)
                .Include(m => m.Proyecto)
                .Include(m => m.CentroCosto)
                .Include(m => m.Usuario)
                .AsQueryable();

            // Buscar por descripción, referencia, módulo o cuenta contable.
            if (!string.IsNullOrWhiteSpace(buscar))
            {
                movimientos = movimientos.Where(m =>
                    (m.Descripcion != null && m.Descripcion.Contains(buscar)) ||
                    (m.Referencia != null && m.Referencia.Contains(buscar)) ||
                    (m.OrigenModulo != null && m.OrigenModulo.Contains(buscar)) ||
                    m.CuentaContable.Nombre.Contains(buscar) ||
                    m.CuentaContable.Codigo.Contains(buscar));
            }

            // Filtrar por cuenta contable.
            if (cuentaContableId.HasValue)
                movimientos = movimientos.Where(m => m.CuentaContableId == cuentaContableId.Value);

            // Filtrar por tipo de movimiento.
            if (!string.IsNullOrWhiteSpace(tipoMovimiento))
                movimientos = movimientos.Where(m => m.TipoMovimiento == tipoMovimiento);

            // Filtrar por módulo de origen.
            if (!string.IsNullOrWhiteSpace(origenModulo))
                movimientos = movimientos.Where(m => m.OrigenModulo == origenModulo);

            // Filtrar por proyecto.
            if (proyectoId.HasValue)
                movimientos = movimientos.Where(m => m.ProyectoId == proyectoId.Value);

            // Filtrar por centro de costo.
            if (centroCostoId.HasValue)
                movimientos = movimientos.Where(m => m.CentroCostoId == centroCostoId.Value);

            // Filtrar por estado del movimiento.
            if (!string.IsNullOrWhiteSpace(estado))
                movimientos = movimientos.Where(m => m.Estado == estado);

            // Filtrar por año contable.
            if (anio.HasValue)
                movimientos = movimientos.Where(m => m.Fecha.Year == anio.Value);

            // Filtrar por mes contable.
            if (mes.HasValue)
                movimientos = movimientos.Where(m => m.Fecha.Month == mes.Value);

            // Filtrar por rango de fechas.
            if (desde.HasValue)
                movimientos = movimientos.Where(m => m.Fecha >= desde.Value);

            if (hasta.HasValue)
                movimientos = movimientos.Where(m => m.Fecha <= hasta.Value);

            // Ordenar los movimientos desde el más reciente.
            var lista = await movimientos
      .OrderByDescending(m => m.IdMovimientoContable)
      .ToListAsync();

            // ===========================================================
            // INDICADORES DEL LIBRO DIARIO
            // ===========================================================

            // Movimientos vigentes dentro de los resultados filtrados.
            var movimientosVigentes = lista
                .Where(m => m.Anulado != true &&
                            m.Estado == "Registrado")
                .ToList();

            // Totales vigentes de los resultados filtrados.
            decimal totalDebeFiltrado =
                movimientosVigentes.Sum(m => m.Debe);

            decimal totalHaberFiltrado =
                movimientosVigentes.Sum(m => m.Haber);

            decimal diferenciaFiltrada =
                totalDebeFiltrado - totalHaberFiltrado;

            ViewBag.TotalDebe = totalDebeFiltrado;
            ViewBag.TotalHaber = totalHaberFiltrado;
            ViewBag.Diferencia = diferenciaFiltrada;

            // Cantidad total de filas encontradas con los filtros.
            ViewBag.TotalMovimientos = lista.Count;

            // Cantidad de movimientos anulados dentro del resultado filtrado.
            ViewBag.TotalAnulados = lista.Count(m =>
                m.Anulado == true ||
                m.Estado == "Anulado");

            // ===========================================================
            // RESUMEN FINANCIERO GENERAL
            // Se obtiene desde el catálogo de cuentas contables.
            // Estos valores representan el estado financiero acumulado.
            // ===========================================================
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

            // Ecuación contable general.
            decimal totalPatrimonio =
                totalActivos - totalPasivos;

            // Resultado financiero acumulado.
            decimal resultadoPeriodo =
                totalIngresos - totalGastos;

            ViewBag.TotalActivos = totalActivos;
            ViewBag.TotalPasivos = totalPasivos;
            ViewBag.TotalPatrimonio = totalPatrimonio;
            ViewBag.TotalIngresos = totalIngresos;
            ViewBag.TotalGastos = totalGastos;
            ViewBag.ResultadoPeriodo = resultadoPeriodo;

            // Obtener la fecha y hora del último movimiento registrado.
            ViewBag.UltimaSincronizacion = await _context.MovimientosContables
                .Where(m => m.FechaCreacion != null)
                .OrderByDescending(m => m.FechaCreacion)
                .Select(m => m.FechaCreacion)
                .FirstOrDefaultAsync();

            // Mantener los filtros aplicados en la consulta.
            ViewBag.Buscar = buscar;
            ViewBag.CuentaContableId = cuentaContableId;
            ViewBag.TipoMovimiento = tipoMovimiento;
            ViewBag.OrigenModulo = origenModulo;
            ViewBag.ProyectoId = proyectoId;
            ViewBag.CentroCostoId = centroCostoId;
            ViewBag.Estado = estado;
            ViewBag.Anio = anio;
            ViewBag.Mes = mes;
            ViewBag.Desde = desde?.ToString("yyyy-MM-dd");
            ViewBag.Hasta = hasta?.ToString("yyyy-MM-dd");

            // Cargar los catálogos utilizados por los filtros.
            ViewBag.CuentasContables = await _context.CuentasContables
                .Where(c => c.Activo)
                .OrderBy(c => c.Codigo)
                .ToListAsync();

            ViewBag.Proyectos = await _context.Proyectos
                .Where(p => p.Activo)
                .OrderBy(p => p.Nombre)
                .ToListAsync();

            ViewBag.CentrosCosto = await _context.CentrosCostos
                .Where(c => c.Activo)
                .OrderBy(c => c.Nombre)
                .ToListAsync();

            ViewBag.Origen = origen;

            return View(lista);
        }

        // ===========================================================
        // DETALLE DEL MOVIMIENTO CONTABLE modificado  al crear libro mayor 

        public async Task<IActionResult> Details(
     int? id,
     int? cuentaContableId,
     string origenVista,
     bool volverLibroMayor = false,
     bool volverDetalleCuenta = false)
        {
            if (id == null)
                return NotFound();

            var movimiento = await _context.MovimientosContables
                .Include(m => m.CuentaContable)
                .Include(m => m.Proyecto)
                .Include(m => m.CentroCosto)
                .Include(m => m.Usuario)
                .Include(m => m.MovimientoReversion)
                .FirstOrDefaultAsync(m => m.IdMovimientoContable == id);

            if (movimiento == null)
                return NotFound();

            ViewBag.CuentaContableId = cuentaContableId;
            ViewBag.OrigenVista = origenVista;
            ViewBag.VolverLibroMayor = volverLibroMayor;
            ViewBag.VolverDetalleCuenta = volverDetalleCuenta;

            return View(movimiento);
        }

        private IQueryable<MovimientosContable> ConstruirConsultaReporte(
    string? buscar,
    int? cuentaContableId,
    string? tipoMovimiento,
    string? origenModulo,
    int? proyectoId,
    int? centroCostoId,
    string? estado,
    int? anio,
    int? mes,
    DateOnly? desde,
    DateOnly? hasta)
        {
            var consulta = _context.MovimientosContables
                .Include(m => m.CuentaContable)
                .Include(m => m.Proyecto)
                .Include(m => m.CentroCosto)
                .Include(m => m.Usuario)
                .AsQueryable();

            if (!string.IsNullOrWhiteSpace(buscar))
            {
                consulta = consulta.Where(m =>
                    (m.Descripcion != null && m.Descripcion.Contains(buscar)) ||
                    (m.Referencia != null && m.Referencia.Contains(buscar)) ||
                    (m.OrigenModulo != null && m.OrigenModulo.Contains(buscar)) ||
                    (m.CuentaContable != null &&
                        (m.CuentaContable.Nombre.Contains(buscar) ||
                         m.CuentaContable.Codigo.Contains(buscar))));
            }

            if (cuentaContableId.HasValue)
            {
                consulta = consulta.Where(m =>
                    m.CuentaContableId == cuentaContableId.Value);
            }

            if (!string.IsNullOrWhiteSpace(tipoMovimiento))
            {
                consulta = consulta.Where(m =>
                    m.TipoMovimiento == tipoMovimiento);
            }

            if (!string.IsNullOrWhiteSpace(origenModulo))
            {
                consulta = consulta.Where(m =>
                    m.OrigenModulo == origenModulo);
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
                if (estado == "Anulado")
                {
                    consulta = consulta.Where(m =>
                        m.Anulado == true ||
                        m.Estado == "Anulado");
                }
                else if (estado == "Reversado")
                {
                    consulta = consulta.Where(m =>
                        m.TipoMovimiento == "Reversión" ||
                        m.Estado == "Reversado");
                }
                else
                {
                    consulta = consulta.Where(m =>
                        m.Estado == estado);
                }
            }

            if (anio.HasValue)
            {
                consulta = consulta.Where(m =>
                    m.Fecha.Year == anio.Value);
            }

            if (mes.HasValue)
            {
                consulta = consulta.Where(m =>
                    m.Fecha.Month == mes.Value);
            }

            if (desde.HasValue)
            {
                consulta = consulta.Where(m =>
                    m.Fecha >= desde.Value);
            }

            if (hasta.HasValue)
            {
                consulta = consulta.Where(m =>
                    m.Fecha <= hasta.Value);
            }

            return consulta;
        }

        public async Task<IActionResult> Reporte(
    string? buscar,
    int? cuentaContableId,
    string? tipoMovimiento,
    string? origenModulo,
    int? proyectoId,
    int? centroCostoId,
    string? estado,
    int? anio,
    int? mes,
    DateOnly? desde,
    DateOnly? hasta)
        {
            var consulta = ConstruirConsultaReporte(
                buscar,
                cuentaContableId,
                tipoMovimiento,
                origenModulo,
                proyectoId,
                centroCostoId,
                estado,
                anio,
                mes,
                desde,
                hasta);

            var movimientos = await consulta
                .OrderByDescending(m => m.Fecha)
                .ThenByDescending(m => m.IdMovimientoContable)
                .ToListAsync();

            var vigentes = movimientos
                .Where(m =>
                    m.Anulado != true &&
                    m.Estado == "Registrado")
                .ToList();

            ViewBag.TotalDebe = vigentes.Sum(m => m.Debe);
            ViewBag.TotalHaber = vigentes.Sum(m => m.Haber);
            ViewBag.Diferencia =
                (decimal)ViewBag.TotalDebe -
                (decimal)ViewBag.TotalHaber;

            ViewBag.TotalRegistros = movimientos.Count;
            ViewBag.FechaGeneracion = DateTime.Now;

            ViewBag.Buscar = buscar;
            ViewBag.CuentaContableId = cuentaContableId;
            ViewBag.TipoMovimiento = tipoMovimiento;
            ViewBag.OrigenModulo = origenModulo;
            ViewBag.ProyectoId = proyectoId;
            ViewBag.CentroCostoId = centroCostoId;
            ViewBag.Estado = estado;
            ViewBag.Anio = anio;
            ViewBag.Mes = mes;
            ViewBag.Desde = desde?.ToString("yyyy-MM-dd");
            ViewBag.Hasta = hasta?.ToString("yyyy-MM-dd");

            await RegistrarAuditoria(
    "MovimientosContables",
    0,
    "Vista previa PDF",
    "Se generó la vista previa del reporte de movimientos contables."
);
            return View(movimientos);
        }

        public async Task<IActionResult> ExportarExcel(
    string? buscar,
    int? cuentaContableId,
    string? tipoMovimiento,
    string? origenModulo,
    int? proyectoId,
    int? centroCostoId,
    string? estado,
    int? anio,
    int? mes,
    DateOnly? desde,
    DateOnly? hasta)
        {
            var consulta = ConstruirConsultaReporte(
    buscar,
    cuentaContableId,
    tipoMovimiento,
    origenModulo,
    proyectoId,
    centroCostoId,
    estado,
    anio,
    mes,
    desde,
    hasta);

            var movimientos = await consulta
                .OrderByDescending(m => m.Fecha)
                .ThenByDescending(m => m.IdMovimientoContable)
                .ToListAsync();

            using var workbook = new XLWorkbook();

            var hoja = workbook.Worksheets.Add("Movimientos Contables");

            // ===========================================================
            // ENCABEZADO DEL REPORTE
            // ===========================================================

            const int totalColumnas = 13;

            hoja.Cell("A1").Value =
                "Unión Cantonal de Asociaciones de Desarrollo de Santa Ana";

            hoja.Range(1, 1, 1, totalColumnas).Merge();

            hoja.Cell("A2").Value =
                "Reporte de Movimientos Contables";

            hoja.Range(2, 1, 2, totalColumnas).Merge();

            hoja.Cell("A3").Value =
                $"Generado el {DateTime.Now:dd/MM/yyyy HH:mm}";

            hoja.Range(3, 1, 3, totalColumnas).Merge();

            // Institución.
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

            encabezadoInstitucion.Style.Alignment.Vertical =
                XLAlignmentVerticalValues.Center;

            // Título.
            var tituloReporte = hoja.Range(
                2,
                1,
                2,
                totalColumnas);

            tituloReporte.Style.Font.Bold = true;
            tituloReporte.Style.Font.FontSize = 13;

            tituloReporte.Style.Alignment.Horizontal =
                XLAlignmentHorizontalValues.Center;

            // Fecha.
            var fechaReporte = hoja.Range(
                3,
                1,
                3,
                totalColumnas);

            fechaReporte.Style.Alignment.Horizontal =
                XLAlignmentHorizontalValues.Center;

            fechaReporte.Style.Font.FontColor =
                XLColor.FromHtml("#666666");

            // ===========================================================
            // FILTROS APLICADOS
            // ===========================================================

            var filtros = new List<string>();

            if (!string.IsNullOrWhiteSpace(buscar))
                filtros.Add($"Búsqueda: {buscar}");

            if (cuentaContableId.HasValue)
                filtros.Add($"Cuenta contable ID: {cuentaContableId}");

            if (!string.IsNullOrWhiteSpace(tipoMovimiento))
                filtros.Add($"Tipo: {tipoMovimiento}");

            if (!string.IsNullOrWhiteSpace(origenModulo))
                filtros.Add($"Origen: {origenModulo}");

            if (proyectoId.HasValue)
                filtros.Add($"Proyecto ID: {proyectoId}");

            if (centroCostoId.HasValue)
                filtros.Add($"Centro de costo ID: {centroCostoId}");

            if (!string.IsNullOrWhiteSpace(estado))
                filtros.Add($"Estado: {estado}");

            if (anio.HasValue)
                filtros.Add($"Año: {anio}");

            if (mes.HasValue)
            {
                string nombreMes = new DateTime(2000, mes.Value, 1)
                    .ToString("MMMM", new System.Globalization.CultureInfo("es-CR"));

                nombreMes = char.ToUpper(nombreMes[0]) + nombreMes.Substring(1);

                filtros.Add($"Mes: {nombreMes}");
            }

            if (desde.HasValue)
                filtros.Add($"Desde: {desde.Value:dd/MM/yyyy}");

            if (hasta.HasValue)
                filtros.Add($"Hasta: {hasta.Value:dd/MM/yyyy}");

            hoja.Cell("A5").Value = "Filtros aplicados:";
            hoja.Cell("A5").Style.Font.Bold = true;

            hoja.Cell("B5").Value =
                filtros.Any()
                    ? string.Join(" | ", filtros)
                    : "Sin filtros";

            hoja.Range("B5:M5").Merge();

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

            // ===========================================================
            // ENCABEZADOS DE LA TABLA
            // ===========================================================

            var filaEncabezado = 7;

            string[] encabezados =
            {
    "Fecha",
    "Código cuenta",
    "Cuenta contable",
    "Tipo",
    "Origen",
    "Referencia",
    "Descripción",
    "Proyecto",
    "Centro de costo",
    "Debe",
    "Haber",
    "Estado",
    "Usuario"
};

            for (int columna = 0; columna < encabezados.Length; columna++)
            {
                hoja.Cell(filaEncabezado, columna + 1).Value =
                    encabezados[columna];
            }

            var rangoEncabezado =
                hoja.Range(
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

            // ===========================================================
            // DATOS
            // ===========================================================

            var fila = filaEncabezado + 1;

            foreach (var movimiento in movimientos)
            {
                string estadoMovimiento;

                if (movimiento.Anulado == true ||
                    movimiento.Estado == "Anulado")
                {
                    estadoMovimiento = "Anulado";
                }
                else if (movimiento.TipoMovimiento == "Reversión")
                {
                    estadoMovimiento = "Reversión";
                }
                else
                {
                    estadoMovimiento =
                        movimiento.Estado ?? "Registrado";
                }

                hoja.Cell(fila, 1).Value =
                    movimiento.Fecha.ToDateTime(TimeOnly.MinValue);

                hoja.Cell(fila, 2).Value =
                    movimiento.CuentaContable?.Codigo ?? "";

                hoja.Cell(fila, 3).Value =
                    movimiento.CuentaContable?.Nombre ?? "";

                hoja.Cell(fila, 4).Value =
                    movimiento.TipoMovimiento ?? "";

                hoja.Cell(fila, 5).Value =
                    movimiento.OrigenModulo ?? "Sin origen";

                hoja.Cell(fila, 6).Value =
                    movimiento.Referencia ?? "Sin referencia";

                hoja.Cell(fila, 7).Value =
                    movimiento.Descripcion ?? "Sin descripción";

                hoja.Cell(fila, 8).Value =
                    movimiento.Proyecto?.Nombre ?? "Sin proyecto";

                hoja.Cell(fila, 9).Value =
                    movimiento.CentroCosto?.Nombre ?? "Sin centro";

                hoja.Cell(fila, 10).Value =
                    movimiento.Debe;

                hoja.Cell(fila, 11).Value =
                    movimiento.Haber;

                hoja.Cell(fila, 12).Value =
                    estadoMovimiento;

                hoja.Cell(fila, 13).Value =
                    movimiento.Usuario?.Nombre ?? "Sistema";

                fila++;
            }

            // ===========================================================
            // FORMATOS
            // ===========================================================

            if (movimientos.Any())
            {
                var rangoTabla = hoja.Range(
    filaEncabezado,
    1,
    fila - 1,
    totalColumnas);

                var tabla = rangoTabla.CreateTable("TablaMovimientosContables");

                tabla.Theme =
                    XLTableTheme.TableStyleMedium2;

                tabla.ShowAutoFilter = true;
                tabla.ShowRowStripes = false;

                hoja.Column(1).Style.DateFormat.Format =
                    "dd/MM/yyyy";

                hoja.Columns(10, 11).Style.NumberFormat.Format =
                    "₡#,##0.00";

                hoja.SheetView.FreezeRows(filaEncabezado);

                hoja.Range(
                    filaEncabezado + 1,
                    1,
                    fila - 1,
                    encabezados.Length)
                    .Style.Alignment.Vertical =
                    XLAlignmentVerticalValues.Center;

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
                // COLORES SEGÚN EL ESTADO DEL MOVIMIENTO
                // ===========================================================

                for (int indice = 0; indice < movimientos.Count; indice++)
                {
                    var movimiento = movimientos[indice];

                    int filaMovimiento =
                        filaEncabezado + 1 + indice;

                    var rangoFila = hoja.Range(
                        filaMovimiento,
                        1,
                        filaMovimiento,
                        totalColumnas);

                    bool esAnulado =
                        movimiento.Anulado == true ||
                        movimiento.Estado == "Anulado";

                    bool esReversion =
                        movimiento.TipoMovimiento == "Reversión" ||
                        movimiento.Estado == "Reversado";

                    if (esAnulado)
                    {
                        rangoFila.Style.Fill.BackgroundColor =
                            XLColor.FromHtml("#F8D7DA");

                        rangoFila.Style.Font.FontColor =
                            XLColor.FromHtml("#842029");
                    }
                    else if (esReversion)
                    {
                        rangoFila.Style.Fill.BackgroundColor =
                            XLColor.FromHtml("#DBEAFE");

                        rangoFila.Style.Font.FontColor =
                            XLColor.FromHtml("#1E40AF");
                    }
                }
            }
            // TOTALES FINANCIEROS

            var vigentes = movimientos
                .Where(m =>
                    m.Anulado != true &&
                    m.Estado == "Registrado")
                .ToList();

            decimal totalDebe =
                vigentes.Sum(m => m.Debe);

            decimal totalHaber =
                vigentes.Sum(m => m.Haber);

            decimal diferencia =
                totalDebe - totalHaber;

            int filaTotales = fila + 1;

            // Franja de totales.
            hoja.Cell(filaTotales, 8).Value =
                "TOTALES CONTABLES";

            hoja.Range(
                    filaTotales,
                    8,
                    filaTotales,
                    9)
                .Merge();

            hoja.Cell(filaTotales, 10).Value =
                totalDebe;

            hoja.Cell(filaTotales, 11).Value =
                totalHaber;

            hoja.Cell(filaTotales, 12).Value =
                diferencia;

            hoja.Cell(filaTotales, 13).Value =
                $"{movimientos.Count} registros";

            var rangoTotales = hoja.Range(
                filaTotales,
                8,
                filaTotales,
                totalColumnas);

            rangoTotales.Style.Fill.BackgroundColor =
                XLColor.FromHtml("#D9EDEF");

            rangoTotales.Style.Font.Bold = true;

            rangoTotales.Style.Font.FontColor =
                XLColor.FromHtml("#123F46");

            rangoTotales.Style.Border.TopBorder =
                XLBorderStyleValues.Medium;

            rangoTotales.Style.Border.TopBorderColor =
                XLColor.FromHtml("#0F5C64");

            rangoTotales.Style.Border.BottomBorder =
                XLBorderStyleValues.Thin;

            rangoTotales.Style.Border.InsideBorder =
                XLBorderStyleValues.Thin;

            rangoTotales.Style.Alignment.Vertical =
                XLAlignmentVerticalValues.Center;

            hoja.Range(
                    filaTotales,
                    10,
                    filaTotales,
                    12)
                .Style.NumberFormat.Format =
                    "₡ #,##0.00";

            hoja.Range(
                    filaTotales,
                    8,
                    filaTotales,
                    totalColumnas)
                .Style.Alignment.Horizontal =
                    XLAlignmentHorizontalValues.Right;
            // ===========================================================
            // RESUMEN DEL ESTADO DE LOS MOVIMIENTOS
            // ===========================================================

            int filaResumen = filaTotales + 3;

            hoja.Cell(filaResumen, 1).Value =
                "RESUMEN DEL ESTADO DE LOS MOVIMIENTOS";

            hoja.Range(
                    filaResumen,
                    1,
                    filaResumen,
                    4)
                .Merge();

            var tituloResumen = hoja.Range(
                filaResumen,
                1,
                filaResumen,
                4);

            tituloResumen.Style.Fill.BackgroundColor =
                XLColor.FromHtml("#0F5C64");

            tituloResumen.Style.Font.Bold = true;
            tituloResumen.Style.Font.FontColor = XLColor.White;

            tituloResumen.Style.Alignment.Horizontal =
                XLAlignmentHorizontalValues.Center;

            string[] encabezadosResumen =
            {
    "Registrados",
    "Reversados",
    "Anulados",
    "Total"
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

            int cantidadRegistrados =
                movimientos.Count(m =>
                    m.Anulado != true &&
                    m.Estado == "Registrado" &&
                    m.TipoMovimiento != "Reversión");

            int cantidadReversados =
                movimientos.Count(m =>
                    m.TipoMovimiento == "Reversión" ||
                    m.Estado == "Reversado");

            int cantidadAnulados =
                movimientos.Count(m =>
                    m.Anulado == true ||
                    m.Estado == "Anulado");

            hoja.Cell(filaResumen + 2, 1).Value =
                cantidadRegistrados;

            hoja.Cell(filaResumen + 2, 2).Value =
                cantidadReversados;

            hoja.Cell(filaResumen + 2, 3).Value =
                cantidadAnulados;

            hoja.Cell(filaResumen + 2, 4).Value =
                movimientos.Count;

            var encabezadoResumen = hoja.Range(
                filaResumen + 1,
                1,
                filaResumen + 1,
                4);

            encabezadoResumen.Style.Fill.BackgroundColor =
                XLColor.FromHtml("#D9EDEF");

            encabezadoResumen.Style.Font.Bold = true;

            encabezadoResumen.Style.Alignment.Horizontal =
                XLAlignmentHorizontalValues.Center;

            var valoresResumen = hoja.Range(
                filaResumen + 2,
                1,
                filaResumen + 2,
                4);

            valoresResumen.Style.Font.Bold = true;

            valoresResumen.Style.Alignment.Horizontal =
                XLAlignmentHorizontalValues.Center;

            var rangoResumen = hoja.Range(
                filaResumen + 1,
                1,
                filaResumen + 2,
                4);

            rangoResumen.Style.Border.OutsideBorder =
                XLBorderStyleValues.Thin;

            rangoResumen.Style.Border.InsideBorder =
                XLBorderStyleValues.Thin;
            // ===========================================================
            // ANCHOS Y AJUSTES
            // ===========================================================

            hoja.Columns().AdjustToContents();

            hoja.Column(3).Width = 28;
            hoja.Column(6).Width = 22;
            hoja.Column(7).Width = 45;
            hoja.Column(8).Width = 30;
            hoja.Column(9).Width = 25;
            hoja.Column(10).Width = 16;
            hoja.Column(11).Width = 16;

            hoja.Column(7).Style.Alignment.WrapText = true;
            hoja.Column(8).Style.Alignment.WrapText = true;
            hoja.Column(9).Style.Alignment.WrapText = true;
            hoja.Columns(10, 11).Style.Alignment.Horizontal =
    XLAlignmentHorizontalValues.Right;

            hoja.SheetView.FreezeRows(filaEncabezado);

            hoja.PageSetup.PageOrientation =
                XLPageOrientation.Landscape;

            hoja.PageSetup.FitToPages(1, 0);

            hoja.PageSetup.SetRowsToRepeatAtTop(
                filaEncabezado,
                filaEncabezado);

            hoja.PageSetup.CenterHorizontally = true;

            // ===========================================================
            // DEVOLVER EL ARCHIVO
            // ===========================================================

            using var stream = new MemoryStream();

            workbook.SaveAs(stream);

            var nombreArchivo =
                $"MovimientosContables_{DateTime.Now:yyyyMMdd_HHmm}.xlsx";
            await RegistrarAuditoria(
    "MovimientosContables",
    0,
    "Exportar Excel",
    "Se exportó el reporte de movimientos contables a Excel."
);
            return File(
                stream.ToArray(),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                nombreArchivo);

        }
        private async Task RegistrarAuditoria(
    string tabla,
    int registroId,
    string accion,
    string descripcion)
        {
            int usuarioId =
                HttpContext.Session.GetInt32("UsuarioId") ?? 1;

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
    }
}