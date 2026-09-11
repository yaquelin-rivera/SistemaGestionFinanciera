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
using System.IO;

namespace SistemaGestionFinanciera.Controllers
{
    public class IngresosController : BaseController
    {
        private readonly SistemaFinancieroContext _context;
        private readonly ContabilidadService _contabilidadService;
        //CONSTRUCTOR
        public IngresosController(
        SistemaFinancieroContext context,
        ContabilidadService contabilidadService)
        {
            _context = context;
            _contabilidadService = contabilidadService;
        }

        // GET: Ingresos
        public async Task<IActionResult> Index(
            string? buscar,
            string? tipo,
            string? estado,
            string? origen,
            int? proyectoId,
            int? anio,
            int? mes,
            DateOnly? fechaInicio,
            DateOnly? fechaFin,
            bool pendientePresupuesto = false)
        {
            int pendientesAsignacionPresupuestaria =
    await _context.Ingresos
        .AsNoTracking()
        .CountAsync(i =>
            i.Activo &&
            i.PresupuestoMensualId == null &&
            !string.IsNullOrWhiteSpace(
                i.MotivoPendientePresupuestario));

            ViewBag.PendientesAsignacionPresupuestaria =
                pendientesAsignacionPresupuestaria;

            bool hayFiltros =
                !string.IsNullOrWhiteSpace(buscar) ||
                !string.IsNullOrWhiteSpace(tipo) ||
                !string.IsNullOrWhiteSpace(estado) ||
                !string.IsNullOrWhiteSpace(origen) ||
                proyectoId.HasValue ||
                anio.HasValue ||
                mes.HasValue ||
                fechaInicio.HasValue ||
                fechaFin.HasValue;

           
            // Obtiene los ingresos con los filtros aplicados.
            var listaIngresos = await ConstruirConsultaReporte(
                buscar,
                tipo,
                estado,
                origen,
                proyectoId,
                anio,
                mes,
                fechaInicio,
                fechaFin);

            
            // ============================================================
            // FILTRO DE INGRESOS PENDIENTES DE ASIGNACIÓN PRESUPUESTARIA
            // ============================================================

            if (pendientePresupuesto)
            {
                listaIngresos = listaIngresos
                    .Where(i =>
                        i.Activo &&
                        i.PresupuestoMensualId == null &&
                        !string.IsNullOrWhiteSpace(
                            i.MotivoPendientePresupuestario))
                    .ToList();
            }
            var ingresosParaTotales = listaIngresos;


            ViewBag.PendientePresupuesto = pendientePresupuesto;
            // Si no hay filtros, las tarjetas muestran únicamente el mes actual.
            if (!hayFiltros)
            {
                var hoy = DateOnly.FromDateTime(DateTime.Today);
                var inicioMes = new DateOnly(hoy.Year, hoy.Month, 1);
                var finMes = inicioMes.AddMonths(1).AddDays(-1);

                ingresosParaTotales = listaIngresos
                    .Where(i =>
                        i.Fecha >= inicioMes &&
                        i.Fecha <= finMes)
                    .ToList();

                ViewBag.TipoResumen = "Mes actual";
            }
            else
            {
                ViewBag.TipoResumen = "Filtro aplicado";
            }

            // Totales financieros.
            ViewBag.TotalEsperado = ingresosParaTotales
                .Where(i =>
                    i.Activo &&
                    i.Estado != "Anulado")
                .Sum(i => i.MontoEsperado ?? 0m);

            ViewBag.TotalRecibido = ingresosParaTotales
                .Where(i =>
                    i.Activo &&
                    i.Estado != "Anulado")
                .Sum(i => i.MontoReal ?? 0m);

            ViewBag.DiferenciaTotal =
                (decimal)ViewBag.TotalEsperado -
                (decimal)ViewBag.TotalRecibido;

            // Cantidades por estado.
            ViewBag.Pendientes = ingresosParaTotales.Count(i =>
                i.Activo &&
                i.Estado == "Pendiente");

            ViewBag.Parciales = ingresosParaTotales.Count(i =>
                i.Activo &&
                i.Estado == "Parcial");

            ViewBag.Pagados = ingresosParaTotales.Count(i =>
                i.Activo &&
                i.Estado == "Pagado");

            ViewBag.Anulados = ingresosParaTotales.Count(i =>
                i.Estado == "Anulado");

            // Mantiene los filtros aplicados en pantalla.
            ViewBag.Buscar = buscar;
            ViewBag.Tipo = tipo;
            ViewBag.Estado = estado;
            ViewBag.Origen = origen;
            ViewBag.ProyectoId = proyectoId;
            ViewBag.Anio = anio;
            ViewBag.Mes = mes;
            ViewBag.FechaInicio = fechaInicio?.ToString("yyyy-MM-dd");
            ViewBag.FechaFin = fechaFin?.ToString("yyyy-MM-dd");

            // Proyectos disponibles para el filtro.
            ViewBag.Proyectos = new SelectList(
                await _context.Proyectos
                    .Where(p => p.Activo)
                    .OrderBy(p => p.Nombre)
                    .ToListAsync(),
                "IdProyecto",
                "Nombre",
                proyectoId);

            // Años existentes en los registros de ingresos.
            ViewBag.Anios = await _context.Ingresos
                .Select(i => i.Fecha.Year)
                .Distinct()
                .OrderByDescending(a => a)
                .ToListAsync();

            ViewBag.PendientePresupuesto =
    pendientePresupuesto;

            return View(listaIngresos);
        }

        // ===========================================================
        // CONSTRUYE LA CONSULTA PARA INDEX, PDF Y EXCEL
        // ===========================================================
        private async Task<List<Ingreso>> ConstruirConsultaReporte(
            string? buscar,
            string? tipo,
            string? estado,
            string? origen,
            int? proyectoId,
            int? anio,
            int? mes,
            DateOnly? fechaInicio,
            DateOnly? fechaFin)
        {
            var consulta = _context.Ingresos
                .Include(i => i.CategoriaIngreso)
                .Include(i => i.ClienteBeneficiario)
                .Include(i => i.Proyecto)
                .Include(i => i.PresupuestoMensualRegistro)
                .Include(i => i.CentroCosto)
                .AsQueryable();

            if (fechaInicio.HasValue)
            {
                consulta = consulta.Where(i =>
                    i.Fecha >= fechaInicio.Value);
            }

            if (fechaFin.HasValue)
            {
                consulta = consulta.Where(i =>
                    i.Fecha <= fechaFin.Value);
            }

            if (!string.IsNullOrWhiteSpace(buscar))
            {
                consulta = consulta.Where(i =>
                    (i.Fuente != null &&
                     i.Fuente.Contains(buscar)) ||

                    (i.Descripcion != null &&
                     i.Descripcion.Contains(buscar)) ||

                    (i.ClienteBeneficiario != null &&
                     i.ClienteBeneficiario.Nombre.Contains(buscar)) ||

                    (i.CategoriaIngreso != null &&
                     i.CategoriaIngreso.Nombre.Contains(buscar)) ||

                    (i.Proyecto != null &&
                     i.Proyecto.Nombre.Contains(buscar)));
            }

            if (!string.IsNullOrWhiteSpace(tipo))
            {
                consulta = consulta.Where(i =>
                    i.TipoIngreso == tipo);
            }

            if (!string.IsNullOrWhiteSpace(estado))
            {
                consulta = consulta.Where(i =>
                    i.Estado == estado);
            }

            if (proyectoId.HasValue)
            {
                consulta = consulta.Where(i =>
                    i.ProyectoId == proyectoId.Value);
            }

            if (anio.HasValue)
            {
                consulta = consulta.Where(i =>
                    i.Fecha.Year == anio.Value);
            }

            if (mes.HasValue)
            {
                consulta = consulta.Where(i =>
                    i.Fecha.Month == mes.Value);
            }

            var ingresos = await consulta
                .OrderByDescending(i => i.IdIngreso)
                .ToListAsync();

            // El origen se determina después de consultar la base de datos,
            // igual que en la lógica actual del Index.
            if (!string.IsNullOrWhiteSpace(origen))
            {
                ingresos = ingresos
                    .Where(i => DeterminarOrigenIngreso(i) == origen)
                    .ToList();
            }

            return ingresos;
        }

    // ===========================================================
    // DETERMINA EL ORIGEN FUNCIONAL DEL INGRESO
    // ===========================================================
    private static string DeterminarOrigenIngreso(Ingreso ingreso)
        {
            return string.IsNullOrWhiteSpace(ingreso.Origen)
                ? "Manual"
                : ingreso.Origen;
        }

        // ===========================================================
        // VISTA PREVIA DEL REPORTE DE INGRESOS
        // ===========================================================
        public async Task<IActionResult> Reporte(
            string? buscar,
            string? tipo,
            string? estado,
            string? origen,
            int? proyectoId,
            int? anio,
            int? mes,
            DateOnly? fechaInicio,
            DateOnly? fechaFin,
           bool pendientePresupuesto = false)
        {
            var ingresos = await ConstruirConsultaReporte(
                buscar,
                tipo,
                estado,
                origen,
                proyectoId,
                anio,
                mes,
                fechaInicio,
                fechaFin);

            if (pendientePresupuesto)
            {
                ingresos = ingresos
                    .Where(i =>
                        i.Activo &&
                        i.Estado != "Anulado" &&
                        i.PresupuestoMensualId == null &&
                        !string.IsNullOrWhiteSpace(
                            i.MotivoPendientePresupuestario))
                    .ToList();
            }

            ViewBag.PendientePresupuesto =
                pendientePresupuesto;
            ingresos = ingresos
                .OrderByDescending(i => i.Fecha)
                .ThenByDescending(i => i.IdIngreso)
                .ToList();

            ViewBag.TotalRegistros = ingresos.Count;

            ViewBag.TotalEsperado = ingresos
                .Where(i =>
                    i.Activo &&
                    i.Estado != "Anulado")
                .Sum(i => i.MontoEsperado ?? 0m);

            ViewBag.TotalRecibido = ingresos
                .Where(i =>
                    i.Activo &&
                    i.Estado != "Anulado")
                .Sum(i => i.MontoReal ?? 0m);

            ViewBag.DiferenciaTotal =
                (decimal)ViewBag.TotalEsperado -
                (decimal)ViewBag.TotalRecibido;

            ViewBag.Pendientes = ingresos.Count(i =>
                i.Activo &&
                i.Estado == "Pendiente");

            ViewBag.Parciales = ingresos.Count(i =>
                i.Activo &&
                i.Estado == "Parcial");

            ViewBag.Pagados = ingresos.Count(i =>
                i.Activo &&
                i.Estado == "Pagado");

            ViewBag.Anulados = ingresos.Count(i =>
                i.Estado == "Anulado");

            ViewBag.FechaGeneracion = DateTime.Now;

            // Conserva los filtros para mostrarlos en el reporte.
            ViewBag.Buscar = buscar;
            ViewBag.Tipo = tipo;
            ViewBag.Estado = estado;
            ViewBag.Origen = origen;
            ViewBag.ProyectoId = proyectoId;
            ViewBag.Anio = anio;
            ViewBag.Mes = mes;
            ViewBag.FechaInicio = fechaInicio?.ToString("yyyy-MM-dd");
            ViewBag.FechaFin = fechaFin?.ToString("yyyy-MM-dd");
            //auditoria
            string descripcionFiltros =
    await ConstruirDescripcionFiltros(
        buscar,
        tipo,
        estado,
        origen,
        proyectoId,
        anio,
        mes,
        fechaInicio,
        fechaFin);

            await RegistrarAuditoria(
                "Vista previa PDF",
                0,
                "Se generó la vista previa del reporte de ingresos. " +
                descripcionFiltros);
            return View(ingresos);
        }

        // ===========================================================
        // EXPORTAR REPORTE DE INGRESOS A EXCEL
        // ===========================================================
        public async Task<IActionResult> ExportarExcel(
            string? buscar,
            string? tipo,
            string? estado,
            string? origen,
            int? proyectoId,
            int? anio,
            int? mes,
            DateOnly? fechaInicio,
            DateOnly? fechaFin,
              bool pendientePresupuesto = false)
        {
            var ingresos = await ConstruirConsultaReporte(
                buscar,
                tipo,
                estado,
                origen,
                proyectoId,
                anio,
                mes,
                fechaInicio,
                fechaFin);

            ingresos = ingresos
                .OrderByDescending(i => i.Fecha)
                .ThenByDescending(i => i.IdIngreso)
                .ToList();

            using var workbook = new XLWorkbook();

            var hoja = workbook.Worksheets.Add("Ingresos");

            const int totalColumnas = 13;

            if (pendientePresupuesto)
            {
                ingresos = ingresos
                    .Where(i =>
                        i.Activo &&
                        i.Estado != "Anulado" &&
                        i.PresupuestoMensualId == null &&
                        !string.IsNullOrWhiteSpace(
                            i.MotivoPendientePresupuestario))
                    .ToList();
            }
            // ===========================================================
            // ENCABEZADO PRINCIPAL
            // ===========================================================

            hoja.Cell("A1").Value = "Unión Cantonal de Asociaciones de Desarrollo de Santa Ana";
            hoja.Range(1, 1, 1, totalColumnas).Merge();

            hoja.Cell("A2").Value = "Reporte de Ingresos";
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

            // ===========================================================
            // FILTROS APLICADOS
            // ===========================================================

            var filtros = new List<string>();

            if (!string.IsNullOrWhiteSpace(buscar))
            {
                filtros.Add($"Búsqueda: {buscar}");
            }

            if (!string.IsNullOrWhiteSpace(tipo))
            {
                filtros.Add($"Tipo: {tipo}");
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
                    .Where(p => p.IdProyecto == proyectoId.Value)
                    .Select(p => p.Nombre)
                    .FirstOrDefaultAsync() ?? "No encontrado";

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
            if (pendientePresupuesto)
            {
                filtros.Add(
                    "Asignación presupuestaria: Pendiente");
            }

            hoja.Cell("A5").Value = filtros.Any()
                ? string.Join(" | ", filtros)
                : "Sin filtros aplicados";

            hoja.Range(5, 1, 5, totalColumnas).Merge();

            hoja.Range(5, 1, 5, totalColumnas)
                .Style.Font.Italic = true;

            hoja.Range(5, 1, 5, totalColumnas)
                .Style.Alignment.WrapText = true;

            // ===========================================================
            // ENCABEZADOS DE LA TABLA
            // ===========================================================

            int filaEncabezado = 7;

            string[] encabezados =
            {
        "Fecha",
        "Fuente",
        "Descripción",
        "Tipo",
        "Categoría",
        "Cliente / Beneficiario",
        "Proyecto",
        "Centro de costo",
        "Origen",
        "Estado",
        "Monto esperado",
        "Monto recibido",
        "Diferencia"
    };

            for (int columna = 1;
                 columna <= encabezados.Length;
                 columna++)
            {
                hoja.Cell(filaEncabezado, columna).Value =
                    encabezados[columna - 1];
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

            foreach (var ingreso in ingresos)
            {
                decimal montoEsperado =
                    ingreso.MontoEsperado ?? 0m;

                decimal montoRecibido =
                    ingreso.MontoReal ?? 0m;

                decimal diferencia =
                    ingreso.Diferencia ??
                    (montoEsperado - montoRecibido);

                hoja.Cell(fila, 1).Value =
                    ingreso.Fecha.ToDateTime(TimeOnly.MinValue);

                hoja.Cell(fila, 2).Value =
                    ingreso.Fuente ?? "Sin fuente";

                hoja.Cell(fila, 3).Value =
                    ingreso.Descripcion ?? "Sin descripción";

                hoja.Cell(fila, 4).Value =
                    ingreso.TipoIngreso ?? "Sin tipo";

                hoja.Cell(fila, 5).Value =
                    ingreso.CategoriaIngreso?.Nombre ??
                    "Sin categoría";

                hoja.Cell(fila, 6).Value =
                    ingreso.ClienteBeneficiario?.Nombre ??
                    "Sin cliente";

                hoja.Cell(fila, 7).Value =
                    ingreso.Proyecto?.Nombre ??
                    "Sin proyecto";

                hoja.Cell(fila, 8).Value =
                    ingreso.CentroCosto?.Nombre ??
                    "Sin centro";

                hoja.Cell(fila, 9).Value =
                    DeterminarOrigenIngreso(ingreso);

                hoja.Cell(fila, 10).Value =
                    ingreso.Estado ?? "Sin estado";

                hoja.Cell(fila, 11).Value =
                    montoEsperado;

                hoja.Cell(fila, 12).Value =
                    montoRecibido;

                hoja.Cell(fila, 13).Value =
                    diferencia;

                fila++;
            }
            // Activa los filtros en los encabezados de Excel.
            if (ingresos.Any())
            {
                hoja.Range(
                    filaEncabezado,
                    1,
                    fila - 1,
                    totalColumnas)
                    .SetAutoFilter();
            }
            // ===========================================================
            // FORMATOS
            // ===========================================================

            if (ingresos.Any())
            {
                hoja.Range(
                        filaEncabezado + 1,
                        1,
                        fila - 1,
                        1)
                    .Style.DateFormat.Format = "dd/MM/yyyy";

                hoja.Range(
                        filaEncabezado + 1,
                        11,
                        fila - 1,
                        13)
                    .Style.NumberFormat.Format =
                        "₡#,##0.00";
            }
            // ===========================================================
            // COLORES SEGÚN EL ESTADO DEL INGRESO
            // ===========================================================

            for (int indice = 0; indice < ingresos.Count; indice++)
            {
                var ingreso = ingresos[indice];

                int filaIngreso =
                    filaEncabezado + 1 + indice;

                var rangoFila = hoja.Range(
                    filaIngreso,
                    1,
                    filaIngreso,
                    totalColumnas);

                if (ingreso.Estado == "Anulado")
                {
                    // Fondo rojo claro.
                    rangoFila.Style.Fill.BackgroundColor =
                        XLColor.FromHtml("#F8D7DA");

                    // Texto rojo oscuro.
                    rangoFila.Style.Font.FontColor =
                        XLColor.FromHtml("#842029");
                }
                else if (ingreso.Estado == "Pendiente")
                {
                    // Fondo amarillo claro.
                    rangoFila.Style.Fill.BackgroundColor =
                        XLColor.FromHtml("#FFF3CD");

                    // Texto amarillo oscuro.
                    rangoFila.Style.Font.FontColor =
                        XLColor.FromHtml("#664D03");
                }
            }
            // ===========================================================
            // TOTALES
            // ===========================================================

            decimal totalEsperado = ingresos
                .Where(i =>
                    i.Activo &&
                    i.Estado != "Anulado")
                .Sum(i => i.MontoEsperado ?? 0m);

            decimal totalRecibido = ingresos
                .Where(i =>
                    i.Activo &&
                    i.Estado != "Anulado")
                .Sum(i => i.MontoReal ?? 0m);

            decimal diferenciaTotal =
                totalEsperado - totalRecibido;

            int filaTotales = fila + 1;

            hoja.Cell(filaTotales, 10).Value = "TOTALES";
            hoja.Cell(filaTotales, 11).Value = totalEsperado;
            hoja.Cell(filaTotales, 12).Value = totalRecibido;
            hoja.Cell(filaTotales, 13).Value = diferenciaTotal;

            var rangoTotales = hoja.Range(
                filaTotales,
                10,
                filaTotales,
                13);

            rangoTotales.Style.Font.Bold = true;

            rangoTotales.Style.Fill.BackgroundColor =
                XLColor.FromHtml("#F2F2F2");

            rangoTotales.Style.NumberFormat.Format =
                "₡#,##0.00";

            // Bordes para toda la tabla.
            if (ingresos.Any())
            {
                hoja.Range(
                        filaEncabezado,
                        1,
                        fila - 1,
                        totalColumnas)
                    .Style.Border.InsideBorder =
                        XLBorderStyleValues.Thin;

                hoja.Range(
                        filaEncabezado,
                        1,
                        fila - 1,
                        totalColumnas)
                    .Style.Border.OutsideBorder =
                        XLBorderStyleValues.Thin;
            }
            if (ingresos.Any())
            {
                var rangoDatos = hoja.Range(
                    filaEncabezado,
                    1,
                    fila - 1,
                    totalColumnas);

                rangoDatos.SetAutoFilter();
            }
            // Ajusta automáticamente el tamaño de las columnas.
            hoja.Columns().AdjustToContents();

            // Limita el ancho de columnas con textos largos.
            hoja.Column(2).Width = 22;
            hoja.Column(3).Width = 35;
            hoja.Column(5).Width = 22;
            hoja.Column(6).Width = 25;
            hoja.Column(7).Width = 25;
            hoja.Column(8).Width = 22;

            hoja.RangeUsed()?.Style.Alignment.SetVertical(
                XLAlignmentVerticalValues.Center);

            hoja.SheetView.FreezeRows(filaEncabezado);

            // ===========================================================
            // GENERAR ARCHIVO
            // ===========================================================

            using var stream = new MemoryStream();

            workbook.SaveAs(stream);

            string nombreArchivo =
                $"Reporte_Ingresos_{DateTime.Now:yyyyMMdd_HHmm}.xlsx";
            //auditoria
            string descripcionFiltros =
    await ConstruirDescripcionFiltros(
        buscar,
        tipo,
        estado,
        origen,
        proyectoId,
        anio,
        mes,
        fechaInicio,
        fechaFin);

            if (pendientePresupuesto)
            {
                filtros.Add(
                    "Asignación presupuestaria=Pendiente");
            }

            await RegistrarAuditoria(
                "Exportar Excel",
                0,
                "Se exportó el reporte de ingresos a Excel. " +
                descripcionFiltros);
            return File(
                stream.ToArray(),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                nombreArchivo);
      
        }

        // GET: Ingresos/Detalles
        public async Task<IActionResult> Details(int? id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var ingreso = await _context.Ingresos
                .Include(i => i.CategoriaIngreso)
                .Include(i => i.ClienteBeneficiario)
                .Include(i => i.Proyecto)
                .Include(i => i.PresupuestoMensualRegistro)
                .Include(i => i.CentroCosto)
                .FirstOrDefaultAsync(m => m.IdIngreso == id);
            if (ingreso == null)
            {
                return NotFound();
            }

            return View(ingreso);
        }
        // CARGA LOS SELECTS DE INGRESOS
        // En Create muestra solamente registros activos.En Edit conserva también el registro actualmente asignado, aunque posteriormente haya sido inactivado.
        private async Task CargarSelectsIngreso(
            Ingreso ingreso,
            bool esEdicion)
        {
            var categorias = await _context.CategoriasIngresos
                .AsNoTracking()
                .Where(c =>
                    c.Activo ||
                    (esEdicion &&
                     c.IdCategoriaIngreso == ingreso.CategoriaIngresoId))
                .OrderBy(c => c.Nombre)
                .ToListAsync();

            var clientes = await _context.ClientesBeneficiarios
                .AsNoTracking()
                .Where(c =>
                    c.Activo ||
                    (esEdicion &&
                     c.IdClienteBeneficiario ==
                     ingreso.ClienteBeneficiarioId))
                .OrderBy(c => c.Nombre)
                .ToListAsync();

            var proyectos = await _context.Proyectos
                .AsNoTracking()
                .Where(p =>
                    p.Activo ||
                    (esEdicion &&
                     p.IdProyecto == ingreso.ProyectoId))
                .OrderBy(p => p.Nombre)
                .ToListAsync();

            var centrosCosto = await _context.CentrosCostos
                .AsNoTracking()
                .Where(c =>
                    c.Activo ||
                    (esEdicion &&
                     c.IdCentroCosto == ingreso.CentroCostoId))
                .OrderBy(c => c.Nombre)
                .ToListAsync();

            ViewData["CategoriaIngresoId"] = new SelectList(
                categorias,
                "IdCategoriaIngreso",
                "Nombre",
                ingreso.CategoriaIngresoId);

            ViewData["ClienteBeneficiarioId"] = new SelectList(
                clientes,
                "IdClienteBeneficiario",
                "Nombre",
                ingreso.ClienteBeneficiarioId);

            ViewData["ProyectoId"] = new SelectList(
                proyectos,
                "IdProyecto",
                "Nombre",
                ingreso.ProyectoId);

            ViewData["CentroCostoId"] = new SelectList(
                centrosCosto,
                "IdCentroCosto",
                "Nombre",
                ingreso.CentroCostoId);

            ViewBag.TiposIngreso = new List<string>
    {
        "Operativo",
        "Extraordinario",
        "Donación",
        "Aporte",
        "Convenio"
    };
        }

        private async Task CargarOpcionesProyectoPeriodoIngreso(
    string? seleccionActual = null)
        {
            int? idPresupuestoActual = null;

            if (!string.IsNullOrWhiteSpace(seleccionActual) &&
                seleccionActual.StartsWith("PRES-"))
            {
                string valorIdActual =
                    seleccionActual.Substring(5);

                if (int.TryParse(valorIdActual, out int idActual))
                {
                    idPresupuestoActual = idActual;
                }
            }
            var cultura = new CultureInfo("es-CR");

            var hoy = DateOnly.FromDateTime(DateTime.Today);
            var fechaMinima = hoy.AddMonths(-3);

            int anioMinimo = fechaMinima.Year;
            int mesMinimo = fechaMinima.Month;
            // ============================================================
            // 1. PERIODOS PRESUPUESTARIOS DISPONIBLES
            // Solo Aprobados o Borrador de proyectos activos.
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

    // En Edit conserva únicamente
    // el presupuesto histórico actual.
    (
        idPresupuestoActual.HasValue &&
        pm.IdPresupuestoMensual ==
            idPresupuestoActual.Value
    ))
             .OrderByDescending(pm => pm.Anio)
.ThenByDescending(pm => pm.Mes)
.ThenBy(pm => pm.Proyecto!.Nombre)
                .ToListAsync();

            var opciones = new List<SelectListItem>();

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

                    Selected =
                        valor == seleccionActual
                });
            }

            // 2. PROYECTOS ACTIVOS SIN UN PERIODO UTILIZABLE
            // Incluye:
            // - proyectos que nunca tuvieron presupuesto; proyectos cuyos únicos presupuestos están rechazados. NO incluye proyectos que tengan al menos un Aprobado/Borrador.

            var proyectosSinPeriodo = await _context.Proyectos
                .AsNoTracking()
                .Where(p =>
                    p.Activo &&
                   !_context.PresupuestosMensuales.Any(pm =>
    pm.ProyectoId == p.IdProyecto &&
    (
        pm.EstadoAprobacion == "Aprobado" ||
        pm.EstadoAprobacion == "Borrador"
    ) &&
            // NUEVO:
            !_context.CierresContables.Any(c =>
                c.Anio == pm.Anio &&
                c.Mes == pm.Mes &&
                c.Estado == "Cerrado") &&
    (
        pm.Anio > anioMinimo ||
        (pm.Anio == anioMinimo && pm.Mes >= mesMinimo)
    )))
                .OrderBy(p => p.Nombre)
                .ToListAsync();

            foreach (var proyecto in proyectosSinPeriodo)
            {
                string valor =
                    $"PROY-{proyecto.IdProyecto}";

                opciones.Add(new SelectListItem
                {
                    Value = valor,

                    Text =
                        $"{proyecto.Nombre} — Sin período presupuestario",

                    Selected =
                        valor == seleccionActual
                });
            }

            ViewBag.OpcionesProyectoPeriodo = opciones;
            ViewBag.SeleccionProyectoPeriodo = seleccionActual;
        }

        private async Task AplicarSeleccionProyectoPeriodoIngreso(
    Ingreso ingreso,
    string? seleccionProyectoPeriodo,
    int? presupuestoActualId = null)
        {
            // Sustituimos la validación Required normal de ProyectoId
            // porque ahora el proyecto se obtiene del selector combinado.
            ModelState.Remove(nameof(ingreso.ProyectoId));

            ingreso.ProyectoId = null;
            ingreso.PresupuestoMensualId = null;
            ingreso.MotivoPendientePresupuestario = null;

            if (string.IsNullOrWhiteSpace(seleccionProyectoPeriodo))
            {
                ModelState.AddModelError(
                    nameof(ingreso.ProyectoId),
                    "Debe seleccionar un proyecto y período presupuestario.");

                return;
            }

            // ============================================================
            // CASO 1:
            // Se seleccionó un presupuesto mensual.
            // Ejemplo: PRES-1025
            // ============================================================

            if (seleccionProyectoPeriodo.StartsWith("PRES-"))
            {
                string valorId =
                    seleccionProyectoPeriodo.Substring(5);

                if (!int.TryParse(
                        valorId,
                        out int idPresupuestoMensual))
                {
                    ModelState.AddModelError(
                        nameof(ingreso.ProyectoId),
                        "La selección del período presupuestario no es válida.");

                    return;
                }

                var presupuesto =
    await _context.PresupuestosMensuales
        .AsNoTracking()
        .Include(pm => pm.Proyecto)
        .FirstOrDefaultAsync(pm =>
            pm.IdPresupuestoMensual ==
                idPresupuestoMensual &&

            (
                // EDIT:
                // puede conservar exactamente
                // el presupuesto que ya tenía.
                pm.IdPresupuestoMensual ==
                    presupuestoActualId

                ||

                // CREATE o cambio en EDIT:
                // debe ser actualmente utilizable.
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
                        c.Estado == "Cerrado")
                )
            ));

                if (presupuesto == null)
                {
                    ModelState.AddModelError(
                        nameof(ingreso.ProyectoId),
                        "El período presupuestario seleccionado ya no está disponible.");

                    return;
                }

                ingreso.ProyectoId =
                    presupuesto.ProyectoId;

                ingreso.PresupuestoMensualId =
                    presupuesto.IdPresupuestoMensual;

                ingreso.MotivoPendientePresupuestario =
                    null;

                return;
            }

            // ============================================================
            // CASO 2:
            // Proyecto activo sin período presupuestario utilizable.
            // Ejemplo: PROY-11
            // ============================================================

            if (seleccionProyectoPeriodo.StartsWith("PROY-"))
            {
                string valorId =
                    seleccionProyectoPeriodo.Substring(5);

                if (!int.TryParse(
                        valorId,
                        out int idProyecto))
                {
                    ModelState.AddModelError(
                        nameof(ingreso.ProyectoId),
                        "El proyecto seleccionado no es válido.");

                    return;
                }

                bool proyectoActivo =
                    await _context.Proyectos
                        .AnyAsync(p =>
                            p.IdProyecto == idProyecto &&
                            p.Activo);

                if (!proyectoActivo)
                {
                    ModelState.AddModelError(
                        nameof(ingreso.ProyectoId),
                        "El proyecto seleccionado está inactivo o no existe.");

                    return;
                }

                // Seguridad:
                // verifica que realmente no exista un presupuesto
                // Aprobado o Borrador para este proyecto.
                bool tienePeriodoDisponible =
                    await _context.PresupuestosMensuales
                        .AnyAsync(pm =>
                            pm.ProyectoId == idProyecto &&
                            (
                                pm.EstadoAprobacion == "Aprobado" ||
                                pm.EstadoAprobacion == "Borrador"
                                  ) &&

            // NUEVO:
            !_context.CierresContables.Any(c =>
                c.Anio == pm.Anio &&
                c.Mes == pm.Mes &&
                c.Estado == "Cerrado"
                            ));

                if (tienePeriodoDisponible)
                {
                    ModelState.AddModelError(
                        nameof(ingreso.ProyectoId),
                        "El proyecto posee un período presupuestario disponible. Debe seleccionar ese período.");

                    return;
                }

                ingreso.ProyectoId =
                    idProyecto;

                ingreso.PresupuestoMensualId =
                    null;

                ingreso.MotivoPendientePresupuestario =
                    "Ingreso registrado sin período presupuestario.";

                return;
            }

            ModelState.AddModelError(
                nameof(ingreso.ProyectoId),
                "La selección de proyecto y período presupuestario no es válida.");
        }
        // GET: Ingresos/Create
        public async Task<IActionResult> Create()
        {
            var ingreso = new Ingreso
            {
                Fecha = DateOnly.FromDateTime(DateTime.Today),
                Estado = "Registrado",
                EsAnticipo = false,
                Activo = true
            };

            await CargarSelectsIngreso(
                ingreso,
                esEdicion: false);

            await CargarOpcionesProyectoPeriodoIngreso();

            return View(ingreso);
        }

        // POST: Ingresos/Crear
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Create(
    [Bind("IdIngreso,Fecha,CategoriaIngresoId,ClienteBeneficiarioId,Fuente,MontoEsperado,MontoReal,EsAnticipo,Descripcion,Estado,TipoIngreso,CentroCostoId,Observaciones,Comprobante")]
    Ingreso ingreso,string? SeleccionProyectoPeriodo)
        {
            // Verifica que la fecha ingresada no sea posterior a la fecha actual
            if (ingreso.Fecha > DateOnly.FromDateTime(DateTime.Today))
            {
                ModelState.AddModelError("Fecha",
                    "La fecha del ingreso no puede ser mayor a la fecha actual.");
            }
            // ============================================================
            // BLOQUEO POR CIERRE CONTABLE
            // ============================================================
            bool periodoCerrado =
                await PeriodoContableCerradoAsync(ingreso.Fecha);

            if (periodoCerrado)
            {
                ModelState.AddModelError(
                    "Fecha",
                    "No se puede registrar el ingreso porque el período contable seleccionado se encuentra cerrado.");
            }
            // Asigna automáticamente el estado del ingreso según el monto recibido.
            if ((ingreso.MontoReal ?? 0) == 0)
            {
                ingreso.Estado = "Pendiente";
            }
            else if ((ingreso.MontoReal ?? 0) < ingreso.MontoEsperado)
            {
                ingreso.Estado = "Parcial";
            }
            else
            {
                ingreso.Estado = "Pagado";
            }

            ingreso.Activo = true;
            // Calcula la diferencia entre el monto esperado y el monto real recibido.
            ingreso.Diferencia = ingreso.MontoEsperado - (ingreso.MontoReal ?? 0);
          
            // PROYECTO / PERIODO PRESUPUESTARIO
            await AplicarSeleccionProyectoPeriodoIngreso(
                ingreso,
                SeleccionProyectoPeriodo);

            // VALIDACIÓN DE RELACIONES
            await ValidarRelacionesIngreso(ingreso);
            if (ModelState.IsValid)
            {
        ingreso.Origen = "Manual";

                _context.Add(ingreso);   
                await _context.SaveChangesAsync();

                // Registra automáticamente los movimientos contables del ingreso manual.
                await _contabilidadService.RegistrarIngresoAsync(
          ingreso,
          HttpContext.Session.GetInt32("UsuarioId") ?? 1);

                // Auditoría: registra la creación del ingreso.
                await RegistrarAuditoria(
                    "Crear",
                    ingreso.IdIngreso,
                    "Se registró el ingreso con fuente " + ingreso.Fuente +
                    " por un monto real de ₡" + (ingreso.MontoReal ?? 0).ToString("N2") + ".");
                // Muestra mensaje de confirmación al usuario.
                TempData["MensajeExito"] =  "Ingreso creado correctamente.";
                return RedirectToAction(nameof(Index));
            }// vulve por validacion muestra solo activos 
            await CargarSelectsIngreso(
     ingreso,
     esEdicion: false);

            await CargarOpcionesProyectoPeriodoIngreso(
    SeleccionProyectoPeriodo);

            return View(ingreso);
        }

        // GET: Ingresos/Editar
        public async Task<IActionResult> Edit(int? id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var ingreso = await _context.Ingresos.FindAsync(id);
            if (ingreso == null)
            {             
                return NotFound();
            }
            // ============================================================
            // BLOQUEO POR CIERRE CONTABLE
            // ============================================================
            bool periodoCerrado =
                await PeriodoContableCerradoAsync(ingreso.Fecha);

            if (periodoCerrado)
            {
                TempData["MensajeError"] =
                    $"No se puede modificar el ingreso porque el período " +
                    $"{ingreso.Fecha:MM/yyyy} se encuentra cerrado. " +
                    $"Para realizar correcciones, primero debe reabrirse el período.";

                return RedirectToAction(
                    nameof(Index),
                    new
                    {
                        anio = ingreso.Fecha.Year,
                        mes = ingreso.Fecha.Month
                    });
            }
            // Los ingresos generados desde facturación solo se modifican desde la factura.
            string origenReal =
       DeterminarOrigenIngreso(ingreso);

            if (origenReal != "Manual")
            {
                TempData["MensajeError"] =
                    $"Este ingreso fue generado automáticamente desde {origenReal}. " +
                    "Debe modificarse desde su módulo de origen.";

                return RedirectToAction(nameof(Index));
            }
            // BLOQUEO DE INGRESOS ANULADOS
            if (ingreso.Estado == "Anulado")
            {
                TempData["MensajeError"] = "No se puede modificar un ingreso anulado.";
                return RedirectToAction(nameof(Index));
            }
            ingreso.MontoEsperado = ingreso.MontoEsperado.HasValue
    ? Math.Truncate(ingreso.MontoEsperado.Value)
    : null;

            ingreso.MontoReal = ingreso.MontoReal.HasValue
    ? Math.Truncate(ingreso.MontoReal.Value)
    : null;

            ingreso.Diferencia = ingreso.Diferencia.HasValue
                ? Math.Truncate(ingreso.Diferencia.Value)
                : null;
            await CargarSelectsIngreso(
                ingreso,
                esEdicion: true);

            string? seleccionActual =
    ingreso.PresupuestoMensualId.HasValue
        ? $"PRES-{ingreso.PresupuestoMensualId.Value}"
        : ingreso.ProyectoId.HasValue
            ? $"PROY-{ingreso.ProyectoId.Value}"
            : null;

            await CargarOpcionesProyectoPeriodoIngreso(
                seleccionActual);

            return View(ingreso);
        }
        // POST: Ingresos/Editar
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Edit(int id, [Bind("IdIngreso,Fecha,CategoriaIngresoId,ClienteBeneficiarioId,Fuente,MontoEsperado,MontoReal,EsAnticipo,Descripcion,Estado,Activo,TipoIngreso,CentroCostoId,Observaciones,Comprobante")]
    Ingreso ingreso,string? SeleccionProyectoPeriodo)
        {
            if (id != ingreso.IdIngreso)
            {
                return NotFound();
            }

            if (ingreso.Fecha > DateOnly.FromDateTime(DateTime.Today))
            {
                ModelState.AddModelError("Fecha",
                    "La fecha del ingreso no puede ser mayor a la fecha actual.");
            }

            // Busca el ingreso actual para validar su origen y estado antes de modificar.
            var ingresoActual = await _context.Ingresos
                .AsNoTracking()
                .FirstOrDefaultAsync(i => i.IdIngreso == id);

            if (ingresoActual == null)
            {
                return NotFound();
            }
            // ============================================================
            // BLOQUEO DEL PERÍODO ORIGINAL
            // ============================================================
            bool periodoOriginalCerrado =
                await PeriodoContableCerradoAsync(ingresoActual.Fecha);

            if (periodoOriginalCerrado)
            {
                TempData["MensajeError"] =
                    $"No se puede modificar el ingreso porque el período " +
                    $"{ingresoActual.Fecha:MM/yyyy} se encuentra cerrado. " +
                    $"Para realizar correcciones, primero debe reabrirse el período.";

                return RedirectToAction(
                    nameof(Index),
                    new
                    {
                        anio = ingresoActual.Fecha.Year,
                        mes = ingresoActual.Fecha.Month
                    });
            }
            // ============================================================
            // IMPEDIR MOVER EL INGRESO A UN PERÍODO CERRADO
            // ============================================================
            bool nuevoPeriodoCerrado =
                await PeriodoContableCerradoAsync(ingreso.Fecha);

            if (nuevoPeriodoCerrado)
            {
                ModelState.AddModelError(
                    "Fecha",
                    "La nueva fecha corresponde a un período contable cerrado.");
            }
            // Solo los ingresos manuales se pueden modificar desde el módulo de ingresos.
            string origenReal =
     DeterminarOrigenIngreso(ingresoActual);

            if (origenReal != "Manual")
            {
                TempData["MensajeError"] =
                    $"Este ingreso fue generado automáticamente desde {origenReal}. " +
                    "Debe modificarse desde su módulo de origen.";

                return RedirectToAction(nameof(Index));
            }

            // No se permite modificar ingresos anulados.
            if (ingresoActual.Estado == "Anulado")
            {
                TempData["MensajeError"] = "No se puede modificar un ingreso anulado.";
                return RedirectToAction(nameof(Index));
            }

            await AplicarSeleccionProyectoPeriodoIngreso(
      ingreso,
      SeleccionProyectoPeriodo,
      ingresoActual.PresupuestoMensualId);

            await ValidarRelacionesIngreso(
    ingreso,
    ingresoActual);

            if (ModelState.IsValid)
            {
                try
                {
                    // Recalcula la diferencia y el estado según el monto recibido.
                    ingreso.Diferencia = ingreso.MontoEsperado - (ingreso.MontoReal ?? 0);

                    if ((ingreso.MontoReal ?? 0) == 0)
                    {
                        ingreso.Estado = "Pendiente";
                        ingreso.Activo = true;
                    }
                    else if ((ingreso.MontoReal ?? 0) < ingreso.MontoEsperado)
                    {
                        ingreso.Estado = "Parcial";
                        ingreso.Activo = true;
                    }
                    else
                    {
                        ingreso.Estado = "Pagado";
                        ingreso.Activo = true;
                    }

                    // Conserva el origen manual para que no se pierda la trazabilidad.
                    ingreso.Origen = "Manual";

                    // Si el ingreso anterior ya estaba contabilizado, reversa el asiento original.
                    await _contabilidadService.ReversarMovimientosAsync(
                        "Ingresos",
                        ingreso.IdIngreso,
                        HttpContext.Session.GetInt32("UsuarioId") ?? 1,
                        "Reversión por modificación de ingreso manual");

                    _context.Update(ingreso);
                    await _context.SaveChangesAsync();

                    // Si después de modificar queda pagado, genera el nuevo asiento contable.
                    await _contabilidadService.RegistrarIngresoAsync(
        ingreso,
        HttpContext.Session.GetInt32("UsuarioId") ?? 1);

                    // Auditoría: registra la modificación del ingreso.
                    await RegistrarAuditoria(
                        "Modificar",
                        ingreso.IdIngreso,
                        "Se modificó el ingreso manual con fuente " + ingreso.Fuente +
                        " y estado " + ingreso.Estado + ".");
                }
                catch (DbUpdateConcurrencyException)
                {
                    if (!IngresoExists(ingreso.IdIngreso))
                    {
                        return NotFound();
                    }
                    else
                    {
                        throw;
                    }
                }

                TempData["MensajeExito"] = "Ingreso modificado correctamente.";
                return RedirectToAction(nameof(Index));
            }

            await CargarSelectsIngreso(
     ingreso,
     esEdicion: true);

            await CargarOpcionesProyectoPeriodoIngreso(
                SeleccionProyectoPeriodo);

            return View(ingreso);
        }
        // ================================================================
        // VALIDA QUE LAS RELACIONES SELECCIONADAS SEAN VÁLIDAS
        // ================================================================
        private async Task ValidarRelacionesIngreso(
      Ingreso ingreso,
      Ingreso? ingresoActual = null)
        {
            bool categoriaValida =
                await _context.CategoriasIngresos.AnyAsync(c =>
                    c.IdCategoriaIngreso == ingreso.CategoriaIngresoId &&
                    (
                        c.Activo ||
                        (
                            ingresoActual != null &&
                            c.IdCategoriaIngreso ==
                                ingresoActual.CategoriaIngresoId
                        )
                    ));

            if (!categoriaValida)
            {
                ModelState.AddModelError(
                    nameof(ingreso.CategoriaIngresoId),
                    "La categoría seleccionada está inactiva o no existe.");
            }

            bool clienteValido =
                await _context.ClientesBeneficiarios.AnyAsync(c =>
                    c.IdClienteBeneficiario ==
                        ingreso.ClienteBeneficiarioId &&
                    (
                        c.Activo ||
                        (
                            ingresoActual != null &&
                            c.IdClienteBeneficiario ==
                                ingresoActual.ClienteBeneficiarioId
                        )
                    ));

            if (!clienteValido)
            {
                ModelState.AddModelError(
                    nameof(ingreso.ClienteBeneficiarioId),
                    "El cliente o beneficiario seleccionado está inactivo o no existe.");
            }

            bool proyectoValido =
                await _context.Proyectos.AnyAsync(p =>
                    p.IdProyecto == ingreso.ProyectoId &&
                    (
                        p.Activo ||
                        (
                            ingresoActual != null &&
                            p.IdProyecto == ingresoActual.ProyectoId
                        )
                    ));

            if (!proyectoValido)
            {
                ModelState.AddModelError(
                    nameof(ingreso.ProyectoId),
                    "El proyecto seleccionado está inactivo o no existe.");
            }

            bool centroValido =
                await _context.CentrosCostos.AnyAsync(c =>
                    c.IdCentroCosto == ingreso.CentroCostoId &&
                    (
                        c.Activo ||
                        (
                            ingresoActual != null &&
                            c.IdCentroCosto ==
                                ingresoActual.CentroCostoId
                        )
                    ));

            if (!centroValido)
            {
                ModelState.AddModelError(
                    nameof(ingreso.CentroCostoId),
                    "El centro de costo seleccionado está inactivo o no existe.");
            }
        }
        // GET: Ingresos/eliminar
        public async Task<IActionResult> Delete(int? id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var ingreso = await _context.Ingresos
                .Include(i => i.CategoriaIngreso)
                .Include(i => i.ClienteBeneficiario)
                .Include(i => i.Proyecto)
                .Include(i => i.PresupuestoMensualRegistro)
                .Include(i => i.CentroCosto)
                .FirstOrDefaultAsync(i =>
                    i.IdIngreso == id);

            if (ingreso == null)
            {
                return NotFound();
            }
            // ============================================================
            // BLOQUEO POR CIERRE CONTABLE
            // ============================================================
            bool periodoCerrado =
                await PeriodoContableCerradoAsync(ingreso.Fecha);

            if (periodoCerrado)
            {
                TempData["MensajeError"] =
                    $"No se puede anular el ingreso porque el período " +
                    $"{ingreso.Fecha:MM/yyyy} se encuentra cerrado. " +
                    $"Para realizar correcciones, primero debe reabrirse el período.";

                return RedirectToAction(
                    nameof(Index),
                    new
                    {
                        anio = ingreso.Fecha.Year,
                        mes = ingreso.Fecha.Month
                    });
            }
            // Solo los ingresos manuales pueden anularse desde este módulo.
            string origenReal =
                DeterminarOrigenIngreso(ingreso);

            if (origenReal != "Manual")
            {
                TempData["MensajeError"] =
                    $"Este ingreso fue generado automáticamente desde {origenReal}. " +
                    "Debe anularse desde su módulo de origen.";

                return RedirectToAction(nameof(Index));
            }

            if (ingreso.Estado == "Anulado" ||
                !ingreso.Activo)
            {
                TempData["MensajeError"] =
                    "El ingreso ya fue anulado.";

                return RedirectToAction(nameof(Index));
            }

            return View(ingreso);
        }
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> DeleteConfirmed(int id)
        {
            var ingreso =
                await _context.Ingresos
                    .FirstOrDefaultAsync(i =>
                        i.IdIngreso == id);

            if (ingreso == null)
            {
                return NotFound();
            }
            // ============================================================
            // BLOQUEO POR CIERRE CONTABLE
            // PROTECCIÓN CONTRA POST DIRECTO
            // ============================================================
            bool periodoCerrado =
                await PeriodoContableCerradoAsync(ingreso.Fecha);

            if (periodoCerrado)
            {
                TempData["MensajeError"] =
                    $"No se puede anular el ingreso porque el período " +
                    $"{ingreso.Fecha:MM/yyyy} se encuentra cerrado. " +
                    $"Para realizar correcciones, primero debe reabrirse el período.";

                return RedirectToAction(
                    nameof(Index),
                    new
                    {
                        anio = ingreso.Fecha.Year,
                        mes = ingreso.Fecha.Month
                    });
            }
            // Vuelve a validar en el POST para impedir
            // anulaciones directas manipulando la solicitud.
            string origenReal =
       DeterminarOrigenIngreso(ingreso);

            if (origenReal != "Manual")
            {
                TempData["MensajeError"] =
                    $"Este ingreso fue generado automáticamente desde {origenReal}. " +
                    "Debe anularse desde su módulo de origen.";

                return RedirectToAction(nameof(Index));
            }

            if (ingreso.Estado == "Anulado" ||
                !ingreso.Activo)
            {
                TempData["MensajeError"] =
                    "El ingreso ya fue anulado.";

                return RedirectToAction(nameof(Index));
            }

            await _contabilidadService
                .ReversarMovimientosAsync(
                    "Ingresos",
                    ingreso.IdIngreso,
                    HttpContext.Session
                        .GetInt32("UsuarioId") ?? 1,
                    "Anulación de ingreso manual");

            ingreso.Activo = false;
            ingreso.Estado = "Anulado";

            _context.Update(ingreso);
            await _context.SaveChangesAsync();

            await RegistrarAuditoria(
                "Anular",
                ingreso.IdIngreso,
                "Se anuló el ingreso manual con fuente " +
                ingreso.Fuente +
                ".");

            TempData["MensajeExito"] =
                "Ingreso anulado correctamente.";

            return RedirectToAction(nameof(Index));
        }
        // ================================================================
        // CONSTRUYE LA DESCRIPCIÓN DE FILTROS PARA AUDITORÍA DE REPORTES
        // ================================================================
        private async Task<string> ConstruirDescripcionFiltros(
            string? buscar,
            string? tipo,
            string? estado,
            string? origen,
            int? proyectoId,
            int? anio,
            int? mes,
            DateOnly? fechaInicio,
            DateOnly? fechaFin,
             bool pendientePresupuesto = false)
        {
            var filtros = new List<string>();

            if (!string.IsNullOrWhiteSpace(buscar))
            {
                filtros.Add($"Búsqueda={buscar}");
            }

            if (!string.IsNullOrWhiteSpace(tipo))
            {
                filtros.Add($"Tipo={tipo}");
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
                    .Where(p => p.IdProyecto == proyectoId.Value)
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
            if (pendientePresupuesto)
            {
                filtros.Add(
                    "Asignación presupuestaria=Pendiente");
            }
            return filtros.Any()
                ? "Filtros aplicados: " +
                  string.Join(", ", filtros) +
                  "."
                : "Sin filtros aplicados.";
        }
        // Registra auditoría de las acciones realizadas en el módulo de ingresos.
        private async Task RegistrarAuditoria(string accion, int registroId, string descripcion)
        {
            int usuarioId = HttpContext.Session.GetInt32("UsuarioId") ?? 1;

            var auditoria = new Auditorium
            {
                UsuarioId = usuarioId,
                Tabla = "Ingresos",
                RegistroId = registroId,
                Accion = accion,
                Descripcion = descripcion,
                Fecha = DateTime.Now
            };

            _context.Auditoria.Add(auditoria);
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
        private bool IngresoExists(int id)
        {
            return _context.Ingresos.Any(e => e.IdIngreso == id);
        }
    }
}
