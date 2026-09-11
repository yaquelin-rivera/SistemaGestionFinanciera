using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using SistemaGestionFinanciera.Data;
using SistemaGestionFinanciera.Models;
using ClosedXML.Excel;
using System.IO;

namespace SistemaGestionFinanciera.Controllers
{
    public class CierresContablesController : BaseController
    {
        private readonly SistemaFinancieroContext _context;

        public CierresContablesController(
            SistemaFinancieroContext context)
        {
            _context = context;
        }
        // CONSULTA PRINCIPAL
        // Mes:
        // 0     = Todos los meses / resumen anual
        // 1-12  = Período mensual
        public async Task<IActionResult> Index(
            int? anio,
            int? mes)
        {
            DateTime hoy = DateTime.Today;

            // Al ingresar al módulo se muestra el mes anterior.
            DateTime periodoPredeterminado =
                new DateTime(
                    hoy.Year,
                    hoy.Month,
                    1)
                .AddMonths(-1);

            int anioSeleccionado =
                anio ?? periodoPredeterminado.Year;

            int mesSeleccionado =
                mes ?? periodoPredeterminado.Month;

            // Validación defensiva.
            if (anioSeleccionado < 2000 ||
                anioSeleccionado > 2100 ||
                mesSeleccionado < 0 ||
                mesSeleccionado > 12)
            {
                anioSeleccionado =
                    periodoPredeterminado.Year;

                mesSeleccionado =
                    periodoPredeterminado.Month;
            }

            // RESUMEN MENSUAL O ANUAL
            // ============================================================
            if (mesSeleccionado == 0)
            {
                await CargarDatosAnuales(
                    anioSeleccionado);
            }
            else
            {
                await CargarDatosPeriodo(
                    anioSeleccionado,
                    mesSeleccionado);
            }

            // DATOS GENERALES DE LA PANTALLA
            // ============================================================
            await CargarAniosDisponibles();

            await CargarHistorialCierres();

            await CargarPeriodosPendientes();

            ViewBag.Anio =
                anioSeleccionado;

            ViewBag.Mes =
                mesSeleccionado;

            ViewBag.EsConsultaAnual =
                mesSeleccionado == 0;

            ViewBag.Rol =
                HttpContext.Session.GetString("Rol")
                ?? "";

            return View();
        }
        // REPORTE DE CIERRE CONTABLE
        // Vista previa para impresión / PDF
        // ================================================================
        [HttpGet]
        public async Task<IActionResult> Reporte(
            int anio,
            int mes)
        {
            // ============================================================
            // VALIDAR PARÁMETROS
            // ============================================================
            if (anio < 2000 ||
                anio > 2100 ||
                mes < 0 ||
                mes > 12)
            {
                MostrarError(
                    "El período seleccionado no es válido.");

                return RedirectToAction(
                    nameof(Index));
            }

            // ============================================================
            // CONSULTA ANUAL
            // ============================================================
            if (mes == 0)
            {
                await CargarDatosAnuales(
                    anio);
            }
            else
            {
                // ========================================================
                // CONSULTA MENSUAL
                // ========================================================
                await CargarDatosPeriodo(
                    anio,
                    mes);
            }

            ViewBag.Anio =
                anio;

            ViewBag.Mes =
                mes;

            ViewBag.EsConsultaAnual =
                mes == 0;

            ViewBag.FechaGeneracion =
                DateTime.Now;

            string nombrePeriodo =
                mes == 0
                    ? $"Año {anio}"
                    : $"{ObtenerNombreMes(mes)} {anio}";

            ViewBag.NombrePeriodo =
                nombrePeriodo;

            // ============================================================
            // AUDITORÍA
            // ============================================================
            await RegistrarAuditoria(
                "Vista previa PDF",
                0,
                $"Se generó la vista previa del reporte de cierre contable " +
                $"correspondiente a {nombrePeriodo}.");

            return View();
        }


        // ================================================================
        // EXPORTAR CIERRE CONTABLE A EXCEL
        // ================================================================
        [HttpGet]
        public async Task<IActionResult> ExportarExcel(
            int anio,
            int mes)
        {
            // ============================================================
            // VALIDAR PARÁMETROS
            // ============================================================
            if (anio < 2000 ||
                anio > 2100 ||
                mes < 0 ||
                mes > 12)
            {
                MostrarError(
                    "El período seleccionado no es válido.");

                return RedirectToAction(
                    nameof(Index));
            }

            // ============================================================
            // OBTENER DATOS
            // ============================================================
            if (mes == 0)
            {
                await CargarDatosAnuales(
                    anio);
            }
            else
            {
                await CargarDatosPeriodo(
                    anio,
                    mes);
            }

            decimal totalIngresos =
                ViewBag.TotalIngresos ?? 0m;

            decimal totalGastos =
                ViewBag.TotalGastos ?? 0m;

            decimal resultadoPeriodo =
                ViewBag.ResultadoPeriodo ?? 0m;

            decimal totalDebe =
                ViewBag.TotalDebe ?? 0m;

            decimal totalHaber =
                ViewBag.TotalHaber ?? 0m;

            decimal diferencia =
                ViewBag.Diferencia ?? 0m;

            int cantidadMovimientos =
                ViewBag.CantidadMovimientos ?? 0;

            string estadoPeriodo =
                mes == 0
                    ? "Consulta anual"
                    : ViewBag.EstadoPeriodo ?? "Abierto";

            string nombrePeriodo =
                mes == 0
                    ? $"Año {anio}"
                    : $"{ObtenerNombreMes(mes)} {anio}";


            // ============================================================
            // CREAR LIBRO
            // ============================================================
            using var workbook =
                new XLWorkbook();

            var hoja =
                workbook.Worksheets.Add(
                    "Cierre contable");

            const int totalColumnas = 4;


            // ============================================================
            // COLORES INSTITUCIONALES
            // ============================================================
            var colorInstitucional =
                XLColor.FromHtml("#0F6B73");

            var colorInstitucionalOscuro =
                XLColor.FromHtml("#0B4F55");

            var colorFondoFiltros =
                XLColor.FromHtml("#F3F7F8");

            var colorVerdeSuave =
                XLColor.FromHtml("#D1E7DD");

            var colorRojoSuave =
                XLColor.FromHtml("#F8D7DA");

            var colorAzulSuave =
                XLColor.FromHtml("#D9EAF7");

            var colorDoradoSuave =
                XLColor.FromHtml("#FFF3CD");

            var colorGrisSuave =
                XLColor.FromHtml("#E2E3E5");


            // ============================================================
            // ENCABEZADO INSTITUCIONAL
            // ============================================================
            hoja.Cell("A1").Value =
                "Unión Cantonal de Asociaciones de Desarrollo de Santa Ana";

            hoja.Range(
                1,
                1,
                1,
                totalColumnas).Merge();

            hoja.Cell("A2").Value =
                "Reporte de Cierre Contable";

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


            // ============================================================
            // ESTILO ENCABEZADO
            // ============================================================
            var rangoInstitucion =
                hoja.Range(
                    1,
                    1,
                    1,
                    totalColumnas);

            rangoInstitucion.Style.Font.Bold =
                true;

            rangoInstitucion.Style.Font.FontSize =
                16;

            rangoInstitucion.Style.Font.FontColor =
                XLColor.FromHtml("#0F5C64");

            rangoInstitucion.Style.Alignment.Horizontal =
                XLAlignmentHorizontalValues.Center;

            rangoInstitucion.Style.Alignment.Vertical =
                XLAlignmentVerticalValues.Center;


            var rangoTitulo =
                hoja.Range(
                    2,
                    1,
                    2,
                    totalColumnas);

            rangoTitulo.Style.Font.Bold =
                true;

            rangoTitulo.Style.Font.FontSize =
                13;

            rangoTitulo.Style.Alignment.Horizontal =
                XLAlignmentHorizontalValues.Center;


            var rangoFecha =
                hoja.Range(
                    3,
                    1,
                    3,
                    totalColumnas);

            rangoFecha.Style.Font.FontColor =
                XLColor.FromHtml("#666666");

            rangoFecha.Style.Alignment.Horizontal =
                XLAlignmentHorizontalValues.Center;

            hoja.Row(1).Height = 24;
            hoja.Row(2).Height = 22;
            hoja.Row(3).Height = 20;


            // ============================================================
            // PERÍODO CONSULTADO
            // ============================================================
            hoja.Cell("A5").Value =
                "Período consultado";

            hoja.Range(
                5,
                1,
                5,
                totalColumnas).Merge();

            var tituloPeriodo =
                hoja.Range(
                    5,
                    1,
                    5,
                    totalColumnas);

            tituloPeriodo.Style.Fill.BackgroundColor =
                colorInstitucional;

            tituloPeriodo.Style.Font.FontColor =
                XLColor.White;

            tituloPeriodo.Style.Font.Bold =
                true;


            hoja.Cell("A6").Value =
                "Período:";

            hoja.Cell("B6").Value =
                nombrePeriodo;

            hoja.Cell("C6").Value =
                "Estado:";

            hoja.Cell("D6").Value =
                estadoPeriodo;

            hoja.Cell("A6").Style.Font.Bold =
                true;

            hoja.Cell("C6").Style.Font.Bold =
                true;

            var rangoPeriodo =
                hoja.Range("A6:D6");

            rangoPeriodo.Style.Fill.BackgroundColor =
                colorFondoFiltros;

            rangoPeriodo.Style.Border.OutsideBorder =
                XLBorderStyleValues.Thin;

            rangoPeriodo.Style.Border.InsideBorder =
                XLBorderStyleValues.Thin;

            rangoPeriodo.Style.Border.OutsideBorderColor =
                colorInstitucional;

            rangoPeriodo.Style.Border.InsideBorderColor =
                XLColor.LightGray;


            // ============================================================
            // RESUMEN FINANCIERO
            // ============================================================
            hoja.Cell("A8").Value =
                "Resumen financiero";

            hoja.Range(
                8,
                1,
                8,
                totalColumnas).Merge();

            var tituloResumen =
                hoja.Range(
                    8,
                    1,
                    8,
                    totalColumnas);

            tituloResumen.Style.Fill.BackgroundColor =
                colorInstitucional;

            tituloResumen.Style.Font.FontColor =
                XLColor.White;

            tituloResumen.Style.Font.Bold =
                true;


            hoja.Cell("A9").Value =
                "Ingresos";

            hoja.Cell("B9").Value =
                totalIngresos;

            hoja.Cell("C9").Value =
                "Gastos";

            hoja.Cell("D9").Value =
                totalGastos;


            hoja.Cell("A10").Value =
                "Resultado";

            hoja.Cell("B10").Value =
                resultadoPeriodo;

            hoja.Cell("C10").Value =
                "Movimientos";

            hoja.Cell("D10").Value =
                cantidadMovimientos;


            hoja.Range("A9:B9")
                .Style.Fill.BackgroundColor =
                colorVerdeSuave;

            hoja.Range("C9:D9")
                .Style.Fill.BackgroundColor =
                colorRojoSuave;

            hoja.Range("A10:B10")
                .Style.Fill.BackgroundColor =
                colorAzulSuave;

            hoja.Range("C10:D10")
                .Style.Fill.BackgroundColor =
                colorDoradoSuave;


            hoja.Range("A9:D10")
                .Style.Font.Bold =
                true;

            hoja.Range("A9:D10")
                .Style.Border.OutsideBorder =
                XLBorderStyleValues.Thin;

            hoja.Range("A9:D10")
                .Style.Border.InsideBorder =
                XLBorderStyleValues.Thin;


            // ============================================================
            // FORMATO MONEDA
            // ============================================================
            hoja.Cell("B9")
                .Style.NumberFormat.Format =
                "₡#,##0.00";

            hoja.Cell("D9")
                .Style.NumberFormat.Format =
                "₡#,##0.00";

            hoja.Cell("B10")
                .Style.NumberFormat.Format =
                "₡#,##0.00";


            // ============================================================
            // VALIDACIÓN CONTABLE
            // ============================================================
            hoja.Cell("A12").Value =
                "Validación contable";

            hoja.Range(
                12,
                1,
                12,
                totalColumnas).Merge();

            var tituloValidacion =
                hoja.Range(
                    12,
                    1,
                    12,
                    totalColumnas);

            tituloValidacion.Style.Fill.BackgroundColor =
                colorInstitucional;

            tituloValidacion.Style.Font.FontColor =
                XLColor.White;

            tituloValidacion.Style.Font.Bold =
                true;


            hoja.Cell("A13").Value =
                "Total Debe";

            hoja.Cell("B13").Value =
                totalDebe;

            hoja.Cell("C13").Value =
                "Total Haber";

            hoja.Cell("D13").Value =
                totalHaber;


            hoja.Cell("A14").Value =
                "Diferencia";

            hoja.Cell("B14").Value =
                diferencia;

            hoja.Cell("C14").Value =
                "Estado contable";

            hoja.Cell("D14").Value =
      cantidadMovimientos == 0
          ? "Sin movimientos"
          : diferencia == 0m
              ? "Balanceado"
              : "Desbalanceado";


            hoja.Cell("B13")
                .Style.NumberFormat.Format =
                "₡#,##0.00";

            hoja.Cell("D13")
                .Style.NumberFormat.Format =
                "₡#,##0.00";

            hoja.Cell("B14")
                .Style.NumberFormat.Format =
                "₡#,##0.00";


            hoja.Range("A13:D14")
                .Style.Border.OutsideBorder =
                XLBorderStyleValues.Thin;

            hoja.Range("A13:D14")
                .Style.Border.InsideBorder =
                XLBorderStyleValues.Thin;

            hoja.Range("A13:D14")
                .Style.Fill.BackgroundColor =
                colorGrisSuave;

            hoja.Range("A13:D14")
                .Style.Font.Bold =
                true;


            if (cantidadMovimientos == 0)
            {
                hoja.Cell("D14")
                    .Style.Fill.BackgroundColor =
                    colorGrisSuave;
            }
            else if (diferencia == 0m)
            {
                hoja.Cell("D14")
                    .Style.Fill.BackgroundColor =
                    colorVerdeSuave;
            }
            else
            {
                hoja.Cell("D14")
                    .Style.Fill.BackgroundColor =
                    colorRojoSuave;
            }
            // ============================================================
            // RESULTADO DEL PERÍODO
            // ============================================================
            string textoResultado;

            if (cantidadMovimientos == 0)
            {
                textoResultado =
                    $"Resultado del período: No existen movimientos financieros " +
                    $"registrados para {nombrePeriodo}.";
            }
            else if (resultadoPeriodo > 0)
            {
                textoResultado =
                    $"Resultado del período: Los ingresos superaron los gastos en " +
                    $"₡{resultadoPeriodo:N2}, generando un resultado positivo durante " +
                    $"{nombrePeriodo}.";
            }
            else if (resultadoPeriodo < 0)
            {
                textoResultado =
                    $"Resultado del período: Los gastos superaron los ingresos en " +
                    $"₡{Math.Abs(resultadoPeriodo):N2}, generando un resultado negativo durante " +
                    $"{nombrePeriodo}.";
            }
            else
            {
                textoResultado =
                    $"Resultado del período: Los ingresos y los gastos fueron equivalentes " +
                    $"durante {nombrePeriodo}, por lo que el resultado del período fue de ₡0.00.";
            }

            hoja.Cell("A16").Value = textoResultado;

            hoja.Range(
                16,
                1,
                16,
                totalColumnas).Merge();

            hoja.Cell("A16")
                .Style.Alignment.WrapText = true;

            hoja.Cell("A16")
                .Style.Font.Bold = false;

            hoja.Cell("A16")
                .Style.Border.OutsideBorder =
                XLBorderStyleValues.Thin;


            // Color según resultado
            if (cantidadMovimientos == 0)
            {
                hoja.Cell("A16")
                    .Style.Fill.BackgroundColor =
                    colorAzulSuave;
            }
            else if (resultadoPeriodo > 0)
            {
                hoja.Cell("A16")
                    .Style.Fill.BackgroundColor =
                    colorVerdeSuave;
            }
            else if (resultadoPeriodo < 0)
            {
                hoja.Cell("A16")
                    .Style.Fill.BackgroundColor =
                    colorDoradoSuave;
            }
            else
            {
                hoja.Cell("A16")
                    .Style.Fill.BackgroundColor =
                    colorAzulSuave;
            }

            hoja.Row(16).Height = 30;
            // ============================================================
            // NOTA
            // ============================================================
            hoja.Cell("A18").Value =
                mes == 0
                    ? "La información corresponde a una consulta acumulada anual. " +
                      "No representa un cierre contable anual."
                    : estadoPeriodo == "Cerrado"
                        ? "La información corresponde a la fotografía contable " +
                          "almacenada al momento del cierre del período."
                        : "La información corresponde al estado actual del período " +
                          "contable consultado.";

            hoja.Range(
                18,
                1,
                18,
                totalColumnas).Merge();

            hoja.Cell("A18")
                .Style.Alignment.WrapText =
                true;

            hoja.Cell("A18")
                .Style.Font.Italic =
                true;

            hoja.Cell("A18")
                .Style.Font.FontColor =
                XLColor.FromHtml("#666666");


            // ============================================================
            // AJUSTES DE HOJA
            // ============================================================
            hoja.Columns()
                .AdjustToContents();

            hoja.Column(1).Width =
                Math.Max(
                    hoja.Column(1).Width,
                    18);

            hoja.Column(2).Width =
                Math.Max(
                    hoja.Column(2).Width,
                    20);

            hoja.Column(3).Width =
                Math.Max(
                    hoja.Column(3).Width,
                    18);

            hoja.Column(4).Width =
                Math.Max(
                    hoja.Column(4).Width,
                    20);

            hoja.SheetView.FreezeRows(3);

            hoja.PageSetup.PageOrientation =
                XLPageOrientation.Portrait;

            hoja.PageSetup.FitToPages(
                1,
                1);

            hoja.PageSetup.Margins.Top =
                0.5;

            hoja.PageSetup.Margins.Bottom =
                0.5;

            hoja.PageSetup.Margins.Left =
                0.5;

            hoja.PageSetup.Margins.Right =
                0.5;


            // ============================================================
            // AUDITORÍA
            // ============================================================
            await RegistrarAuditoria(
                "Exportar Excel",
                0,
                $"Se exportó a Excel el reporte de cierre contable " +
                $"correspondiente a {nombrePeriodo}.");


            // ============================================================
            // DESCARGAR ARCHIVO
            // ============================================================
            using var stream =
                new MemoryStream();

            workbook.SaveAs(
                stream);

            stream.Position =
                0;

            string nombreArchivo =
                mes == 0
                    ? $"Cierre_Contable_{anio}.xlsx"
                    : $"Cierre_Contable_{ObtenerNombreMes(mes)}_{anio}.xlsx";

            return File(
                stream.ToArray(),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                nombreArchivo);
        }
        // ================================================================
        // CERRAR PERÍODO
        // ================================================================
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Cerrar(
            int anio,
            int mes)
        {
            // ============================================================
            // 1. VALIDAR PERÍODO
            // ============================================================
            if (!PeriodoMensualValido(
                anio,
                mes))
            {
                MostrarError(
                    "El período seleccionado no es válido.");

                return RedirectToAction(
                    nameof(Index),
                    new { anio, mes });
            }

            DateTime hoy =
                DateTime.Today;

            DateTime periodoSeleccionado =
                new DateTime(
                    anio,
                    mes,
                    1);

            DateTime periodoActual =
                new DateTime(
                    hoy.Year,
                    hoy.Month,
                    1);

            // ============================================================
            // 2. SOLO PERÍODOS ANTERIORES
            // ============================================================
            if (periodoSeleccionado >=
                periodoActual)
            {
                MostrarAdvertencia(
                    "Solo se pueden cerrar períodos anteriores al mes actual.");

                return RedirectToAction(
                    nameof(Index),
                    new { anio, mes });
            }

            // ============================================================
            // 3. PERMISOS
            // Administrador y Contador pueden cerrar.
            // ============================================================
            string rol =
                HttpContext.Session.GetString("Rol")
                ?? "";

            bool puedeCerrar =
                rol.Equals(
                    "Administrador",
                    StringComparison.OrdinalIgnoreCase)
                ||
                rol.Equals(
                    "Contador",
                    StringComparison.OrdinalIgnoreCase);

            if (!puedeCerrar)
            {
                MostrarAdvertencia(
                    "No tiene permisos para realizar el cierre contable.");

                return RedirectToAction(
                    nameof(Index),
                    new { anio, mes });
            }

            // ============================================================
            // 4. BUSCAR CIERRE EXISTENTE
            // ============================================================
            var cierreExistente =
                await _context.CierresContables
                    .FirstOrDefaultAsync(c =>
                        c.Anio == anio &&
                        c.Mes == mes);

            if (cierreExistente != null &&
                cierreExistente.Estado ==
                "Cerrado")
            {
                MostrarAdvertencia(
                    "El período seleccionado ya se encuentra cerrado.");

                return RedirectToAction(
                    nameof(Index),
                    new { anio, mes });
            }

            // ============================================================
            // 5. CALCULAR RESUMEN ACTUAL
            // ============================================================
            var resumen =
                await CalcularResumenPeriodo(
                    anio,
                    mes);

            decimal diferencia =
                resumen.TotalDebe -
                resumen.TotalHaber;

            // ============================================================
            // 6. VALIDAR PARTIDA DOBLE
            // ============================================================
            if (diferencia != 0m)
            {
                MostrarError(
                    "No se puede cerrar el período porque el Debe y el Haber " +
                    $"presentan una diferencia de ₡{Math.Abs(diferencia):N2}.");

                return RedirectToAction(
                    nameof(Index),
                    new { anio, mes });
            }

            // ============================================================
            // 7. USUARIO
            // ============================================================
            int usuarioId =
                HttpContext.Session
                    .GetInt32("UsuarioId")
                ?? 0;

            if (usuarioId <= 0)
            {
                MostrarError(
                    "No se pudo identificar al usuario que realiza el cierre.");

                return RedirectToAction(
                    nameof(Index),
                    new { anio, mes });
            }

            // ============================================================
            // 8. CREAR O ACTUALIZAR CIERRE
            // ============================================================
            CierreContable cierre;

            if (cierreExistente == null)
            {
                cierre =
                    new CierreContable
                    {
                        Anio =
                            anio,

                        Mes =
                            mes,

                        Estado =
                            "Cerrado",

                        TotalIngresos =
                            resumen.TotalIngresos,

                        TotalGastos =
                            resumen.TotalGastos,

                        ResultadoPeriodo =
                            resumen.ResultadoPeriodo,

                        TotalDebe =
                            resumen.TotalDebe,

                        TotalHaber =
                            resumen.TotalHaber,

                        CantidadMovimientos =
                            resumen.CantidadMovimientos,

                        FechaCierre =
                            DateTime.Now,

                        UsuarioCierreId =
                            usuarioId
                    };

                _context.CierresContables
                    .Add(cierre);
            }
            else
            {
                // El período estaba Reabierto.
                cierre =
                    cierreExistente;

                cierre.Estado =
                    "Cerrado";

                cierre.TotalIngresos =
                    resumen.TotalIngresos;

                cierre.TotalGastos =
                    resumen.TotalGastos;

                cierre.ResultadoPeriodo =
                    resumen.ResultadoPeriodo;

                cierre.TotalDebe =
                    resumen.TotalDebe;

                cierre.TotalHaber =
                    resumen.TotalHaber;

                cierre.CantidadMovimientos =
                    resumen.CantidadMovimientos;

                cierre.FechaCierre =
                    DateTime.Now;

                cierre.UsuarioCierreId =
                    usuarioId;

                // Los datos de la reapertura anterior
                // se conservan como referencia.
                // Auditoría mantiene toda la trazabilidad.
            }

            await _context.SaveChangesAsync();

            // ============================================================
            // 9. AUDITORÍA DEL CIERRE
            // ============================================================
            string nombreMes =
                ObtenerNombreMes(
                    mes);

            await RegistrarAuditoria(
                "Cerrar período",
                cierre.IdCierreContable,
                $"Se cerró el período contable {nombreMes} {anio}. " +
                $"Ingresos: ₡{resumen.TotalIngresos:N2}. " +
                $"Gastos: ₡{resumen.TotalGastos:N2}. " +
                $"Resultado: ₡{resumen.ResultadoPeriodo:N2}. " +
                $"Debe: ₡{resumen.TotalDebe:N2}. " +
                $"Haber: ₡{resumen.TotalHaber:N2}. " +
                $"Movimientos contables: {resumen.CantidadMovimientos}.");

            MostrarExito(
                $"El período {nombreMes} {anio} fue cerrado correctamente.");

            return RedirectToAction(
                nameof(Index),
                new { anio, mes });
        }

        // ================================================================
        // REABRIR PERÍODO
        // ================================================================
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Reabrir(
            int anio,
            int mes,
            string motivoReapertura)
        {
            // ============================================================
            // 1. SOLO ADMINISTRADOR
            // ============================================================
            string rol =
                HttpContext.Session.GetString("Rol")
                ?? "";

            if (!rol.Equals(
                "Administrador",
                StringComparison.OrdinalIgnoreCase))
            {
                MostrarAdvertencia(
                    "Solo el Administrador puede reabrir un período contable.");

                return RedirectToAction(
                    nameof(Index),
                    new { anio, mes });
            }

            // ============================================================
            // 2. VALIDAR PERÍODO
            // ============================================================
            if (!PeriodoMensualValido(
                anio,
                mes))
            {
                MostrarError(
                    "El período seleccionado no es válido.");

                return RedirectToAction(
                    nameof(Index),
                    new { anio, mes });
            }

            // ============================================================
            // 3. MOTIVO OBLIGATORIO
            // ============================================================
            if (string.IsNullOrWhiteSpace(
                motivoReapertura))
            {
                MostrarAdvertencia(
                    "Debe indicar el motivo de la reapertura.");

                return RedirectToAction(
                    nameof(Index),
                    new { anio, mes });
            }

            motivoReapertura =
                motivoReapertura.Trim();

            if (motivoReapertura.Length < 5)
            {
                MostrarAdvertencia(
                    "El motivo de reapertura debe contener al menos 5 caracteres.");

                return RedirectToAction(
                    nameof(Index),
                    new { anio, mes });
            }

            if (motivoReapertura.Length > 500)
            {
                MostrarAdvertencia(
                    "El motivo de reapertura no puede superar los 500 caracteres.");

                return RedirectToAction(
                    nameof(Index),
                    new { anio, mes });
            }

            // ============================================================
            // 4. BUSCAR CIERRE
            // ============================================================
            var cierre =
                await _context.CierresContables
                    .FirstOrDefaultAsync(c =>
                        c.Anio == anio &&
                        c.Mes == mes);

            if (cierre == null)
            {
                MostrarAdvertencia(
                    "El período seleccionado nunca ha sido cerrado.");

                return RedirectToAction(
                    nameof(Index),
                    new { anio, mes });
            }

            if (cierre.Estado !=
                "Cerrado")
            {
                MostrarAdvertencia(
                    "El período seleccionado ya se encuentra abierto o reabierto.");

                return RedirectToAction(
                    nameof(Index),
                    new { anio, mes });
            }

            // ============================================================
            // 5. IDENTIFICAR USUARIO
            // ============================================================
            int usuarioId =
                HttpContext.Session
                    .GetInt32("UsuarioId")
                ?? 0;

            if (usuarioId <= 0)
            {
                MostrarError(
                    "No se pudo identificar al usuario que realiza la reapertura.");

                return RedirectToAction(
                    nameof(Index),
                    new { anio, mes });
            }

            // ============================================================
            // 6. REABRIR
            // ============================================================
            cierre.Estado =
                "Reabierto";

            cierre.FechaReapertura =
                DateTime.Now;

            cierre.UsuarioReaperturaId =
                usuarioId;

            cierre.MotivoReapertura =
                motivoReapertura;

            await _context.SaveChangesAsync();

            // ============================================================
            // 7. AUDITORÍA
            // ============================================================
            string nombreMes =
                ObtenerNombreMes(
                    mes);

            await RegistrarAuditoria(
                "Reabrir período",
                cierre.IdCierreContable,
                $"Se reabrió el período contable {nombreMes} {anio}. " +
                $"Motivo: {motivoReapertura}");

            MostrarExito(
                $"El período {nombreMes} {anio} fue reabierto correctamente.");

            return RedirectToAction(
                nameof(Index),
                new { anio, mes });
        }

        // ================================================================
        // CARGAR DATOS DE UN MES
        // ================================================================
        private async Task CargarDatosPeriodo(
            int anio,
            int mes)
        {
            var cierre =
                await _context.CierresContables
                    .AsNoTracking()
                    .Include(c =>
                        c.UsuarioCierre)
                    .Include(c =>
                        c.UsuarioReapertura)
                    .FirstOrDefaultAsync(c =>
                        c.Anio == anio &&
                        c.Mes == mes);

            var resumenActual =
                await CalcularResumenPeriodo(
                    anio,
                    mes);

            string estadoPeriodo =
                cierre?.Estado
                ?? "Abierto";

            ViewBag.EstadoPeriodo =
                estadoPeriodo;

            ViewBag.Cierre =
                cierre;

            ViewBag.EsConsultaAnual =
                false;

            // ============================================================
            // PERÍODO CERRADO:
            // mostrar fotografía almacenada.
            // ============================================================
            if (cierre != null &&
                cierre.Estado ==
                "Cerrado")
            {
                ViewBag.TotalIngresos =
                    cierre.TotalIngresos;

                ViewBag.TotalGastos =
                    cierre.TotalGastos;

                ViewBag.ResultadoPeriodo =
                    cierre.ResultadoPeriodo;

                ViewBag.TotalDebe =
                    cierre.TotalDebe;

                ViewBag.TotalHaber =
                    cierre.TotalHaber;

                ViewBag.CantidadMovimientos =
                    cierre.CantidadMovimientos;

                ViewBag.Diferencia =
                    cierre.TotalDebe -
                    cierre.TotalHaber;
            }
            else
            {
                // Abierto o Reabierto:
                // mostrar situación actual.
                ViewBag.TotalIngresos =
                    resumenActual.TotalIngresos;

                ViewBag.TotalGastos =
                    resumenActual.TotalGastos;

                ViewBag.ResultadoPeriodo =
                    resumenActual.ResultadoPeriodo;

                ViewBag.TotalDebe =
                    resumenActual.TotalDebe;

                ViewBag.TotalHaber =
                    resumenActual.TotalHaber;

                ViewBag.CantidadMovimientos =
                    resumenActual.CantidadMovimientos;

                ViewBag.Diferencia =
                    resumenActual.TotalDebe -
                    resumenActual.TotalHaber;
            }

            DateTime hoy =
                DateTime.Today;

            DateTime periodoActual =
                new DateTime(
                    hoy.Year,
                    hoy.Month,
                    1);

            DateTime periodoSeleccionado =
                new DateTime(
                    anio,
                    mes,
                    1);

            bool esPeriodoAnterior =
                periodoSeleccionado <
                periodoActual;

            bool estaCerrado =
                cierre != null &&
                cierre.Estado ==
                "Cerrado";

            ViewBag.EsPeriodoAnterior =
                esPeriodoAnterior;

            ViewBag.PuedeCerrar =
                esPeriodoAnterior &&
                !estaCerrado;

            ViewBag.EsPeriodoActual =
                periodoSeleccionado ==
                periodoActual;

            ViewBag.EsPeriodoFuturo =
                periodoSeleccionado >
                periodoActual;

            ViewBag.NombreMes =
                ObtenerNombreMes(
                    mes);

            ViewBag.NombrePeriodo =
                $"{ObtenerNombreMes(mes)} {anio}";
        }

        // ================================================================
        // CARGAR RESUMEN DE TODO EL AÑO
        //
        // Es únicamente una consulta.
        // No representa un cierre anual.
        // ================================================================
        private async Task CargarDatosAnuales(
            int anio)
        {
            var resumen =
                await CalcularResumenAnual(
                    anio);

            ViewBag.EstadoPeriodo =
                "Consulta anual";

            ViewBag.Cierre =
                null;

            ViewBag.TotalIngresos =
                resumen.TotalIngresos;

            ViewBag.TotalGastos =
                resumen.TotalGastos;

            ViewBag.ResultadoPeriodo =
                resumen.ResultadoPeriodo;

            ViewBag.TotalDebe =
                resumen.TotalDebe;

            ViewBag.TotalHaber =
                resumen.TotalHaber;

            ViewBag.Diferencia =
                resumen.TotalDebe -
                resumen.TotalHaber;

            ViewBag.CantidadMovimientos =
                resumen.CantidadMovimientos;

            ViewBag.EsPeriodoAnterior =
                false;

            ViewBag.PuedeCerrar =
                false;

            ViewBag.EsPeriodoActual =
                false;

            ViewBag.EsPeriodoFuturo =
                false;

            ViewBag.NombreMes =
                "Todos los meses";

            ViewBag.NombrePeriodo =
                $"Año {anio}";

            ViewBag.EsConsultaAnual =
                true;
        }

        // ================================================================
        // CALCULAR RESUMEN DE UN MES
        // ================================================================
        private async Task<ResumenCierrePeriodo>
            CalcularResumenPeriodo(
                int anio,
                int mes)
        {
            var movimientos =
                await _context.MovimientosContables
                    .AsNoTracking()
                    .Include(m =>
                        m.CuentaContable)
                    .Where(m =>
                        m.Fecha.Year == anio &&
                        m.Fecha.Month == mes &&
                        m.Anulado != true &&
                        m.Estado == "Registrado")
                    .ToListAsync();

            return CalcularResumenDesdeMovimientos(
                movimientos);
        }

        // ================================================================
        // CALCULAR RESUMEN ANUAL
        // ================================================================
        private async Task<ResumenCierrePeriodo>
            CalcularResumenAnual(
                int anio)
        {
            var movimientos =
                await _context.MovimientosContables
                    .AsNoTracking()
                    .Include(m =>
                        m.CuentaContable)
                    .Where(m =>
                        m.Fecha.Year == anio &&
                        m.Anulado != true &&
                        m.Estado == "Registrado")
                    .ToListAsync();

            return CalcularResumenDesdeMovimientos(
                movimientos);
        }

        // ================================================================
        // CÁLCULO COMÚN DEL RESUMEN
        // ================================================================
        private ResumenCierrePeriodo
            CalcularResumenDesdeMovimientos(
                List<MovimientosContable> movimientos)
        {
            decimal totalDebe =
                movimientos.Sum(m =>
                    m.Debe);

            decimal totalHaber =
                movimientos.Sum(m =>
                    m.Haber);

            decimal totalIngresos =
                movimientos
                    .Where(m =>
                        m.CuentaContable != null &&
                        m.CuentaContable.TipoCuenta ==
                        "Ingreso")
                    .Sum(m =>
                        m.Haber -
                        m.Debe);

            decimal totalGastos =
                movimientos
                    .Where(m =>
                        m.CuentaContable != null &&
                        m.CuentaContable.TipoCuenta ==
                        "Gasto")
                    .Sum(m =>
                        m.Debe -
                        m.Haber);

            return new ResumenCierrePeriodo
            {
                TotalIngresos =
                    totalIngresos,

                TotalGastos =
                    totalGastos,

                ResultadoPeriodo =
                    totalIngresos -
                    totalGastos,

                TotalDebe =
                    totalDebe,

                TotalHaber =
                    totalHaber,

                CantidadMovimientos =
                    movimientos.Count
            };
        }

        // ================================================================
        // HISTORIAL DE CIERRES
        //
        // Se muestra primero el cierre realizado más recientemente.
        // DataTables permitirá después ordenar por cualquier columna.
        // ================================================================
        private async Task CargarHistorialCierres()
        {
            var historial =
                await _context.CierresContables
                    .AsNoTracking()
                    .Include(c =>
                        c.UsuarioCierre)
                    .Include(c =>
                        c.UsuarioReapertura)
                    .OrderByDescending(c =>
                        c.FechaCierre)
                    .ThenByDescending(c =>
                        c.Anio)
                    .ThenByDescending(c =>
                        c.Mes)
                    .ToListAsync();

            ViewBag.HistorialCierres =
                historial;

            ViewBag.TotalCierres =
                historial.Count;

            ViewBag.TotalPeriodosCerrados =
                historial.Count(c =>
                    c.Estado == "Cerrado");

            ViewBag.TotalPeriodosReabiertos =
                historial.Count(c =>
                    c.Estado == "Reabierto");
        }

        // ================================================================
        // PERÍODOS PENDIENTES DE CIERRE
        //
        // Solo se toman meses anteriores al actual que hayan tenido
        // movimientos contables.
        //
        // Un período Reabierto también se considera pendiente.
        // ================================================================
        private async Task CargarPeriodosPendientes()
        {
            DateTime hoy =
                DateTime.Today;

            DateOnly inicioMesActual =
                new DateOnly(
                    hoy.Year,
                    hoy.Month,
                    1);

            // Meses históricos con actividad contable.
            var periodosConMovimientos =
                await _context.MovimientosContables
                    .AsNoTracking()
                    .Where(m =>
                        m.Fecha <
                        inicioMesActual)
                    .Select(m =>
                        new
                        {
                            Anio =
                                m.Fecha.Year,

                            Mes =
                                m.Fecha.Month
                        })
                    .Distinct()
                    .ToListAsync();

            // Estados de cierre existentes.
            var cierres =
                await _context.CierresContables
                    .AsNoTracking()
                    .Select(c =>
                        new
                        {
                            c.Anio,
                            c.Mes,
                            c.Estado
                        })
                    .ToListAsync();

            var pendientes =
                periodosConMovimientos
                    .Where(p =>
                    {
                        var cierre =
                            cierres
                                .FirstOrDefault(c =>
                                    c.Anio ==
                                    p.Anio &&
                                    c.Mes ==
                                    p.Mes);

                        // No existe cierre:
                        // sigue pendiente.
                        if (cierre == null)
                        {
                            return true;
                        }

                        // Reabierto:
                        // vuelve a estar pendiente.
                        return cierre.Estado !=
                               "Cerrado";
                    })
                    .OrderByDescending(p =>
                        p.Anio)
                    .ThenByDescending(p =>
                        p.Mes)
                    .Select(p =>
                        new PeriodoPendienteCierre
                        {
                            Anio =
                                p.Anio,

                            Mes =
                                p.Mes,

                            NombreMes =
                                ObtenerNombreMes(
                                    p.Mes),

                            NombrePeriodo =
                                $"{ObtenerNombreMes(p.Mes)} {p.Anio}"
                        })
                    .ToList();

            ViewBag.PeriodosPendientes =
                pendientes;

            ViewBag.CantidadPeriodosPendientes =
                pendientes.Count;
        }

        // CARGAR AÑOS DISPONIBLES. Si en 2027 aparece el primer movimiento, 2027 aparecerá automáticamente.
        private async Task CargarAniosDisponibles()
        {
            int anioActual =
                DateTime.Today.Year;

            var aniosMovimientos =
                await _context.MovimientosContables
                    .AsNoTracking()
                    .Select(m =>
                        m.Fecha.Year)
                    .Distinct()
                    .ToListAsync();

            var aniosCierres =
                await _context.CierresContables
                    .AsNoTracking()
                    .Select(c =>
                        c.Anio)
                    .Distinct()
                    .ToListAsync();

            var anios =
                aniosMovimientos
                    .Concat(
                        aniosCierres)
                    .Append(
                        anioActual)
                    .Append(
                        anioActual - 1)
                    .Distinct()
                    .OrderByDescending(a =>
                        a)
                    .ToList();

            ViewBag.Anios =
                anios;
        }

        // ================================================================
        // VALIDAR PERÍODO MENSUAL
        // ================================================================
        private bool PeriodoMensualValido(
            int anio,
            int mes)
        {
            return
                anio >= 2000 &&
                anio <= 2100 &&
                mes >= 1 &&
                mes <= 12;
        }

        // ================================================================
        // NOMBRE DEL MES
        // ================================================================
        private string ObtenerNombreMes(
            int mes)
        {
            return mes switch
            {
                1 => "Enero",
                2 => "Febrero",
                3 => "Marzo",
                4 => "Abril",
                5 => "Mayo",
                6 => "Junio",
                7 => "Julio",
                8 => "Agosto",
                9 => "Septiembre",
                10 => "Octubre",
                11 => "Noviembre",
                12 => "Diciembre",
                _ => ""
            };
        }

        // AUDITORÍA
        // Se registran las operaciones que alteran el estado del período: Cerrar y Reabrir.
        // No registre cada carga del Index para evitar llenar

        private async Task RegistrarAuditoria(
            string accion,
            int registroId,
            string descripcion)
        {
            int usuarioId =
                HttpContext.Session
                    .GetInt32("UsuarioId")
                ?? 0;

            if (usuarioId <= 0)
            {
                return;
            }

            var auditoria =
                new Auditorium
                {
                    UsuarioId =
                        usuarioId,

                    Tabla =
                        "CierresContables",

                    RegistroId =
                        registroId,

                    Accion =
                        accion,

                    Descripcion =
                        descripcion,

                    Fecha =
                        DateTime.Now
                };

            _context.Auditoria
                .Add(auditoria);

            await _context
                .SaveChangesAsync();
        }

        // ================================================================
        // CLASE AUXILIAR INTERNA PARA CÁLCULOS
        // No crea ninguna tabla.
        // ================================================================
        private class ResumenCierrePeriodo
        {
            public decimal TotalIngresos
            {
                get;
                set;
            }

            public decimal TotalGastos
            {
                get;
                set;
            }

            public decimal ResultadoPeriodo
            {
                get;
                set;
            }

            public decimal TotalDebe
            {
                get;
                set;
            }

            public decimal TotalHaber
            {
                get;
                set;
            }

            public int CantidadMovimientos
            {
                get;
                set;
            }
        }
    }

    // AUXILIAR PARA MOSTRAR LOS PERÍODOS PENDIENTES EN LA VISTA
    // No es una entidad de EF Core.
    // No crea tabla en SQL Server.
    public class PeriodoPendienteCierre
    {
        public int Anio
        {
            get;
            set;
        }

        public int Mes
        {
            get;
            set;
        }

        public string NombreMes
        {
            get;
            set;
        } = "";

        public string NombrePeriodo
        {
            get;
            set;
        } = "";
    }
}