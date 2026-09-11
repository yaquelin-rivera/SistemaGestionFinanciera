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
using System.Globalization;
using System.IO;

namespace SistemaGestionFinanciera.Controllers
{
    public class PresupuestosMensualesController : BaseController
    {
        private readonly SistemaFinancieroContext _context;

        public PresupuestosMensualesController(SistemaFinancieroContext context)
        {
            _context = context;
        }
        // ===========================================================
        // CONSULTA COMPARTIDA PARA INDEX, REPORTE Y EXCEL
        // ===========================================================
        private async Task<List<PresupuestosMensuale>>
            ConstruirConsultaReporte(
                int? proyectoId,
                int? anio,
                int? mes,
                string? estado,
                string? estadoAprobacion)
        {
            // Se cargan todos los presupuestos con su proyecto.
            var presupuestos = await _context.PresupuestosMensuales
                .Include(p => p.Proyecto)
                .ToListAsync();

            // Se conserva la lógica actual:
            // antes de mostrar o exportar, se recalculan los datos reales.
            foreach (var presupuesto in presupuestos)
            {
                await RecalcularPresupuesto(presupuesto);
            }

            await _context.SaveChangesAsync();

            // Filtro por proyecto.
            if (proyectoId.HasValue &&
                proyectoId.Value > 0)
            {
                presupuestos = presupuestos
                    .Where(p =>
                        p.ProyectoId == proyectoId.Value)
                    .ToList();
            }

            // Filtro por año.
            if (anio.HasValue)
            {
                presupuestos = presupuestos
                    .Where(p =>
                        p.Anio == anio.Value)
                    .ToList();
            }

            // Filtro por mes.
            if (mes.HasValue)
            {
                presupuestos = presupuestos
                    .Where(p =>
                        p.Mes == mes.Value)
                    .ToList();
            }

            // Filtro por estado de ejecución.
            if (!string.IsNullOrWhiteSpace(estado))
            {
                presupuestos = presupuestos
                    .Where(p =>
                        p.Estado == estado)
                    .ToList();
            }

            // Filtro por estado de aprobación.
            if (!string.IsNullOrWhiteSpace(
                estadoAprobacion))
            {
                presupuestos = presupuestos
                    .Where(p =>
                        p.EstadoAprobacion ==
                        estadoAprobacion)
                    .ToList();
            }

            // Registros más recientes primero.
            return presupuestos
       .OrderByDescending(p => p.Anio)
       .ThenByDescending(p => p.Mes)
       .ThenBy(p => p.Proyecto != null
           ? p.Proyecto.Nombre
           : "")
       .ToList();
        }
        // ===========================================================
        // LISTADO DE PRESUPUESTOS MENSUALES
        // ===========================================================
        public async Task<IActionResult> Index(
            int? proyectoId,
            int? anio,
            int? mes,
            string? estado,
            string? estadoAprobacion,
            bool desdeProyectos = false)
        {
            if (string.IsNullOrEmpty(
                HttpContext.Session.GetString("Usuario")))
            {
                return RedirectToAction(
                    "Login",
                    "Acceso");
            }

            var presupuestos =
                await ConstruirConsultaReporte(
                    proyectoId,
                    anio,
                    mes,
                    estado,
                    estadoAprobacion);

            ViewBag.DesdeProyectos =
                desdeProyectos;

            // Conserva los filtros en pantalla.
            ViewBag.ProyectoId =
                proyectoId;

            ViewBag.Anio =
                anio;

            ViewBag.Mes =
                mes;

            ViewBag.Estado =
                estado;

            ViewBag.EstadoAprobacion =
                estadoAprobacion;

            ViewBag.Rol =
                HttpContext.Session.GetString("Rol")
                ?? "";

            // Proyectos activos.
            ViewData["Proyectos"] = new SelectList(
                await _context.Proyectos
                    .Where(p => p.Activo == true)
                    .OrderBy(p => p.Nombre)
                    .ToListAsync(),
                "IdProyecto",
                "Nombre",
                proyectoId);

            // Años disponibles.
            ViewBag.Anios =
                await _context.PresupuestosMensuales
                    .Select(p => p.Anio)
                    .Distinct()
                    .OrderByDescending(a => a)
                    .ToListAsync();

            return View(presupuestos);
        }
        // ===========================================================
        // VISTA PREVIA DEL REPORTE DE PRESUPUESTOS
        // ===========================================================
        public async Task<IActionResult> Reporte(
            int? proyectoId,
            int? anio,
            int? mes,
            string? estado,
            string? estadoAprobacion)
        {
            if (string.IsNullOrEmpty(
                HttpContext.Session.GetString("Usuario")))
            {
                return RedirectToAction(
                    "Login",
                    "Acceso");
            }

            var presupuestos =
                await ConstruirConsultaReporte(
                    proyectoId,
                    anio,
                    mes,
                    estado,
                    estadoAprobacion);

            // =======================================================
            // TOTALES GENERALES
            // =======================================================

            ViewBag.TotalPresupuestos =
                presupuestos.Count;

            ViewBag.TotalIngresoPlanificado =
                presupuestos.Sum(p =>
                    p.MontoIngresadoPlanificado ?? 0m);

            ViewBag.TotalGastoPlanificado =
                presupuestos.Sum(p =>
                    p.MontoGastoPlanificado ?? 0m);

            ViewBag.TotalIngresoReal =
                presupuestos.Sum(p =>
                    p.TotalIngresoReal ?? 0m);

            ViewBag.TotalGastoReal =
                presupuestos.Sum(p =>
                    p.TotalGastoReal ?? 0m);

            ViewBag.TotalDiferenciaIngreso =
                presupuestos.Sum(p =>
                    p.DiferenciaIngreso ?? 0m);

            ViewBag.TotalDiferenciaGasto =
                presupuestos.Sum(p =>
                    p.DiferenciaGasto ?? 0m);

            ViewBag.TotalBalanceReal =
                presupuestos.Sum(p =>
                    p.SaldoDisponible ?? 0m);

            // =======================================================
            // RESUMEN POR ESTADO DE APROBACIÓN
            // =======================================================

            ViewBag.CantidadBorradores =
                presupuestos.Count(p =>
                    p.EstadoAprobacion == "Borrador");

            ViewBag.CantidadAprobados =
                presupuestos.Count(p =>
                    p.EstadoAprobacion == "Aprobado");

            ViewBag.CantidadRechazados =
                presupuestos.Count(p =>
                    p.EstadoAprobacion == "Rechazado");

            // =======================================================
            // RESUMEN POR ESTADO DE EJECUCIÓN
            // =======================================================

            ViewBag.CantidadSinMovimientos =
                presupuestos.Count(p =>
                    p.Estado == "Sin movimientos");

            ViewBag.CantidadEnEjecucion =
                presupuestos.Count(p =>
                    p.Estado == "En ejecución");

            ViewBag.CantidadIngresosBajos =
                presupuestos.Count(p =>
                    p.Estado == "Ingresos bajos");

            ViewBag.CantidadGastoExcedido =
                presupuestos.Count(p =>
                    p.Estado == "Gasto excedido");

            // =======================================================
            // PRESUPUESTOS EN ALERTA
            // =======================================================

            ViewBag.CantidadAlertas =
                presupuestos.Count(p =>
                    (p.SaldoDisponible ?? 0m) < 0m ||

                    (p.TotalGastoReal ?? 0m) >
                    (p.MontoGastoPlanificado ?? 0m) ||

                    (p.TotalIngresoReal ?? 0m) <
                    (p.MontoIngresadoPlanificado ?? 0m));

            ViewBag.CantidadDeficit =
                presupuestos.Count(p =>
                    (p.SaldoDisponible ?? 0m) < 0m);

            // =======================================================
            // DATOS DEL REPORTE
            // =======================================================

            ViewBag.FechaGeneracion =
                DateTime.Now;

            // Conserva los filtros.
            ViewBag.ProyectoId =
                proyectoId;

            ViewBag.Anio =
                anio;

            ViewBag.Mes =
                mes;

            ViewBag.Estado =
                estado;

            ViewBag.EstadoAprobacion =
                estadoAprobacion;

            // Nombre del proyecto seleccionado.
            ViewBag.NombreProyecto =
                proyectoId.HasValue &&
                proyectoId.Value > 0
                    ? await _context.Proyectos
                        .Where(p =>
                            p.IdProyecto ==
                            proyectoId.Value)
                        .Select(p => p.Nombre)
                        .FirstOrDefaultAsync()
                    : null;
            await RegistrarAuditoria(
    "Vista previa PDF",
    0,
    "Se generó la vista previa del reporte de presupuestos mensuales. " +
    ConstruirDescripcionFiltrosReporte(
        proyectoId,
        anio,
        mes,
        estado,
        estadoAprobacion)
);
            return View(presupuestos);
        }
        // ===========================================================
        // EXPORTAR REPORTE DE PRESUPUESTOS A EXCEL
        // ===========================================================
        public async Task<IActionResult> ExportarExcel(
            int? proyectoId,
            int? anio,
            int? mes,
            string? estado,
            string? estadoAprobacion)
        {
            if (string.IsNullOrEmpty(HttpContext.Session.GetString("Usuario")))
            {
                return RedirectToAction("Login", "Acceso");
            }

            var presupuestos = await ConstruirConsultaReporte(
                proyectoId,
                anio,
                mes,
                estado,
                estadoAprobacion);

            using var workbook = new XLWorkbook();

            var ws = workbook.Worksheets.Add("Presupuesto");

            int fila = 1;

            ws.Cell(fila, 1).Value =
                "UNIÓN CANTONAL DE ASOCIACIONES DE DESARROLLO DE SANTA ANA";

            ws.Range(fila, 1, fila, 15).Merge();

            ws.Cell(fila, 1).Style.Font.Bold = true;
            ws.Cell(fila, 1).Style.Font.FontSize = 16;
            ws.Cell(fila, 1).Style.Alignment.Horizontal =
                XLAlignmentHorizontalValues.Center;

            fila++;

            ws.Cell(fila, 1).Value =
                "Reporte General de Presupuestos Mensuales";

            ws.Range(fila, 1, fila, 15).Merge();

            ws.Cell(fila, 1).Style.Font.Bold = true;
            ws.Cell(fila, 1).Style.Alignment.Horizontal =
                XLAlignmentHorizontalValues.Center;

            fila++;

            ws.Cell(fila, 1).Value =
                "Fecha de generación:";

            ws.Cell(fila, 2).Value =
    DateTime.Now.ToString("dd/MM/yyyy HH:mm");

            ws.Range(3, 2, 3, 3).Merge();

            ws.Cell(3, 2).Style.Alignment.WrapText = false;
            ws.Cell(3, 2).Style.Alignment.Horizontal =
                XLAlignmentHorizontalValues.Left;
            ws.Cell(3, 2).Style.Alignment.Vertical =
                XLAlignmentVerticalValues.Center;

            fila += 2;

            // ================================
            // FILTROS
            // ================================

            ws.Cell(fila, 1).Value = "Proyecto";
            ws.Cell(fila, 2).Value =
                proyectoId.HasValue
                    ? await _context.Proyectos
                        .Where(x => x.IdProyecto == proyectoId)
                        .Select(x => x.Nombre)
                        .FirstOrDefaultAsync()
                    : "Todos";

            fila++;

            ws.Cell(fila, 1).Value = "Año";
            ws.Cell(fila, 2).Value =
                anio?.ToString() ?? "Todos";

            fila++;

            ws.Cell(fila, 1).Value = "Mes";
            ws.Cell(fila, 2).Value =
                mes?.ToString() ?? "Todos";

            fila++;

            ws.Cell(fila, 1).Value = "Estado ejecución";
            ws.Cell(fila, 2).Value =
                estado ?? "Todos";

            fila++;

            ws.Cell(fila, 1).Value = "Estado aprobación";
            ws.Cell(fila, 2).Value =
                estadoAprobacion ?? "Todos";

            fila += 2;

            // ================================
            // ENCABEZADOS
            // ================================

            string[] encabezados =
            {
        "Proyecto",
        "Año",
        "Mes",
        "Ingreso Planificado",
        "Gasto Planificado",
        "Ingreso Real",
        "Gasto Real",
        "Dif. Ingreso",
        "Saldo Gasto",
        "Balance",
        "Estado",
        "Aprobación",
        "Observación",
        "Alerta",
        "Fecha"
    };

            for (int i = 0; i < encabezados.Length; i++)
            {
                ws.Cell(fila, i + 1).Value = encabezados[i];

                ws.Cell(fila, i + 1).Style.Font.Bold = true;
                ws.Cell(fila, i + 1).Style.Fill.BackgroundColor =
                    XLColor.FromHtml("#0F5C64");

                ws.Cell(fila, i + 1).Style.Font.FontColor =
                    XLColor.White;
            }

            fila++;
            // ================================
            // DATOS DEL REPORTE
            // ================================

            int filaInicioDatos = fila;

            foreach (var item in presupuestos)
            {
                string nombreMes = item.Mes switch
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
                    _ => "No definido"
                };

                decimal ingresoPlanificado =
                    item.MontoIngresadoPlanificado ?? 0m;

                decimal gastoPlanificado =
                    item.MontoGastoPlanificado ?? 0m;

                decimal ingresoReal =
                    item.TotalIngresoReal ?? 0m;

                decimal gastoReal =
                    item.TotalGastoReal ?? 0m;

                decimal diferenciaIngreso =
                    item.DiferenciaIngreso ?? 0m;

                decimal diferenciaGasto =
                    item.DiferenciaGasto ?? 0m;

                decimal balance =
                    item.SaldoDisponible ?? 0m;

                string alerta;

                if (balance < 0)
                {
                    alerta = "Déficit presupuestario";
                }
                else if (gastoReal > gastoPlanificado)
                {
                    alerta = "Gasto supera lo planificado";
                }
                else if (ingresoReal < ingresoPlanificado)
                {
                    alerta = "Ingreso menor al planificado";
                }
                else
                {
                    alerta = "Presupuesto controlado";
                }

                ws.Cell(fila, 1).Value =
                    item.Proyecto?.Nombre ?? "Sin proyecto";

                ws.Cell(fila, 2).Value =
                    item.Anio;

                ws.Cell(fila, 3).Value =
                    nombreMes;

                ws.Cell(fila, 4).Value =
                    ingresoPlanificado;

                ws.Cell(fila, 5).Value =
                    gastoPlanificado;

                ws.Cell(fila, 6).Value =
                    ingresoReal;

                ws.Cell(fila, 7).Value =
                    gastoReal;

                ws.Cell(fila, 8).Value =
                    diferenciaIngreso;

                ws.Cell(fila, 9).Value =
                    diferenciaGasto;

                ws.Cell(fila, 10).Value =
                    balance;

                ws.Cell(fila, 11).Value =
                    item.Estado ?? "Sin definir";

                ws.Cell(fila, 12).Value =
                    item.EstadoAprobacion ?? "Sin definir";

                ws.Cell(fila, 13).Value =
                    string.IsNullOrWhiteSpace(
                        item.ObservacionRevision)
                        ? "—"
                        : item.ObservacionRevision;

                ws.Cell(fila, 14).Value =
                    alerta;

                if (item.FechaCreacion.HasValue)
                {
                    ws.Cell(fila, 15).Value = item.FechaCreacion.Value;

                    ws.Cell(fila, 15)
                        .Style.DateFormat.Format = "dd/MM/yyyy";
                }
                else
                {
                    ws.Cell(fila, 15).Value = "";
                }

                // ================================
                // COLORES SEGÚN ESTADO DE EJECUCIÓN
                // ================================

                if (item.Estado == "Gasto excedido")
                {
                    ws.Cell(fila, 11)
                        .Style.Fill.BackgroundColor =
                        XLColor.FromHtml("#F8D7DA");

                    ws.Cell(fila, 11)
                        .Style.Font.FontColor =
                        XLColor.FromHtml("#842029");

                    ws.Cell(fila, 11)
                        .Style.Font.Bold = true;
                }
                else if (item.Estado == "Ingresos bajos")
                {
                    ws.Cell(fila, 11)
                        .Style.Fill.BackgroundColor =
                        XLColor.FromHtml("#FFF3CD");

                    ws.Cell(fila, 11)
                        .Style.Font.FontColor =
                        XLColor.FromHtml("#664D03");

                    ws.Cell(fila, 11)
                        .Style.Font.Bold = true;
                }
                else if (item.Estado == "En ejecución")
                {
                    ws.Cell(fila, 11)
                        .Style.Fill.BackgroundColor =
                        XLColor.FromHtml("#D1E7DD");

                    ws.Cell(fila, 11)
                        .Style.Font.FontColor =
                        XLColor.FromHtml("#0F5132");

                    ws.Cell(fila, 11)
                        .Style.Font.Bold = true;
                }
                else
                {
                    ws.Cell(fila, 11)
                        .Style.Fill.BackgroundColor =
                        XLColor.FromHtml("#E2E3E5");

                    ws.Cell(fila, 11)
                        .Style.Font.FontColor =
                        XLColor.FromHtml("#41464B");
                }

                // ================================
                // COLORES SEGÚN APROBACIÓN
                // ================================

                if (item.EstadoAprobacion == "Aprobado")
                {
                    ws.Cell(fila, 12)
                        .Style.Fill.BackgroundColor =
                        XLColor.FromHtml("#D1E7DD");

                    ws.Cell(fila, 12)
                        .Style.Font.FontColor =
                        XLColor.FromHtml("#0F5132");

                    ws.Cell(fila, 12)
                        .Style.Font.Bold = true;
                }
                else if (item.EstadoAprobacion == "Rechazado")
                {
                    ws.Cell(fila, 12)
                        .Style.Fill.BackgroundColor =
                        XLColor.FromHtml("#F8D7DA");

                    ws.Cell(fila, 12)
                        .Style.Font.FontColor =
                        XLColor.FromHtml("#842029");

                    ws.Cell(fila, 12)
                        .Style.Font.Bold = true;
                }
                else
                {
                    ws.Cell(fila, 12)
                        .Style.Fill.BackgroundColor =
                        XLColor.FromHtml("#E2E3E5");

                    ws.Cell(fila, 12)
                        .Style.Font.FontColor =
                        XLColor.FromHtml("#41464B");
                }

                // ================================
                // COLORES SEGÚN ALERTA
                // ================================

                if (alerta == "Déficit presupuestario" ||
                    alerta == "Gasto supera lo planificado")
                {
                    ws.Cell(fila, 14)
                        .Style.Fill.BackgroundColor =
                        XLColor.FromHtml("#F8D7DA");

                    ws.Cell(fila, 14)
                        .Style.Font.FontColor =
                        XLColor.FromHtml("#842029");

                    ws.Cell(fila, 14)
                        .Style.Font.Bold = true;
                }
                else if (alerta ==
                         "Ingreso menor al planificado")
                {
                    ws.Cell(fila, 14)
                        .Style.Fill.BackgroundColor =
                        XLColor.FromHtml("#FFF3CD");

                    ws.Cell(fila, 14)
                        .Style.Font.FontColor =
                        XLColor.FromHtml("#664D03");

                    ws.Cell(fila, 14)
                        .Style.Font.Bold = true;
                }
                else
                {
                    ws.Cell(fila, 14)
                        .Style.Fill.BackgroundColor =
                        XLColor.FromHtml("#D1E7DD");

                    ws.Cell(fila, 14)
                        .Style.Font.FontColor =
                        XLColor.FromHtml("#0F5132");

                    ws.Cell(fila, 14)
                        .Style.Font.Bold = true;
                }

                // ================================
                // COLORES DE DIFERENCIAS Y BALANCE
                // ================================

                if (diferenciaIngreso < 0)
                {
                    ws.Cell(fila, 8)
                        .Style.Font.FontColor =
                        XLColor.FromHtml("#842029");

                    ws.Cell(fila, 8)
                        .Style.Font.Bold = true;
                }
                else
                {
                    ws.Cell(fila, 8)
                        .Style.Font.FontColor =
                        XLColor.FromHtml("#0F5132");

                    ws.Cell(fila, 8)
                        .Style.Font.Bold = true;
                }

                if (diferenciaGasto < 0)
                {
                    ws.Cell(fila, 9)
                        .Style.Font.FontColor =
                        XLColor.FromHtml("#842029");

                    ws.Cell(fila, 9)
                        .Style.Font.Bold = true;
                }
                else
                {
                    ws.Cell(fila, 9)
                        .Style.Font.FontColor =
                        XLColor.FromHtml("#0F5132");

                    ws.Cell(fila, 9)
                        .Style.Font.Bold = true;
                }

                if (balance < 0)
                {
                    ws.Cell(fila, 10)
                        .Style.Font.FontColor =
                        XLColor.FromHtml("#842029");

                    ws.Cell(fila, 10)
                        .Style.Font.Bold = true;
                }
                else
                {
                    ws.Cell(fila, 10)
                        .Style.Font.FontColor =
                        XLColor.FromHtml("#0D6EFD");

                    ws.Cell(fila, 10)
                        .Style.Font.Bold = true;
                }

                fila++;
            }

            int filaFinDatos = fila - 1;

            // ================================
            // FORMATO MONEDA
            // ================================

            if (filaFinDatos >= filaInicioDatos)
            {
                ws.Range(
                        filaInicioDatos,
                        4,
                        filaFinDatos,
                        10)
                    .Style.NumberFormat.Format =
                    "₡#,##0.00;[Red]-₡#,##0.00";
            }

            // ================================
            // FILA DE TOTALES
            // ================================

            ws.Cell(fila, 1).Value =
                "TOTALES";

            ws.Range(fila, 1, fila, 3)
                .Merge();

            ws.Cell(fila, 4).Value =
                presupuestos.Sum(x =>
                    x.MontoIngresadoPlanificado ?? 0m);

            ws.Cell(fila, 5).Value =
                presupuestos.Sum(x =>
                    x.MontoGastoPlanificado ?? 0m);

            ws.Cell(fila, 6).Value =
                presupuestos.Sum(x =>
                    x.TotalIngresoReal ?? 0m);

            ws.Cell(fila, 7).Value =
                presupuestos.Sum(x =>
                    x.TotalGastoReal ?? 0m);

            ws.Cell(fila, 8).Value =
                presupuestos.Sum(x =>
                    x.DiferenciaIngreso ?? 0m);

            ws.Cell(fila, 9).Value =
                presupuestos.Sum(x =>
                    x.DiferenciaGasto ?? 0m);

            ws.Cell(fila, 10).Value =
                presupuestos.Sum(x =>
                    x.SaldoDisponible ?? 0m);

            ws.Range(fila, 1, fila, 15)
                .Style.Fill.BackgroundColor =
                XLColor.FromHtml("#CFE5E8");

            ws.Range(fila, 1, fila, 15)
                .Style.Font.Bold = true;

            ws.Range(fila, 4, fila, 10)
                .Style.NumberFormat.Format =
                "₡#,##0.00;[Red]-₡#,##0.00";

            fila += 2;

            // ================================
            // RESUMEN FINAL
            // ================================

            int totalPresupuestos =
                presupuestos.Count;

            int totalBorradores =
                presupuestos.Count(x =>
                    x.EstadoAprobacion == "Borrador");

            int totalAprobados =
                presupuestos.Count(x =>
                    x.EstadoAprobacion == "Aprobado");

            int totalRechazados =
                presupuestos.Count(x =>
                    x.EstadoAprobacion == "Rechazado");

            int totalAlertas =
                presupuestos.Count(x =>
                    (x.SaldoDisponible ?? 0m) < 0m ||

                    (x.TotalGastoReal ?? 0m) >
                    (x.MontoGastoPlanificado ?? 0m) ||

                    (x.TotalIngresoReal ?? 0m) <
                    (x.MontoIngresadoPlanificado ?? 0m));

            ws.Cell(fila, 1).Value =
                "RESUMEN DEL REPORTE";

            ws.Range(fila, 1, fila, 4)
                .Merge();

            ws.Cell(fila, 1)
                .Style.Font.Bold = true;

            ws.Cell(fila, 1)
                .Style.Fill.BackgroundColor =
                XLColor.FromHtml("#0F5C64");

            ws.Cell(fila, 1)
                .Style.Font.FontColor =
                XLColor.White;

            fila++;

            ws.Cell(fila, 1).Value =
                "Total de presupuestos";

            ws.Cell(fila, 2).Value =
                totalPresupuestos;

            fila++;

            ws.Cell(fila, 1).Value =
                "Borradores";

            ws.Cell(fila, 2).Value =
                totalBorradores;

            fila++;

            ws.Cell(fila, 1).Value =
                "Aprobados";

            ws.Cell(fila, 2).Value =
                totalAprobados;

            fila++;

            ws.Cell(fila, 1).Value =
                "Rechazados";

            ws.Cell(fila, 2).Value =
                totalRechazados;

            fila++;

            ws.Cell(fila, 1).Value =
                "Presupuestos en alerta";

            ws.Cell(fila, 2).Value =
                totalAlertas;

            // ================================
            // BORDES Y ALINEACIÓN
            // ================================

            int ultimaFilaUsada =
                ws.LastRowUsed()?.RowNumber() ?? fila;

            var rangoGeneral =
                ws.Range(1, 1, ultimaFilaUsada, 15);

            rangoGeneral.Style.Alignment.Vertical =
                XLAlignmentVerticalValues.Center;

            rangoGeneral.Style.Border.OutsideBorder =
                XLBorderStyleValues.Thin;

            rangoGeneral.Style.Border.InsideBorder =
                XLBorderStyleValues.Thin;

            // Ajuste de texto.
            ws.Range(1, 1, ultimaFilaUsada, 15)
                .Style.Alignment.WrapText = true;

            // Alineación centrada.
            ws.Range(
                    filaInicioDatos,
                    2,
                    Math.Max(filaFinDatos, filaInicioDatos),
                    3)
                .Style.Alignment.Horizontal =
                XLAlignmentHorizontalValues.Center;

            ws.Range(
                    filaInicioDatos,
                    11,
                    Math.Max(filaFinDatos, filaInicioDatos),
                    15)
                .Style.Alignment.Horizontal =
                XLAlignmentHorizontalValues.Center;

            // ================================
            // AUTOFILTRO
            // ================================

            int filaEncabezados =
                filaInicioDatos - 1;

            if (filaFinDatos >= filaInicioDatos)
            {
                ws.Range(
                        filaEncabezados,
                        1,
                        filaFinDatos,
                        15)
                    .SetAutoFilter();
            }

            // ================================
            // ANCHOS DE COLUMNAS
            // ================================

            ws.Column(1).Width = 26;
            ws.Column(2).Width = 10;
            ws.Column(3).Width = 13;
            ws.Column(4).Width = 18;
            ws.Column(5).Width = 18;
            ws.Column(6).Width = 18;
            ws.Column(7).Width = 18;
            ws.Column(8).Width = 16;
            ws.Column(9).Width = 16;
            ws.Column(10).Width = 18;
            ws.Column(11).Width = 18;
            ws.Column(12).Width = 18;
            ws.Column(13).Width = 35;
            ws.Column(14).Width = 30;
            ws.Column(15).Width = 14;

            // Congelar encabezados.
            ws.SheetView.FreezeRows(
                filaEncabezados);

            // Configuración de impresión.
            ws.PageSetup.PageOrientation =
                XLPageOrientation.Landscape;

            ws.PageSetup.PaperSize =
                XLPaperSize.A4Paper;

            ws.PageSetup.FitToPages(
                1,
                0);

            ws.PageSetup.Margins.Top = 0.25;
            ws.PageSetup.Margins.Bottom = 0.25;
            ws.PageSetup.Margins.Left = 0.20;
            ws.PageSetup.Margins.Right = 0.20;

            // ================================
            // DESCARGA
            // ================================

            using var stream =
                new MemoryStream();

            workbook.SaveAs(stream);

            stream.Position = 0;

            string nombreArchivo =
                $"Reporte_Presupuestos_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";
            await RegistrarAuditoria(
    "Exportar Excel",
    0,
    "Se exportó el reporte de presupuestos mensuales a Excel. " +
    ConstruirDescripcionFiltrosReporte(
        proyectoId,
        anio,
        mes,
        estado,
        estadoAprobacion)
);
            return File(
                stream.ToArray(),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                nombreArchivo);
        }
        // GET: PresupuestosMensuales/Details/5
        public async Task<IActionResult> Details(int? id)
        {

            if (id == null)
            {
                return NotFound();
            }
            var presupuestosMensuale = await _context.PresupuestosMensuales
                        .Include(p => p.Proyecto)
                        .Include(p => p.PresupuestoDetalles)
                        .ThenInclude(d => d.CategoriaIngreso)
                        .Include(p => p.PresupuestoDetalles)
                        .ThenInclude(d => d.CategoriaGasto)
                        .FirstOrDefaultAsync(m => m.IdPresupuestoMensual == id);

            if (presupuestosMensuale == null)
            {
                return NotFound();
            }
            await RecalcularPresupuesto(presupuestosMensuale);

            // Se guardan los cambios calculados.
            await _context.SaveChangesAsync();

            return View(presupuestosMensuale);
        }
        
        // ===========================================================
        // PRESUPUESTO COMPROMETIDO DEL PROYECTO
        // Borrador y Aprobado reservan presupuesto.
        // Rechazado no consume. escrito el 4/8/26
        // ===========================================================
        private async Task<decimal> CalcularGastoComprometidoProyecto(
            int proyectoId,
            int? presupuestoExcluirId = null)
        {
            var consulta = _context.PresupuestosMensuales
                .Where(p =>
                    p.ProyectoId == proyectoId &&
                    (
                        p.EstadoAprobacion == "Borrador" ||
                        p.EstadoAprobacion == "Aprobado"
                    ));

            if (presupuestoExcluirId.HasValue)
            {
                consulta = consulta.Where(p =>
                    p.IdPresupuestoMensual !=
                    presupuestoExcluirId.Value);
            }

            return await consulta.SumAsync(p =>
                p.MontoGastoPlanificado ?? 0m);
        }
        // ===========================================================
        // VALIDAR QUE EL MES PERTENEZCA A LA VIGENCIA DEL PROYECTO escrito el 4/8/26
        // ===========================================================
        private void ValidarVigenciaProyecto(
            Proyecto proyecto,
            int anio,
            int mes)
        {
            if (mes < 1 || mes > 12)
            {
                return;
            }

            if (anio < 1)
            {
                return;
            }

            DateOnly inicioPeriodo;

            try
            {
                inicioPeriodo =
                    new DateOnly(anio, mes, 1);
            }
            catch
            {
                return;
            }

            DateOnly finPeriodo =
                inicioPeriodo
                    .AddMonths(1)
                    .AddDays(-1);

            if (!proyecto.FechaInicio.HasValue)
            {
                ModelState.AddModelError(
                    "ProyectoId",
                    "El proyecto seleccionado no tiene una fecha de inicio válida.");

                return;
            }

            bool periodoAnteriorAlProyecto =
                finPeriodo < proyecto.FechaInicio.Value;

            bool periodoPosteriorAlProyecto =
                proyecto.FechaFin.HasValue &&
                inicioPeriodo > proyecto.FechaFin.Value;

            if (periodoAnteriorAlProyecto ||
                periodoPosteriorAlProyecto)
            {
                string vigencia =
                    proyecto.FechaFin.HasValue
                        ? $"del {proyecto.FechaInicio.Value:dd/MM/yyyy} " +
                          $"al {proyecto.FechaFin.Value:dd/MM/yyyy}"
                        : $"desde el {proyecto.FechaInicio.Value:dd/MM/yyyy}";

                string nombreMes =
                    new DateTime(anio, mes, 1)
                        .ToString(
                            "MMMM 'de' yyyy",
                            new CultureInfo("es-CR"));

                ModelState.AddModelError(
                    "Mes",
                    $"No puede crear un presupuesto para {nombreMes}, " +
                    $"porque el proyecto tiene vigencia {vigencia}.");
            }
        }
        // ===========================================================
        // INFORMACIÓN PRESUPUESTARIA DEL PROYECTO endpoint creado el 4/8/26
        // ===========================================================
        [HttpGet]
        public async Task<IActionResult> ObtenerResumenProyecto(
            int proyectoId)
        {
            var proyecto = await _context.Proyectos
                .AsNoTracking()
                .FirstOrDefaultAsync(p =>
                    p.IdProyecto == proyectoId &&
                    p.Activo);

            if (proyecto == null)
            {
                return NotFound(new
                {
                    mensaje =
                        "El proyecto no existe o se encuentra inactivo."
                });
            }

            decimal presupuestoGlobal =
                proyecto.PresupuestoAsignado ?? 0m;

            decimal comprometido =
                await CalcularGastoComprometidoProyecto(
                    proyectoId);

            decimal disponible =
                presupuestoGlobal - comprometido;

            int presupuestosCreados =
                await _context.PresupuestosMensuales
                    .CountAsync(p =>
                        p.ProyectoId == proyectoId &&
                        p.EstadoAprobacion != "Rechazado");

            int? totalPeriodos = null;
            int? periodosPendientes = null;
            decimal? distribucionOrientativa = null;

            if (proyecto.FechaInicio.HasValue &&
                proyecto.FechaFin.HasValue)
            {
                DateOnly inicio =
                    new DateOnly(
                        proyecto.FechaInicio.Value.Year,
                        proyecto.FechaInicio.Value.Month,
                        1);

                DateOnly fin =
                    new DateOnly(
                        proyecto.FechaFin.Value.Year,
                        proyecto.FechaFin.Value.Month,
                        1);

                totalPeriodos =
                    ((fin.Year - inicio.Year) * 12) +
                    fin.Month -
                    inicio.Month +
                    1;

                periodosPendientes =
                    Math.Max(
                        totalPeriodos.Value -
                        presupuestosCreados,
                        0);

                if (periodosPendientes.Value > 0 &&
                    disponible > 0)
                {
                    distribucionOrientativa =
                        disponible /
                        periodosPendientes.Value;
                }
            }

            return Json(new
            {
                proyecto.IdProyecto,
                proyecto.Nombre,

                fechaInicio =
                    proyecto.FechaInicio?.ToString(
                        "dd/MM/yyyy"),

                fechaFin =
                    proyecto.FechaFin?.ToString(
                        "dd/MM/yyyy"),

                presupuestoGlobal,
                comprometido,
                disponible,
                presupuestosCreados,
                totalPeriodos,
                periodosPendientes,
                distribucionOrientativa,

                presupuestoAgotado =
                    disponible <= 0
            });
        }
        // GET: PresupuestosMensuales/Create/6/8/26
        public async Task<IActionResult> Create(
            int? proyectoId,
            int? anio,
            int? mes,
            bool desdeProyectos = false)
        {
            string rol =
                HttpContext.Session.GetString("Rol") ?? "";

            if (!PuedeAccederCreate(
                    desdeProyectos,
                    proyectoId))
            {
                TempData["MensajeError"] =
                    "Debe registrar el presupuesto desde la vista de Proyectos.";

                return RedirectToAction(nameof(Index));
            }

            int anioSeleccionado =
                anio ?? DateTime.Today.Year;

            int mesSeleccionado =
                mes ?? DateTime.Today.Month;
            // ============================================================
            // BLOQUEO POR CIERRE CONTABLE
            // ============================================================
            bool periodoCerrado =
                await PeriodoContableCerradoAsync(
                    anioSeleccionado,
                    mesSeleccionado);

            if (periodoCerrado)
            {
                TempData["MensajeError"] =
                    $"No se puede registrar un presupuesto para el período " +
                    $"{mesSeleccionado:D2}/{anioSeleccionado} porque se encuentra cerrado.";

                return RedirectToAction(nameof(Index));
            }
            /*
             * Cuando se llega desde Proyectos, se valida en servidor
             * que el proyecto y el período realmente sean válidos.
             * Esto evita alterar manualmente la URL.
             */
            if (desdeProyectos && proyectoId.HasValue)
            {
                var proyectoSeleccionado =
                    await _context.Proyectos
                        .FirstOrDefaultAsync(p =>
                            p.IdProyecto == proyectoId.Value);

                if (proyectoSeleccionado == null)
                {
                    TempData["MensajeError"] =
                        "El proyecto seleccionado no existe.";

                    return RedirectToAction(
                        "Index",
                        "Proyectos");
                }

                if (!proyectoSeleccionado.Activo)
                {
                    TempData["MensajeError"] =
                        "No se puede registrar un presupuesto para un proyecto inactivo.";

                    return RedirectToAction(
                        "Index",
                        "Proyectos");
                }

                /*
                 * Reutiliza la misma validación de vigencia
                 * empleada en Create POST, Edit y Aprobar. 
                 */
                ValidarVigenciaProyecto(
                    proyectoSeleccionado,
                    anioSeleccionado,
                    mesSeleccionado);

                if (!ModelState.IsValid)
                {
                    string mensaje =
                        ModelState.Values
                            .SelectMany(v => v.Errors)
                            .Select(e => e.ErrorMessage)
                            .FirstOrDefault()
                        ?? "El período seleccionado no pertenece a la vigencia del proyecto.";

                    ModelState.Clear();

                    TempData["MensajeError"] =
                        mensaje;

                    return RedirectToAction(
                        "Index",
                        "Proyectos");
                }

                bool existePresupuesto =
                    await _context.PresupuestosMensuales
                        .AnyAsync(p =>
                            p.ProyectoId == proyectoId.Value &&
                            p.Anio == anioSeleccionado &&
                            p.Mes == mesSeleccionado);

                if (existePresupuesto)
                {
                    TempData["MensajeError"] =
                        "Ya existe un presupuesto registrado para este proyecto, año y mes.";

                    return RedirectToAction(
                        nameof(Index),
                        new
                        {
                            proyectoId = proyectoId.Value
                        });
                }
            }

            ViewBag.DesdeProyectos =
                desdeProyectos;

            var proyectosActivos =
                await _context.Proyectos
                    .Where(p => p.Activo)
                    .OrderBy(p => p.Nombre)
                    .ToListAsync();

            ViewData["ProyectoId"] =
                new SelectList(
                    proyectosActivos,
                    "IdProyecto",
                    "Nombre",
                    proyectoId);

            var presupuesto =
                new PresupuestosMensuale
                {
                    ProyectoId = proyectoId,
                    Anio = anioSeleccionado,
                    Mes = mesSeleccionado,
                    MontoIngresadoPlanificado = 0m,
                    MontoGastoPlanificado = 0m
                };

            return View(presupuesto);
        }
        // POST: PresupuestosMensuales/Crear
        [HttpPost]
        [ValidateAntiForgeryToken]
       public async Task<IActionResult> Create([Bind("ProyectoId,Anio,Mes," + "MontoIngresadoPlanificado," +"MontoGastoPlanificado")]PresupuestosMensuale presupuesto,bool desdeProyectos = false)
        {
            string rol = HttpContext.Session.GetString("Rol") ?? "";

            if (!PuedeAccederCreate(
         desdeProyectos,
         presupuesto.ProyectoId))
            {
                TempData["MensajeError"] =
                    "No tiene permisos para registrar el presupuesto desde esta ubicación.";

                return RedirectToAction(nameof(Index));
            }

            // Validar proyecto.
            if (presupuesto.ProyectoId == null)
            {
                ModelState.AddModelError(
                    "ProyectoId",
                    "Debe seleccionar un proyecto."
                );
            }
            else
            {
                bool proyectoActivo = await _context.Proyectos.AnyAsync(p =>
                    p.IdProyecto == presupuesto.ProyectoId &&
                    p.Activo == true);

                if (!proyectoActivo)
                {
                    ModelState.AddModelError(
                        "ProyectoId",
                        "El proyecto seleccionado no existe o está inactivo."
                    );
                }
            }

            // Validar año.
            int anioMinimo = 2000;
            int anioMaximo = DateTime.Today.Year + 10;

            if (presupuesto.Anio < anioMinimo ||
                presupuesto.Anio > anioMaximo)
            {
                ModelState.AddModelError(
                    "Anio",
                    $"El año debe estar entre {anioMinimo} y {anioMaximo}."
                );
            }

            // Validar mes.
            if (presupuesto.Mes < 1 || presupuesto.Mes > 12)
            {
                ModelState.AddModelError(
                    "Mes",
                    "El mes debe estar entre 1 y 12."
                );
            }

            // Validar ingreso planificado.
            if ((presupuesto.MontoIngresadoPlanificado ?? 0) < 0)
            {
                ModelState.AddModelError(
                    "MontoIngresadoPlanificado",
                    "El monto de ingresos planificados no puede ser negativo."
                );
            }

            // Validar gasto planificado.
            if ((presupuesto.MontoGastoPlanificado ?? 0) < 0)
            {
                ModelState.AddModelError(
                    "MontoGastoPlanificado",
                    "El monto de gastos planificados no puede ser negativo."
                );
            }

            // Evitar duplicados por proyecto, año y mes.
            bool existePresupuesto =
                await _context.PresupuestosMensuales.AnyAsync(p =>
                    p.ProyectoId == presupuesto.ProyectoId &&
                    p.Anio == presupuesto.Anio &&
                    p.Mes == presupuesto.Mes);

            if (existePresupuesto)
            {
                ModelState.AddModelError(
                    "",
                    "Ya existe un presupuesto registrado para este proyecto, año y mes."
                );
            }
            // ===========================================================
            // VALIDAR PROYECTO, VIGENCIA Y PRESUPUESTO GLOBAL
            // ===========================================================
            if (presupuesto.ProyectoId.HasValue)
            {
                var proyectoSeleccionado =
                    await _context.Proyectos
                        .FirstOrDefaultAsync(p =>
                            p.IdProyecto ==
                                presupuesto.ProyectoId.Value &&
                            p.Activo);

                if (proyectoSeleccionado == null)
                {
                    ModelState.AddModelError(
                        "ProyectoId",
                        "El proyecto seleccionado no existe o está inactivo.");
                }
                else
                {
                    // El mes debe encontrarse dentro de la
                    // vigencia del proyecto.
                    ValidarVigenciaProyecto(
                        proyectoSeleccionado,
                        presupuesto.Anio,
                        presupuesto.Mes);

                    decimal presupuestoGlobal =
                        proyectoSeleccionado
                            .PresupuestoAsignado ?? 0m;

                    decimal gastoComprometido =
                        await CalcularGastoComprometidoProyecto(
                            proyectoSeleccionado.IdProyecto);

                    decimal nuevoGasto =
                        presupuesto.MontoGastoPlanificado ?? 0m;

                    decimal disponible =
                        presupuestoGlobal -
                        gastoComprometido;

                    decimal totalLuegoDeCrear =
                        gastoComprometido +
                        nuevoGasto;

                    if (totalLuegoDeCrear >
                        presupuestoGlobal)
                    {
                        ModelState.AddModelError(
                            "MontoGastoPlanificado",
                            $"El gasto planificado supera el presupuesto " +
                            $"disponible del proyecto. " +
                            $"Presupuesto global: ₡{presupuestoGlobal:N2}. " +
                            $"Comprometido en presupuestos vigentes: " +
                            $"₡{gastoComprometido:N2}. " +
                            $"Disponible: ₡{disponible:N2}.");
                    }
                }
            }
            // Campos calculados por el sistema.
            ModelState.Remove("TotalIngresoReal");
            ModelState.Remove("TotalGastoReal");
            ModelState.Remove("DiferenciaIngreso");
            ModelState.Remove("DiferenciaGasto");
            ModelState.Remove("SaldoDisponible");
            ModelState.Remove("Estado");
            ModelState.Remove("EstadoAprobacion");
            ModelState.Remove("ObservacionRevision");
            ModelState.Remove("FechaCreacion");
            ModelState.Remove("Proyecto");
            ModelState.Remove("PresupuestoDetalles");

            // ============================================================
            // BLOQUEO POR CIERRE CONTABLE
            // ============================================================
            if (presupuesto.Anio > 0 &&
                presupuesto.Mes >= 1 &&
                presupuesto.Mes <= 12)
            {
                bool periodoCerrado =
                    await PeriodoContableCerradoAsync(
                        presupuesto.Anio,
                        presupuesto.Mes);

                if (periodoCerrado)
                {
                    ModelState.AddModelError(
                        "Mes",
                        $"No se puede registrar el presupuesto porque el período " +
                        $"{presupuesto.Mes:D2}/{presupuesto.Anio} se encuentra cerrado.");
                }
            }

            if (!ModelState.IsValid)
            {
                ViewData["ProyectoId"] = new SelectList(
                    _context.Proyectos.Where(p => p.Activo == true),
                    "IdProyecto",
                    "Nombre",
                    presupuesto.ProyectoId
                );
                ViewBag.DesdeProyectos =
    desdeProyectos;
                return View(presupuesto);
            }

            // Valores iniciales.
            presupuesto.FechaCreacion = DateTime.Now;
            presupuesto.EstadoAprobacion = "Borrador";
            presupuesto.ObservacionRevision = null;

            // Calcula ingresos reales, gastos reales, diferencias,
            // saldo disponible y estado de ejecución.
            await RecalcularPresupuesto(presupuesto);

            _context.Add(presupuesto);
            await _context.SaveChangesAsync();

            await RegistrarAuditoria(
                "Crear",
                presupuesto.IdPresupuestoMensual,
                "Se creó el presupuesto mensual del proyecto ID " +
                presupuesto.ProyectoId +
                " para el mes " +
                presupuesto.Mes +
                " del año " +
                presupuesto.Anio
            );

            TempData["MensajeExito"] =
                "El presupuesto mensual fue registrado correctamente.";

            return RedirectToAction(
                nameof(Index),
                new { proyectoId = presupuesto.ProyectoId }
            );
        }

        // Método para recalcular automáticamente los datos reales del presupuesto.
        private async Task RecalcularPresupuesto(PresupuestosMensuale presupuesto)
        {
            // Se obtiene la fecha inicial y final del mes presupuestado.
            var fechaInicio = new DateOnly(presupuesto.Anio, presupuesto.Mes, 1);
            var fechaFin = fechaInicio.AddMonths(1).AddDays(-1);

            // Se calculan los ingresos reales correspondientes al período presupuestario
            //. Registros nuevos: se relacionan directamente mediante PresupuestoMensualId.
            // Registros históricos:si todavía no tenían PresupuestoMensualId y tampoco
            // están marcados como pendientes de asignación presupuestaria,conservan la lógica anterior por proyecto + fecha.
            presupuesto.TotalIngresoReal =
                await _context.Ingresos
                    .Where(i =>
                        i.Activo == true &&
                        i.Estado != "Anulado" &&
                        (
                            i.PresupuestoMensualId ==
                                presupuesto.IdPresupuestoMensual

                            ||

                            (
                                i.PresupuestoMensualId == null &&
                                string.IsNullOrWhiteSpace(
                                    i.MotivoPendientePresupuestario) &&
                                i.ProyectoId ==
                                    presupuesto.ProyectoId &&
                                i.Fecha >= fechaInicio &&
                                i.Fecha <= fechaFin
                            )
                        ))
                    .SumAsync(i =>
                        i.MontoReal ?? 0);

            // Solo los gastos aprobados afectan el presupuesto real.
            // Los registros nuevos se relacionan directamente con el período
            // mediante PresupuestoMensualId.
            // Los registros históricos conservan la lógica anterior por fecha.
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

            // Se calcula la diferencia de ingresos.
            presupuesto.DiferenciaIngreso =
             (presupuesto.TotalIngresoReal ?? 0) -
             (presupuesto.MontoIngresadoPlanificado ?? 0);

            // Se calcula la diferencia de gastos.
            presupuesto.DiferenciaGasto =
                (presupuesto.MontoGastoPlanificado ?? 0) -
                (presupuesto.TotalGastoReal ?? 0);

            // Se calcula el saldo disponible.
            presupuesto.SaldoDisponible =
                (presupuesto.TotalIngresoReal ?? 0) -
                (presupuesto.TotalGastoReal ?? 0);

            // Se asigna el estado del presupuesto.
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
        }
        // Método para recalcular el presupuesto mensual usando detalles planificados
        // y movimientos reales de ingresos y gastos.
        private async Task RecalcularPresupuestoDesdeDetalles(PresupuestosMensuale presupuesto)
        {
            // Se suman los ingresos planificados desde el detalle de presupuesto.
            presupuesto.MontoIngresadoPlanificado = await _context.PresupuestoDetalles
                .Where(d =>
                    d.PresupuestoMensualId == presupuesto.IdPresupuestoMensual &&
                    d.Tipo == "Ingreso")
                .SumAsync(d => d.MontoPlanificado);

            // Se suman los gastos planificados desde el detalle de presupuesto.
            presupuesto.MontoGastoPlanificado = await _context.PresupuestoDetalles
                .Where(d =>
                    d.PresupuestoMensualId == presupuesto.IdPresupuestoMensual &&
                    d.Tipo == "Gasto")
                .SumAsync(d => d.MontoPlanificado);

            // Se recalculan los montos reales, diferencias, balance y estado.
            await RecalcularPresupuesto(presupuesto);
        }

        // Método para verificar si el presupuesto ya tiene ingresos o gastos reales.
        // Si existen movimientos reales, no se debe cambiar Proyecto, Año ni Mes.
        private async Task<bool> TieneMovimientosReales(
            PresupuestosMensuale presupuesto)
        {
            bool tieneIngresos =
                await TieneIngresosReales(presupuesto);

            bool tieneGastos =
                await TieneGastosAprobados(presupuesto);

            return tieneIngresos || tieneGastos;
        }
        private async Task<bool> TieneIngresosReales(
    PresupuestosMensuale presupuesto)
        {
            DateOnly fechaInicio =
                new DateOnly(
                    presupuesto.Anio,
                    presupuesto.Mes,
                    1);

            DateOnly fechaFin =
                fechaInicio
                    .AddMonths(1)
                    .AddDays(-1);

            return await _context.Ingresos.AnyAsync(i =>
    i.Activo == true &&
    i.Estado != "Anulado" &&
    (
        i.PresupuestoMensualId ==
            presupuesto.IdPresupuestoMensual

        ||

        (
            i.PresupuestoMensualId == null &&
            string.IsNullOrWhiteSpace(
                i.MotivoPendientePresupuestario) &&
            i.ProyectoId ==
                presupuesto.ProyectoId &&
            i.Fecha >= fechaInicio &&
            i.Fecha <= fechaFin
        )
    ));
        }

        private async Task<bool> TieneGastosAprobados(
            PresupuestosMensuale presupuesto)
        {
            DateOnly fechaInicio =
                new DateOnly(
                    presupuesto.Anio,
                    presupuesto.Mes,
                    1);

            DateOnly fechaFin =
                fechaInicio
                    .AddMonths(1)
                    .AddDays(-1);

            return await _context.Gastos.AnyAsync(g =>
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
          ));
        }
        // Método para recalcular el monto ejecutado de cada detalle del presupuesto mensual.
        private async Task RecalcularDetallePresupuestoMensual(
            PresupuestoDetalle detalle,
            PresupuestosMensuale presupuesto)
        {
            // Se obtiene la fecha inicial y final del mes presupuestado.
            var fechaInicio = new DateOnly(presupuesto.Anio, presupuesto.Mes, 1);
            var fechaFin = fechaInicio.AddMonths(1).AddDays(-1);

            // Si el detalle es de ingreso, se calcula desde el módulo de ingresos.
            if (detalle.Tipo == "Ingreso")
            {
                detalle.MontoEjecutado = await _context.Ingresos
                    .Where(i =>
                        i.ProyectoId == presupuesto.ProyectoId &&
                        i.CategoriaIngresoId == detalle.CategoriaIngresoId &&
                        i.Fecha >= fechaInicio &&
                        i.Fecha <= fechaFin &&
                        i.Estado != "Anulado")
                    .SumAsync(i => i.MontoReal ?? 0);
            }

            // Si el detalle es de gasto, se calcula desde el módulo de gastos.
            if(detalle.Tipo == "Gasto")
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
                            g.MontoTotal ?? 0);
            }

            // Se calcula la diferencia entre lo planificado y lo ejecutado.
            detalle.Diferencia =
                detalle.MontoPlanificado -
                (detalle.MontoEjecutado ?? 0);
        }

        private async Task<string> ConstruirDescripcionFiltrosReporte(
    int? proyectoId,
    int? anio,
    int? mes,
    string? estado,
    string? estadoAprobacion)
        {
            var filtros = new List<string>();

            if (proyectoId.HasValue && proyectoId.Value > 0)
            {
                string? nombreProyecto = await _context.Proyectos
                    .Where(p => p.IdProyecto == proyectoId.Value)
                    .Select(p => p.Nombre)
                    .FirstOrDefaultAsync();

                filtros.Add(
                    "Proyecto: " +
                    (nombreProyecto ?? proyectoId.Value.ToString()));
            }

            if (anio.HasValue)
            {
                filtros.Add("Año: " + anio.Value);
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

                filtros.Add("Mes: " + nombreMes);
            }

            if (!string.IsNullOrWhiteSpace(estado))
            {
                filtros.Add("Estado de ejecución: " + estado);
            }

            if (!string.IsNullOrWhiteSpace(estadoAprobacion))
            {
                filtros.Add(
                    "Estado de aprobación: " +
                    estadoAprobacion);
            }

            return filtros.Any()
                ? "Filtros aplicados: " + string.Join(" | ", filtros)
                : "Sin filtros aplicados.";
        }

        // GET: PresupuestosMensuales/Edit/5
        public async Task<IActionResult> Edit(int? id)
        {
            string rol =
                HttpContext.Session.GetString("Rol") ?? "";

            if (!PuedePrepararPresupuestos())
            {
                TempData["MensajeError"] =
                    "No tiene permisos para modificar presupuestos mensuales.";

                return RedirectToAction(nameof(Index));
            }

            if (id == null)
            {
                return NotFound();
            }

            var presupuesto =
                await _context.PresupuestosMensuales
                    .Include(p => p.Proyecto)
                    .FirstOrDefaultAsync(p =>
                        p.IdPresupuestoMensual == id);

            if (presupuesto == null)
            {
                return NotFound();
            }
            // ============================================================
            // BLOQUEO POR CIERRE CONTABLE
            // ============================================================
            bool periodoCerrado =
                await PeriodoContableCerradoAsync(
                    presupuesto.Anio,
                    presupuesto.Mes);

            if (periodoCerrado)
            {
                TempData["MensajeError"] =
                    $"No se puede modificar el presupuesto porque el período " +
                    $"{presupuesto.Mes:D2}/{presupuesto.Anio} se encuentra cerrado. " +
                    $"Para realizar correcciones, primero debe reabrirse el período.";

                return RedirectToAction(nameof(Index));
            }
            if (presupuesto.EstadoAprobacion == "Aprobado")
            {
                TempData["MensajeError"] =
                    "No es posible modificar un presupuesto aprobado.";

                return RedirectToAction(nameof(Index));
            }

            await RecalcularPresupuesto(presupuesto);
            await _context.SaveChangesAsync();

            bool tieneIngresos =
                await TieneIngresosReales(presupuesto);

            bool tieneGastos =
                await TieneGastosAprobados(presupuesto);

            /*
             * Si hay gastos aprobados se bloquea todo.
             * Si solo hay ingresos, se bloquea el período,
             * pero se permiten los montos.
             */
            bool bloqueadoTotal =
                tieneGastos;

            bool bloquearPeriodo =
                tieneIngresos || tieneGastos;

            ViewBag.TieneIngresosReales =
                tieneIngresos;

            ViewBag.TieneGastosAprobados =
                tieneGastos;

            ViewBag.BloqueadoTotal =
                bloqueadoTotal;

            ViewBag.BloquearPeriodo =
                bloquearPeriodo;

            ViewData["ProyectoId"] =
                new SelectList(
                    _context.Proyectos
                        .Where(p =>
                            p.Activo == true ||
                            p.IdProyecto ==
                                presupuesto.ProyectoId)
                        .OrderBy(p => p.Nombre),
                    "IdProyecto",
                    "Nombre",
                    presupuesto.ProyectoId);

            return View(presupuesto);
        }

        // POST: PresupuestosMensuales/Edit/5
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Edit(int id, [Bind("IdPresupuestoMensual,ProyectoId,Anio,Mes," + "MontoIngresadoPlanificado,MontoGastoPlanificado")]PresupuestosMensuale presupuesto)
        {
            string rol =
                HttpContext.Session.GetString("Rol") ?? "";
            if (!PuedePrepararPresupuestos())
            {
                TempData["MensajeError"] =
                    "No tiene permisos para modificar presupuestos mensuales.";

                return RedirectToAction(nameof(Index));
            }

            if (id != presupuesto.IdPresupuestoMensual)
            {
                return NotFound();
            }

            var presupuestoActual =
                await _context.PresupuestosMensuales
                    .Include(p => p.Proyecto)
                    .FirstOrDefaultAsync(p =>
                        p.IdPresupuestoMensual == id);

            if (presupuestoActual == null)
            {
                return NotFound();
            }
            // ============================================================
            // BLOQUEO DEL PERÍODO ORIGINAL POR CIERRE CONTABLE
            // ============================================================
            bool periodoOriginalCerrado =
                await PeriodoContableCerradoAsync(
                    presupuestoActual.Anio,
                    presupuestoActual.Mes);

            if (periodoOriginalCerrado)
            {
                TempData["MensajeError"] =
                    $"No se puede modificar el presupuesto porque el período " +
                    $"{presupuestoActual.Mes:D2}/{presupuestoActual.Anio} " +
                    $"se encuentra cerrado. " +
                    $"Para realizar correcciones, primero debe reabrirse el período.";

                return RedirectToAction(nameof(Index));
            }
            if (presupuestoActual.EstadoAprobacion ==
                "Aprobado")
            {
                TempData["MensajeError"] =
                    "No es posible modificar un presupuesto aprobado.";

                return RedirectToAction(nameof(Index));
            }

            bool tieneIngresos =
                await TieneIngresosReales(
                    presupuestoActual);

            bool tieneGastos =
                await TieneGastosAprobados(
                    presupuestoActual);

            bool bloqueadoTotal =
                tieneGastos;

            bool bloquearPeriodo =
                tieneIngresos || tieneGastos;

            /*
             * Si existen gastos aprobados no se permite
             * cambiar ningún dato.
             */
            if (bloqueadoTotal)
            {
                TempData["MensajeError"] =
                    "No es posible modificar el presupuesto porque posee gastos aprobados asociados.";

                return RedirectToAction(nameof(Index));
            }

            /*
             * Si existen ingresos, se conservan en servidor
             * el proyecto, año y mes originales.
             */
            if (bloquearPeriodo)
            {
                presupuesto.ProyectoId =
                    presupuestoActual.ProyectoId;

                presupuesto.Anio =
                    presupuestoActual.Anio;

                presupuesto.Mes =
                    presupuestoActual.Mes;

                ModelState.Remove(
                    nameof(presupuesto.ProyectoId));

                ModelState.Remove(
                    nameof(presupuesto.Anio));

                ModelState.Remove(
                    nameof(presupuesto.Mes));
            }
            // ============================================================
            // IMPEDIR MOVER EL PRESUPUESTO A UN PERÍODO CERRADO
            // Solo aplica cuando el período todavía puede modificarse.
            // ============================================================
            if (!bloquearPeriodo &&
                presupuesto.Anio > 0 &&
                presupuesto.Mes >= 1 &&
                presupuesto.Mes <= 12 &&
                (
                    presupuesto.Anio != presupuestoActual.Anio ||
                    presupuesto.Mes != presupuestoActual.Mes
                ))
            {
                bool nuevoPeriodoCerrado =
                    await PeriodoContableCerradoAsync(
                        presupuesto.Anio,
                        presupuesto.Mes);

                if (nuevoPeriodoCerrado)
                {
                    ModelState.AddModelError(
                        nameof(presupuesto.Mes),
                        $"No se puede cambiar el presupuesto al período " +
                        $"{presupuesto.Mes:D2}/{presupuesto.Anio} porque ese " +
                        $"período contable se encuentra cerrado.");
                }
            }
            if (!presupuesto.ProyectoId.HasValue)
            {
                ModelState.AddModelError(
                    nameof(presupuesto.ProyectoId),
                    "Debe seleccionar un proyecto.");
            }

            int anioMinimo = 2000;
            int anioMaximo =
                DateTime.Today.Year + 10;

            if (presupuesto.Anio < anioMinimo ||
                presupuesto.Anio > anioMaximo)
            {
                ModelState.AddModelError(
                    nameof(presupuesto.Anio),
                    $"El año debe estar entre {anioMinimo} y {anioMaximo}.");
            }

            if (presupuesto.Mes < 1 ||
                presupuesto.Mes > 12)
            {
                ModelState.AddModelError(
                    nameof(presupuesto.Mes),
                    "El mes debe estar entre 1 y 12.");
            }

            if ((presupuesto.MontoIngresadoPlanificado ?? 0m) < 0m)
            {
                ModelState.AddModelError(
                    nameof(
                        presupuesto.MontoIngresadoPlanificado),
                    "El monto de ingresos planificados no puede ser negativo.");
            }

            if ((presupuesto.MontoGastoPlanificado ?? 0m) < 0m)
            {
                ModelState.AddModelError(
                    nameof(
                        presupuesto.MontoGastoPlanificado),
                    "El monto de gastos planificados no puede ser negativo.");
            }

            var proyectoSeleccionado =
                presupuesto.ProyectoId.HasValue
                    ? await _context.Proyectos
                        .FirstOrDefaultAsync(p =>
                            p.IdProyecto ==
                                presupuesto.ProyectoId.Value &&
                            (
                                p.Activo ||
                                p.IdProyecto ==
                                    presupuestoActual.ProyectoId
                            ))
                    : null;

            if (presupuesto.ProyectoId.HasValue &&
                proyectoSeleccionado == null)
            {
                ModelState.AddModelError(
                    nameof(presupuesto.ProyectoId),
                    "El proyecto seleccionado no existe o está inactivo.");
            }

            if (proyectoSeleccionado != null)
            {
                ValidarVigenciaProyecto(
                    proyectoSeleccionado,
                    presupuesto.Anio,
                    presupuesto.Mes);

                decimal presupuestoGlobal =
                    proyectoSeleccionado
                        .PresupuestoAsignado ?? 0m;

                /*
                 * Excluye este presupuesto y suma únicamente
                 * los demás Borradores y Aprobados.
                 */
                decimal comprometidoOtros =
                    await CalcularGastoComprometidoProyecto(
                        proyectoSeleccionado.IdProyecto,
                        presupuesto.IdPresupuestoMensual);

                decimal gastoEditado =
                    presupuesto.MontoGastoPlanificado ?? 0m;

                decimal totalComprometido =
                    comprometidoOtros +
                    gastoEditado;

                decimal disponibleParaEsteMes =
                    presupuestoGlobal -
                    comprometidoOtros;

                if (totalComprometido >
                    presupuestoGlobal)
                {
                    ModelState.AddModelError(
                        nameof(
                            presupuesto.MontoGastoPlanificado),
                        $"El gasto planificado supera el presupuesto disponible del proyecto. " +
                        $"Presupuesto global: ₡{presupuestoGlobal:N2}. " +
                        $"Comprometido en otros meses: ₡{comprometidoOtros:N2}. " +
                        $"Disponible para este mes: ₡{disponibleParaEsteMes:N2}.");
                }
            }

            bool existePresupuesto =
                await _context.PresupuestosMensuales
                    .AnyAsync(p =>
                        p.IdPresupuestoMensual !=
                            presupuesto.IdPresupuestoMensual &&
                        p.ProyectoId ==
                            presupuesto.ProyectoId &&
                        p.Anio == presupuesto.Anio &&
                        p.Mes == presupuesto.Mes);

            if (existePresupuesto)
            {
                ModelState.AddModelError(
                    string.Empty,
                    "Ya existe otro presupuesto registrado para este proyecto, año y mes.");
            }

            ModelState.Remove("TotalIngresoReal");
            ModelState.Remove("TotalGastoReal");
            ModelState.Remove("DiferenciaIngreso");
            ModelState.Remove("DiferenciaGasto");
            ModelState.Remove("SaldoDisponible");
            ModelState.Remove("Estado");
            ModelState.Remove("EstadoAprobacion");
            ModelState.Remove("ObservacionRevision");
            ModelState.Remove("FechaCreacion");
            ModelState.Remove("Proyecto");
            ModelState.Remove("PresupuestoDetalles");

           

            if (!ModelState.IsValid)
            {
                ViewBag.TieneIngresosReales =
                    tieneIngresos;

                ViewBag.TieneGastosAprobados =
                    tieneGastos;

                ViewBag.BloqueadoTotal =
                    bloqueadoTotal;

                ViewBag.BloquearPeriodo =
                    bloquearPeriodo;

                ViewData["ProyectoId"] =
                    new SelectList(
                        _context.Proyectos
                            .Where(p =>
                                p.Activo ||
                                p.IdProyecto ==
                                    presupuestoActual.ProyectoId)
                            .OrderBy(p => p.Nombre),
                        "IdProyecto",
                        "Nombre",
                        presupuesto.ProyectoId);

                return View(presupuesto);
            }

            bool estabaRechazado =
                presupuestoActual.EstadoAprobacion ==
                "Rechazado";

            presupuestoActual
                .MontoIngresadoPlanificado =
                    presupuesto
                        .MontoIngresadoPlanificado;

            presupuestoActual
                .MontoGastoPlanificado =
                    presupuesto
                        .MontoGastoPlanificado;

            /*
             * Proyecto, año y mes solamente cambian
             * si no existen movimientos reales.
             */
            if (!bloquearPeriodo)
            {
                presupuestoActual.ProyectoId =
                    presupuesto.ProyectoId;

                presupuestoActual.Anio =
                    presupuesto.Anio;

                presupuestoActual.Mes =
                    presupuesto.Mes;
            }

            if (estabaRechazado)
            {
                presupuestoActual.EstadoAprobacion =
                    "Borrador";

                presupuestoActual.ObservacionRevision =
                    null;
            }

            await RecalcularPresupuesto(
                presupuestoActual);

            _context.Update(presupuestoActual);
            await _context.SaveChangesAsync();

            string detalleAuditoria =
                $"Se modificó el presupuesto mensual del proyecto ID " +
                $"{presupuestoActual.ProyectoId}, período " +
                $"{presupuestoActual.Mes}/{presupuestoActual.Anio}.";

            if (tieneIngresos)
            {
                detalleAuditoria +=
                    " El proyecto y el período se conservaron porque existían ingresos reales.";
            }

            if (estabaRechazado)
            {
                detalleAuditoria +=
                    " El presupuesto rechazado volvió al estado Borrador.";
            }

            await RegistrarAuditoria(
                "Modificar",
                presupuestoActual.IdPresupuestoMensual,
                detalleAuditoria);

            TempData["MensajeExito"] =
                estabaRechazado
                    ? "El presupuesto fue corregido y volvió al estado Borrador."
                    : "El presupuesto mensual fue modificado correctamente.";

            return RedirectToAction(
                nameof(Index),
                new
                {
                    proyectoId =
                        presupuestoActual.ProyectoId
                });
        }
        //modal con resumen presupuestario
        [HttpGet]
        public async Task<IActionResult> ObtenerResumenAprobacion(int id)
        {
            var presupuesto =
                await _context.PresupuestosMensuales
                    .Include(p => p.Proyecto)
                    .AsNoTracking()
                    .FirstOrDefaultAsync(p =>
                        p.IdPresupuestoMensual == id);

            if (presupuesto == null ||
                presupuesto.Proyecto == null)
            {
                return NotFound(new
                {
                    mensaje = "No fue posible encontrar el presupuesto."
                });
            }

            var proyecto = presupuesto.Proyecto;

            decimal presupuestoGlobal =
                proyecto.PresupuestoAsignado ?? 0m;

            decimal comprometidoOtros =
                await CalcularGastoComprometidoProyecto(
                    proyecto.IdProyecto,
                    presupuesto.IdPresupuestoMensual);

            decimal gastoPeriodo =
                presupuesto.MontoGastoPlanificado ?? 0m;

            decimal comprometidoDespues =
                comprometidoOtros + gastoPeriodo;

            decimal disponibleDespues =
                presupuestoGlobal - comprometidoDespues;

            int? totalPeriodos = null;
            int? periodosPresupuestados = null;
            int? periodosPendientes = null;
            decimal? distribucionOrientativa = null;

            if (proyecto.FechaInicio.HasValue &&
                proyecto.FechaFin.HasValue)
            {
                DateOnly inicioProyecto =
                    new DateOnly(
                        proyecto.FechaInicio.Value.Year,
                        proyecto.FechaInicio.Value.Month,
                        1);

                DateOnly finProyecto =
                    new DateOnly(
                        proyecto.FechaFin.Value.Year,
                        proyecto.FechaFin.Value.Month,
                        1);

                totalPeriodos =
                    ((finProyecto.Year - inicioProyecto.Year) * 12) +
                    finProyecto.Month -
                    inicioProyecto.Month +
                    1;

                /*
                 * Cuenta los períodos Borrador o Aprobados,
                 * incluyendo el presupuesto actual.
                 */
                periodosPresupuestados =
                    await _context.PresupuestosMensuales
                        .CountAsync(p =>
                            p.ProyectoId == proyecto.IdProyecto &&
                            (
                                p.EstadoAprobacion == "Borrador" ||
                                p.EstadoAprobacion == "Aprobado"
                            ));

                periodosPendientes =
                    Math.Max(
                        totalPeriodos.Value -
                        periodosPresupuestados.Value,
                        0);

                if (periodosPendientes.Value > 0 &&
                    disponibleDespues > 0)
                {
                    distribucionOrientativa =
                        disponibleDespues /
                        periodosPendientes.Value;
                }
            }

            string nombreMes =
                new DateTime(
                    presupuesto.Anio,
                    presupuesto.Mes,
                    1)
                .ToString(
                    "MMMM",
                    new CultureInfo("es-CR"));

            nombreMes =
                char.ToUpper(nombreMes[0]) +
                nombreMes.Substring(1);

            return Json(new
            {
                presupuesto.IdPresupuestoMensual,

                proyecto =
                    proyecto.Nombre,

                periodo =
                    $"{nombreMes} {presupuesto.Anio}",

                ingresoPlanificado =
                    presupuesto.MontoIngresadoPlanificado ?? 0m,

                gastoPlanificado =
                    gastoPeriodo,

                presupuestoGlobal,
                comprometidoAntes =
                    comprometidoOtros,

                comprometidoDespues,
                disponibleDespues,
                totalPeriodos,
                periodosPresupuestados,
                periodosPendientes,
                distribucionOrientativa,

                presupuestoEnCero =
                    (presupuesto.MontoIngresadoPlanificado ?? 0m) == 0m &&
                    gastoPeriodo == 0m
            });
        }
        // POST: PresupuestosMensuales/Aprobar/5
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Aprobar(int id)
        {

            string rol = HttpContext.Session.GetString("Rol") ?? "";

            if (!PuedeAprobarPresupuestos())
            {
                TempData["MensajeError"] =
                    "Solo el Administrador puede aprobar presupuestos.";

                return RedirectToAction(nameof(Index));
            }

            var presupuesto = await _context.PresupuestosMensuales
                .Include(p => p.Proyecto)
                .FirstOrDefaultAsync(p => p.IdPresupuestoMensual == id);


            if (presupuesto == null)
            {
                return NotFound();
            }
            // ============================================================
            // BLOQUEO POR CIERRE CONTABLE
            // ============================================================
            bool periodoCerrado =
                await PeriodoContableCerradoAsync(
                    presupuesto.Anio,
                    presupuesto.Mes);

            if (periodoCerrado)
            {
                TempData["MensajeError"] =
                    $"No se puede aprobar el presupuesto porque el período " +
                    $"{presupuesto.Mes:D2}/{presupuesto.Anio} se encuentra cerrado. " +
                    $"Para realizar cambios, primero debe reabrirse el período.";

                return RedirectToAction(nameof(Index));
            }
            if (presupuesto.EstadoAprobacion != "Borrador")
            {
                TempData["MensajeError"] =
                    "Solo los presupuestos en estado Borrador pueden aprobarse.";

                return RedirectToAction(nameof(Index));
            }
            if (presupuesto.Proyecto == null ||
    !presupuesto.Proyecto.Activo)
            {
                TempData["MensajeError"] =
                    "No puede aprobarse el presupuesto porque el proyecto está inactivo o no existe.";

                return RedirectToAction(nameof(Index));
            }

            ValidarVigenciaProyecto(
                presupuesto.Proyecto,
                presupuesto.Anio,
                presupuesto.Mes);

            if (!ModelState.IsValid)
            {
                string mensaje =
                    ModelState.Values
                        .SelectMany(v => v.Errors)
                        .Select(e => e.ErrorMessage)
                        .FirstOrDefault()
                    ?? "El período no pertenece a la vigencia del proyecto.";

                TempData["MensajeError"] =
                    mensaje;

                return RedirectToAction(nameof(Index));
            }

            decimal presupuestoGlobal =
                presupuesto.Proyecto
                    .PresupuestoAsignado ?? 0m;

            decimal comprometidoOtros =
                await CalcularGastoComprometidoProyecto(
                    presupuesto.ProyectoId!.Value,
                    presupuesto.IdPresupuestoMensual);

            decimal gastoActual =
                presupuesto.MontoGastoPlanificado ?? 0m;

            decimal totalComprometido =
                comprometidoOtros +
                gastoActual;

            if (totalComprometido >
                presupuestoGlobal)
            {
                TempData["MensajeError"] =
                    $"No puede aprobarse el presupuesto. " +
                    $"Presupuesto global: ₡{presupuestoGlobal:N2}. " +
                    $"Comprometido en otros períodos: ₡{comprometidoOtros:N2}. " +
                    $"Monto de este presupuesto: ₡{gastoActual:N2}.";

                return RedirectToAction(nameof(Index));
            }
            presupuesto.EstadoAprobacion = "Aprobado";
            presupuesto.ObservacionRevision = null;

            await RecalcularPresupuesto(presupuesto);

            _context.Update(presupuesto);
            await _context.SaveChangesAsync();

            await RegistrarAuditoria(
                "Aprobar",
                presupuesto.IdPresupuestoMensual,
                "Se aprobó el presupuesto mensual del proyecto " +
                (presupuesto.Proyecto?.Nombre ?? "sin nombre") +
                " para el mes " + presupuesto.Mes +
                " del año " + presupuesto.Anio + "."
            );

            TempData["MensajeExito"] =
                "El presupuesto mensual fue aprobado correctamente.";

            return RedirectToAction(nameof(Index));
        }
        // ===========================================================
        // RESUMEN DE MOVIMIENTOS ANTES DE RECHAZAR PRESUPUESTO
        // ===========================================================
        [HttpGet]
        public async Task<IActionResult> ObtenerResumenRechazo(int id)
        {
            string rol =
                HttpContext.Session.GetString("Rol") ?? "";

            if (rol != "Administrador")
            {
                return Forbid();
            }

            var presupuesto =
                await _context.PresupuestosMensuales
                    .Include(p => p.Proyecto)
                    .FirstOrDefaultAsync(p =>
                        p.IdPresupuestoMensual == id);

            if (presupuesto == null)
            {
                return NotFound(new
                {
                    mensaje = "El presupuesto no existe."
                });
            }

            var ingresosAsociados =
                await _context.Ingresos
                    .Where(i =>
                        i.PresupuestoMensualId ==
                            presupuesto.IdPresupuestoMensual &&
                        i.Activo)
                    .ToListAsync();

            int cantidadIngresos =
                ingresosAsociados.Count;

            decimal montoIngresos =
                ingresosAsociados.Sum(i =>
                    i.MontoReal ?? 0m);

            bool tieneGastosAprobados =
                await TieneGastosAprobados(presupuesto);

            return Json(new
            {
                id = presupuesto.IdPresupuestoMensual,

                proyecto =
                    presupuesto.Proyecto?.Nombre ??
                    "Sin proyecto",

                anio = presupuesto.Anio,
                mes = presupuesto.Mes,

                ingresoPlanificado =
                    presupuesto.MontoIngresadoPlanificado ?? 0m,

                gastoPlanificado =
                    presupuesto.MontoGastoPlanificado ?? 0m,

                cantidadIngresos,
                montoIngresos,

                tieneGastosAprobados
            });
        }
        // POST: PresupuestosMensuales/Rechazar/5
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Rechazar( int id, string? observacionRevision)
        {


            string rol = HttpContext.Session.GetString("Rol") ?? "";

            if (!PuedeAprobarPresupuestos())
            {
                TempData["MensajeError"] =
                    "Solo el Administrador puede rechazar presupuestos.";

                return RedirectToAction(nameof(Index));
            }

            var presupuesto = await _context.PresupuestosMensuales
                .Include(p => p.Proyecto)
                .FirstOrDefaultAsync(p => p.IdPresupuestoMensual == id);

            if (presupuesto == null)
            {
                return NotFound();
            }
            // ============================================================
            // BLOQUEO POR CIERRE CONTABLE
            // ============================================================
            bool periodoCerrado =
                await PeriodoContableCerradoAsync(
                    presupuesto.Anio,
                    presupuesto.Mes);

            if (periodoCerrado)
            {
                TempData["MensajeError"] =
                    $"No se puede rechazar el presupuesto porque el período " +
                    $"{presupuesto.Mes:D2}/{presupuesto.Anio} se encuentra cerrado. " +
                    $"Para realizar cambios, primero debe reabrirse el período.";

                return RedirectToAction(nameof(Index));
            }
            observacionRevision = observacionRevision?.Trim();

            if (string.IsNullOrWhiteSpace(observacionRevision))
            {
                TempData["MensajeError"] =
                    "Debe indicar la observación o motivo del rechazo.";

                return RedirectToAction(nameof(Index));
            }

            if (observacionRevision.Length > 500)
            {
                TempData["MensajeError"] =
                    "La observación del rechazo no puede superar los 500 caracteres.";

                return RedirectToAction(nameof(Index));
            }

            if (presupuesto.EstadoAprobacion != "Borrador")
            {
                TempData["MensajeError"] =
                    "Solo los presupuestos en estado Borrador pueden rechazarse.";

                return RedirectToAction(nameof(Index));
            }
            //valida si puede rechazarse
 bool tieneGastosAprobados =
    await TieneGastosAprobados(presupuesto);

            if (tieneGastosAprobados)
            {
                TempData["MensajeError"] =
                    "No puede rechazarse el presupuesto porque ya posee gastos aprobados asociados.";

                return RedirectToAction(nameof(Index));
            }
            //==== LIBERAR INGRESOS ASOCIADOS AL PERÍODO RECHAZADO

            var ingresosAsociados = await _context.Ingresos
                .Where(i =>
                    i.PresupuestoMensualId == presupuesto.IdPresupuestoMensual &&
                    i.Activo)
                .ToListAsync();
            foreach (var ingreso in ingresosAsociados)
            {
                // Mantiene la relación con el presupuesto rechazado
                ingreso.PresupuestoMensualId =
                    presupuesto.IdPresupuestoMensual;

                ingreso.MotivoPendientePresupuestario =
                    $"El período presupuestario {presupuesto.Mes}/{presupuesto.Anio} " +
                    $"fue rechazado. Motivo: {observacionRevision}";
            }

            presupuesto.EstadoAprobacion = "Rechazado";
            presupuesto.ObservacionRevision = observacionRevision;

           

            _context.Update(presupuesto);
            await _context.SaveChangesAsync();

            await RegistrarAuditoria(
         "Rechazar",
         presupuesto.IdPresupuestoMensual,
         "Se rechazó el presupuesto mensual del proyecto " +
         (presupuesto.Proyecto?.Nombre ?? "sin nombre") +
         " para el mes " + presupuesto.Mes +
         " del año " + presupuesto.Anio +
         ". Motivo: " + observacionRevision + "."
     );

            TempData["MensajeExito"] =
                "El presupuesto mensual fue rechazado correctamente.";

            return RedirectToAction(nameof(Index));
        }

        // GET: PresupuestosMensuales/Delete/5
        public IActionResult Delete(int? id)
        {
            TempData["MensajeError"] =
                "La eliminación de presupuestos no está disponible desde esta vista.";

            return RedirectToAction(nameof(Index));
        }
       

        // POST: PresupuestosMensuales/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> DeleteConfirmed(int id)
        {
            string rol = HttpContext.Session.GetString("Rol") ?? "";

            if (rol != "Administrador")
            {
                TempData["MensajeError"] =
                    "No tiene permisos para eliminar presupuestos.";

                return RedirectToAction(nameof(Index));
            }

            var presupuesto = await _context.PresupuestosMensuales
                .Include(p => p.Proyecto)
                .FirstOrDefaultAsync(p => p.IdPresupuestoMensual == id);

            if (presupuesto == null)
            {
                return NotFound();
            }
            // ============================================================
            // BLOQUEO POR CIERRE CONTABLE
            // PROTECCIÓN CONTRA POST DIRECTO
            // ============================================================
            bool periodoCerrado =
                await PeriodoContableCerradoAsync(
                    presupuesto.Anio,
                    presupuesto.Mes);

            if (periodoCerrado)
            {
                TempData["MensajeError"] =
                    $"No se puede eliminar el presupuesto porque el período " +
                    $"{presupuesto.Mes:D2}/{presupuesto.Anio} " +
                    $"se encuentra cerrado. " +
                    $"Para realizar cambios, primero debe reabrirse el período.";

                return RedirectToAction(nameof(Index));
            }
            if (presupuesto.EstadoAprobacion != "Borrador")
            {
                TempData["MensajeError"] =
                    "Solo puede eliminarse un presupuesto que permanezca en Borrador.";

                return RedirectToAction(nameof(Index));
            }

            bool tieneMovimientos = await TieneMovimientosReales(presupuesto);

            if (tieneMovimientos)
            {
                TempData["MensajeError"] =
                    "No es posible eliminar un presupuesto que tiene movimientos reales.";

                return RedirectToAction(nameof(Index));
            }

            bool tieneDetalles = await _context.PresupuestoDetalles
                .AnyAsync(d => d.PresupuestoMensualId == id);

            if (tieneDetalles)
            {
                TempData["MensajeError"] =
                    "No es posible eliminar el presupuesto porque tiene detalles registrados.";

                return RedirectToAction(nameof(Index));
            }

            _context.PresupuestosMensuales.Remove(presupuesto);
            await _context.SaveChangesAsync();

            await RegistrarAuditoria(
                "Eliminar",
                presupuesto.IdPresupuestoMensual,
                "Se eliminó el presupuesto mensual del proyecto " +
                (presupuesto.Proyecto?.Nombre ?? "sin nombre") +
                " para el mes " + presupuesto.Mes +
                " del año " + presupuesto.Anio + "."
            );

            TempData["MensajeExito"] =
                "El presupuesto mensual fue eliminado correctamente.";

            return RedirectToAction(nameof(Index));
        }
        //funciones por rol 
        private string ObtenerRolActual()
        {
            return HttpContext.Session.GetString("Rol") ?? "";
        }

        private bool PuedePrepararPresupuestos()
        {
            string rol = ObtenerRolActual();

            return rol == "Administrador" ||
                   rol == "Contador" ||
                   rol == "Auxiliar Contable";
        }

        private bool PuedeAprobarPresupuestos()
        {
            return ObtenerRolActual() == "Administrador";
        }
        private bool PuedeAccederCreate(
    bool desdeProyectos,
    int? proyectoId)
        {
            string rol =
                HttpContext.Session.GetString("Rol")
                ?? "";

            if (rol == "Administrador" ||
                rol == "Contador")
            {
                return true;
            }

            if (rol == "Auxiliar Contable")
            {
                return desdeProyectos &&
                       proyectoId.HasValue;
            }

            return false;
        }
        // ============================================================
        // VALIDAR SI UN PERÍODO CONTABLE ESTÁ CERRADO
        // ============================================================
        private async Task<bool> PeriodoContableCerradoAsync(
            int anio,
            int mes)
        {
            return await _context.CierresContables
                .AsNoTracking()
                .AnyAsync(c =>
                    c.Anio == anio &&
                    c.Mes == mes &&
                    c.Estado == "Cerrado");
        }
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
                    Tabla = "PresupuestosMensuales",
                    RegistroId = registroId,
                    Accion = accion,
                    Descripcion = descripcion,
                    Fecha = DateTime.Now
                };

                _context.Auditoria.Add(auditoria);
                await _context.SaveChangesAsync();
            }
        }


        private bool PresupuestosMensualeExists(int id)
        {
            return _context.PresupuestosMensuales.Any(e => e.IdPresupuestoMensual == id);
        }
    }
}
