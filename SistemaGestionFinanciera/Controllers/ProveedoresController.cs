using System;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Http;
using Microsoft.EntityFrameworkCore;
using SistemaGestionFinanciera.Data;
using SistemaGestionFinanciera.Models;
using ClosedXML.Excel;
using System.IO;
using System.Collections.Generic;

namespace SistemaGestionFinanciera.Controllers
{
    public class ProveedoresController : BaseController
    {
        private readonly SistemaFinancieroContext _context;

        public ProveedoresController(SistemaFinancieroContext context)
        {
            _context = context;
        }

        // INDEX - LISTADO DE PROVEEDORES
        // ================================================================
        public async Task<IActionResult> Index(
            string? buscar,
            string? estado,
            string? situacionFinanciera)
        {
            var proveedores =
                await ObtenerProveedoresConSituacionFinanciera(
                    buscar,
                    estado,
                    situacionFinanciera);

            CargarResumenFinancieroProveedores(
                proveedores);

            ViewBag.Buscar = buscar;
            ViewBag.Estado = estado;

            ViewBag.SituacionFinanciera =
                situacionFinanciera;

            return View(proveedores);
        }
        // ================================================================
        // REPORTE DE PROVEEDORES
        // ================================================================
        [HttpGet]
        public async Task<IActionResult> Reporte(
            string? buscar,
            string? estado,
            string? situacionFinanciera)
        {
            var proveedores =
                await ObtenerProveedoresConSituacionFinanciera(
                    buscar,
                    estado,
                    situacionFinanciera,
                    incluirContactosAdicionales: true);

            ViewBag.TotalRegistros =
                proveedores.Count;

            ViewBag.TotalActivos =
                proveedores.Count(p =>
                    p.Activo);

            ViewBag.TotalInactivos =
                proveedores.Count(p =>
                    !p.Activo);

            CargarResumenFinancieroProveedores(
                proveedores);

            ViewBag.Buscar = buscar;
            ViewBag.Estado = estado;

            ViewBag.SituacionFinanciera =
                situacionFinanciera;

            ViewBag.FechaGeneracion =
                DateTime.Now;

            await RegistrarAuditoria(
                "Vista previa PDF",
                "Proveedores",
                0,
                "Se generó la vista previa del reporte de proveedores.");

            return View(proveedores);
        }
        // ================================================================
        // EXPORTAR REPORTE DE PROVEEDORES A EXCEL
        // ================================================================
        [HttpGet]
        public async Task<IActionResult> ExportarExcel(
            string? buscar,
            string? estado,
            string? situacionFinanciera)
        {
            var proveedores =
                await ObtenerProveedoresConSituacionFinanciera(
                    buscar,
                    estado,
                    situacionFinanciera,
                    incluirContactosAdicionales: true);

            using var workbook = new XLWorkbook();

            var hoja = workbook.Worksheets.Add(
                "Proveedores");

            // ============================================================
            // COLUMNAS DEL REPORTE
            // ============================================================
            const int totalColumnas = 14;

            // ============================================================
            // COLORES
            // ============================================================
            var colorInstitucional =
                XLColor.FromHtml("#0F6B73");

            var colorInstitucionalOscuro =
                XLColor.FromHtml("#0B4F55");

            var colorFondoSuave =
                XLColor.FromHtml("#E7F1F2");

            var colorFondoFiltros =
                XLColor.FromHtml("#F3F7F8");

            var colorVerde =
                XLColor.FromHtml("#198754");

            var colorVerdeSuave =
                XLColor.FromHtml("#D1E7DD");

            var colorGris =
                XLColor.FromHtml("#6C757D");

            var colorGrisSuave =
                XLColor.FromHtml("#E2E3E5");

            var colorAzulSuave =
                XLColor.FromHtml("#D9EAF7");

            var colorDoradoSuave =
                XLColor.FromHtml("#FFF3CD");

            var colorRojoSuave =
                XLColor.FromHtml("#F8D7DA");

            var colorRojoTexto =
                XLColor.FromHtml("#842029");

            var colorAmarilloTexto =
                XLColor.FromHtml("#664D03");

            // ============================================================
            // TÍTULO
            // ============================================================
            hoja.Cell("A1").Value =
                "Unión Cantonal de Asociaciones de Desarrollo de Santa Ana";

            hoja.Range(
                1,
                1,
                1,
                totalColumnas).Merge();

            hoja.Cell("A2").Value =
                "Reporte de Proveedores";

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

            var rangoInstitucion = hoja.Range(
                1,
                1,
                1,
                totalColumnas);

            rangoInstitucion.Style.Font.Bold = true;
            rangoInstitucion.Style.Font.FontSize = 16;

            rangoInstitucion.Style.Font.FontColor =
                XLColor.FromHtml("#0F5C64");

            rangoInstitucion.Style.Alignment.Horizontal =
                XLAlignmentHorizontalValues.Center;

            rangoInstitucion.Style.Alignment.Vertical =
                XLAlignmentVerticalValues.Center;

            var rangoTitulo = hoja.Range(
                2,
                1,
                2,
                totalColumnas);

            rangoTitulo.Style.Font.Bold = true;
            rangoTitulo.Style.Font.FontSize = 13;
            rangoTitulo.Style.Font.FontColor = XLColor.Black;

            rangoTitulo.Style.Alignment.Horizontal =
                XLAlignmentHorizontalValues.Center;

            rangoTitulo.Style.Alignment.Vertical =
                XLAlignmentVerticalValues.Center;

            var rangoFecha = hoja.Range(
                3,
                1,
                3,
                totalColumnas);

            rangoFecha.Style.Alignment.Horizontal =
                XLAlignmentHorizontalValues.Center;

            rangoFecha.Style.Alignment.Vertical =
                XLAlignmentVerticalValues.Center;

            rangoFecha.Style.Font.FontColor =
                XLColor.FromHtml("#666666");

            rangoFecha.Style.Fill.BackgroundColor =
                XLColor.White;

            rangoFecha.Style.Font.Bold = false;

            hoja.Row(1).Height = 24;
            hoja.Row(2).Height = 22;
            hoja.Row(3).Height = 20;

            // ============================================================
            // FILTROS APLICADOS
            // ============================================================
            hoja.Cell("A5").Value =
                "Filtros aplicados";

            hoja.Range(
                5,
                1,
                5,
                totalColumnas).Merge();

            var tituloFiltros = hoja.Range(
                5,
                1,
                5,
                totalColumnas);

            tituloFiltros.Style.Fill.BackgroundColor =
                colorInstitucional;

            tituloFiltros.Style.Font.FontColor =
                XLColor.White;

            tituloFiltros.Style.Font.Bold = true;

            tituloFiltros.Style.Alignment.Horizontal =
                XLAlignmentHorizontalValues.Left;

            hoja.Cell("A6").Value = "Buscar:";

            hoja.Cell("B6").Value =
                string.IsNullOrWhiteSpace(buscar)
                    ? "Todos"
                    : buscar;

            hoja.Cell("C6").Value = "Estado:";

            hoja.Cell("D6").Value =
                string.IsNullOrWhiteSpace(estado)
                    ? "Todos"
                    : estado;

            hoja.Cell("E6").Value =
                "Situación financiera:";

            hoja.Cell("F6").Value =
                string.IsNullOrWhiteSpace(
                    situacionFinanciera)
                    ? "Todas"
                    : situacionFinanciera;

            hoja.Cell("A6").Style.Font.Bold = true;
            hoja.Cell("C6").Style.Font.Bold = true;
            hoja.Cell("E6").Style.Font.Bold = true;

            // Espacio restante de la fila de filtros
            hoja.Range("G6:N6").Merge();

            var rangoFiltros = hoja.Range(
                6,
                1,
                6,
                totalColumnas);

            rangoFiltros.Style.Fill.BackgroundColor =
                colorFondoFiltros;

            rangoFiltros.Style.Border.OutsideBorder =
                XLBorderStyleValues.Thin;

            rangoFiltros.Style.Border.InsideBorder =
                XLBorderStyleValues.Thin;

            rangoFiltros.Style.Border.OutsideBorderColor =
                colorInstitucional;

            rangoFiltros.Style.Border.InsideBorderColor =
                XLColor.LightGray;

            rangoFiltros.Style.Alignment.Vertical =
                XLAlignmentVerticalValues.Center;

            hoja.Row(6).Height = 22;

            // ============================================================
            // RESUMEN
            // ============================================================
            hoja.Cell("A8").Value = "Resumen";

            hoja.Range(
                8,
                1,
                8,
                totalColumnas).Merge();

            var tituloResumen = hoja.Range(
                8,
                1,
                8,
                totalColumnas);

            tituloResumen.Style.Fill.BackgroundColor =
                colorInstitucional;

            tituloResumen.Style.Font.FontColor =
                XLColor.White;

            tituloResumen.Style.Font.Bold = true;

            // Primera fila del resumen
            hoja.Cell("A9").Value =
                "Total proveedores";

            hoja.Cell("B9").Value =
                proveedores.Count;

            hoja.Cell("C9").Value =
                "Activos";

            hoja.Cell("D9").Value =
                proveedores.Count(p => p.Activo);

            hoja.Cell("E9").Value =
                "Inactivos";

            hoja.Cell("F9").Value =
                proveedores.Count(p => !p.Activo);

            hoja.Cell("G9").Value =
                "Proveedores con saldo";

            hoja.Cell("H9").Value =
                proveedores.Count(p =>
                    p.SaldoTotalPendiente > 0);

            hoja.Cell("I9").Value =
                "Proveedores vencidos";

            hoja.Cell("J9").Value =
                proveedores.Count(p =>
                    p.TotalCuentasVencidas > 0);

            // Segunda fila del resumen
            hoja.Cell("A10").Value =
                "CxP pendientes";

            hoja.Cell("B10").Value =
                proveedores.Sum(p =>
                    p.TotalCuentasPendientes);

            hoja.Cell("C10").Value =
                "CxP vencidas";

            hoja.Cell("D10").Value =
                proveedores.Sum(p =>
                    p.TotalCuentasVencidas);

            hoja.Cell("E10").Value =
                "Saldo pendiente";

            hoja.Cell("F10").Value =
                proveedores.Sum(p =>
                    p.SaldoTotalPendiente);

            hoja.Cell("G10").Value =
                "Saldo vencido";

            hoja.Cell("H10").Value =
                proveedores.Sum(p =>
                    p.SaldoTotalVencido);

            hoja.Range("I10:N10").Merge();

            // Formato general del resumen
            var rangoResumen =
                hoja.Range("A9:N10");

            rangoResumen.Style.Font.Bold = true;

            rangoResumen.Style.Border.OutsideBorder =
                XLBorderStyleValues.Thin;

            rangoResumen.Style.Border.InsideBorder =
                XLBorderStyleValues.Thin;

            rangoResumen.Style.Alignment.Vertical =
                XLAlignmentVerticalValues.Center;

            // Colores primera fila
            hoja.Range("A9:B9")
                .Style.Fill.BackgroundColor =
                colorAzulSuave;

            hoja.Range("C9:D9")
                .Style.Fill.BackgroundColor =
                colorVerdeSuave;

            hoja.Range("E9:F9")
                .Style.Fill.BackgroundColor =
                colorGrisSuave;

            hoja.Range("G9:H9")
                .Style.Fill.BackgroundColor =
                colorDoradoSuave;

            hoja.Range("I9:J9")
                .Style.Fill.BackgroundColor =
                colorRojoSuave;

            hoja.Range("K9:N9")
                .Style.Fill.BackgroundColor =
                colorFondoSuave;

            // Colores segunda fila
            hoja.Range("A10:B10")
                .Style.Fill.BackgroundColor =
                colorDoradoSuave;

            hoja.Range("C10:D10")
                .Style.Fill.BackgroundColor =
                colorRojoSuave;

            hoja.Range("E10:F10")
                .Style.Fill.BackgroundColor =
                colorDoradoSuave;

            hoja.Range("G10:H10")
                .Style.Fill.BackgroundColor =
                colorRojoSuave;

            hoja.Range("I10:N10")
                .Style.Fill.BackgroundColor =
                colorFondoSuave;

            // Formato monetario
            hoja.Cell("F10")
                .Style.NumberFormat.Format =
                "₡#,##0.00";

            hoja.Cell("H10")
                .Style.NumberFormat.Format =
                "₡#,##0.00";

            hoja.Row(9).Height = 22;
            hoja.Row(10).Height = 22;

            // ============================================================
            // ENCABEZADOS DE LA TABLA
            // ============================================================
            const int filaEncabezado = 12;

            string[] encabezados =
            {
        "Nombre",
        "Identificación",
        "Teléfono principal",
        "Correo principal",
        "Dirección principal",
        "Teléfonos adicionales",
        "Correos adicionales",
        "Direcciones adicionales",
        "Estado",
        "Situación financiera",
        "CxP pendientes",
        "CxP vencidas",
        "Saldo pendiente",
        "Saldo vencido"
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

            var encabezadoTabla = hoja.Range(
                filaEncabezado,
                1,
                filaEncabezado,
                totalColumnas);

            encabezadoTabla.Style.Fill.BackgroundColor =
                colorInstitucional;

            encabezadoTabla.Style.Font.FontColor =
                XLColor.White;

            encabezadoTabla.Style.Font.Bold = true;

            encabezadoTabla.Style.Alignment.Horizontal =
                XLAlignmentHorizontalValues.Center;

            encabezadoTabla.Style.Alignment.Vertical =
                XLAlignmentVerticalValues.Center;

            encabezadoTabla.Style.Alignment.WrapText = true;

            encabezadoTabla.Style.Border.OutsideBorder =
                XLBorderStyleValues.Thin;

            encabezadoTabla.Style.Border.InsideBorder =
                XLBorderStyleValues.Thin;

            encabezadoTabla.Style.Border.OutsideBorderColor =
                colorInstitucionalOscuro;

            encabezadoTabla.Style.Border.InsideBorderColor =
                colorInstitucionalOscuro;

            hoja.Row(filaEncabezado).Height = 30;

            // ============================================================
            // DATOS
            // ============================================================
            int fila = filaEncabezado + 1;

            foreach (var proveedor in proveedores)
            {
                hoja.Cell(fila, 1).Value =
                    proveedor.Nombre;

                hoja.Cell(fila, 2).Value =
                    proveedor.Identificacion;

                hoja.Cell(fila, 3).Value =
                    string.IsNullOrWhiteSpace(
                        proveedor.Telefono)
                        ? "No registrado"
                        : proveedor.Telefono;

                hoja.Cell(fila, 4).Value =
                    string.IsNullOrWhiteSpace(
                        proveedor.Email)
                        ? "No registrado"
                        : proveedor.Email;
                hoja.Cell(fila, 5).Value =
    string.IsNullOrWhiteSpace(
        proveedor.Direccion)
        ? "No registrada"
        : proveedor.Direccion;

                // ========================================================
                // CONTACTOS ADICIONALES
                // ========================================================
                var telefonosAdicionales =
                    proveedor.ContactosAdicionales
                        .Where(c =>
                            c.Activo &&
                            c.TipoContacto == "Telefono")
                        .OrderBy(c =>
                            c.IdContactoProveedor)
                        .Select(c =>
                            string.IsNullOrWhiteSpace(
                                c.Descripcion)
                                ? c.Valor
                                : $"{c.Valor} ({c.Descripcion})")
                        .ToList();

                var correosAdicionales =
                    proveedor.ContactosAdicionales
                        .Where(c =>
                            c.Activo &&
                            c.TipoContacto == "Correo")
                        .OrderBy(c =>
                            c.IdContactoProveedor)
                        .Select(c =>
                            string.IsNullOrWhiteSpace(
                                c.Descripcion)
                                ? c.Valor
                                : $"{c.Valor} ({c.Descripcion})")
                        .ToList();

                var direccionesAdicionales =
                    proveedor.ContactosAdicionales
                        .Where(c =>
                            c.Activo &&
                            c.TipoContacto == "Direccion")
                        .OrderBy(c =>
                            c.IdContactoProveedor)
                        .Select(c =>
                            string.IsNullOrWhiteSpace(
                                c.Descripcion)
                                ? c.Valor
                                : $"{c.Valor} ({c.Descripcion})")
                        .ToList();

                hoja.Cell(fila, 6).Value =
     telefonosAdicionales.Any()
         ? string.Join(
             Environment.NewLine,
             telefonosAdicionales)
         : "No registrado";

                hoja.Cell(fila, 7).Value =
                    correosAdicionales.Any()
                        ? string.Join(
                            Environment.NewLine,
                            correosAdicionales)
                        : "No registrado";

                hoja.Cell(fila, 8).Value =
                    direccionesAdicionales.Any()
                        ? string.Join(
                            Environment.NewLine,
                            direccionesAdicionales)
                        : "No registrado";

                // ========================================================
                // ESTADO Y SITUACIÓN FINANCIERA
                // ========================================================
                hoja.Cell(fila, 9).Value =
       proveedor.Activo
           ? "Activo"
           : "Inactivo";

                hoja.Cell(fila, 10).Value =
                    proveedor.SituacionFinanciera;

                hoja.Cell(fila, 11).Value =
                    proveedor.TotalCuentasPendientes;

                hoja.Cell(fila, 12).Value =
                    proveedor.TotalCuentasVencidas;

                hoja.Cell(fila, 13).Value =
                    proveedor.SaldoTotalPendiente;

                hoja.Cell(fila, 14).Value =
                    proveedor.SaldoTotalVencido;

                // Teléfonos tratados como texto
                hoja.Cell(fila, 3)
                    .Style.NumberFormat.Format = "@";

                hoja.Cell(fila, 6)
                    .Style.NumberFormat.Format = "@";

                // Formato monetario
                hoja.Cell(fila, 13)
                    .Style.NumberFormat.Format =
                    "₡#,##0.00";

                hoja.Cell(fila, 14)
                    .Style.NumberFormat.Format =
                    "₡#,##0.00";

                // Filas alternas
                if (fila % 2 == 0)
                {
                    hoja.Range(
                            fila,
                            1,
                            fila,
                            totalColumnas)
                        .Style.Fill.BackgroundColor =
                        XLColor.FromHtml("#F7F9FA");
                }

                // ========================================================
                // COLOR DEL ESTADO ADMINISTRATIVO
                // ========================================================
                if (proveedor.Activo)
                {
                    hoja.Cell(fila, 9)
                        .Style.Fill.BackgroundColor =
                        colorVerdeSuave;

                    hoja.Cell(fila, 9)
                        .Style.Font.FontColor =
                        colorVerde;
                }
                else
                {
                    hoja.Cell(fila, 9)
                        .Style.Fill.BackgroundColor =
                        colorGrisSuave;

                    hoja.Cell(fila, 9)
                        .Style.Font.FontColor =
                        colorGris;
                }

                hoja.Cell(fila, 9)
                    .Style.Font.Bold = true;

                // ========================================================
                // COLOR DE SITUACIÓN FINANCIERA
                // ========================================================
                switch (proveedor.SituacionFinanciera)
                {
                    case "Pago vencido":

                        hoja.Cell(fila, 10)
                            .Style.Fill.BackgroundColor =
                            colorRojoSuave;

                        hoja.Cell(fila, 10)
                            .Style.Font.FontColor =
                            colorRojoTexto;

                        break;

                    case "Pago pendiente":

                        hoja.Cell(fila, 10)
                            .Style.Fill.BackgroundColor =
                            colorDoradoSuave;

                        hoja.Cell(fila, 10)
                            .Style.Font.FontColor =
                            colorAmarilloTexto;

                        break;

                    case "Al día":

                        hoja.Cell(fila, 10)
                            .Style.Fill.BackgroundColor =
                            colorVerdeSuave;

                        hoja.Cell(fila, 10)
                            .Style.Font.FontColor =
                            colorVerde;

                        break;

                    default:

                        hoja.Cell(fila, 10)
                            .Style.Fill.BackgroundColor =
                            colorGrisSuave;

                        hoja.Cell(fila, 10)
                            .Style.Font.FontColor =
                            colorGris;

                        break;
                }

                hoja.Cell(fila, 10)
                    .Style.Font.Bold = true;

                // Resaltar saldo vencido
                if (proveedor.SaldoTotalVencido > 0)
                {
                    hoja.Cell(fila, 14)
                        .Style.Font.FontColor =
                        XLColor.FromHtml("#DC3545");

                    hoja.Cell(fila, 14)
                        .Style.Font.Bold = true;
                }

                fila++;
            }

            // ============================================================
            // TABLA Y FILTROS POR COLUMNA
            // ============================================================
            if (proveedores.Any())
            {
                int filaFinDatos = fila - 1;

                var rangoTabla = hoja.Range(
                    filaEncabezado,
                    1,
                    filaFinDatos,
                    totalColumnas);

                var tabla = rangoTabla.CreateTable(
                    "TablaProveedores");

                tabla.Theme = XLTableTheme.None;
                tabla.ShowAutoFilter = true;
                tabla.ShowRowStripes = false;

                var rangoDatos = hoja.Range(
                    filaEncabezado + 1,
                    1,
                    filaFinDatos,
                    totalColumnas);

                rangoDatos.Style.Alignment.Vertical =
                    XLAlignmentVerticalValues.Center;

                rangoDatos.Style.Border.BottomBorder =
                    XLBorderStyleValues.Thin;

                rangoDatos.Style.Border.BottomBorderColor =
                    XLColor.LightGray;

                // Centrar identificación y teléfono
                hoja.Range(
                        filaEncabezado + 1,
                        2,
                        filaFinDatos,
                        3)
                    .Style.Alignment.Horizontal =
                    XLAlignmentHorizontalValues.Center;

                // Centrar estado, situación y cantidades
                hoja.Range(
                        filaEncabezado + 1,
                        9,
                        filaFinDatos,
                        12)
                    .Style.Alignment.Horizontal =
                    XLAlignmentHorizontalValues.Center;

                // Contactos adicionales con salto de línea
                hoja.Range(
                        filaEncabezado + 1,
                        5,
                        filaFinDatos,
                        8)
                    .Style.Alignment.WrapText = true;
            }
            else
            {
                hoja.Cell(fila, 1).Value =
                    "No se encontraron proveedores con los filtros aplicados.";

                hoja.Range(
                    fila,
                    1,
                    fila,
                    totalColumnas).Merge();

                hoja.Range(
                        fila,
                        1,
                        fila,
                        totalColumnas)
                    .Style.Alignment.Horizontal =
                    XLAlignmentHorizontalValues.Center;

                hoja.Range(
                        fila,
                        1,
                        fila,
                        totalColumnas)
                    .Style.Font.Italic = true;

                hoja.Range(
                        fila,
                        1,
                        fila,
                        totalColumnas)
                    .Style.Fill.BackgroundColor =
                    colorFondoSuave;

                fila++;
            }

            // ============================================================
            // TOTAL FINAL
            // ============================================================
            int filaTotal = fila + 1;

            hoja.Cell(filaTotal, 13).Value =
                "Total de registros:";

            hoja.Cell(filaTotal, 14).Value =
                proveedores.Count;

            var rangoTotal = hoja.Range(
                filaTotal,
                1,
                filaTotal,
                totalColumnas);

            rangoTotal.Style.Fill.BackgroundColor =
                colorFondoSuave;

            rangoTotal.Style.Border.TopBorder =
                XLBorderStyleValues.Medium;

            rangoTotal.Style.Border.TopBorderColor =
                colorInstitucional;

            hoja.Cell(filaTotal, 13)
                .Style.Font.Bold = true;

            hoja.Cell(filaTotal, 14)
                .Style.Font.Bold = true;

            hoja.Cell(filaTotal, 13)
                .Style.Font.FontColor =
                colorInstitucionalOscuro;

            hoja.Cell(filaTotal, 14)
                .Style.Font.FontColor =
                colorInstitucionalOscuro;

            hoja.Cell(filaTotal, 14)
                .Style.Alignment.Horizontal =
                XLAlignmentHorizontalValues.Center;

            // ============================================================
            // CONFIGURACIÓN FINAL
            // ============================================================
            hoja.Column(1).Width = 32;  // Nombre
            hoja.Column(2).Width = 20;  // Identificación
            hoja.Column(3).Width = 20;  // Teléfono principal
            hoja.Column(4).Width = 36;  // Correo principal
            hoja.Column(5).Width = 45;  // Dirección principal
            hoja.Column(6).Width = 30;  // Teléfonos adicionales
            hoja.Column(7).Width = 38;  // Correos adicionales
            hoja.Column(8).Width = 45;  // Direcciones adicionales
            hoja.Column(9).Width = 14;  // Estado
            hoja.Column(10).Width = 22; // Situación financiera
            hoja.Column(11).Width = 18; // CxP pendientes
            hoja.Column(12).Width = 18; // CxP vencidas
            hoja.Column(13).Width = 20; // Saldo pendiente
            hoja.Column(14).Width = 20; // Saldo vencido;

            hoja.Column(1)
        .Style.Alignment.WrapText = true;

            hoja.Column(4)
                .Style.Alignment.WrapText = true;

            hoja.Column(5)
                .Style.Alignment.WrapText = true;

            hoja.Column(6)
                .Style.Alignment.WrapText = true;

            hoja.Column(7)
                .Style.Alignment.WrapText = true;

            hoja.Column(8)
                .Style.Alignment.WrapText = true;

            hoja.Column(10)
                .Style.Alignment.WrapText = true;

            hoja.SheetView.FreezeRows(
                filaEncabezado);

            hoja.PageSetup.PageOrientation =
                XLPageOrientation.Landscape;

            hoja.PageSetup.PaperSize =
                XLPaperSize.A4Paper;

            hoja.PageSetup.FitToPages(1, 0);

            hoja.PageSetup.CenterHorizontally = true;

            hoja.PageSetup.SetRowsToRepeatAtTop(
                filaEncabezado,
                filaEncabezado);

            // ============================================================
            // GENERAR ARCHIVO
            // ============================================================
            using var stream = new MemoryStream();

            workbook.SaveAs(stream);

            string nombreArchivo =
                $"Reporte_Proveedores_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";

            await RegistrarAuditoria(
                "Exportar Excel",
                "Proveedores",
                0,
                "Se exportó el reporte de proveedores a Excel.");

            return File(
                stream.ToArray(),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                nombreArchivo);
        }
        // ================================================================
        // DETAILS
        // ================================================================
        public async Task<IActionResult> Details(int? id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var proveedor = await _context.Proveedores
     .Include(p => p.ContactosAdicionales)
     .Include(p => p.CuentasPorPagars)
     .Include(p => p.Facturas)
     .Include(p => p.Gastos)
     .FirstOrDefaultAsync(p => p.IdProveedor == id);

            if (proveedor == null)
            {
                return NotFound();
            }

            ViewBag.TieneRegistros = await TieneRegistrosAsociados(proveedor.IdProveedor);
            ViewBag.PuedeDesactivar = await PuedeDesactivar(proveedor.IdProveedor);

            return View(proveedor);
        }

        // ================================================================
        // CREATE GET
        // ================================================================
        public IActionResult Create()
        {
            var proveedor = new Proveedore
            {
                Activo = true
            };

            return View(proveedor);
        }

        // ================================================================
        // CREATE POST
        // ================================================================
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Create(Proveedore proveedor)
        {
            proveedor.Activo = true;

            NormalizarProveedor(proveedor);
            ValidarContactosAdicionales(proveedor);

            await ValidarDuplicados(proveedor);

            if (!ModelState.IsValid)
            {
                TempData["Error"] = "No fue posible registrar el proveedor. Revise los datos ingresados.";
                return View(proveedor);
            }

            _context.Proveedores.Add(proveedor);
            await _context.SaveChangesAsync();
            // ================================================================
            // GUARDAR CONTACTOS ADICIONALES OPCIONALES
            // ================================================================

            if (!string.IsNullOrWhiteSpace(proveedor.OtroTelefono))
            {
                _context.ContactosProveedores.Add(
                    new ContactoProveedor
                    {
                        IdProveedor = proveedor.IdProveedor,
                        TipoContacto = "Telefono",
                        Valor = proveedor.OtroTelefono.Trim(),
                        Descripcion = string.IsNullOrWhiteSpace(
                            proveedor.DescripcionOtroTelefono)
                                ? "Adicional"
                                : proveedor.DescripcionOtroTelefono.Trim(),
                        Activo = true
                    });
            }

            if (!string.IsNullOrWhiteSpace(proveedor.OtroTelefono2))
            {
                _context.ContactosProveedores.Add(
                    new ContactoProveedor
                    {
                        IdProveedor = proveedor.IdProveedor,
                        TipoContacto = "Telefono",
                        Valor = proveedor.OtroTelefono2.Trim(),
                        Descripcion = string.IsNullOrWhiteSpace(
                            proveedor.DescripcionOtroTelefono2)
                                ? "Adicional 2"
                                : proveedor.DescripcionOtroTelefono2.Trim(),
                        Activo = true
                    });
            }

            if (!string.IsNullOrWhiteSpace(proveedor.OtroCorreo))
            {
                _context.ContactosProveedores.Add(
                    new ContactoProveedor
                    {
                        IdProveedor = proveedor.IdProveedor,
                        TipoContacto = "Correo",
                        Valor = proveedor.OtroCorreo.Trim(),
                        Descripcion = string.IsNullOrWhiteSpace(
                            proveedor.DescripcionOtroCorreo)
                                ? "Adicional"
                                : proveedor.DescripcionOtroCorreo.Trim(),
                        Activo = true
                    });
            }

            if (!string.IsNullOrWhiteSpace(proveedor.OtroCorreo2))
            {
                _context.ContactosProveedores.Add(
                    new ContactoProveedor
                    {
                        IdProveedor = proveedor.IdProveedor,
                        TipoContacto = "Correo",
                        Valor = proveedor.OtroCorreo2.Trim(),
                        Descripcion = string.IsNullOrWhiteSpace(
                            proveedor.DescripcionOtroCorreo2)
                                ? "Adicional 2"
                                : proveedor.DescripcionOtroCorreo2.Trim(),
                        Activo = true
                    });
            }

            if (!string.IsNullOrWhiteSpace(proveedor.OtraDireccion))
            {
                _context.ContactosProveedores.Add(
                    new ContactoProveedor
                    {
                        IdProveedor = proveedor.IdProveedor,
                        TipoContacto = "Direccion",
                        Valor = proveedor.OtraDireccion.Trim(),
                        Descripcion = string.IsNullOrWhiteSpace(
                            proveedor.DescripcionOtraDireccion)
                                ? "Adicional"
                                : proveedor.DescripcionOtraDireccion.Trim(),
                        Activo = true
                    });
            }

            if (!string.IsNullOrWhiteSpace(proveedor.OtraDireccion2))
            {
                _context.ContactosProveedores.Add(
                    new ContactoProveedor
                    {
                        IdProveedor = proveedor.IdProveedor,
                        TipoContacto = "Direccion",
                        Valor = proveedor.OtraDireccion2.Trim(),
                        Descripcion = string.IsNullOrWhiteSpace(
                            proveedor.DescripcionOtraDireccion2)
                                ? "Adicional 2"
                                : proveedor.DescripcionOtraDireccion2.Trim(),
                        Activo = true
                    });
            }

            await _context.SaveChangesAsync();
            await RegistrarAuditoria(
                "Crear",
                "Proveedores",
                proveedor.IdProveedor,
                $"Se registró el proveedor: {proveedor.Nombre}, identificación: {proveedor.Identificacion}.");

            TempData["Exito"] = "Proveedor registrado correctamente.";
            return RedirectToAction(nameof(Index));
        }

        // ================================================================
        // EDIT GET
        // Carga datos principales y contactos adicionales.
        // ================================================================
        public async Task<IActionResult> Edit(int? id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var proveedor = await _context.Proveedores
                .Include(p => p.ContactosAdicionales)
                .Include(p => p.CuentasPorPagars)
                .Include(p => p.Facturas)
                .Include(p => p.Gastos)
                .FirstOrDefaultAsync(p => p.IdProveedor == id);

            if (proveedor == null)
            {
                return NotFound();
            }

            if (!proveedor.Activo)
            {
                TempData["Error"] =
                    "No se puede modificar un proveedor inactivo. Primero debe reactivarlo.";

                return RedirectToAction(nameof(Index));
            }

            // Teléfonos adicionales activos.
            var telefonos = proveedor.ContactosAdicionales
                .Where(c =>
                    c.Activo &&
                    c.TipoContacto == "Telefono")
                .OrderBy(c => c.IdContactoProveedor)
                .Take(2)
                .ToList();

            if (telefonos.Count > 0)
            {
                proveedor.OtroTelefono =
                    telefonos[0].Valor;

                proveedor.DescripcionOtroTelefono =
                    telefonos[0].Descripcion;
            }

            if (telefonos.Count > 1)
            {
                proveedor.OtroTelefono2 =
                    telefonos[1].Valor;

                proveedor.DescripcionOtroTelefono2 =
                    telefonos[1].Descripcion;
            }

            // Correos adicionales activos.
            var correos = proveedor.ContactosAdicionales
                .Where(c =>
                    c.Activo &&
                    c.TipoContacto == "Correo")
                .OrderBy(c => c.IdContactoProveedor)
                .Take(2)
                .ToList();

            if (correos.Count > 0)
            {
                proveedor.OtroCorreo =
                    correos[0].Valor;

                proveedor.DescripcionOtroCorreo =
                    correos[0].Descripcion;
            }

            if (correos.Count > 1)
            {
                proveedor.OtroCorreo2 =
                    correos[1].Valor;

                proveedor.DescripcionOtroCorreo2 =
                    correos[1].Descripcion;
            }

            // Direcciones adicionales activas.
            var direcciones = proveedor.ContactosAdicionales
                .Where(c =>
                    c.Activo &&
                    c.TipoContacto == "Direccion")
                .OrderBy(c => c.IdContactoProveedor)
                .Take(2)
                .ToList();

            if (direcciones.Count > 0)
            {
                proveedor.OtraDireccion =
                    direcciones[0].Valor;

                proveedor.DescripcionOtraDireccion =
                    direcciones[0].Descripcion;
            }

            if (direcciones.Count > 1)
            {
                proveedor.OtraDireccion2 =
                    direcciones[1].Valor;

                proveedor.DescripcionOtraDireccion2 =
                    direcciones[1].Descripcion;
            }

            ViewBag.TieneRegistros =
                await TieneRegistrosAsociados(
                    proveedor.IdProveedor);

            return View(proveedor);
        }
        // ================================================================
        // EDIT POST
        // Sin registros asociados: permite modificar todo.
        // Con registros asociados: solo datos de contacto.
        // ================================================================
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Edit(
            int id,
            Proveedore proveedorFormulario)
        {
            if (id != proveedorFormulario.IdProveedor)
            {
                return NotFound();
            }

            NormalizarProveedor(proveedorFormulario);
            ValidarContactosAdicionales(proveedorFormulario);

            var proveedorBD = await _context.Proveedores
                .Include(p => p.ContactosAdicionales)
                .Include(p => p.CuentasPorPagars)
                .Include(p => p.Facturas)
                .Include(p => p.Gastos)
                .FirstOrDefaultAsync(
                    p => p.IdProveedor == id);

            if (proveedorBD == null)
            {
                return NotFound();
            }

            if (!proveedorBD.Activo)
            {
                TempData["Error"] =
                    "No se puede modificar un proveedor inactivo. Primero debe reactivarlo.";

                return RedirectToAction(nameof(Index));
            }

            bool tieneRegistros =
                await TieneRegistrosAsociados(
                    proveedorBD.IdProveedor);

            await ValidarDuplicados(
                proveedorFormulario,
                id);

            if (!ModelState.IsValid)
            {
                ViewBag.TieneRegistros = tieneRegistros;

                TempData["Error"] =
                    "No fue posible modificar el proveedor. Revise los datos ingresados.";

                return View(proveedorFormulario);
            }

            if (tieneRegistros)
            {
                // Protege nombre e identificación.
                proveedorBD.Telefono =
                    proveedorFormulario.Telefono;

                proveedorBD.Email =
                    proveedorFormulario.Email;

                proveedorBD.Direccion =
                    proveedorFormulario.Direccion;
            }
            else
            {
                proveedorBD.Nombre =
                    proveedorFormulario.Nombre;

                proveedorBD.Identificacion =
                    proveedorFormulario.Identificacion;

                proveedorBD.Telefono =
                    proveedorFormulario.Telefono;

                proveedorBD.Email =
                    proveedorFormulario.Email;

                proveedorBD.Direccion =
                    proveedorFormulario.Direccion;
            }

            await SincronizarContactosAdicionales(
                proveedorBD.IdProveedor,
                proveedorFormulario);

            await _context.SaveChangesAsync();

            await RegistrarAuditoria(
                "Modificar",
                "Proveedores",
                proveedorBD.IdProveedor,
                tieneRegistros
                    ? $"Se modificaron datos de contacto del proveedor {proveedorBD.Nombre}. No se permitió cambiar nombre ni identificación por tener registros asociados."
                    : $"Se modificó el proveedor {proveedorBD.Nombre}.");

            TempData["Exito"] = tieneRegistros
                ? "Proveedor modificado correctamente. Por tener registros asociados, solo se actualizaron los datos de contacto."
                : "Proveedor modificado correctamente.";

            return RedirectToAction(nameof(Index));
        }

        // ================================================================
        // DELETE GET - ACCESO DIRECTO BLOQUEADO
        // La desactivación se realiza únicamente desde el listado mediante
        // SweetAlert y una solicitud POST.
   
        [HttpGet]
        public IActionResult Delete(int? id)
        {
            TempData["Error"] =
                "La desactivación de proveedores debe realizarse desde el listado.";

            return RedirectToAction(nameof(Index));
        }

        // ================================================================
        // DELETE POST - DESACTIVAR
        // ================================================================
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> DeleteConfirmed(int id)
        {
            var proveedor = await _context.Proveedores
                .Include(p => p.CuentasPorPagars)
                .Include(p => p.Facturas)
                .Include(p => p.Gastos)
                .FirstOrDefaultAsync(p => p.IdProveedor == id);

            if (proveedor == null)
            {
                return NotFound();
            }

            if (!await PuedeDesactivar(proveedor.IdProveedor))
            {
                TempData["Error"] = "No se puede desactivar el proveedor porque posee cuentas por pagar activas o pendientes.";
                return RedirectToAction(nameof(Index));
            }

            proveedor.Activo = false;
            await _context.SaveChangesAsync();

            await RegistrarAuditoria(
                "Desactivar",
                "Proveedores",
                proveedor.IdProveedor,
                $"Se desactivó el proveedor: {proveedor.Nombre}.");

            TempData["Exito"] = "Proveedor desactivado correctamente.";
            return RedirectToAction(nameof(Index));
        }

        // ================================================================
        // REACTIVAR
        // ================================================================
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Reactivar(int id)
        {
            var proveedor = await _context.Proveedores
                .FirstOrDefaultAsync(p => p.IdProveedor == id);

            if (proveedor == null)
            {
                return NotFound();
            }

            proveedor.Activo = true;
            await _context.SaveChangesAsync();

            await RegistrarAuditoria(
                "Reactivar",
                "Proveedores",
                proveedor.IdProveedor,
                $"Se reactivó el proveedor: {proveedor.Nombre}.");

            TempData["Exito"] = "Proveedor reactivado correctamente.";
            return RedirectToAction(nameof(Index));
        }
        // ================================================================
        // OBTENER PROVEEDORES CON INFORMACIÓN FINANCIERA
        // Se reutiliza en Index, PDF y Excel.
        // ================================================================
        private async Task<List<Proveedore>>
            ObtenerProveedoresConSituacionFinanciera(
                string? buscar,
                string? estado,
                string? situacionFinanciera,
                bool incluirContactosAdicionales = false)
        {
            var consulta = ConstruirConsultaReporte(
                buscar,
                estado,
                incluirContactosAdicionales);

            consulta = consulta
                .Include(p => p.CuentasPorPagars);

            var proveedores = await consulta
                .OrderByDescending(p => p.IdProveedor)
                .ToListAsync();

            CalcularSituacionFinancieraProveedores(
                proveedores);

            if (!string.IsNullOrWhiteSpace(
                situacionFinanciera))
            {
                situacionFinanciera =
                    situacionFinanciera.Trim();

                proveedores = proveedores
                    .Where(p =>
                        p.SituacionFinanciera ==
                        situacionFinanciera)
                    .ToList();
            }

            return proveedores;
        }


        // ================================================================
        // CALCULAR SITUACIÓN FINANCIERA DE PROVEEDORES
        // ================================================================
        private static void
            CalcularSituacionFinancieraProveedores(
                IEnumerable<Proveedore> proveedores)
        {
            var fechaHoy =
                DateOnly.FromDateTime(DateTime.Today);

            foreach (var proveedor in proveedores)
            {
                // Las cuentas anuladas no representan obligaciones vigentes.
                var cuentasValidas =
                    proveedor.CuentasPorPagars
                        .Where(c =>
                            c.Estado != "Anulada")
                        .ToList();

                var cuentasPendientes =
                    cuentasValidas
                        .Where(c =>
                            c.SaldoPendiente > 0)
                        .ToList();

                var cuentasVencidas =
                    cuentasPendientes
                        .Where(c =>
                            c.FechaVencimiento < fechaHoy)
                        .ToList();

                proveedor.TotalCuentasValidas =
                    cuentasValidas.Count;

                proveedor.TotalCuentasPendientes =
                    cuentasPendientes.Count;

                proveedor.TotalCuentasVencidas =
                    cuentasVencidas.Count;

                proveedor.SaldoTotalPendiente =
                    cuentasPendientes.Sum(c =>
                        c.SaldoPendiente);

                proveedor.SaldoTotalVencido =
                    cuentasVencidas.Sum(c =>
                        c.SaldoPendiente);

                // Prioridad:
                // Pago vencido > pago pendiente
                // > al día > sin movimientos.
                if (cuentasVencidas.Any())
                {
                    proveedor.SituacionFinanciera =
                        "Pago vencido";
                }
                else if (cuentasPendientes.Any())
                {
                    proveedor.SituacionFinanciera =
                        "Pago pendiente";
                }
                else if (cuentasValidas.Any())
                {
                    proveedor.SituacionFinanciera =
                        "Al día";
                }
                else
                {
                    proveedor.SituacionFinanciera =
                        "Sin movimientos";
                }
            }
        }


        // ================================================================
        // CARGAR RESUMEN FINANCIERO EN VIEWBAG
        // ================================================================
        private void CargarResumenFinancieroProveedores(
            IEnumerable<Proveedore> proveedores)
        {
            var lista = proveedores.ToList();

            ViewBag.ProveedoresConSaldo =
                lista.Count(p =>
                    p.SaldoTotalPendiente > 0);

            ViewBag.ProveedoresConCuentasVencidas =
                lista.Count(p =>
                    p.TotalCuentasVencidas > 0);

            ViewBag.SaldoTotalPendiente =
                lista.Sum(p =>
                    p.SaldoTotalPendiente);

            ViewBag.SaldoTotalVencido =
                lista.Sum(p =>
                    p.SaldoTotalVencido);

            ViewBag.TotalCuentasPendientes =
                lista.Sum(p =>
                    p.TotalCuentasPendientes);

            ViewBag.TotalCuentasVencidas =
                lista.Sum(p =>
                    p.TotalCuentasVencidas);
        }
        // ================================================================
        // CONSULTA COMPARTIDA PARA INDEX, REPORTE Y EXCEL
        // ================================================================
        private IQueryable<Proveedore> ConstruirConsultaReporte(
            string? buscar,
            string? estado,
            bool incluirContactosAdicionales = false)
        {
            IQueryable<Proveedore> consulta =
                _context.Proveedores
                    .AsNoTracking();

            if (incluirContactosAdicionales)
            {
                consulta = consulta
                    .Include(p => p.ContactosAdicionales);
            }

            if (!string.IsNullOrWhiteSpace(buscar))
            {
                buscar = buscar.Trim();

                consulta = consulta.Where(p =>
                    p.Nombre.Contains(buscar) ||
                    p.Identificacion.Contains(buscar) ||
                    (p.Email != null &&
                     p.Email.Contains(buscar)) ||
                    (p.Telefono != null &&
                     p.Telefono.Contains(buscar)) ||
                    (p.Direccion != null &&
                     p.Direccion.Contains(buscar)));
            }

            if (!string.IsNullOrWhiteSpace(estado))
            {
                if (estado == "Activo")
                {
                    consulta = consulta.Where(p =>
                        p.Activo);
                }
                else if (estado == "Inactivo")
                {
                    consulta = consulta.Where(p =>
                        !p.Activo);
                }
            }

            return consulta;
        }
        // ================================================================
        // MÉTODOS PRIVADOS
        // ================================================================

        private async Task<bool> TieneRegistrosAsociados(int idProveedor)
        {
            bool tieneCuentasPorPagar = await _context.CuentasPorPagars
                .AnyAsync(c => c.ProveedorId == idProveedor);

            bool tieneFacturas = await _context.Facturas
                .AnyAsync(f => f.ProveedorId == idProveedor);

            bool tieneGastos = await _context.Gastos
                .AnyAsync(g => g.ProveedorId == idProveedor);

            return tieneCuentasPorPagar || tieneFacturas || tieneGastos;
        }

        private async Task<bool> PuedeDesactivar(int idProveedor)
        {
            bool tieneCxPActivas = await _context.CuentasPorPagars
                .AnyAsync(c =>
                    c.ProveedorId == idProveedor &&
                    c.Estado != "Anulada" &&
                    c.Estado != "Cancelada");

            return !tieneCxPActivas;
        }

        private void NormalizarProveedor(Proveedore proveedor)
        {
            proveedor.Nombre =
                NormalizarTexto(proveedor.Nombre) ?? string.Empty;

            proveedor.Identificacion =
                NormalizarTexto(proveedor.Identificacion);

            proveedor.Telefono =
                NormalizarTexto(proveedor.Telefono);

            proveedor.Email =
                NormalizarTexto(proveedor.Email)?.ToLowerInvariant();

            proveedor.Direccion =
                NormalizarTexto(proveedor.Direccion);

            proveedor.OtroTelefono =
                NormalizarTexto(proveedor.OtroTelefono);

            proveedor.DescripcionOtroTelefono =
                NormalizarTexto(proveedor.DescripcionOtroTelefono);

            proveedor.OtroTelefono2 =
                NormalizarTexto(proveedor.OtroTelefono2);

            proveedor.DescripcionOtroTelefono2 =
                NormalizarTexto(proveedor.DescripcionOtroTelefono2);

            proveedor.OtroCorreo =
                NormalizarTexto(proveedor.OtroCorreo)?.ToLowerInvariant();

            proveedor.DescripcionOtroCorreo =
                NormalizarTexto(proveedor.DescripcionOtroCorreo);

            proveedor.OtroCorreo2 =
                NormalizarTexto(proveedor.OtroCorreo2)?.ToLowerInvariant();

            proveedor.DescripcionOtroCorreo2 =
                NormalizarTexto(proveedor.DescripcionOtroCorreo2);

            proveedor.OtraDireccion =
                NormalizarTexto(proveedor.OtraDireccion);

            proveedor.DescripcionOtraDireccion =
                NormalizarTexto(proveedor.DescripcionOtraDireccion);

            proveedor.OtraDireccion2 =
                NormalizarTexto(proveedor.OtraDireccion2);

            proveedor.DescripcionOtraDireccion2 =
                NormalizarTexto(proveedor.DescripcionOtraDireccion2);
        }
        private void ValidarContactosAdicionales(
    Proveedore proveedor)
        {
            ValidarContactoAdicional(
                proveedor.OtroTelefono,
                proveedor.DescripcionOtroTelefono,
                nameof(proveedor.OtroTelefono),
                nameof(proveedor.DescripcionOtroTelefono),
                "teléfono adicional");

            ValidarContactoAdicional(
                proveedor.OtroTelefono2,
                proveedor.DescripcionOtroTelefono2,
                nameof(proveedor.OtroTelefono2),
                nameof(proveedor.DescripcionOtroTelefono2),
                "segundo teléfono adicional");

            ValidarContactoAdicional(
                proveedor.OtroCorreo,
                proveedor.DescripcionOtroCorreo,
                nameof(proveedor.OtroCorreo),
                nameof(proveedor.DescripcionOtroCorreo),
                "correo adicional");

            ValidarContactoAdicional(
                proveedor.OtroCorreo2,
                proveedor.DescripcionOtroCorreo2,
                nameof(proveedor.OtroCorreo2),
                nameof(proveedor.DescripcionOtroCorreo2),
                "segundo correo adicional");

            ValidarContactoAdicional(
                proveedor.OtraDireccion,
                proveedor.DescripcionOtraDireccion,
                nameof(proveedor.OtraDireccion),
                nameof(proveedor.DescripcionOtraDireccion),
                "dirección adicional");

            ValidarContactoAdicional(
                proveedor.OtraDireccion2,
                proveedor.DescripcionOtraDireccion2,
                nameof(proveedor.OtraDireccion2),
                nameof(proveedor.DescripcionOtraDireccion2),
                "segunda dirección adicional");
        }

        private void ValidarContactoAdicional(
            string? valor,
            string? descripcion,
            string campoValor,
            string campoDescripcion,
            string nombreContacto)
        {
            if (string.IsNullOrWhiteSpace(valor) &&
                !string.IsNullOrWhiteSpace(descripcion))
            {
                ModelState.AddModelError(
                    campoValor,
                    $"Debe ingresar el {nombreContacto} antes de agregar su descripción.");

                ModelState.AddModelError(
                    campoDescripcion,
                    $"La descripción no puede registrarse sin el {nombreContacto}.");
            }
        }
        private static string? NormalizarTexto(string? valor)
        {
            if (string.IsNullOrWhiteSpace(valor))
            {
                return null;
            }

            return valor.Trim();
        }

        private async Task ValidarDuplicados(Proveedore proveedor, int? idActual = null)
        {
            bool identificacionDuplicada = await _context.Proveedores.AnyAsync(p =>
                p.Identificacion == proveedor.Identificacion &&
                (!idActual.HasValue || p.IdProveedor != idActual.Value));

            if (identificacionDuplicada)
            {
                ModelState.AddModelError("Identificacion", "Ya existe un proveedor registrado con esta identificación.");
            }

            if (!string.IsNullOrWhiteSpace(proveedor.Email))
            {
                bool emailDuplicado = await _context.Proveedores.AnyAsync(p =>
                    p.Email == proveedor.Email &&
                    (!idActual.HasValue || p.IdProveedor != idActual.Value));

                if (emailDuplicado)
                {
                    ModelState.AddModelError("Email", "Ya existe un proveedor registrado con este correo electrónico.");
                }
            }
        }
        private async Task SincronizarContactosAdicionales(
        int idProveedor,
        Proveedore formulario)
        {
            var contactos =
                await _context.ContactosProveedores
                    .Where(c =>
                        c.IdProveedor == idProveedor)
                    .OrderBy(c =>
                        c.IdContactoProveedor)
                    .ToListAsync();

            SincronizarTipoContacto(
                contactos,
                idProveedor,
                "Telefono",
                formulario.OtroTelefono,
                formulario.DescripcionOtroTelefono,
                formulario.OtroTelefono2,
                formulario.DescripcionOtroTelefono2);

            SincronizarTipoContacto(
                contactos,
                idProveedor,
                "Correo",
                formulario.OtroCorreo,
                formulario.DescripcionOtroCorreo,
                formulario.OtroCorreo2,
                formulario.DescripcionOtroCorreo2);

            SincronizarTipoContacto(
                contactos,
                idProveedor,
                "Direccion",
                formulario.OtraDireccion,
                formulario.DescripcionOtraDireccion,
                formulario.OtraDireccion2,
                formulario.DescripcionOtraDireccion2);
        }
        private void SincronizarTipoContacto(
    List<ContactoProveedor> contactos,
    int idProveedor,
    string tipoContacto,
    string? valor1,
    string? descripcion1,
    string? valor2,
    string? descripcion2)
        {
            var existentes = contactos
                .Where(c =>
                    c.TipoContacto == tipoContacto)
                .OrderBy(c =>
                    c.IdContactoProveedor)
                .ToList();

            ActualizarOCrearContacto(
                existentes,
                0,
                idProveedor,
                tipoContacto,
                valor1,
                descripcion1);

            ActualizarOCrearContacto(
                existentes,
                1,
                idProveedor,
                tipoContacto,
                valor2,
                descripcion2);

            // Si existieran contactos extra,
            // se conservan para mantener historial.
        }
        private void ActualizarOCrearContacto(
    List<ContactoProveedor> existentes,
    int posicion,
    int idProveedor,
    string tipoContacto,
    string? valor,
    string? descripcion)
        {
            valor = valor?.Trim();
            descripcion = descripcion?.Trim();

            if (posicion < existentes.Count)
            {
                var contacto =
                    existentes[posicion];

                if (string.IsNullOrWhiteSpace(valor))
                {
                    contacto.Activo = false;
                    return;
                }

                contacto.Valor = valor;

                contacto.Descripcion =
                    string.IsNullOrWhiteSpace(descripcion)
                        ? "Adicional"
                        : descripcion;

                contacto.Activo = true;
            }
            else if (!string.IsNullOrWhiteSpace(valor))
            {
                _context.ContactosProveedores.Add(
                    new ContactoProveedor
                    {
                        IdProveedor = idProveedor,
                        TipoContacto = tipoContacto,
                        Valor = valor,
                        Descripcion =
                            string.IsNullOrWhiteSpace(descripcion)
                                ? "Adicional"
                                : descripcion,
                        Activo = true
                    });
            }
        }
        private async Task RegistrarAuditoria(string accion, string tabla, int registroId, string descripcion)
        {
            int usuarioId = HttpContext.Session.GetInt32("UsuarioId") ?? 1;

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

        private bool ProveedoreExists(int id)
        {
            return _context.Proveedores.Any(e => e.IdProveedor == id);
        }
    }
}