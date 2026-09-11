using System;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using SistemaGestionFinanciera.Data;
using SistemaGestionFinanciera.Models;
using ClosedXML.Excel;
using System.Collections.Generic;
using System.IO;

namespace SistemaGestionFinanciera.Controllers
{
    public class ClientesBeneficiariosController : BaseController
    {
        private readonly SistemaFinancieroContext _context;

        public ClientesBeneficiariosController(SistemaFinancieroContext context)
        {
            _context = context;
        }

        // INDEX - LISTADO DE CLIENTES / BENEFICIARIOS
        public async Task<IActionResult> Index(
            string? buscar,
            string? tipo,
            string? estado,
            string? situacionFinanciera)
        {
            var clientes =
                await ObtenerClientesConSituacionFinanciera(
                    buscar,
                    tipo,
                    estado,
                    situacionFinanciera);

            CargarResumenFinancieroClientes(clientes);

            ViewBag.Buscar = buscar;
            ViewBag.Tipo = tipo;
            ViewBag.Estado = estado;

            ViewBag.SituacionFinanciera =
                situacionFinanciera;

            return View(clientes);
        }
        // ================================================================
        // REPORTE DE CLIENTES / BENEFICIARIOS
        // ================================================================
        [HttpGet]
        public async Task<IActionResult> Reporte(
            string? buscar,
            string? tipo,
            string? estado,
            string? situacionFinanciera)
        {
            var clientes =
                await ObtenerClientesConSituacionFinanciera(
                    buscar,
                    tipo,
                    estado,
                    situacionFinanciera,
                    incluirContactosAdicionales: true);

            ViewBag.TotalRegistros =
                clientes.Count;

            ViewBag.TotalActivos =
                clientes.Count(c => c.Activo);

            ViewBag.TotalInactivos =
                clientes.Count(c => !c.Activo);

            ViewBag.TotalClientes =
                clientes.Count(c =>
                    c.Tipo == "Cliente");

            ViewBag.TotalBeneficiarios =
                clientes.Count(c =>
                    c.Tipo == "Beneficiario");

            CargarResumenFinancieroClientes(clientes);

            ViewBag.Buscar = buscar;
            ViewBag.Tipo = tipo;
            ViewBag.Estado = estado;

            ViewBag.SituacionFinanciera =
                situacionFinanciera;

            ViewBag.FechaGeneracion =
                DateTime.Now;

            await RegistrarAuditoria(
                "Vista previa PDF",
                "ClientesBeneficiarios",
                0,
                "Se generó la vista previa del reporte de clientes y beneficiarios.");

            return View(clientes);
        }
        // ================================================================
        // EXPORTAR REPORTE DE CLIENTES / BENEFICIARIOS A EXCEL
        // ================================================================
        [HttpGet]
        public async Task<IActionResult> ExportarExcel(
    string? buscar,
    string? tipo,
    string? estado,
    string? situacionFinanciera)
        {
            var clientes =
      await ObtenerClientesConSituacionFinanciera(
          buscar,
          tipo,
          estado,
          situacionFinanciera,
          incluirContactosAdicionales: true);

            using var workbook = new XLWorkbook();

            var hoja = workbook.Worksheets.Add(
                "Clientes y beneficiarios");

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

            // ============================================================
            // TÍTULO
            // ============================================================
            hoja.Cell("A1").Value =
                "Unión Cantonal de Asociaciones de Desarrollo de Santa Ana";

            hoja.Range(1, 1, 1, totalColumnas).Merge();

            hoja.Cell("A2").Value =
                "Reporte de Clientes y Beneficiarios";

            hoja.Range(2, 1, 2, totalColumnas).Merge();

            hoja.Cell("A3").Value =
                $"Fecha de generación: {DateTime.Now:dd/MM/yyyy hh:mm tt}";

            hoja.Range(3, 1, 3, totalColumnas).Merge();

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

            rangoTitulo.Style.Font.FontColor =
                XLColor.Black;

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

            hoja.Range(5, 1, 5, totalColumnas).Merge();

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

            hoja.Cell("C6").Value = "Tipo:";

            hoja.Cell("D6").Value =
                string.IsNullOrWhiteSpace(tipo)
                    ? "Todos"
                    : tipo;

            hoja.Cell("E6").Value = "Estado:";

            hoja.Cell("F6").Value =
                string.IsNullOrWhiteSpace(estado)
                    ? "Todos"
                    : estado;
            hoja.Cell("G6").Value =
    "Situación financiera:";

            hoja.Cell("H6").Value =
                string.IsNullOrWhiteSpace(
                    situacionFinanciera)
                    ? "Todas"
                    : situacionFinanciera;

            hoja.Cell("G6").Style.Font.Bold = true;

            hoja.Cell("A6").Style.Font.Bold = true;
            hoja.Cell("C6").Style.Font.Bold = true;
            hoja.Cell("E6").Style.Font.Bold = true;

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
            hoja.Cell("A9").Value = "Total";
            hoja.Cell("B9").Value = clientes.Count;

            hoja.Cell("C9").Value = "Activos";
            hoja.Cell("D9").Value =
                clientes.Count(c => c.Activo);

            hoja.Cell("E9").Value = "Inactivos";
            hoja.Cell("F9").Value =
                clientes.Count(c => !c.Activo);

            hoja.Cell("G9").Value =
                "Clientes con saldo";

            hoja.Cell("H9").Value =
                clientes.Count(c =>
                    c.SaldoTotalPendiente > 0);

            hoja.Cell("I9").Value =
                "Clientes vencidos";

            hoja.Cell("J9").Value =
                clientes.Count(c =>
                    c.TotalCuentasVencidas > 0);


            // Segunda fila del resumen
            hoja.Cell("A10").Value = "Clientes";
            hoja.Cell("B10").Value =
                clientes.Count(c =>
                    c.Tipo == "Cliente");

            hoja.Cell("C10").Value =
                "Beneficiarios";

            hoja.Cell("D10").Value =
                clientes.Count(c =>
                    c.Tipo == "Beneficiario");

            hoja.Cell("E10").Value =
                "Saldo pendiente";

            hoja.Cell("F10").Value =
                clientes.Sum(c =>
                    c.SaldoTotalPendiente);

            hoja.Cell("G10").Value =
                "Saldo vencido";

            hoja.Cell("H10").Value =
                clientes.Sum(c =>
                    c.SaldoTotalVencido);

            hoja.Cell("I10").Value =
                "Cuentas vencidas";

            hoja.Cell("J10").Value =
                clientes.Sum(c =>
                    c.TotalCuentasVencidas);


            // Formato general
            var rangoResumen =
                hoja.Range("A9:J10");

            rangoResumen.Style.Font.Bold = true;

            rangoResumen.Style.Border.OutsideBorder =
                XLBorderStyleValues.Thin;

            rangoResumen.Style.Border.InsideBorder =
                XLBorderStyleValues.Thin;

            rangoResumen.Style.Alignment.Vertical =
                XLAlignmentVerticalValues.Center;


            // Fondos
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
                XLColor.FromHtml("#F8D7DA");

            hoja.Range("A10:B10")
                .Style.Fill.BackgroundColor =
                colorDoradoSuave;

            hoja.Range("C10:D10")
                .Style.Fill.BackgroundColor =
                colorVerdeSuave;

            hoja.Range("E10:F10")
                .Style.Fill.BackgroundColor =
                colorDoradoSuave;

            hoja.Range("G10:H10")
                .Style.Fill.BackgroundColor =
                XLColor.FromHtml("#F8D7DA");

            hoja.Range("I10:J10")
                .Style.Fill.BackgroundColor =
                XLColor.FromHtml("#F8D7DA");


            // Formato monetario
            hoja.Cell("F10").Style.NumberFormat.Format =
                "₡#,##0.00";

            hoja.Cell("H10").Style.NumberFormat.Format =
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
    "Tipo",
    "Identificación",
    "Teléfono principal",
    "Correo principal",
    "Teléfonos adicionales",
    "Correos adicionales",
    "Direcciones adicionales",
    "Estado",
    "Situación financiera",
    "Cuentas pendientes",
    "Cuentas vencidas",
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

            encabezadoTabla.Style.Border.OutsideBorder =
                XLBorderStyleValues.Thin;

            encabezadoTabla.Style.Border.InsideBorder =
                XLBorderStyleValues.Thin;

            encabezadoTabla.Style.Border.OutsideBorderColor =
                colorInstitucionalOscuro;

            encabezadoTabla.Style.Border.InsideBorderColor =
                colorInstitucionalOscuro;

            hoja.Row(filaEncabezado).Height = 28;

            // ============================================================
            // DATOS
            // ============================================================
            int fila = filaEncabezado + 1;

            foreach (var cliente in clientes)
            {
                hoja.Cell(fila, 1).Value =
                    cliente.Nombre;

                hoja.Cell(fila, 2).Value =
                    cliente.Tipo;

                hoja.Cell(fila, 3).Value =
                    cliente.Identificacion;

                hoja.Cell(fila, 4).Value =
                    string.IsNullOrWhiteSpace(cliente.Telefono)
                        ? "No registrado"
                        : cliente.Telefono;

                hoja.Cell(fila, 5).Value =
                    string.IsNullOrWhiteSpace(cliente.Email)
                        ? "No registrado"
                        : cliente.Email;

                var telefonosAdicionales =
                    cliente.ContactosAdicionales
                        .Where(c =>
                            c.Activo &&
                            c.TipoContacto == "Telefono")
                        .OrderBy(c => c.IdContactoCliente)
                        .Select(c =>
                            string.IsNullOrWhiteSpace(c.Descripcion)
                                ? c.Valor
                                : $"{c.Valor} ({c.Descripcion})")
                        .ToList();

                var correosAdicionales =
                    cliente.ContactosAdicionales
                        .Where(c =>
                            c.Activo &&
                            c.TipoContacto == "Correo")
                        .OrderBy(c => c.IdContactoCliente)
                        .Select(c =>
                            string.IsNullOrWhiteSpace(c.Descripcion)
                                ? c.Valor
                                : $"{c.Valor} ({c.Descripcion})")
                        .ToList();

                var direccionesAdicionales =
                    cliente.ContactosAdicionales
                        .Where(c =>
                            c.Activo &&
                            c.TipoContacto == "Direccion")
                        .OrderBy(c => c.IdContactoCliente)
                        .Select(c =>
                            string.IsNullOrWhiteSpace(c.Descripcion)
                                ? c.Valor
                                : $"{c.Valor} ({c.Descripcion})")
                        .ToList();

                hoja.Cell(fila, 6).Value =
                    telefonosAdicionales.Any()
                        ? string.Join(Environment.NewLine, telefonosAdicionales)
                        : "No registrado";

                hoja.Cell(fila, 7).Value =
                    correosAdicionales.Any()
                        ? string.Join(Environment.NewLine, correosAdicionales)
                        : "No registrado";

                hoja.Cell(fila, 8).Value =
                    direccionesAdicionales.Any()
                        ? string.Join(Environment.NewLine, direccionesAdicionales)
                        : "No registrado";

                hoja.Cell(fila, 9).Value =
                    cliente.Activo
                        ? "Activo"
                        : "Inactivo";
                hoja.Cell(fila, 10).Value =
    cliente.SituacionFinanciera;

                hoja.Cell(fila, 11).Value =
                    cliente.TotalCuentasPendientes;

                hoja.Cell(fila, 12).Value =
                    cliente.TotalCuentasVencidas;

                hoja.Cell(fila, 13).Value =
                    cliente.SaldoTotalPendiente;

                hoja.Cell(fila, 14).Value =
                    cliente.SaldoTotalVencido;

                hoja.Cell(fila, 13)
                    .Style.NumberFormat.Format =
                    "₡#,##0.00";

                hoja.Cell(fila, 14)
                    .Style.NumberFormat.Format =
                    "₡#,##0.00";

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

                if (cliente.Activo)
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

                hoja.Cell(fila, 9).Style.Font.Bold = true;

                switch (cliente.SituacionFinanciera)
                {
                    case "Cuenta vencida":

                        hoja.Cell(fila, 10)
                            .Style.Fill.BackgroundColor =
                            XLColor.FromHtml("#F8D7DA");

                        hoja.Cell(fila, 10)
                            .Style.Font.FontColor =
                            XLColor.FromHtml("#842029");

                        break;

                    case "Saldo pendiente":

                        hoja.Cell(fila, 10)
                            .Style.Fill.BackgroundColor =
                            XLColor.FromHtml("#FFF3CD");

                        hoja.Cell(fila, 10)
                            .Style.Font.FontColor =
                            XLColor.FromHtml("#664D03");

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

                hoja.Cell(fila, 10).Style.Font.Bold = true;

                if (cliente.SaldoTotalVencido > 0)
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
                if (clientes.Any())
            {
                int filaFinDatos = fila - 1;

                var rangoTabla = hoja.Range(
                    filaEncabezado,
                    1,
                    filaFinDatos,
                    totalColumnas);

                var tabla = rangoTabla.CreateTable(
                    "TablaClientesBeneficiarios");

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

                hoja.Range(
                        filaEncabezado + 1,
                        2,
                        filaFinDatos,
                        3)
                    .Style.Alignment.Horizontal =
                    XLAlignmentHorizontalValues.Center;

                hoja.Range(
                        filaEncabezado + 1,
                        9,
                        filaFinDatos,
                        9)
                    .Style.Alignment.Horizontal =
                    XLAlignmentHorizontalValues.Center;
                hoja.Range(
        filaEncabezado + 1,
        6,
        filaFinDatos,
        8)
    .Style.Alignment.WrapText = true;
            }
            else
            {
                hoja.Cell(fila, 1).Value =
                    "No se encontraron clientes o beneficiarios con los filtros aplicados.";

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
                clientes.Count;

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

            hoja.Cell(filaTotal, 13).Style.Font.Bold = true;
            hoja.Cell(filaTotal, 14).Style.Font.Bold = true;

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
            hoja.Column(1).Width = 30;
            hoja.Column(2).Width = 16;
            hoja.Column(3).Width = 20;
            hoja.Column(4).Width = 20;
            hoja.Column(5).Width = 34;
            hoja.Column(6).Width = 28;
            hoja.Column(7).Width = 34;
            hoja.Column(8).Width = 40;
            hoja.Column(9).Width = 14;
            hoja.Column(10).Width = 22;
            hoja.Column(11).Width = 18;
            hoja.Column(12).Width = 18;
            hoja.Column(13).Width = 20;
            hoja.Column(14).Width = 20;

            hoja.Column(1).Style.Alignment.WrapText = true;
            hoja.Column(5).Style.Alignment.WrapText = true;
            hoja.Column(6).Style.Alignment.WrapText = true;
            hoja.Column(7).Style.Alignment.WrapText = true;
            hoja.Column(8).Style.Alignment.WrapText = true;
            hoja.Column(10).Style.Alignment.WrapText = true;

            hoja.SheetView.FreezeRows(filaEncabezado);

            hoja.PageSetup.PageOrientation =
                XLPageOrientation.Landscape;

            hoja.PageSetup.PaperSize =
                XLPaperSize.A4Paper;

            hoja.PageSetup.FitToPages(1, 0);

            hoja.PageSetup.CenterHorizontally = true;

            hoja.PageSetup.SetRowsToRepeatAtTop(
                filaEncabezado,
                filaEncabezado);

            using var stream = new MemoryStream();

            workbook.SaveAs(stream);

            string nombreArchivo =
                $"Reporte_ClientesBeneficiarios_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";
            await RegistrarAuditoria(
"Exportar Excel",
"ClientesBeneficiarios",
0,
"Se exportó el reporte de clientes y beneficiarios a Excel."
);
            return File(
                stream.ToArray(),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                nombreArchivo);
        }

        // DETAILS - DETALLE DEL CLIENTE // Muestra datos personales, contactos adicionales y resumen de CxC.
        public async Task<IActionResult> Details(int? id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var cliente = await _context.ClientesBeneficiarios
                .Include(c => c.ContactosAdicionales)
                .Include(c => c.CuentasPorCobrars)
                .Include(c => c.Facturas)
                .Include(c => c.Ingresos)
                .FirstOrDefaultAsync(c => c.IdClienteBeneficiario == id);

            if (cliente == null)
            {
                return NotFound();
            }

            var fechaHoy = DateOnly.FromDateTime(DateTime.Today);

            // Cuentas válidas para el resumen financiero.
            // No toma en cuenta cuentas anuladas.
            var cuentasValidas = cliente.CuentasPorCobrars
                .Where(c => c.Estado != "Anulada")
                .ToList();

            var cuentasPendientes = cuentasValidas
                .Where(c => c.SaldoPendiente > 0)
                .ToList();

            var cuentasVencidas = cuentasPendientes
                .Where(c => c.FechaVencimiento < fechaHoy)
                .ToList();

            ViewBag.TotalCuentasPorCobrar = cuentasValidas.Count;
            ViewBag.CuentasPendientes = cuentasPendientes.Count;
            ViewBag.CuentasVencidas = cuentasVencidas.Count;
            ViewBag.SaldoPendiente = cuentasPendientes.Sum(c => c.SaldoPendiente);

            ViewBag.TieneRegistros =
                await TieneRegistrosAsociados(cliente.IdClienteBeneficiario);

            ViewBag.PuedeDesactivar =
                await PuedeDesactivar(cliente.IdClienteBeneficiario);

            return View(cliente);
        }

        // ================================================================
        // CREATE GET
        // ================================================================
        public IActionResult Create()
        {
            var cliente = new ClientesBeneficiario
            {
                Activo = true
            };

            return View(cliente);
        }

        // ================================================================
        // CREATE POST
        // Crea cliente activo automáticamente.
        // Valida duplicidad de identificación y correo.
        // ================================================================
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Create(ClientesBeneficiario cliente)
        {
            cliente.Activo = true;

            NormalizarCliente(cliente);
            ValidarContactosAdicionales(cliente);

            await ValidarDuplicados(cliente);

            if (!ModelState.IsValid)
            {
                TempData["Error"] = "No fue posible registrar el cliente. Revise los datos ingresados.";
                return View(cliente);
            }

            _context.ClientesBeneficiarios.Add(cliente);
            await _context.SaveChangesAsync();
            // guardar contactos adicioonales 
            // ================================================================
            // GUARDAR CONTACTOS ADICIONALES OPCIONALES
            // ================================================================

            if (!string.IsNullOrWhiteSpace(cliente.OtroTelefono))
            {
                _context.ContactosClientesBeneficiarios.Add(
                    new ContactoClienteBeneficiario
                    {
                        IdClienteBeneficiario = cliente.IdClienteBeneficiario,
                        TipoContacto = "Telefono",
                        Valor = cliente.OtroTelefono,
                        Descripcion = string.IsNullOrWhiteSpace(cliente.DescripcionOtroTelefono)
                            ? "Adicional"
                            : cliente.DescripcionOtroTelefono,
                        Activo = true
                    });
            }
            if (!string.IsNullOrWhiteSpace(cliente.OtroTelefono2))
            {
                _context.ContactosClientesBeneficiarios.Add(
                    new ContactoClienteBeneficiario
                    {
                        IdClienteBeneficiario = cliente.IdClienteBeneficiario,
                        TipoContacto = "Telefono",
                        Valor = cliente.OtroTelefono2,
                        Descripcion = string.IsNullOrWhiteSpace(cliente.DescripcionOtroTelefono2)
                            ? "Adicional 2"
                            : cliente.DescripcionOtroTelefono2,
                        Activo = true
                    });
            }
            if (!string.IsNullOrWhiteSpace(cliente.OtroCorreo))
            {
                _context.ContactosClientesBeneficiarios.Add(
                    new ContactoClienteBeneficiario
                    {
                        IdClienteBeneficiario = cliente.IdClienteBeneficiario,
                        TipoContacto = "Correo",
                        Valor = cliente.OtroCorreo,
                        Descripcion = string.IsNullOrWhiteSpace(cliente.DescripcionOtroCorreo)
                            ? "Adicional"
                            : cliente.DescripcionOtroCorreo,
                        Activo = true
                    });
            }
            if (!string.IsNullOrWhiteSpace(cliente.OtroCorreo2))
            {
                _context.ContactosClientesBeneficiarios.Add(
                    new ContactoClienteBeneficiario
                    {
                        IdClienteBeneficiario = cliente.IdClienteBeneficiario,
                        TipoContacto = "Correo",
                        Valor = cliente.OtroCorreo2,
                        Descripcion = string.IsNullOrWhiteSpace(cliente.DescripcionOtroCorreo2)
                            ? "Adicional 2"
                            : cliente.DescripcionOtroCorreo2,
                        Activo = true
                    });
            }

            if (!string.IsNullOrWhiteSpace(cliente.OtraDireccion))
            {
                _context.ContactosClientesBeneficiarios.Add(
                    new ContactoClienteBeneficiario
                    {
                        IdClienteBeneficiario = cliente.IdClienteBeneficiario,
                        TipoContacto = "Direccion",
                        Valor = cliente.OtraDireccion,
                        Descripcion = string.IsNullOrWhiteSpace(cliente.DescripcionOtraDireccion)
                            ? "Adicional"
                            : cliente.DescripcionOtraDireccion,
                        Activo = true
                    });
            }
            if (!string.IsNullOrWhiteSpace(cliente.OtraDireccion2))
            {
                _context.ContactosClientesBeneficiarios.Add(
                    new ContactoClienteBeneficiario
                    {
                        IdClienteBeneficiario = cliente.IdClienteBeneficiario,
                        TipoContacto = "Direccion",
                        Valor = cliente.OtraDireccion2,
                        Descripcion = string.IsNullOrWhiteSpace(cliente.DescripcionOtraDireccion2)
                            ? "Adicional 2"
                            : cliente.DescripcionOtraDireccion2,
                        Activo = true
                    });
            }


            await _context.SaveChangesAsync();

            //auditoria
            await RegistrarAuditoria(
       "Crear",
       "ClientesBeneficiarios",
       cliente.IdClienteBeneficiario,
       $"Se registró el cliente/beneficiario: {cliente.Nombre}, identificación: {cliente.Identificacion}.");

            TempData["Exito"] = "Cliente registrado correctamente.";
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

            var cliente = await _context.ClientesBeneficiarios
                .Include(c => c.ContactosAdicionales)
                .Include(c => c.CuentasPorCobrars)
                .Include(c => c.Facturas)
                .Include(c => c.Ingresos)
                .FirstOrDefaultAsync(c => c.IdClienteBeneficiario == id);

            if (cliente == null)
            {
                return NotFound();
            }

            if (!cliente.Activo)
            {
                TempData["Error"] =
                    "No se puede modificar un cliente inactivo. Primero debe reactivarlo.";

                return RedirectToAction(nameof(Index));
            }

            // Teléfonos adicionales activos
            var telefonos = cliente.ContactosAdicionales
                .Where(c => c.Activo && c.TipoContacto == "Telefono")
                .OrderBy(c => c.IdContactoCliente)
                .Take(2)
                .ToList();

            if (telefonos.Count > 0)
            {
                cliente.OtroTelefono = telefonos[0].Valor;
                cliente.DescripcionOtroTelefono = telefonos[0].Descripcion;
            }

            if (telefonos.Count > 1)
            {
                cliente.OtroTelefono2 = telefonos[1].Valor;
                cliente.DescripcionOtroTelefono2 = telefonos[1].Descripcion;
            }

            // Correos adicionales activos
            var correos = cliente.ContactosAdicionales
                .Where(c => c.Activo && c.TipoContacto == "Correo")
                .OrderBy(c => c.IdContactoCliente)
                .Take(2)
                .ToList();

            if (correos.Count > 0)
            {
                cliente.OtroCorreo = correos[0].Valor;
                cliente.DescripcionOtroCorreo = correos[0].Descripcion;
            }

            if (correos.Count > 1)
            {
                cliente.OtroCorreo2 = correos[1].Valor;
                cliente.DescripcionOtroCorreo2 = correos[1].Descripcion;
            }

            // Direcciones adicionales activas
            var direcciones = cliente.ContactosAdicionales
                .Where(c => c.Activo && c.TipoContacto == "Direccion")
                .OrderBy(c => c.IdContactoCliente)
                .Take(2)
                .ToList();

            if (direcciones.Count > 0)
            {
                cliente.OtraDireccion = direcciones[0].Valor;
                cliente.DescripcionOtraDireccion = direcciones[0].Descripcion;
            }

            if (direcciones.Count > 1)
            {
                cliente.OtraDireccion2 = direcciones[1].Valor;
                cliente.DescripcionOtraDireccion2 = direcciones[1].Descripcion;
            }

            ViewBag.TieneRegistros =
                await TieneRegistrosAsociados(cliente.IdClienteBeneficiario);

            return View(cliente);
        }


        // ================================================================
        // EDIT POST
        // Si NO tiene registros asociados: permite modificar todo.
        // Si SÍ tiene registros asociados: solo permite datos generales.
        // ================================================================
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Edit(int id, ClientesBeneficiario clienteFormulario)
        {
            if (id != clienteFormulario.IdClienteBeneficiario)
            {
                return NotFound();
            }

            NormalizarCliente(clienteFormulario);

            ValidarContactosAdicionales(clienteFormulario);

            var clienteBD = await _context.ClientesBeneficiarios
                .Include(c => c.ContactosAdicionales)
                .Include(c => c.CuentasPorCobrars)
                .Include(c => c.Facturas)
                .Include(c => c.Ingresos)
               
                .FirstOrDefaultAsync(c => c.IdClienteBeneficiario == id);

            if (clienteBD == null)
            {
                return NotFound();
            }

            if (!clienteBD.Activo)
            {
                TempData["Error"] = "No se puede modificar un cliente inactivo. Primero debe reactivarlo.";
                return RedirectToAction(nameof(Index));
            }

            bool tieneRegistros = await TieneRegistrosAsociados(clienteBD.IdClienteBeneficiario);

            await ValidarDuplicados(clienteFormulario, id);

            if (!ModelState.IsValid)
            {
                ViewBag.TieneRegistros = tieneRegistros;
                TempData["Error"] = "No fue posible modificar el cliente. Revise los datos ingresados.";
                return View(clienteFormulario);
            }

            if (tieneRegistros)
            {
                clienteBD.Telefono = clienteFormulario.Telefono;
                clienteBD.Email = clienteFormulario.Email;
                clienteBD.Direccion = clienteFormulario.Direccion; ;
            }
            else
            {
                clienteBD.Nombre = clienteFormulario.Nombre;
                clienteBD.Tipo = clienteFormulario.Tipo;
                clienteBD.Identificacion = clienteFormulario.Identificacion;
                clienteBD.Telefono = clienteFormulario.Telefono;
                clienteBD.Email = clienteFormulario.Email;
                clienteBD.Direccion = clienteFormulario.Direccion;
            }

            await SincronizarContactosAdicionales(
    clienteBD.IdClienteBeneficiario,
    clienteFormulario);
            await _context.SaveChangesAsync();

            // auditoria 
            await RegistrarAuditoria(
         "Modificar",
         "ClientesBeneficiarios",
         clienteBD.IdClienteBeneficiario,
         tieneRegistros
           ? $"Se modificaron datos de contacto del cliente {clienteBD.Nombre}. No se permitió cambiar nombre, tipo ni identificación por tener registros asociados."
             : $"Se modificó el cliente {clienteBD.Nombre}.");

            TempData["Exito"] = tieneRegistros
                ? "Cliente modificado correctamente. Por tener registros asociados, solo se actualizaron los datos generales."
                : "Cliente modificado correctamente.";

            return RedirectToAction(nameof(Index));
        }

        // ================================================================
        // DELETE GET - ACCESO DIRECTO BLOQUEADO
        // La desactivación se realiza únicamente desde el listado mediante
        // SweetAlert y una solicitud POST.
        // ================================================================
        [HttpGet]
        public IActionResult Delete(int? id)
        {
            TempData["Error"] =
                "La desactivación de clientes y beneficiarios debe realizarse desde el listado.";

            return RedirectToAction(nameof(Index));
        }

        // ================================================================
        // DELETE POST - DESACTIVAR
        // No borra registros. Cambia Activo=false si cumple reglas.
        // ================================================================
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> DeleteConfirmed(int id)
        {
            var cliente = await _context.ClientesBeneficiarios
                .Include(c => c.CuentasPorCobrars)
                .Include(c => c.Facturas)
                .Include(c => c.Ingresos)
                .FirstOrDefaultAsync(c => c.IdClienteBeneficiario == id);

            if (cliente == null)
            {
                return NotFound();
            }

            if (!await PuedeDesactivar(cliente.IdClienteBeneficiario))
            {
                TempData["Error"] = "No se puede desactivar el cliente porque posee registros financieros activos o pendientes.";
                return RedirectToAction(nameof(Index));
            }

            cliente.Activo = false;
            await _context.SaveChangesAsync();

            //auditoria
            await RegistrarAuditoria(
     "Desactivar",
     "ClientesBeneficiarios",
     cliente.IdClienteBeneficiario,
     $"Se desactivó el cliente/beneficiario: {cliente.Nombre}.");

            TempData["Exito"] = "Cliente desactivado correctamente.";
            return RedirectToAction(nameof(Index));
        }

        // ================================================================
        // REACTIVAR CLIENTE
        // ================================================================
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Reactivar(int id)
        {
            var cliente = await _context.ClientesBeneficiarios
                .FirstOrDefaultAsync(c => c.IdClienteBeneficiario == id);

            if (cliente == null)
            {
                return NotFound();
            }

            cliente.Activo = true;
            await _context.SaveChangesAsync();
             //auditoria
            await RegistrarAuditoria(
     "Reactivar",
     "ClientesBeneficiarios",
     cliente.IdClienteBeneficiario,
     $"Se reactivó el cliente/beneficiario: {cliente.Nombre}.");

            TempData["Exito"] = "Cliente reactivado correctamente.";
            return RedirectToAction(nameof(Index));
        }
        // ================================================================
        // OBTENER CLIENTES CON INFORMACIÓN FINANCIERA
        // Se reutiliza en Index, Reporte PDF y Exportación Excel.
        // ================================================================
        private async Task<List<ClientesBeneficiario>>
            ObtenerClientesConSituacionFinanciera(
                string? buscar,
                string? tipo,
                string? estado,
                string? situacionFinanciera,
                bool incluirContactosAdicionales = false)
        {
            var consulta = ConstruirConsultaReporte(
                buscar,
                tipo,
                estado,
                incluirContactosAdicionales);

            consulta = consulta
                .Include(c => c.CuentasPorCobrars);

            var clientes = await consulta
                .OrderByDescending(c => c.IdClienteBeneficiario)
                .ToListAsync();

            CalcularSituacionFinanciera(clientes);

            if (!string.IsNullOrWhiteSpace(situacionFinanciera))
            {
                situacionFinanciera =
                    situacionFinanciera.Trim();

                clientes = clientes
                    .Where(c =>
                        c.SituacionFinanciera ==
                        situacionFinanciera)
                    .ToList();
            }

            return clientes;
        }


        // ================================================================
        // CALCULAR SITUACIÓN FINANCIERA
        // ================================================================
        private static void CalcularSituacionFinanciera(
            IEnumerable<ClientesBeneficiario> clientes)
        {
            var fechaHoy =
                DateOnly.FromDateTime(DateTime.Today);

            foreach (var cliente in clientes)
            {
                // Se excluyen las cuentas anuladas porque no representan
                // obligaciones financieras vigentes.
                var cuentasValidas =
                    cliente.CuentasPorCobrars
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

                cliente.TotalCuentasValidas =
                    cuentasValidas.Count;

                cliente.TotalCuentasPendientes =
                    cuentasPendientes.Count;

                cliente.TotalCuentasVencidas =
                    cuentasVencidas.Count;

                cliente.SaldoTotalPendiente =
                    cuentasPendientes
                        .Sum(c => c.SaldoPendiente);

                cliente.SaldoTotalVencido =
                    cuentasVencidas
                        .Sum(c => c.SaldoPendiente);

                // Prioridad:
                // Cuenta vencida > saldo pendiente
                // > al día > sin movimientos
                if (cuentasVencidas.Any())
                {
                    cliente.SituacionFinanciera =
                        "Cuenta vencida";
                }
                else if (cuentasPendientes.Any())
                {
                    cliente.SituacionFinanciera =
                        "Saldo pendiente";
                }
                else if (cuentasValidas.Any())
                {
                    cliente.SituacionFinanciera =
                        "Al día";
                }
                else
                {
                    cliente.SituacionFinanciera =
                        "Sin movimientos";
                }
            }
        }


        // ================================================================
        // CARGAR RESUMEN FINANCIERO EN VIEWBAG
        // ================================================================
        private void CargarResumenFinancieroClientes(
            IEnumerable<ClientesBeneficiario> clientes)
        {
            var lista = clientes.ToList();

            ViewBag.ClientesConSaldo =
                lista.Count(c =>
                    c.SaldoTotalPendiente > 0);

            ViewBag.ClientesConCuentasVencidas =
                lista.Count(c =>
                    c.TotalCuentasVencidas > 0);

            ViewBag.SaldoTotalPendiente =
                lista.Sum(c =>
                    c.SaldoTotalPendiente);

            ViewBag.SaldoTotalVencido =
                lista.Sum(c =>
                    c.SaldoTotalVencido);

            ViewBag.TotalCuentasPendientes =
                lista.Sum(c =>
                    c.TotalCuentasPendientes);

            ViewBag.TotalCuentasVencidas =
                lista.Sum(c =>
                    c.TotalCuentasVencidas);
        }
        // ================================================================
        // CONSULTA COMPARTIDA PARA INDEX, REPORTE Y EXCEL
        // ================================================================
        // CONSULTA COMPARTIDA PARA INDEX, REPORTE Y EXCEL
        // ================================================================
        private IQueryable<ClientesBeneficiario> ConstruirConsultaReporte(
            string? buscar,
            string? tipo,
            string? estado,
            bool incluirContactosAdicionales = false)
        {
            IQueryable<ClientesBeneficiario> consulta =
                _context.ClientesBeneficiarios
                    .AsNoTracking();

            if (incluirContactosAdicionales)
            {
                consulta = consulta
                    .Include(c => c.ContactosAdicionales);
            }

            if (!string.IsNullOrWhiteSpace(buscar))
            {
                buscar = buscar.Trim();

                consulta = consulta.Where(c =>
                    c.Nombre.Contains(buscar) ||
                    c.Identificacion.Contains(buscar) ||
                    (c.Email != null &&
                     c.Email.Contains(buscar)) ||
                    (c.Telefono != null &&
                     c.Telefono.Contains(buscar)) ||

                    c.ContactosAdicionales.Any(contacto =>
                        contacto.Activo &&
                        contacto.Valor.Contains(buscar)));
            }

            if (!string.IsNullOrWhiteSpace(tipo))
            {
                consulta = consulta.Where(c =>
                    c.Tipo == tipo);
            }

            if (!string.IsNullOrWhiteSpace(estado))
            {
                if (estado == "Activo")
                {
                    consulta = consulta.Where(c =>
                        c.Activo);
                }
                else if (estado == "Inactivo")
                {
                    consulta = consulta.Where(c =>
                        !c.Activo);
                }
            }

            return consulta;
        }
        // MÉTODOS PRIVADOS DE APOYO
        // ================================================================
        private async Task<bool> TieneRegistrosAsociados(int idCliente)
        {
            bool tieneCuentasPorCobrar = await _context.CuentasPorCobrars
                .AnyAsync(c => c.ClienteBeneficiarioId == idCliente);

            bool tieneFacturas = await _context.Facturas
                .AnyAsync(f => f.ClienteBeneficiarioId == idCliente);

            bool tieneIngresos = await _context.Ingresos
              .AnyAsync(i => i.ClienteBeneficiarioId == idCliente);

            return tieneCuentasPorCobrar || tieneFacturas || tieneIngresos;
        }

        private async Task<bool> PuedeDesactivar(int idCliente)
        {
            bool tieneCxCActivas = await _context.CuentasPorCobrars
                .AnyAsync(c =>
                    c.ClienteBeneficiarioId == idCliente &&
                    c.Estado != "Anulada" &&
                    c.Estado != "Cancelada");

            return !tieneCxCActivas;
        }

        private void NormalizarCliente(ClientesBeneficiario cliente)
        {
            cliente.Nombre = NormalizarTexto(cliente.Nombre) ?? string.Empty;
            cliente.Tipo = NormalizarTexto(cliente.Tipo) ?? string.Empty;
            cliente.Identificacion = NormalizarTexto(cliente.Identificacion);

            cliente.Telefono = NormalizarTexto(cliente.Telefono);
            cliente.Email = NormalizarTexto(cliente.Email)?.ToLowerInvariant();
            cliente.Direccion = NormalizarTexto(cliente.Direccion);

            cliente.OtroTelefono = NormalizarTexto(cliente.OtroTelefono);
            cliente.DescripcionOtroTelefono =
                NormalizarTexto(cliente.DescripcionOtroTelefono);

            cliente.OtroTelefono2 = NormalizarTexto(cliente.OtroTelefono2);
            cliente.DescripcionOtroTelefono2 =
                NormalizarTexto(cliente.DescripcionOtroTelefono2);

            cliente.OtroCorreo =
                NormalizarTexto(cliente.OtroCorreo)?.ToLowerInvariant();

            cliente.DescripcionOtroCorreo =
                NormalizarTexto(cliente.DescripcionOtroCorreo);

            cliente.OtroCorreo2 =
                NormalizarTexto(cliente.OtroCorreo2)?.ToLowerInvariant();

            cliente.DescripcionOtroCorreo2 =
                NormalizarTexto(cliente.DescripcionOtroCorreo2);

            cliente.OtraDireccion =
                NormalizarTexto(cliente.OtraDireccion);

            cliente.DescripcionOtraDireccion =
                NormalizarTexto(cliente.DescripcionOtraDireccion);

            cliente.OtraDireccion2 =
                NormalizarTexto(cliente.OtraDireccion2);

            cliente.DescripcionOtraDireccion2 =
                NormalizarTexto(cliente.DescripcionOtraDireccion2);
        }

        private static string? NormalizarTexto(string? valor)
        {
            if (string.IsNullOrWhiteSpace(valor))
            {
                return null;
            }

            return valor.Trim();
        }
        private void ValidarContactosAdicionales(
    ClientesBeneficiario cliente)
        {
            ValidarContactoAdicional(
                cliente.OtroTelefono,
                cliente.DescripcionOtroTelefono,
                nameof(cliente.OtroTelefono),
                nameof(cliente.DescripcionOtroTelefono),
                "teléfono adicional");

            ValidarContactoAdicional(
                cliente.OtroTelefono2,
                cliente.DescripcionOtroTelefono2,
                nameof(cliente.OtroTelefono2),
                nameof(cliente.DescripcionOtroTelefono2),
                "segundo teléfono adicional");

            ValidarContactoAdicional(
                cliente.OtroCorreo,
                cliente.DescripcionOtroCorreo,
                nameof(cliente.OtroCorreo),
                nameof(cliente.DescripcionOtroCorreo),
                "correo adicional");

            ValidarContactoAdicional(
                cliente.OtroCorreo2,
                cliente.DescripcionOtroCorreo2,
                nameof(cliente.OtroCorreo2),
                nameof(cliente.DescripcionOtroCorreo2),
                "segundo correo adicional");

            ValidarContactoAdicional(
                cliente.OtraDireccion,
                cliente.DescripcionOtraDireccion,
                nameof(cliente.OtraDireccion),
                nameof(cliente.DescripcionOtraDireccion),
                "dirección adicional");

            ValidarContactoAdicional(
                cliente.OtraDireccion2,
                cliente.DescripcionOtraDireccion2,
                nameof(cliente.OtraDireccion2),
                nameof(cliente.DescripcionOtraDireccion2),
                "segunda dirección adicional");
        }

        private void ValidarContactoAdicional(
            string? valor,
            string? descripcion,
            string campoValor,
            string campoDescripcion,
            string nombreContacto)
        {
            /*
             * No debe permitirse una descripción sin haber ingresado
             * el contacto al cual pertenece.
             */
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
        private async Task ValidarDuplicados(ClientesBeneficiario cliente, int? idActual = null)
        {
            bool identificacionDuplicada = await _context.ClientesBeneficiarios.AnyAsync(c =>
                c.Identificacion == cliente.Identificacion &&
                (!idActual.HasValue || c.IdClienteBeneficiario != idActual.Value));

            if (identificacionDuplicada)
            {
                ModelState.AddModelError("Identificacion", "Ya existe un cliente registrado con esta identificación.");
            }

            if (!string.IsNullOrWhiteSpace(cliente.Email))
            {
                bool emailDuplicado = await _context.ClientesBeneficiarios.AnyAsync(c =>
                    c.Email == cliente.Email &&
                    (!idActual.HasValue || c.IdClienteBeneficiario != idActual.Value));

                if (emailDuplicado)
                {
                    ModelState.AddModelError("Email", "Ya existe un cliente registrado con este correo electrónico.");
                }
            }
        }
        private async Task SincronizarContactosAdicionales(
    int idCliente,
    ClientesBeneficiario formulario)
        {
            var contactos = await _context.ContactosClientesBeneficiarios
                .Where(c => c.IdClienteBeneficiario == idCliente)
                .OrderBy(c => c.IdContactoCliente)
                .ToListAsync();

            SincronizarTipoContacto(
                contactos,
                idCliente,
                "Telefono",
                formulario.OtroTelefono,
                formulario.DescripcionOtroTelefono,
                formulario.OtroTelefono2,
                formulario.DescripcionOtroTelefono2);

            SincronizarTipoContacto(
                contactos,
                idCliente,
                "Correo",
                formulario.OtroCorreo,
                formulario.DescripcionOtroCorreo,
                formulario.OtroCorreo2,
                formulario.DescripcionOtroCorreo2);

            SincronizarTipoContacto(
                contactos,
                idCliente,
                "Direccion",
                formulario.OtraDireccion,
                formulario.DescripcionOtraDireccion,
                formulario.OtraDireccion2,
                formulario.DescripcionOtraDireccion2);
        }
        private void SincronizarTipoContacto(
    List<ContactoClienteBeneficiario> contactos,
    int idCliente,
    string tipoContacto,
    string? valor1,
    string? descripcion1,
    string? valor2,
    string? descripcion2)
        {
            var existentes = contactos
                .Where(c => c.TipoContacto == tipoContacto)
                .OrderBy(c => c.IdContactoCliente)
                .ToList();

            ActualizarOCrearContacto(
                existentes,
                0,
                idCliente,
                tipoContacto,
                valor1,
                descripcion1);

            ActualizarOCrearContacto(
                existentes,
                1,
                idCliente,
                tipoContacto,
                valor2,
                descripcion2);

            // Si hubiera contactos extra de ese tipo, se conservan.
        }

        private void ActualizarOCrearContacto(
    List<ContactoClienteBeneficiario> existentes,
    int posicion,
    int idCliente,
    string tipoContacto,
    string? valor,
    string? descripcion)
        {
            valor = valor?.Trim();
            descripcion = descripcion?.Trim();

            if (posicion < existentes.Count)
            {
                var contacto = existentes[posicion];

                if (string.IsNullOrWhiteSpace(valor))
                {
                    contacto.Activo = false;
                    return;
                }

                contacto.Valor = valor;
                contacto.Descripcion = string.IsNullOrWhiteSpace(descripcion)
                    ? "Adicional"
                    : descripcion;

                contacto.Activo = true;
            }
            else if (!string.IsNullOrWhiteSpace(valor))
            {
                _context.ContactosClientesBeneficiarios.Add(
                    new ContactoClienteBeneficiario
                    {
                        IdClienteBeneficiario = idCliente,
                        TipoContacto = tipoContacto,
                        Valor = valor,
                        Descripcion = string.IsNullOrWhiteSpace(descripcion)
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

       

        private bool ClientesBeneficiarioExists(int id)
        {
            return _context.ClientesBeneficiarios.Any(e => e.IdClienteBeneficiario == id);
        }
    }
}