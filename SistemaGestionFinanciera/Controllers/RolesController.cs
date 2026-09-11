using System;
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
    public class RolesController : BaseController
    {
        private readonly SistemaFinancieroContext _context;

        public RolesController(SistemaFinancieroContext context)
        {
            _context = context;
        }
        // INDEX - ROLES
        // ================================================================
        [HttpGet]
        public async Task<IActionResult> Index(
            string? buscar,
            string? estado)
        {
            var roles = await ConstruirConsultaReporte(
                    buscar,
                    estado)
                .OrderByDescending(r => r.IdRol)
                .ToListAsync();

            ViewBag.Buscar = buscar;
            ViewBag.Estado = estado;

            return View(roles);
        }
        // ================================================================
        // REPORTE DE ROLES
        // ================================================================
        [HttpGet]
        public async Task<IActionResult> Reporte(
            string? buscar,
            string? estado)
        {
            var roles = await ConstruirConsultaReporte(
                    buscar,
                    estado)
                .OrderByDescending(r => r.IdRol)
                .ToListAsync();

            ViewBag.TotalRoles =
                roles.Count;

            ViewBag.RolesActivos =
                roles.Count(r => r.Activo);

            ViewBag.RolesInactivos =
                roles.Count(r => !r.Activo);

            ViewBag.Buscar = buscar;
            ViewBag.Estado = estado;

            ViewBag.FechaGeneracion =
                DateTime.Now;
            await RegistrarAuditoria(
    "Vista previa PDF",
    "Roles",
    0,
    "Se generó la vista previa del reporte de roles."
);
            return View(roles);
        }
        // ================================================================
        // EXPORTAR ROLES A EXCEL
        // ================================================================
        [HttpGet]
        public async Task<IActionResult> ExportarExcel(
            string? buscar,
            string? estado)
        {
            var roles = await ConstruirConsultaReporte(
                    buscar,
                    estado)
                .OrderByDescending(r => r.IdRol)
                .ToListAsync();

            int totalRoles =
                roles.Count;

            int rolesActivos =
                roles.Count(r => r.Activo);

            int rolesInactivos =
                roles.Count(r => !r.Activo);

            using var workbook =
                new XLWorkbook();

            var hoja =
                workbook.Worksheets.Add("Roles");

            const int totalColumnas = 4;

            var colorInstitucional =
                XLColor.FromHtml("#0F6B73");

            var colorInstitucionalOscuro =
                XLColor.FromHtml("#0B4F55");

            var colorFondoSuave =
                XLColor.FromHtml("#E7F1F2");

            var colorFondoFiltros =
                XLColor.FromHtml("#F3F7F8");

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
                "Reporte de Roles del Sistema";

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

            var rangoInstitucion =
                hoja.Range(1, 1, 1, totalColumnas);

            rangoInstitucion.Style.Font.Bold = true;
            rangoInstitucion.Style.Font.FontSize = 16;
            rangoInstitucion.Style.Font.FontColor =
                XLColor.FromHtml("#0F5C64");

            rangoInstitucion.Style.Alignment.Horizontal =
                XLAlignmentHorizontalValues.Center;

            var rangoTitulo =
                hoja.Range(2, 1, 2, totalColumnas);

            rangoTitulo.Style.Font.Bold = true;
            rangoTitulo.Style.Font.FontSize = 13;
            rangoTitulo.Style.Font.FontColor = XLColor.Black;

            rangoTitulo.Style.Alignment.Horizontal =
                XLAlignmentHorizontalValues.Center;

            var rangoFecha =
                hoja.Range(3, 1, 3, totalColumnas);

            rangoFecha.Style.Font.FontColor =
                XLColor.FromHtml("#666666");

            rangoFecha.Style.Alignment.Horizontal =
                XLAlignmentHorizontalValues.Center;

            hoja.Row(1).Height = 24;
            hoja.Row(2).Height = 22;
            hoja.Row(3).Height = 20;

            // ============================================================
            // FILTROS
            // ============================================================
            hoja.Cell("A5").Value =
                "Filtros aplicados";

            hoja.Range(
                5,
                1,
                5,
                totalColumnas).Merge();

            var tituloFiltros =
                hoja.Range(5, 1, 5, totalColumnas);

            tituloFiltros.Style.Fill.BackgroundColor =
                colorInstitucional;

            tituloFiltros.Style.Font.FontColor =
                XLColor.White;

            tituloFiltros.Style.Font.Bold = true;

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

            hoja.Cell("A6").Style.Font.Bold = true;
            hoja.Cell("C6").Style.Font.Bold = true;

            var rangoFiltros =
                hoja.Range(6, 1, 6, totalColumnas);

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

            // ============================================================
            // RESUMEN
            // ============================================================
            hoja.Cell("A8").Value =
                "Resumen";

            hoja.Range(
                8,
                1,
                8,
                totalColumnas).Merge();

            var tituloResumen =
                hoja.Range(8, 1, 8, totalColumnas);

            tituloResumen.Style.Fill.BackgroundColor =
                colorInstitucional;

            tituloResumen.Style.Font.FontColor =
                XLColor.White;

            tituloResumen.Style.Font.Bold = true;

            hoja.Cell("A9").Value =
                "Total de roles";

            hoja.Cell("B9").Value =
                totalRoles;

            hoja.Cell("C9").Value =
                "Roles activos";

            hoja.Cell("D9").Value =
                rolesActivos;

            hoja.Cell("A10").Value =
                "Roles inactivos";

            hoja.Cell("B10").Value =
                rolesInactivos;

            hoja.Range("A9:D10")
                .Style.Font.Bold = true;

            hoja.Range("A9:D10")
                .Style.Fill.BackgroundColor =
                colorFondoSuave;

            hoja.Range("A9:D10")
                .Style.Border.OutsideBorder =
                XLBorderStyleValues.Thin;

            hoja.Range("A9:D10")
                .Style.Border.InsideBorder =
                XLBorderStyleValues.Thin;

            // ============================================================
            // ENCABEZADO DE TABLA
            // ============================================================
            const int filaEncabezado = 12;

            string[] encabezados =
            {
        "Nombre",
        "Descripción",
        "Estado",
        "Fecha de creación"
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

            var encabezadoTabla =
                hoja.Range(
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
            int fila =
                filaEncabezado + 1;

            foreach (var rol in roles)
            {
                hoja.Cell(fila, 1).Value =
                    rol.Nombre;

                hoja.Cell(fila, 2).Value =
                    string.IsNullOrWhiteSpace(rol.Descripcion)
                        ? "Sin descripción"
                        : rol.Descripcion;

                hoja.Cell(fila, 3).Value =
                    rol.Activo
                        ? "Activo"
                        : "Inactivo";

                if (rol.FechaCreacion.HasValue)
                {
                    hoja.Cell(fila, 4).Value =
                        rol.FechaCreacion.Value;

                    hoja.Cell(fila, 4)
                        .Style.DateFormat.Format =
                        "dd/MM/yyyy hh:mm AM/PM";
                }
                else
                {
                    hoja.Cell(fila, 4).Value =
                        "Sin fecha";
                }

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

                fila++;
            }

            // ============================================================
            // TABLA
            // ============================================================
            if (roles.Any())
            {
                int filaFinDatos =
                    fila - 1;

                var rangoTabla =
                    hoja.Range(
                        filaEncabezado,
                        1,
                        filaFinDatos,
                        totalColumnas);

                var tabla =
                    rangoTabla.CreateTable(
                        "TablaRoles");

                tabla.Theme = XLTableTheme.None;
                tabla.ShowAutoFilter = true;
                tabla.ShowRowStripes = false;

                var rangoDatos =
                    hoja.Range(
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

                rangoDatos.Style.Alignment.WrapText = true;
            }
            else
            {
                hoja.Cell(fila, 1).Value =
                    "No se encontraron roles con los filtros aplicados.";

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
            int filaTotal =
                fila + 1;

            hoja.Cell(filaTotal, 3).Value =
                "Total de registros:";

            hoja.Cell(filaTotal, 4).Value =
                roles.Count;

            var rangoTotal =
                hoja.Range(
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

            hoja.Cell(filaTotal, 3)
                .Style.Font.Bold = true;

            hoja.Cell(filaTotal, 4)
                .Style.Font.Bold = true;

            // ============================================================
            // CONFIGURACIÓN FINAL
            // ============================================================
            hoja.Column(1).Width = 28;
            hoja.Column(2).Width = 55;
            hoja.Column(3).Width = 18;
            hoja.Column(4).Width = 25;

            hoja.Columns(1, totalColumnas)
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

            using var stream =
                new MemoryStream();

            workbook.SaveAs(stream);

            string nombreArchivo =
                $"Reporte_Roles_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";
            await RegistrarAuditoria(
    "Exportar Excel",
    "Roles",
    0,
    "Se exportó el reporte de roles a Excel."
);
            return File(
                stream.ToArray(),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                nombreArchivo);
        }
        // DETAILS
        public async Task<IActionResult> Details(int? id)
        {
            if (id == null) return NotFound();

            var role = await _context.Roles
                .FirstOrDefaultAsync(r => r.IdRol == id);

            if (role == null) return NotFound();

            return View(role);
        }

        // CREATE GET
        public IActionResult Create()
        {
            return View(new Role
            {
                Activo = true,
                FechaCreacion = DateTime.Now
            });
        }

        // CREATE POST
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Create(
            [Bind("Nombre,Descripcion")] Role role)
        {
            NormalizarDatos(role);
            await ValidarRol(role);

            if (ModelState.IsValid)
            {
                role.Activo = true;
                role.FechaCreacion = DateTime.Now;

                _context.Roles.Add(role);
                await _context.SaveChangesAsync();

                await RegistrarAuditoria(
                    "Crear",
                    "Roles",
                    role.IdRol,
                    $"Se creó el rol: {role.Nombre}.");

                TempData["MensajeExito"] = "Rol registrado correctamente.";
                return RedirectToAction(nameof(Index));
            }

            return View(role);
        }

        // EDIT GET
        public async Task<IActionResult> Edit(int? id)
        {
            if (id == null) return NotFound();

            var role = await _context.Roles.FindAsync(id);

            if (role == null) return NotFound();

            if (!role.Activo)
            {
                TempData["MensajeError"] = "El rol está inactivo. Para modificarlo primero debe activarlo.";
                return RedirectToAction(nameof(Index));
            }

            return View(role);
        }

        // EDIT POST
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Edit(
            int id,
            [Bind("IdRol,Nombre,Descripcion")] Role role)
        {
            if (id != role.IdRol) return NotFound();

            var rolActual = await _context.Roles
                .AsNoTracking()
                .FirstOrDefaultAsync(r => r.IdRol == id);

            if (rolActual == null) return NotFound();

            if (!rolActual.Activo)
            {
                TempData["MensajeError"] = "El rol está inactivo. Para modificarlo primero debe activarlo.";
                return RedirectToAction(nameof(Index));
            }

            NormalizarDatos(role);
            await ValidarRol(role);

            if (ModelState.IsValid)
            {
                try
                {
                    role.Activo = rolActual.Activo;
                    role.FechaCreacion = rolActual.FechaCreacion;

                    _context.Roles.Update(role);
                    await _context.SaveChangesAsync();

                    await RegistrarAuditoria(
                        "Modificar",
                        "Roles",
                        role.IdRol,
                        $"Se modificó el rol: {role.Nombre}.");

                    TempData["MensajeExito"] = "Rol modificado correctamente.";
                    return RedirectToAction(nameof(Index));
                }
                catch (DbUpdateConcurrencyException)
                {
                    if (!RoleExists(role.IdRol)) return NotFound();
                    throw;
                }
            }

            return View(role);
        }

        // DESACTIVAR
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> DeleteConfirmed(int id)
        {
            var role = await _context.Roles.FindAsync(id);

            if (role == null) return NotFound();

            if (!role.Activo)
            {
                TempData["MensajeError"] = "El rol ya se encuentra inactivo.";
                return RedirectToAction(nameof(Index));
            }

            bool tieneUsuarios = await _context.Usuarios
                .AnyAsync(u => u.RolId == id);

            // Se permite desactivar aunque tenga usuarios,
            // pero no se elimina físicamente para conservar historial.
            role.Activo = false;

            _context.Roles.Update(role);
            await _context.SaveChangesAsync();

            await RegistrarAuditoria(
                "Desactivar",
                "Roles",
                role.IdRol,
                tieneUsuarios
                    ? $"Se desactivó el rol con usuarios asociados: {role.Nombre}."
                    : $"Se desactivó el rol: {role.Nombre}.");

            TempData["MensajeExito"] = "Rol desactivado correctamente.";
            return RedirectToAction(nameof(Index));
        }

        // ACTIVAR
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Activar(int id)
        {
            var role = await _context.Roles.FindAsync(id);

            if (role == null) return NotFound();

            if (role.Activo)
            {
                TempData["MensajeError"] = "El rol ya se encuentra activo.";
                return RedirectToAction(nameof(Index));
            }

            role.Activo = true;

            _context.Roles.Update(role);
            await _context.SaveChangesAsync();

            await RegistrarAuditoria(
                "Activar",
                "Roles",
                role.IdRol,
                $"Se activó el rol: {role.Nombre}.");

            TempData["MensajeExito"] = "Rol activado correctamente.";
            return RedirectToAction(nameof(Index));
        }

        // MÉTODOS PRIVADOS
        // ================================================================
        // CONSULTA COMPARTIDA PARA INDEX, PDF Y EXCEL
        // ================================================================
        private IQueryable<Role> ConstruirConsultaReporte(
            string? buscar,
            string? estado)
        {
            var consulta = _context.Roles
                .AsNoTracking()
                .AsQueryable();

            if (!string.IsNullOrWhiteSpace(buscar))
            {
                buscar = buscar.Trim();

                consulta = consulta.Where(r =>
                    r.Nombre.Contains(buscar) ||
                    (
                        r.Descripcion != null &&
                        r.Descripcion.Contains(buscar)
                    ));
            }

            if (!string.IsNullOrWhiteSpace(estado))
            {
                if (estado.Equals(
                    "Activo",
                    StringComparison.OrdinalIgnoreCase))
                {
                    consulta = consulta.Where(r =>
                        r.Activo);
                }
                else if (estado.Equals(
                    "Inactivo",
                    StringComparison.OrdinalIgnoreCase))
                {
                    consulta = consulta.Where(r =>
                        !r.Activo);
                }
            }

            return consulta;
        }
        private void NormalizarDatos(Role role)
        {
            role.Nombre = role.Nombre?.Trim() ?? "";
            role.Descripcion = role.Descripcion?.Trim() ?? "";
        }

        private async Task ValidarRol(Role role)
        {
            if (string.IsNullOrWhiteSpace(role.Nombre))
            {
                ModelState.AddModelError("Nombre", "El nombre del rol es obligatorio.");
            }
            else if (role.Nombre.Length < 3)
            {
                ModelState.AddModelError("Nombre", "El nombre debe tener al menos 3 caracteres.");
            }
            else if (role.Nombre.Length > 100)
            {
                ModelState.AddModelError("Nombre", "El nombre no puede superar los 100 caracteres.");
            }
            else if (!System.Text.RegularExpressions.Regex.IsMatch(
                role.Nombre,
                @"^[A-Za-zÁÉÍÓÚáéíóúÑñ0-9\s-]+$"))
            {
                ModelState.AddModelError("Nombre", "El nombre solo puede contener letras, números, espacios y guiones.");
            }

            bool nombreDuplicado = await _context.Roles
                .AnyAsync(r =>
                    r.Nombre.ToLower() == role.Nombre.ToLower()
                    && r.IdRol != role.IdRol);

            if (nombreDuplicado)
            {
                ModelState.AddModelError("Nombre", "Ya existe un rol registrado con ese nombre.");
            }

            if (!string.IsNullOrWhiteSpace(role.Descripcion))
            {
                if (role.Descripcion.Length > 250)
                {
                    ModelState.AddModelError("Descripcion", "La descripción no puede superar los 250 caracteres.");
                }
                else if (!System.Text.RegularExpressions.Regex.IsMatch(
                    role.Descripcion,
                    @"^[A-Za-zÁÉÍÓÚáéíóúÑñ0-9\s.,;:()/-]+$"))
                {
                    ModelState.AddModelError("Descripcion", "La descripción contiene caracteres no permitidos.");
                }
            }
        }

        private async Task RegistrarAuditoria(
            string accion,
            string tabla,
            int registroId,
            string descripcion)
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

        private bool RoleExists(int id)
        {
            return _context.Roles.Any(e => e.IdRol == id);
        }
    }
}
