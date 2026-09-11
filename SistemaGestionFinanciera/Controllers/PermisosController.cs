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
    public class PermisosController : BaseController
    {
        private readonly SistemaFinancieroContext _context;

        public PermisosController(SistemaFinancieroContext context)
        {
            _context = context;
        }

        public async Task<IActionResult> Index(
    string buscar,
    string estado)
        {
            var permisos =
                ConstruirConsultaReporte(
                    buscar,
                    estado);

            var lista =
                await permisos
                    .ToListAsync();

            ViewBag.Buscar = buscar;
            ViewBag.Estado = estado;

            ViewBag.TotalPermisos =
                lista.Count;

            ViewBag.PermisosActivos =
                lista.Count(p => p.Activo);

            ViewBag.PermisosInactivos =
                lista.Count(p => !p.Activo);

            return View(lista);
        }
        public async Task<IActionResult> Reporte(
    string buscar,
    string estado)
        {
            var lista =
                await ConstruirConsultaReporte(
                    buscar,
                    estado)
                .ToListAsync();

            ViewBag.Buscar = buscar;
            ViewBag.Estado = estado;

            ViewBag.TotalPermisos =
                lista.Count;

            ViewBag.PermisosActivos =
                lista.Count(p => p.Activo);

            ViewBag.PermisosInactivos =
                lista.Count(p => !p.Activo);

            ViewBag.FechaGeneracion =
                DateTime.Now;
            await RegistrarAuditoria(
    "Vista previa PDF",
    "Permisos",
    0,
    "Se generó la vista previa del reporte de permisos."
);
            return View(lista);
        }
        public async Task<IActionResult> ExportarExcel(
    string buscar,
    string estado)
        {
            var lista =
                await ConstruirConsultaReporte(
                    buscar,
                    estado)
                .ToListAsync();

            using var workbook = new XLWorkbook();

            var ws =
                workbook.Worksheets.Add("Permisos");

            int fila = 1;

            ws.Cell(fila, 1).Value =
                "Unión Cantonal de Asociaciones de Desarrollo de Santa Ana";

            ws.Range(fila, 1, fila, 4).Merge();

            ws.Cell(fila, 1).Style.Font.Bold = true;
            ws.Cell(fila, 1).Style.Font.FontSize = 16;
            ws.Cell(fila, 1).Style.Font.FontColor = XLColor.DarkCyan;

            fila++;

            ws.Cell(fila, 1).Value =
                "Reporte de Permisos";

            ws.Range(fila, 1, fila, 4).Merge();

            ws.Cell(fila, 1).Style.Font.Bold = true;
            ws.Cell(fila, 1).Style.Font.FontSize = 14;

            fila++;

            ws.Cell(fila, 1).Value =
                $"Fecha: {DateTime.Now:dd/MM/yyyy HH:mm}";

            ws.Range(fila, 1, fila, 4).Merge();

            ws.Cell(fila, 1).Style.Font.FontColor =
                XLColor.Gray;

            fila += 2;

            ws.Cell(fila, 1).Value = "Buscar";
            ws.Cell(fila, 2).Value =
                string.IsNullOrWhiteSpace(buscar)
                    ? "Todos"
                    : buscar;

            fila++;

            ws.Cell(fila, 1).Value = "Estado";
            ws.Cell(fila, 2).Value =
                string.IsNullOrWhiteSpace(estado)
                    ? "Todos"
                    : estado;

            ws.Range(fila - 1, 1, fila, 2)
                .Style.Fill.BackgroundColor =
                XLColor.DarkCyan;

            ws.Range(fila - 1, 1, fila, 2)
                .Style.Font.FontColor =
                XLColor.White;

            fila += 2;

            ws.Cell(fila, 1).Value = "Nombre";
            ws.Cell(fila, 2).Value = "Descripción";
            ws.Cell(fila, 3).Value = "Estado";
            ws.Cell(fila, 4).Value = "Fecha creación";

            ws.Range(fila, 1, fila, 4)
                .Style.Fill.BackgroundColor =
                XLColor.DarkCyan;

            ws.Range(fila, 1, fila, 4)
                .Style.Font.FontColor =
                XLColor.White;

            ws.Range(fila, 1, fila, 4)
                .Style.Font.Bold = true;

            fila++;

            foreach (var item in lista)
            {
                ws.Cell(fila, 1).Value = item.Nombre;

                ws.Cell(fila, 2).Value =
                    string.IsNullOrWhiteSpace(item.Descripcion)
                        ? "Sin descripción"
                        : item.Descripcion;

                ws.Cell(fila, 3).Value =
                    item.Activo
                        ? "Activo"
                        : "Inactivo";

                ws.Cell(fila, 4).Value =
                    item.FechaCreacion;

                ws.Cell(fila, 4)
                    .Style.DateFormat.Format =
                    "dd/MM/yyyy HH:mm";

                fila++;
            }

            fila++;

            ws.Cell(fila, 1).Value = "Resumen";

            ws.Range(fila, 1, fila, 2)
                .Merge();

            ws.Range(fila, 1, fila, 2)
                .Style.Fill.BackgroundColor =
                XLColor.DarkCyan;

            ws.Range(fila, 1, fila, 2)
                .Style.Font.FontColor =
                XLColor.White;

            fila++;

            ws.Cell(fila, 1).Value = "Total";
            ws.Cell(fila, 2).Value = lista.Count;

            fila++;

            ws.Cell(fila, 1).Value = "Activos";
            ws.Cell(fila, 2).Value =
                lista.Count(x => x.Activo);

            fila++;

            ws.Cell(fila, 1).Value = "Inactivos";
            ws.Cell(fila, 2).Value =
                lista.Count(x => !x.Activo);

            ws.Columns().AdjustToContents();

            ws.SheetView.FreezeRows(7);

            using var stream =
                new MemoryStream();

            workbook.SaveAs(stream);
            await RegistrarAuditoria(
    "Exportar Excel",
    "Permisos",
    0,
    "Se exportó el reporte de permisos a Excel."
);
            return File(
                stream.ToArray(),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                $"ReportePermisos_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");
        }
        // DETAILS
        public async Task<IActionResult> Details(int? id)
        {
            if (id == null) return NotFound();

            var permiso = await _context.Permisos
                .FirstOrDefaultAsync(p => p.IdPermiso == id);

            if (permiso == null) return NotFound();

            return View(permiso);
        }

        // CREATE GET
        public IActionResult Create()
        {
            return View(new Permiso
            {
                Activo = true,
                FechaCreacion = DateTime.Now
            });
        }

        // CREATE POST
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Create(
            [Bind("Nombre,Descripcion")] Permiso permiso)
        {
            NormalizarDatos(permiso);
            await ValidarPermiso(permiso);

            if (ModelState.IsValid)
            {
                permiso.Activo = true;
                permiso.FechaCreacion = DateTime.Now;

                _context.Permisos.Add(permiso);
                await _context.SaveChangesAsync();

                await RegistrarAuditoria(
                    "Crear",
                    "Permisos",
                    permiso.IdPermiso,
                    $"Se creó el permiso: {permiso.Nombre}.");

                TempData["MensajeExito"] = "Permiso registrado correctamente.";
                return RedirectToAction(nameof(Index));
            }

            return View(permiso);
        }

        // EDIT GET
        public async Task<IActionResult> Edit(int? id)
        {
            if (id == null) return NotFound();

            var permiso = await _context.Permisos.FindAsync(id);

            if (permiso == null) return NotFound();

            if (!permiso.Activo)
            {
                TempData["MensajeError"] = "El permiso está inactivo. Para modificarlo primero debe activarlo.";
                return RedirectToAction(nameof(Index));
            }

            return View(permiso);
        }

        // EDIT POST
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Edit(
            int id,
            [Bind("IdPermiso,Nombre,Descripcion")] Permiso permiso)
        {
            if (id != permiso.IdPermiso) return NotFound();

            var permisoActual = await _context.Permisos
                .AsNoTracking()
                .FirstOrDefaultAsync(p => p.IdPermiso == id);

            if (permisoActual == null) return NotFound();

            if (!permisoActual.Activo)
            {
                TempData["MensajeError"] = "El permiso está inactivo. Para modificarlo primero debe activarlo.";
                return RedirectToAction(nameof(Index));
            }

            NormalizarDatos(permiso);
            await ValidarPermiso(permiso);

            if (ModelState.IsValid)
            {
                try
                {
                    permiso.Activo = permisoActual.Activo;
                    permiso.FechaCreacion = permisoActual.FechaCreacion;

                    _context.Permisos.Update(permiso);
                    await _context.SaveChangesAsync();

                    await RegistrarAuditoria(
                        "Modificar",
                        "Permisos",
                        permiso.IdPermiso,
                        $"Se modificó el permiso: {permiso.Nombre}.");

                    TempData["MensajeExito"] = "Permiso modificado correctamente.";
                    return RedirectToAction(nameof(Index));
                }
                catch (DbUpdateConcurrencyException)
                {
                    if (!PermisoExists(permiso.IdPermiso)) return NotFound();
                    throw;
                }
            }

            return View(permiso);
        }

        // DESACTIVAR
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> DeleteConfirmed(int id)
        {
            var permiso = await _context.Permisos.FindAsync(id);

            if (permiso == null) return NotFound();

            if (!permiso.Activo)
            {
                TempData["MensajeError"] = "El permiso ya se encuentra inactivo.";
                return RedirectToAction(nameof(Index));
            }

            permiso.Activo = false;

            _context.Permisos.Update(permiso);
            await _context.SaveChangesAsync();

            await RegistrarAuditoria(
                "Desactivar",
                "Permisos",
                permiso.IdPermiso,
                $"Se desactivó el permiso: {permiso.Nombre}.");

            TempData["MensajeExito"] = "Permiso desactivado correctamente.";
            return RedirectToAction(nameof(Index));
        }

        // ACTIVAR
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Activar(int id)
        {
            var permiso = await _context.Permisos.FindAsync(id);

            if (permiso == null) return NotFound();

            if (permiso.Activo)
            {
                TempData["MensajeError"] = "El permiso ya se encuentra activo.";
                return RedirectToAction(nameof(Index));
            }

            permiso.Activo = true;

            _context.Permisos.Update(permiso);
            await _context.SaveChangesAsync();

            await RegistrarAuditoria(
                "Activar",
                "Permisos",
                permiso.IdPermiso,
                $"Se activó el permiso: {permiso.Nombre}.");

            TempData["MensajeExito"] = "Permiso activado correctamente.";
            return RedirectToAction(nameof(Index));
        }

        // MÉTODOS PRIVADOS
        private IQueryable<Permiso> ConstruirConsultaReporte(
    string buscar,
    string estado)
        {
            var consulta =
                _context.Permisos
                .AsNoTracking()
                .AsQueryable();

            if (!string.IsNullOrWhiteSpace(buscar))
            {
                consulta = consulta.Where(p =>

                    p.Nombre.Contains(buscar)

                    ||

                    (p.Descripcion != null &&
                     p.Descripcion.Contains(buscar))
                );
            }

            if (estado == "Activo")
            {
                consulta = consulta.Where(p => p.Activo);
            }
            else if (estado == "Inactivo")
            {
                consulta = consulta.Where(p => !p.Activo);
            }

            return consulta
                .OrderByDescending(p => p.IdPermiso);
        }
        private void NormalizarDatos(Permiso permiso)
        {
            permiso.Nombre = permiso.Nombre?.Trim() ?? "";
            permiso.Descripcion = permiso.Descripcion?.Trim() ?? "";
        }

        private async Task ValidarPermiso(Permiso permiso)
        {
            if (string.IsNullOrWhiteSpace(permiso.Nombre))
            {
                ModelState.AddModelError("Nombre", "El nombre del permiso es obligatorio.");
            }
            else if (permiso.Nombre.Length < 3)
            {
                ModelState.AddModelError("Nombre", "El nombre debe tener al menos 3 caracteres.");
            }
            else if (permiso.Nombre.Length > 100)
            {
                ModelState.AddModelError("Nombre", "El nombre no puede superar los 100 caracteres.");
            }
            else if (!System.Text.RegularExpressions.Regex.IsMatch(
                permiso.Nombre,
                @"^[A-Za-zÁÉÍÓÚáéíóúÑñ0-9\s-]+$"))
            {
                ModelState.AddModelError("Nombre", "El nombre solo puede contener letras, números, espacios y guiones.");
            }

            bool nombreDuplicado = await _context.Permisos
                .AnyAsync(p =>
                    p.Nombre.ToLower() == permiso.Nombre.ToLower()
                    && p.IdPermiso != permiso.IdPermiso);

            if (nombreDuplicado)
            {
                ModelState.AddModelError("Nombre", "Ya existe un permiso registrado con ese nombre.");
            }

            if (!string.IsNullOrWhiteSpace(permiso.Descripcion))
            {
                if (permiso.Descripcion.Length > 250)
                {
                    ModelState.AddModelError("Descripcion", "La descripción no puede superar los 250 caracteres.");
                }
                else if (!System.Text.RegularExpressions.Regex.IsMatch(
                    permiso.Descripcion,
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

        private bool PermisoExists(int id)
        {
            return _context.Permisos.Any(e => e.IdPermiso == id);
        }
    }
}