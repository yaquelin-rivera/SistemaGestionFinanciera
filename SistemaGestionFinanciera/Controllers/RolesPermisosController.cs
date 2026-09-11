using System;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.EntityFrameworkCore;
using SistemaGestionFinanciera.Data;
using SistemaGestionFinanciera.Models;

namespace SistemaGestionFinanciera.Controllers
{
    public class RolesPermisosController : BaseController
    {
        private readonly SistemaFinancieroContext _context;

        public RolesPermisosController(SistemaFinancieroContext context)
        {
            _context = context;
        }

        // PANTALLA DE ASIGNACIÓN
        public async Task<IActionResult> Index(int? rolId)
        {
            if (string.IsNullOrEmpty(HttpContext.Session.GetString("Usuario")))
            {
                return RedirectToAction("Login", "Acceso");
            }

            var roles = await _context.Roles
                .Where(r => r.Activo)
                .OrderBy(r => r.Nombre)
                .ToListAsync();

            var permisos = await _context.Permisos
                .Where(p => p.Activo)
                .OrderBy(p => p.Nombre)
                .ToListAsync();

            var permisosAsignados = new int[0];
            string nombreRolSeleccionado = "";

            if (rolId.HasValue)
            {
                var rolSeleccionado = roles.FirstOrDefault(r => r.IdRol == rolId.Value);

                if (rolSeleccionado != null)
                {
                    nombreRolSeleccionado = rolSeleccionado.Nombre;
                }

                permisosAsignados = await _context.RolesPermisos
                    .Where(rp => rp.RolId == rolId.Value)
                    .Select(rp => rp.PermisoId)
                    .ToArrayAsync();
            }

            ViewBag.RolId = rolId;
            ViewBag.Roles = new SelectList(roles, "IdRol", "Nombre", rolId);
            ViewBag.Permisos = permisos;
            ViewBag.PermisosAsignados = permisosAsignados;
            ViewBag.TotalRoles = roles.Count;
            ViewBag.TotalPermisos = permisos.Count;
            ViewBag.TotalAsignados = permisosAsignados.Length;

            ViewBag.NombreRolSeleccionado = nombreRolSeleccionado;

            return View();
        }

        // GUARDAR ASIGNACIÓN
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Guardar(int rolId, int[] permisosSeleccionados)
        {
            if (rolId <= 0)
            {
                TempData["MensajeError"] = "Debe seleccionar un rol válido.";
                return RedirectToAction(nameof(Index));
            }

            var rol = await _context.Roles.FirstOrDefaultAsync(r => r.IdRol == rolId);

            if (rol == null)
            {
                TempData["MensajeError"] = "El rol seleccionado no existe.";
                return RedirectToAction(nameof(Index));
            }

            if (!rol.Activo)
            {
                TempData["MensajeError"] = "No se pueden asignar permisos a un rol inactivo.";
                return RedirectToAction(nameof(Index));
            }

            permisosSeleccionados ??= new int[0];

            var permisosValidos = await _context.Permisos
                .Where(p => p.Activo && permisosSeleccionados.Contains(p.IdPermiso))
                .Select(p => p.IdPermiso)
                .ToListAsync();

            var asignacionesActuales = await _context.RolesPermisos
                .Where(rp => rp.RolId == rolId)
                .ToListAsync();

            _context.RolesPermisos.RemoveRange(asignacionesActuales);

            foreach (var permisoId in permisosValidos.Distinct())
            {
                _context.RolesPermisos.Add(new RolesPermiso
                {
                    RolId = rolId,
                    PermisoId = permisoId
                });
            }

            await _context.SaveChangesAsync();

            await RegistrarAuditoria(
                "Asignar permisos",
                "RolesPermisos",
                rolId,
                $"Se actualizaron los permisos asignados al rol: {rol.Nombre}. Total permisos asignados: {permisosValidos.Count}.");

            TempData["MensajeExito"] = "Permisos asignados correctamente.";
            return RedirectToAction(nameof(Index), new { rolId });
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
    }
}
