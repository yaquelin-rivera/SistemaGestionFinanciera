using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using SistemaGestionFinanciera.Data;
using SistemaGestionFinanciera.Models;
using SistemaGestionFinanciera.ViewModels;
using System.Text.Json;

namespace SistemaGestionFinanciera.Controllers
{
    public class AccesoController : Controller
    {
        private readonly SistemaFinancieroContext _context;

        public AccesoController(SistemaFinancieroContext context)
        {
            _context = context;
        }

        // ================================================================
        // LOGIN - GET
        // ================================================================
        [HttpGet]
        public IActionResult Login()
        {
            return View();
        }

        // ================================================================
        // LOGIN - POST
        // ================================================================
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Login(LoginViewModel modelo)
        {
            if (!ModelState.IsValid)
            {
                return View(modelo);
            }

            string correo = modelo.Email
                .Trim()
                .ToLower();

            var usuario = await _context.Usuarios
                .Include(u => u.Rol)
                .FirstOrDefaultAsync(u =>
                    u.Email.ToLower() == correo);

            if (usuario == null)
            {
                ViewBag.LoginMensaje =
                    "El correo electrónico o la contraseña son incorrectos.";

                ViewBag.LoginTipo = "error";

                return View(modelo);
            }

            if (!usuario.Activo)
            {
                ViewBag.LoginMensaje =
                    "El usuario se encuentra inactivo. Comuníquese con el administrador del sistema.";

                ViewBag.LoginTipo = "warning";

                return View(modelo);
            }

            if (usuario.Rol == null || !usuario.Rol.Activo)
            {
                ViewBag.LoginMensaje =
                    "El rol asignado al usuario se encuentra inactivo. Comuníquese con el administrador del sistema.";

                ViewBag.LoginTipo = "warning";

                return View(modelo);
            }

            if (usuario.PasswordHash != modelo.Password)
            {
                ViewBag.LoginMensaje =
                    "El correo electrónico o la contraseña son incorrectos.";

                ViewBag.LoginTipo = "error";

                return View(modelo);
            }

            // ============================================================
            // CARGAR PERMISOS ACTIVOS ASIGNADOS AL ROL
            // ============================================================
            var permisos = await _context.RolesPermisos
                .AsNoTracking()
                .Where(rp =>
                    rp.RolId == usuario.RolId &&
                    rp.Permiso.Activo)
                .Select(rp => rp.Permiso.Nombre)
                .Distinct()
                .OrderBy(nombre => nombre)
                .ToListAsync();

            // ============================================================
            // GUARDAR DATOS DEL USUARIO EN SESIÓN
            // ============================================================
            HttpContext.Session.SetInt32(
                "UsuarioId",
                usuario.IdUsuario);

            HttpContext.Session.SetString(
                "Usuario",
                usuario.Nombre);

            HttpContext.Session.SetInt32(
                "RolId",
                usuario.RolId);

            HttpContext.Session.SetString(
                "Rol",
                usuario.Rol.Nombre);

            HttpContext.Session.SetString(
                "Permisos",
                JsonSerializer.Serialize(permisos));

            // ============================================================
            // REGISTRAR ÚLTIMO ACCESO
            // ============================================================
            usuario.UltimoAcceso = DateTime.Now;

            // ============================================================
            // AUDITORÍA DEL INICIO DE SESIÓN
            // ============================================================
            _context.Auditoria.Add(new Auditorium
            {
                UsuarioId = usuario.IdUsuario,
                Tabla = "Acceso",
                RegistroId = usuario.IdUsuario,
                Accion = "Inicio de sesión",
                Descripcion =
                    $"El usuario inició sesión con el rol {usuario.Rol.Nombre}.",
                Fecha = DateTime.Now
            });

            await _context.SaveChangesAsync();

            return RedirectToAction(
                "Index",
                "Home");
        }

        // ================================================================
        // LOGOUT
        // ================================================================
        [HttpGet]
        public async Task<IActionResult> Logout()
        {
            int? usuarioId =
                HttpContext.Session.GetInt32("UsuarioId");

            if (usuarioId.HasValue)
            {
                var usuario = await _context.Usuarios
                    .AsNoTracking()
                    .FirstOrDefaultAsync(u =>
                        u.IdUsuario == usuarioId.Value);

                if (usuario != null)
                {
                    _context.Auditoria.Add(new Auditorium
                    {
                        UsuarioId = usuario.IdUsuario,
                        Tabla = "Acceso",
                        RegistroId = usuario.IdUsuario,
                        Accion = "Cierre de sesión",
                        Descripcion =
                            "El usuario cerró sesión en el sistema.",
                        Fecha = DateTime.Now
                    });

                    await _context.SaveChangesAsync();
                }
            }

            HttpContext.Session.Clear();

            return RedirectToAction(
                nameof(Login),
                "Acceso");
        }
    }
}