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
using System.IO;

namespace SistemaGestionFinanciera.Controllers
{
    public class UsuariosController : BaseController
    {
        private readonly SistemaFinancieroContext _context;

        public UsuariosController(SistemaFinancieroContext context)
        {
            _context = context;
        }

        // INDEX - USUARIOS
        // ================================================================
        [HttpGet]
        public async Task<IActionResult> Index(
            string? buscar,
            int? rolId,
            string? estado)
        {
            if (string.IsNullOrEmpty(
                HttpContext.Session.GetString("Usuario")))
            {
                return RedirectToAction(
                    "Login",
                    "Acceso");
            }

            var consulta = ConstruirConsultaReporte(
                buscar,
                rolId,
                estado);

            var usuarios = await consulta
                .OrderByDescending(u => u.IdUsuario)
                .ToListAsync();

            ViewBag.Roles = await _context.Roles
                .AsNoTracking()
                .OrderBy(r => r.Nombre)
                .ToListAsync();

            ViewBag.Buscar = buscar;
            ViewBag.RolId = rolId;
            ViewBag.Estado = estado;

            return View(usuarios);
        }
        // ================================================================
        // REPORTE DE USUARIOS
        // ================================================================
        [HttpGet]
        public async Task<IActionResult> Reporte(
            string? buscar,
            int? rolId,
            string? estado)
        {
            var usuarios = await ConstruirConsultaReporte(
                    buscar,
                    rolId,
                    estado)
                .OrderByDescending(u => u.IdUsuario)
                .ToListAsync();

            ViewBag.TotalUsuarios =
                usuarios.Count;

            ViewBag.UsuariosActivos =
                usuarios.Count(u => u.Activo);

            ViewBag.UsuariosInactivos =
                usuarios.Count(u => !u.Activo);

            ViewBag.Administradores =
                usuarios.Count(u =>
                    u.Rol != null &&
                    u.Rol.Nombre == "Administrador");

            ViewBag.Buscar = buscar;
            ViewBag.RolId = rolId;
            ViewBag.Estado = estado;

            ViewBag.NombreRol = rolId.HasValue
                ? await _context.Roles
                    .Where(r => r.IdRol == rolId.Value)
                    .Select(r => r.Nombre)
                    .FirstOrDefaultAsync()
                : null;

            ViewBag.FechaGeneracion =
                DateTime.Now;
            await RegistrarAuditoria(
    "Vista previa PDF",
    "Usuarios",
    0,
    "Se generó la vista previa del reporte de usuarios."
);
            return View(usuarios);
        }
        // ================================================================
        // EXPORTAR USUARIOS A EXCEL
        // ================================================================
        [HttpGet]
        public async Task<IActionResult> ExportarExcel(
            string? buscar,
            int? rolId,
            string? estado)
        {
            var usuarios = await ConstruirConsultaReporte(
                    buscar,
                    rolId,
                    estado)
                .OrderByDescending(u => u.IdUsuario)
                .ToListAsync();

            string? nombreRol = null;

            if (rolId.HasValue)
            {
                nombreRol = await _context.Roles
                    .Where(r => r.IdRol == rolId.Value)
                    .Select(r => r.Nombre)
                    .FirstOrDefaultAsync();
            }

            int totalUsuarios =
                usuarios.Count;

            int usuariosActivos =
                usuarios.Count(u => u.Activo);

            int usuariosInactivos =
                usuarios.Count(u => !u.Activo);

            int administradores =
                usuarios.Count(u =>
                    u.Rol != null &&
                    u.Rol.Nombre == "Administrador");

            using var workbook =
                new XLWorkbook();

            var hoja =
                workbook.Worksheets.Add("Usuarios");

            const int totalColumnas = 6;

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
                "Reporte de Usuarios del Sistema";

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
            // FILTROS APLICADOS
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

            hoja.Cell("C6").Value = "Rol:";

            hoja.Cell("D6").Value =
                string.IsNullOrWhiteSpace(nombreRol)
                    ? "Todos"
                    : nombreRol;

            hoja.Cell("E6").Value = "Estado:";

            hoja.Cell("F6").Value =
                string.IsNullOrWhiteSpace(estado)
                    ? "Todos"
                    : estado;

            foreach (int columna in new[] { 1, 3, 5 })
            {
                hoja.Cell(6, columna)
                    .Style.Font.Bold = true;
            }

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
                "Total de usuarios";

            hoja.Cell("B9").Value =
                totalUsuarios;

            hoja.Cell("C9").Value =
                "Activos";

            hoja.Cell("D9").Value =
                usuariosActivos;

            hoja.Cell("E9").Value =
                "Inactivos";

            hoja.Cell("F9").Value =
                usuariosInactivos;

            hoja.Cell("A10").Value =
                "Administradores";

            hoja.Cell("B10").Value =
                administradores;

            hoja.Range("A9:F10")
                .Style.Font.Bold = true;

            hoja.Range("A9:F10")
                .Style.Fill.BackgroundColor =
                colorFondoSuave;

            hoja.Range("A9:F10")
                .Style.Border.OutsideBorder =
                XLBorderStyleValues.Thin;

            hoja.Range("A9:F10")
                .Style.Border.InsideBorder =
                XLBorderStyleValues.Thin;

            // ============================================================
            // ENCABEZADO DE TABLA
            // ============================================================
            const int filaEncabezado = 12;

            string[] encabezados =
            {
        "Nombre",
        "Correo electrónico",
        "Rol",
        "Último acceso",
        "Fecha de creación",
        "Estado"
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

            foreach (var usuario in usuarios)
            {
                hoja.Cell(fila, 1).Value =
                    usuario.Nombre;

                hoja.Cell(fila, 2).Value =
                    usuario.Email;

                hoja.Cell(fila, 3).Value =
                    usuario.Rol?.Nombre ?? "Sin rol";

                if (usuario.UltimoAcceso.HasValue)
                {
                    hoja.Cell(fila, 4).Value =
                        usuario.UltimoAcceso.Value;

                    hoja.Cell(fila, 4)
                        .Style.DateFormat.Format =
                        "dd/MM/yyyy hh:mm AM/PM";
                }
                else
                {
                    hoja.Cell(fila, 4).Value =
                        "Sin acceso";
                }

                if (usuario.FechaCreacion.HasValue)
                {
                    hoja.Cell(fila, 5).Value =
                        usuario.FechaCreacion.Value;

                    hoja.Cell(fila, 5)
                        .Style.DateFormat.Format =
                        "dd/MM/yyyy";
                }
                else
                {
                    hoja.Cell(fila, 5).Value =
                        "Sin fecha";
                }

                hoja.Cell(fila, 6).Value =
                    usuario.Activo
                        ? "Activo"
                        : "Inactivo";

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
            if (usuarios.Any())
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
                        "TablaUsuarios");

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
                    "No se encontraron usuarios con los filtros aplicados.";

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

            hoja.Cell(filaTotal, 5).Value =
                "Total de registros:";

            hoja.Cell(filaTotal, 6).Value =
                usuarios.Count;

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

            hoja.Cell(filaTotal, 5)
                .Style.Font.Bold = true;

            hoja.Cell(filaTotal, 6)
                .Style.Font.Bold = true;

            // ============================================================
            // CONFIGURACIÓN FINAL
            // ============================================================
            hoja.Column(1).Width = 27;
            hoja.Column(2).Width = 35;
            hoja.Column(3).Width = 23;
            hoja.Column(4).Width = 24;
            hoja.Column(5).Width = 22;
            hoja.Column(6).Width = 15;

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
                $"Reporte_Usuarios_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";
            await RegistrarAuditoria(
    "Exportar Excel",
    "Usuarios",
    0,
    "Se exportó el reporte de usuarios a Excel."
);
            return File(
                stream.ToArray(),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                nombreArchivo);
        }
        // GET: Usuarios/Detalle/5
        public async Task<IActionResult> Details(int? id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var usuario = await _context.Usuarios
                .Include(u => u.Rol)
                .FirstOrDefaultAsync(m => m.IdUsuario == id);
            if (usuario == null)
            {
                return NotFound();
            }

            return View(usuario);
        }

        // GET: Usuarios/Crear
        public IActionResult Create()
        {
            ViewData["RolId"] = new SelectList(
                _context.Roles
                    .Where(r => r.Activo)
                    .OrderBy(r => r.Nombre),
                "IdRol",
                "Nombre"
            );

            return View();
        }

        // POST: Usuarios/Crear nuevo usuario en el sistema 
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Create([Bind("Nombre,Email,PasswordHash,RolId")] Usuario usuario,string? ConfirmarPassword)
        {
            // Asignar valores iniciales al usuario
            usuario.Activo = true;
            usuario.FechaCreacion = DateTime.Now;
            usuario.UltimoAcceso = null;

            // Eliminar espacios al inicio y al final
            usuario.Nombre = usuario.Nombre?.Trim() ?? "";
            usuario.Email = usuario.Email?.Trim().ToLower() ?? "";


            ModelState.Remove("Activo");
            ModelState.Remove("FechaCreacion");
            ModelState.Remove("UltimoAcceso");
            ModelState.Remove("Rol");
            ModelState.Remove("RolId");   // <-- Agregar
            ModelState.Remove("ConfirmarPassword");

            if (usuario.RolId <= 0)
            {
                ModelState.AddModelError("RolId", "Debe seleccionar un rol.");
            }

            if (string.IsNullOrWhiteSpace(usuario.Nombre))
            { 
                ModelState.AddModelError("Nombre", "El nombre es obligatorio.");
            }
            else if (usuario.Nombre.Length < 3)
            {
                ModelState.AddModelError("Nombre", "El nombre debe tener al menos 3 caracteres.");
            }
          

            if (string.IsNullOrWhiteSpace(ConfirmarPassword))
            { 
                ModelState.AddModelError("ConfirmarPassword",
                    "Debe confirmar la contraseña.");
            }
            else if (usuario.PasswordHash != ConfirmarPassword) 
            {
                ModelState.AddModelError("ConfirmarPassword",
                    "Las contraseñas no coinciden.");
            }

            // validar correo repetido
            if (ModelState.IsValid)
            {
                bool correoExiste = await _context.Usuarios
                    .AnyAsync(u => u.Email == usuario.Email);

                if (correoExiste)
                {
                    ModelState.AddModelError("Email",
                        "Ya existe un usuario registrado con ese correo.");
                }
                else
                {
                    // Guardar usuario en la base de datos y muestra msj de comfirmacion
                    _context.Add(usuario);
                    await _context.SaveChangesAsync();

                    // Registrar auditoría
                    var usuarioSesionNombre = HttpContext.Session.GetString("Usuario");

                    var usuarioSesion = await _context.Usuarios
                        .FirstOrDefaultAsync(u => u.Nombre == usuarioSesionNombre);

                    if (usuarioSesion != null)
                    {
                        var auditoria = new Auditorium
                        {
                            UsuarioId = usuarioSesion.IdUsuario,
                            Tabla = "Usuarios",
                            RegistroId = usuario.IdUsuario,
                            Accion = "Crear",
                            Descripcion = "Se creó el usuario: " + usuario.Nombre,
                            Fecha = DateTime.Now
                        };
                        // nombre de DbSet
                        _context.Auditoria.Add(auditoria);
                        await _context.SaveChangesAsync();
                    }
                    TempData["MensajeExito"] = "Usuario creado correctamente.";

                    return RedirectToAction(nameof(Index));
                }
            }

            ViewData["RolId"] = new SelectList(
      _context.Roles
          .Where(r => r.Activo)
          .OrderBy(r => r.Nombre),
      "IdRol",
      "Nombre",
      usuario.RolId
  );
            return View(usuario);
        }

        // GET: Usuarios/modificar o editar
        public async Task<IActionResult> Edit(int? id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var usuario = await _context.Usuarios.FindAsync(id);

            if (usuario == null)
            {
                return NotFound();
            }

            if (!usuario.Activo)
            {
                TempData["MensajeError"] = "No puede modificar usuarios inactivos. Debe activarlo primero.";
                return RedirectToAction(nameof(Index));
            }

            ViewData["RolId"] = new SelectList(
                _context.Roles
                    .Where(r => r.Activo || r.IdRol == usuario.RolId)
                    .OrderBy(r => r.Nombre),
                "IdRol",
                "Nombre",
                usuario.RolId
            );

            return View(usuario);
        }

        // POST: Usuarios/modificar o editar
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Edit(int id, [Bind("IdUsuario,Nombre,Email,PasswordHash,RolId")] Usuario usuario)
        {
            if (id != usuario.IdUsuario)
            {
                return NotFound();
            }

            usuario.Nombre = usuario.Nombre?.Trim() ?? "";
            usuario.Email = usuario.Email?.Trim().ToLower() ?? "";

            ModelState.Remove("Activo");
            ModelState.Remove("FechaCreacion");
            ModelState.Remove("UltimoAcceso");
            ModelState.Remove("Rol");
            ModelState.Remove("RolId");

            if (usuario.RolId <= 0)
            {
                ModelState.AddModelError("RolId", "Debe seleccionar un rol.");
            }

            if (string.IsNullOrWhiteSpace(usuario.Nombre))
            {
                ModelState.AddModelError("Nombre", "El nombre es obligatorio.");
            }
            else if (usuario.Nombre.Length < 3)
            {
                ModelState.AddModelError("Nombre", "El nombre debe tener al menos 3 caracteres.");
            }

            if (ModelState.IsValid)
            {
                try
                {
                    bool correoExiste = await _context.Usuarios
                        .AnyAsync(u => u.Email == usuario.Email && u.IdUsuario != usuario.IdUsuario);

                    if (correoExiste)
                    {
                        ModelState.AddModelError("Email", "Ya existe otro usuario registrado con ese correo.");

                        ViewData["RolId"] = new SelectList(
     _context.Roles
         .Where(r => r.Activo || r.IdRol == usuario.RolId)
         .OrderBy(r => r.Nombre),
     "IdRol",
     "Nombre",
     usuario.RolId
 );

                        return View(usuario);
                    }

                    var usuarioExistente = await _context.Usuarios.FindAsync(id);

                    if (usuarioExistente == null)
                    {
                        return NotFound();
                    }

                    if (!usuarioExistente.Activo)
                    {
                        TempData["MensajeError"] = "No puede modificar usuarios inactivos. Debe activarlo primero.";
                        return RedirectToAction(nameof(Index));
                    }

                    usuarioExistente.Nombre = usuario.Nombre;
                    usuarioExistente.Email = usuario.Email;
                    usuarioExistente.RolId = usuario.RolId;

                    if (!string.IsNullOrWhiteSpace(usuario.PasswordHash))
                    {
                        usuarioExistente.PasswordHash = usuario.PasswordHash;
                    }

                    await _context.SaveChangesAsync();

                    var usuarioSesionNombre = HttpContext.Session.GetString("Usuario");

                    var usuarioSesionAud = await _context.Usuarios
                        .FirstOrDefaultAsync(u => u.Nombre == usuarioSesionNombre);

                    if (usuarioSesionAud != null)
                    {
                        _context.Auditoria.Add(new Auditorium
                        {
                            UsuarioId = usuarioSesionAud.IdUsuario,
                            Tabla = "Usuarios",
                            RegistroId = usuarioExistente.IdUsuario,
                            Accion = "Modificar",
                            Descripcion = "Se modificó el usuario: " + usuarioExistente.Nombre,
                            Fecha = DateTime.Now
                        });

                        await _context.SaveChangesAsync();
                    }

                    TempData["MensajeExito"] = "Usuario modificado correctamente.";
                    return RedirectToAction(nameof(Index));
                }
                catch (DbUpdateConcurrencyException)
                {
                    if (!UsuarioExists(usuario.IdUsuario))
                    {
                        return NotFound();
                    }

                    throw;
                }
            }

            ViewData["RolId"] = new SelectList(
        _context.Roles
            .Where(r => r.Activo || r.IdRol == usuario.RolId)
            .OrderBy(r => r.Nombre),
        "IdRol",
        "Nombre",
        usuario.RolId
    );

            return View(usuario);
        }


        public async Task<IActionResult> CambiarEstado(int id)
        {
            var usuario = await _context.Usuarios
                .Include(u => u.Rol)
                .FirstOrDefaultAsync(u => u.IdUsuario == id);

            if (usuario == null)
            {
                return NotFound();
            }

            var usuarioSesionNombre = HttpContext.Session.GetString("Usuario");

            if (usuario.Nombre.Trim().ToLower() == usuarioSesionNombre?.Trim().ToLower())
            {
                TempData["MensajeError"] = "No puede desactivar su propia cuenta.";
                return RedirectToAction(nameof(Index));
            }

            if (usuario.Activo && usuario.Rol != null && usuario.Rol.Nombre == "Administrador")
            {
                var cantidadAdminsActivos = await _context.Usuarios
                    .Include(u => u.Rol)
                    .CountAsync(u => u.Activo && u.Rol.Nombre == "Administrador");

                if (cantidadAdminsActivos <= 1)
                {
                    TempData["MensajeError"] = "No puede desactivar el último administrador activo del sistema.";
                    return RedirectToAction(nameof(Index));
                }
            }

            usuario.Activo = !usuario.Activo;

            var usuarioSesion = await _context.Usuarios
                .FirstOrDefaultAsync(u => u.Nombre == usuarioSesionNombre);

            if (usuarioSesion != null)
            {
                _context.Auditoria.Add(new Auditorium
                {
                    UsuarioId = usuarioSesion.IdUsuario,
                    Tabla = "Usuarios",
                    RegistroId = usuario.IdUsuario,
                    Accion = usuario.Activo ? "Activar" : "Desactivar",
                    Descripcion = usuario.Activo
                        ? "Se activó el usuario: " + usuario.Nombre
                        : "Se desactivó el usuario: " + usuario.Nombre,
                    Fecha = DateTime.Now
                });
            }

            await _context.SaveChangesAsync();

            TempData["MensajeExito"] = usuario.Activo
                ? "Usuario activado correctamente."
                : "Usuario desactivado correctamente.";

            return RedirectToAction(nameof(Index));
        }
        // ================================================================
        // CONSULTA COMPARTIDA PARA INDEX, PDF Y EXCEL
        // ================================================================
        private IQueryable<Usuario> ConstruirConsultaReporte(
            string? buscar,
            int? rolId,
            string? estado)
        {
            var consulta = _context.Usuarios
                .Include(u => u.Rol)
                .AsNoTracking()
                .AsQueryable();

            if (!string.IsNullOrWhiteSpace(buscar))
            {
                buscar = buscar.Trim();

                consulta = consulta.Where(u =>
                    (u.Nombre != null &&
                     u.Nombre.Contains(buscar)) ||

                    (u.Email != null &&
                     u.Email.Contains(buscar)) ||

                    (u.Rol != null &&
                     u.Rol.Nombre.Contains(buscar)));
            }

            if (rolId.HasValue)
            {
                consulta = consulta.Where(u =>
                    u.RolId == rolId.Value);
            }

            if (!string.IsNullOrWhiteSpace(estado))
            {
                if (estado.Equals(
                    "Activo",
                    StringComparison.OrdinalIgnoreCase))
                {
                    consulta = consulta.Where(u =>
                        u.Activo);
                }
                else if (estado.Equals(
                    "Inactivo",
                    StringComparison.OrdinalIgnoreCase))
                {
                    consulta = consulta.Where(u =>
                        !u.Activo);
                }
            }

            return consulta;
        }
        private async Task RegistrarAuditoria(
    string accion,
    string tabla,
    int registroId,
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
        private bool UsuarioExists(int id)
        {
            return _context.Usuarios.Any(e => e.IdUsuario == id);
        }
    }
}
