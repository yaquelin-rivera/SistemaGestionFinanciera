using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Filters;
using System;
using System.Collections.Generic;
using System.Text.Json;

namespace SistemaGestionFinanciera.Controllers
{
    public class BaseController : Controller
    {
        
        // RELACIÓN ENTRE CONTROLADORES Y PERMISOS
        // los permisos deben coincidir exactamente con los registrados en la tabla Permisos.

        private static readonly Dictionary<string, string> PermisosPorControlador =
     new(StringComparer.OrdinalIgnoreCase)
     {
         // ========================================================
         // MANTENIMIENTOS Y CATÁLOGOS
         ["Proyectos"] =
             "Gestionar Proyectos",

         ["ClientesBeneficiarios"] =
             "Gestionar Clientes",

         ["Proveedores"] =
             "Gestionar Proveedores",

         ["CategoriasIngresos"] =
             "Gestionar Categorías",

         ["CategoriasGastos"] =
             "Gestionar Categorías",

         ["CentrosCostos"] =
             "Gestionar Centros de Costo",

         // ========================================================
         // CONTABILIDAD
         // ========================================================
         ["CuentasContables"] =
    "Gestionar Cuentas Contables",

         ["MovimientosContables"] =
    "Consultar Movimientos Contables",

         ["CierresContables"] =
    "Consultar Movimientos Contables",
         // ========================================================
         // OPERACIONES
         // ========================================================
         ["Ingresos"] =
             "Registrar Ingresos",

         ["Gastos"] =
             "Registrar Gastos",

         ["Facturas"] =
             "Gestionar Facturación",

         ["FacturasDetalles"] =
             "Gestionar Facturación",

         // ========================================================
         // CUENTAS POR COBRAR
         // ========================================================
         ["CuentasPorCobrar"] =
             "Gestionar Cuentas por Cobrar",

         ["PagosCuentasPorCobrar"] =
             "Gestionar Cuentas por Cobrar",

         // ========================================================
         // CUENTAS POR PAGAR
         // ========================================================
         ["CuentasPorPagar"] =
             "Gestionar Cuentas por Pagar",

         ["PagosCuentasPorPagar"] =
             "Gestionar Cuentas por Pagar",

         // ========================================================
         // PRESUPUESTO
         // ========================================================
         ["PresupuestosMensuales"] =
             "Gestionar Presupuesto",

         ["PresupuestosDetalles"] =
             "Gestionar Presupuesto"
     };

        // ================================================================
        // CONTROLADORES EXCLUSIVOS DEL ADMINISTRADOR
        // ================================================================
        private static readonly HashSet<string> ControladoresSoloAdministrador =
            new(StringComparer.OrdinalIgnoreCase)
            {
                "Usuarios",
                "Roles",
                "Permisos",
                "RolesPermisos",
                "Auditoriums"
            };

        // ================================================================
        // VALIDACIÓN GENERAL ANTES DE EJECUTAR CUALQUIER ACCIÓN
        // ================================================================
        public override void OnActionExecuting(
            ActionExecutingContext context)
        {
            string? usuario =
                HttpContext.Session.GetString("Usuario");

            string? rol =
                HttpContext.Session.GetString("Rol");

            string? controller =
                context.RouteData.Values["controller"]?.ToString();

            // ============================================================
            // 1. VERIFICAR SESIÓN
            // ============================================================
            if (string.IsNullOrWhiteSpace(usuario))
            {
                context.Result = new RedirectToActionResult(
                    "Login",
                    "Acceso",
                    null);

                return;
            }

            // ============================================================
            // 2. VALIDAR QUE EXISTA EL CONTROLADOR
            // ============================================================
            if (string.IsNullOrWhiteSpace(controller))
            {
                DenegarAcceso(context);
                return;
            }

            // ============================================================
            // 3. HOME ES ACCESIBLE PARA TODO USUARIO AUTENTICADO
            // ============================================================
            if (controller.Equals(
                "Home",
                StringComparison.OrdinalIgnoreCase))
            {
                base.OnActionExecuting(context);
                return;
            }

            bool esAdministrador =
                rol?.Equals(
                    "Administrador",
                    StringComparison.OrdinalIgnoreCase)
                == true;

            // ============================================================
            // 4. SEGURIDAD Y AUDITORÍA:
            //    EXCLUSIVAMENTE ADMINISTRADOR
            // ============================================================
            if (ControladoresSoloAdministrador.Contains(controller))
            {
                if (!esAdministrador)
                {
                    DenegarAcceso(context);
                    return;
                }

                base.OnActionExecuting(context);
                return;
            }

            // ============================================================
            // 5. ADMINISTRADOR TIENE ACCESO COMPLETO
            // ============================================================
            if (esAdministrador)
            {
                base.OnActionExecuting(context);
                return;
            }

            // ============================================================
            // 6. BUSCAR EL PERMISO REQUERIDO PARA EL CONTROLADOR
            // ============================================================
            if (!PermisosPorControlador.TryGetValue(
                controller,
                out string? permisoRequerido))
            {
                // Por seguridad, todo controlador que no esté registrado
                // queda bloqueado para usuarios no administradores.
                DenegarAcceso(context);
                return;
            }

            // ============================================================
            // 7. VALIDAR EL PERMISO CARGADO EN LA SESIÓN
            // ============================================================
            if (!TienePermiso(permisoRequerido))
            {
                DenegarAcceso(context);
                return;
            }

            base.OnActionExecuting(context);
        }

        // ================================================================
        // COMPROBAR SI LA SESIÓN CONTIENE UN PERMISO
        // ================================================================
        protected bool TienePermiso(string nombrePermiso)
        {
            string? rol =
                HttpContext.Session.GetString("Rol");

            // El Administrador conserva acceso completo.
            if (rol?.Equals(
                "Administrador",
                StringComparison.OrdinalIgnoreCase)
                == true)
            {
                return true;
            }

            string? permisosJson =
                HttpContext.Session.GetString("Permisos");

            if (string.IsNullOrWhiteSpace(permisosJson))
            {
                return false;
            }

            try
            {
                var permisos =
                    JsonSerializer.Deserialize<List<string>>(
                        permisosJson)
                    ?? new List<string>();

                return permisos.Exists(
                    permiso =>
                        permiso.Equals(
                            nombrePermiso,
                            StringComparison.OrdinalIgnoreCase));
            }
            catch (JsonException)
            {
                // Si la información de permisos está dañada,
                // se deniega el acceso por seguridad.
                return false;
            }
        }

        // ================================================================
        // REDIRECCIÓN POR ACCESO DENEGADO
        // ================================================================
        private void DenegarAcceso(
            ActionExecutingContext context)
        {
            TempData["Warning"] =
                "No tiene permiso para acceder a esta funcionalidad.";

            context.Result = new RedirectToActionResult(
                "Index",
                "Home",
                null);
        }

        // ================================================================
        // MENSAJES GENERALES DEL SISTEMA
        // Se muestran automáticamente mediante SweetAlert desde _Layout
        // ================================================================
        protected void MostrarExito(string mensaje)
        {
            TempData["Success"] = mensaje;
        }

        protected void MostrarError(string mensaje)
        {
            TempData["Error"] = mensaje;
        }

        protected void MostrarAdvertencia(string mensaje)
        {
            TempData["Warning"] = mensaje;
        }

        protected void MostrarInformacion(string mensaje)
        {
            TempData["Info"] = mensaje;
        }
    }
}