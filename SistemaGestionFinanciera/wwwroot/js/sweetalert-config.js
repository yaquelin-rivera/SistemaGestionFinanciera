// ================================================================
// CONFIGURACIÓN GENERAL DE SWEETALERT2
// Sistema Web de Gestión Financiera - UCAD Santa Ana
// ================================================================

const colorPrincipalSweetAlert = "#0F7C90";
const colorPrincipalOscuroSweetAlert = "#0F5A63";
const colorCancelarSweetAlert = "#6C757D";
const colorPeligroSweetAlert = "#DC3545";
const colorAdvertenciaSweetAlert = "#F0AD4E";
const colorExitoSweetAlert = "#198754";
const colorInformacionSweetAlert = "#0D6EFD"; 


// ================================================================
// CONFIGURACIÓN VISUAL GENERAL
// ================================================================

const configuracionVisualSweetAlert = {
    width: "430px",
    padding: "1.5rem",
    background: "#FFFFFF",
    color: "#34454B",

    buttonsStyling: true,
    heightAuto: false,

    allowOutsideClick: false,
    allowEscapeKey: true,

    backdrop: "rgba(3, 43, 48, 0.52)",

    customClass: {
        popup: "sweetalert-ucad-popup",
        title: "sweetalert-ucad-title",
        htmlContainer: "sweetalert-ucad-text",
        confirmButton: "sweetalert-ucad-confirm",
        cancelButton: "sweetalert-ucad-cancel",
        icon: "sweetalert-ucad-icon"
    }
};


// ================================================================
// FUNCIÓN INTERNA PARA UNIFICAR MENSAJES
// ================================================================

function crearSweetAlert(configuracion) {

    return Swal.fire({
        ...configuracionVisualSweetAlert,
        ...configuracion
    });
}


// ================================================================
// MENSAJE DE ÉXITO
// ================================================================

function mostrarExito(
    mensaje,
    titulo = "Operación realizada correctamente"
) {

    return crearSweetAlert({
        icon: "success",
        iconColor: colorExitoSweetAlert,

        title: titulo,
        text: mensaje,

        confirmButtonText: "Aceptar",
        confirmButtonColor: colorPrincipalSweetAlert
    });
}


// ================================================================
// MENSAJE DE ERROR
// ================================================================

function mostrarError(
    mensaje,
    titulo = "No fue posible realizar la operación"
) {

    return crearSweetAlert({
        icon: "error",
        iconColor: colorPeligroSweetAlert,

        title: titulo,
        text: mensaje,

        confirmButtonText: "Entendido",
        confirmButtonColor: colorPrincipalSweetAlert
    });
}


// ================================================================
// MENSAJE DE ADVERTENCIA
// ================================================================

function mostrarAdvertencia(
    mensaje,
    titulo = "Revise la información"
) {

    return crearSweetAlert({
        icon: "warning",
        iconColor: colorAdvertenciaSweetAlert,

        title: titulo,
        text: mensaje,

        confirmButtonText: "Entendido",
        confirmButtonColor: colorPrincipalSweetAlert
    });
}


// ================================================================
// MENSAJE INFORMATIVO
// ================================================================

function mostrarInformacion(
    mensaje,
    titulo = "Información del sistema"
) {

    return crearSweetAlert({
        icon: "info",
        iconColor: colorInformacionSweetAlert,

        title: titulo,
        text: mensaje,

        confirmButtonText: "Aceptar",
        confirmButtonColor: colorPrincipalSweetAlert
    });
}


// ================================================================
// CONFIRMACIÓN GENERAL
// Devuelve true cuando el usuario confirma.
// ================================================================

async function confirmarAccion(
    mensaje,
    titulo = "¿Desea continuar?",
    textoConfirmar = "Continuar",
    textoCancelar = "Cancelar",
    tipoIcono = "warning"
) {

    const resultado = await crearSweetAlert({
        icon: tipoIcono,
        iconColor: colorAdvertenciaSweetAlert,

        title: titulo,
        text: mensaje,

        showCancelButton: true,

        confirmButtonText: textoConfirmar,
        cancelButtonText: textoCancelar,

        confirmButtonColor: colorPrincipalSweetAlert,
        cancelButtonColor: colorCancelarSweetAlert,

        reverseButtons: true,
        focusCancel: true
    });

    return resultado.isConfirmed;
}


// ================================================================
// CONFIRMACIÓN PARA ELIMINAR
// ================================================================

async function confirmarEliminacion(
    mensaje = "El registro será eliminado permanentemente y esta acción no se podrá deshacer."
) {

    const resultado = await crearSweetAlert({
        icon: "warning",
        iconColor: colorPeligroSweetAlert,

        title: "¿Eliminar el registro?",
        text: mensaje,

        showCancelButton: true,

        confirmButtonText: "Eliminar",
        cancelButtonText: "Cancelar",

        confirmButtonColor: colorPeligroSweetAlert,
        cancelButtonColor: colorCancelarSweetAlert,

        reverseButtons: true,
        focusCancel: true
    });

    return resultado.isConfirmed;
}


// ================================================================
// CONFIRMACIÓN PARA ANULAR
// ================================================================

async function confirmarAnulacion(
    mensaje = "El registro será anulado y podrían revertirse sus movimientos asociados."
) {

    const resultado = await crearSweetAlert({
        icon: "warning",
        iconColor: colorPeligroSweetAlert,

        title: "¿Anular el registro?",
        text: mensaje,

        showCancelButton: true,

        confirmButtonText: "Anular",
        cancelButtonText: "Cancelar",

        confirmButtonColor: colorPeligroSweetAlert,
        cancelButtonColor: colorCancelarSweetAlert,

        reverseButtons: true,
        focusCancel: true
    });

    return resultado.isConfirmed;
}


// ================================================================
// CONFIRMACIÓN PARA DESACTIVAR
// ================================================================

async function confirmarDesactivacion(
    mensaje = "El registro dejará de estar disponible para nuevas operaciones."
) {

    const resultado = await crearSweetAlert({
        icon: "warning",
        iconColor: colorAdvertenciaSweetAlert,

        title: "¿Desactivar el registro?",
        text: mensaje,

        showCancelButton: true,

        confirmButtonText: "Desactivar",
        cancelButtonText: "Cancelar",

        confirmButtonColor: colorPeligroSweetAlert,
        cancelButtonColor: colorCancelarSweetAlert,

        reverseButtons: true,
        focusCancel: true
    });

    return resultado.isConfirmed;
}


// ================================================================
// CONFIRMACIÓN PARA REACTIVAR
// ================================================================

async function confirmarReactivacion(
    mensaje = "El registro volverá a estar disponible en el sistema."
) {

    const resultado = await crearSweetAlert({
        icon: "success",
        iconColor: colorExitoSweetAlert,

        title: "¿Reactivar el registro?",
        text: mensaje,

        showCancelButton: true,

        confirmButtonText: "Reactivar",
        cancelButtonText: "Cancelar",

        confirmButtonColor: colorPrincipalSweetAlert,
        cancelButtonColor: colorCancelarSweetAlert,

        reverseButtons: true,
        focusCancel: true
    });

    return resultado.isConfirmed;
}

// ================================================================
// CONFIRMACIÓN PROFESIONAL PARA CERRAR SESIÓN
// ================================================================
window.confirmarCerrarSesion = function (urlCerrarSesion) {

    Swal.fire({
        width: "455px",
        padding: 0,

        // Ya no usamos el icono predeterminado de SweetAlert
        icon: undefined,

        html: `
            <div class="swal-logout-header">
                <div class="swal-logout-icon">
                    <i class="fa-solid fa-right-from-bracket"></i>
                </div>

                <h2 class="swal-logout-title">
                    Cerrar sesión
                </h2>
            </div>

            <div class="swal-logout-body">
                <p class="swal-logout-question">
                    ¿Desea finalizar la sesión actual?
                </p>

                <p class="swal-logout-description">
                    Para volver a acceder al sistema deberá ingresar nuevamente
                    su correo electrónico y contraseña.
                </p>
            </div>
        `,

        showCancelButton: true,

        confirmButtonText:
            '<i class="fa-solid fa-right-from-bracket me-2"></i>Cerrar sesión',

        cancelButtonText:
            '<i class="fa-solid fa-xmark me-2"></i>Cancelar',

        confirmButtonColor: "#0F7C90",
        cancelButtonColor: "#6C757D",

        reverseButtons: true,
        focusCancel: true,

        allowOutsideClick: false,
        allowEscapeKey: true,

        backdrop: "rgba(3, 43, 48, 0.78)",

        customClass: {
            popup: "swal-logout-popup",
            actions: "swal-logout-actions",
            confirmButton: "swal-logout-confirm",
            cancelButton: "swal-logout-cancel"
        }
    }).then(function (resultado) {

        if (resultado.isConfirmed) {
            window.location.href = urlCerrarSesion;
        }
    });

    return false;
};