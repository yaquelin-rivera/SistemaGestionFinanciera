using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace SistemaGestionFinanciera.Models;

public partial class Proveedore
{
    public int IdProveedor { get; set; }

    [Required(ErrorMessage = "El nombre es obligatorio.")]
    [StringLength(150, ErrorMessage = "El nombre no puede superar los 150 caracteres.")]
    [RegularExpression(@"^[A-Za-zÁÉÍÓÚáéíóúÑñ0-9&().,/\-\s]+$",
        ErrorMessage = "El nombre contiene caracteres no permitidos.")]
    public string Nombre { get; set; } = null!;

    [Required(ErrorMessage = "La identificación es obligatoria.")]
    [StringLength(50, ErrorMessage = "La identificación no puede superar los 50 caracteres.")]
    [MinLength(5, ErrorMessage = "La identificación debe tener al menos 5 caracteres.")]
    [RegularExpression(@"^(?=.*\d)[A-Za-z0-9\-]+$",
        ErrorMessage = "La identificación debe contener al menos un número y solo puede incluir letras, números y guiones.")]
    public string? Identificacion { get; set; }

    [RegularExpression(@"^[0-9]{8,15}$", ErrorMessage = "El teléfono debe contener solo números, entre 8 y 15 dígitos.")]
    public string? Telefono { get; set; }

    [EmailAddress(ErrorMessage = "Debe ingresar un correo electrónico válido.")]
    [StringLength(150)]
    public string? Email { get; set; }

    [StringLength(300, ErrorMessage = "La dirección no puede superar los 300 caracteres.")]
    [RegularExpression(@"^[A-Za-zÁÉÍÓÚáéíóúÑñ0-9#°&().,/\-\s]+$",
        ErrorMessage = "La dirección contiene caracteres no permitidos.")]
    public string? Direccion { get; set; }

    public bool Activo { get; set; } = true;

    // ================================================================
    // DATOS DE CONTACTO ADICIONALES se lamacenan en la tabl hija de otros contactos proveedores
    // ================================================================

    [NotMapped]
    [RegularExpression(
        @"^[0-9]{8,15}$",
        ErrorMessage = "El teléfono adicional debe contener solo números, entre 8 y 15 dígitos.")]
    public string? OtroTelefono { get; set; }

    [NotMapped]
    [StringLength(50, ErrorMessage = "La descripción no puede superar los 50 caracteres.")]
    [RegularExpression(
        @"^(?=.*[A-Za-zÁÉÍÓÚáéíóúÑñ])[A-Za-zÁÉÍÓÚáéíóúÑñ0-9 .,'()/#°º\-]+$",
        ErrorMessage = "La descripción debe contener al menos una letra y solo puede incluir texto, números y signos básicos.")]
    public string? DescripcionOtroTelefono { get; set; }


    [NotMapped]
    [EmailAddress(
        ErrorMessage = "Debe ingresar un correo electrónico adicional válido.")]
    [StringLength(
        150,
        ErrorMessage = "El correo adicional no puede superar los 150 caracteres.")]
    public string? OtroCorreo { get; set; }

    [NotMapped]
    [StringLength(50, ErrorMessage = "La descripción no puede superar los 50 caracteres.")]
    [RegularExpression(
        @"^(?=.*[A-Za-zÁÉÍÓÚáéíóúÑñ])[A-Za-zÁÉÍÓÚáéíóúÑñ0-9 .,'()/#°º\-]+$",
        ErrorMessage = "La descripción debe contener al menos una letra y solo puede incluir texto, números y signos básicos.")]
    public string? DescripcionOtroCorreo { get; set; }


    [NotMapped]
    [StringLength(
        300,
        ErrorMessage = "La dirección adicional no puede superar los 300 caracteres.")]
    [RegularExpression(
        @"^(?=.*[A-Za-zÁÉÍÓÚáéíóúÑñ])[A-Za-zÁÉÍÓÚáéíóúÑñ0-9 .,'()/#°º\-]+$",
        ErrorMessage = "La dirección adicional debe contener al menos una letra y solo puede incluir texto, números y signos básicos.")]
    public string? OtraDireccion { get; set; }

    [NotMapped]
    [StringLength(50, ErrorMessage = "La descripción no puede superar los 50 caracteres.")]
    [RegularExpression(
        @"^(?=.*[A-Za-zÁÉÍÓÚáéíóúÑñ])[A-Za-zÁÉÍÓÚáéíóúÑñ0-9 .,'()/#°º\-]+$",
        ErrorMessage = "La descripción debe contener al menos una letra y solo puede incluir texto, números y signos básicos.")]
    public string? DescripcionOtraDireccion { get; set; }


    [NotMapped]
    [RegularExpression(
        @"^[0-9]{8,15}$",
        ErrorMessage = "El segundo teléfono adicional debe contener solo números, entre 8 y 15 dígitos.")]
    public string? OtroTelefono2 { get; set; }

    [NotMapped]
    [StringLength(50, ErrorMessage = "La descripción no puede superar los 50 caracteres.")]
    [RegularExpression(
        @"^(?=.*[A-Za-zÁÉÍÓÚáéíóúÑñ])[A-Za-zÁÉÍÓÚáéíóúÑñ0-9 .,'()/#°º\-]+$",
        ErrorMessage = "La descripción debe contener al menos una letra y solo puede incluir texto, números y signos básicos.")]
    public string? DescripcionOtroTelefono2 { get; set; }


    [NotMapped]
    [EmailAddress(
        ErrorMessage = "Debe ingresar un segundo correo electrónico adicional válido.")]
    [StringLength(
        150,
        ErrorMessage = "El segundo correo adicional no puede superar los 150 caracteres.")]
    public string? OtroCorreo2 { get; set; }

    [NotMapped]
    [StringLength(50, ErrorMessage = "La descripción no puede superar los 50 caracteres.")]
    [RegularExpression(
        @"^(?=.*[A-Za-zÁÉÍÓÚáéíóúÑñ])[A-Za-zÁÉÍÓÚáéíóúÑñ0-9 .,'()/#°º\-]+$",
        ErrorMessage = "La descripción debe contener al menos una letra y solo puede incluir texto, números y signos básicos.")]
    public string? DescripcionOtroCorreo2 { get; set; }


    [NotMapped]
    [StringLength(
        300,
        ErrorMessage = "La segunda dirección adicional no puede superar los 300 caracteres.")]
    [RegularExpression(
        @"^(?=.*[A-Za-zÁÉÍÓÚáéíóúÑñ])[A-Za-zÁÉÍÓÚáéíóúÑñ0-9 .,'()/#°º\-]+$",
        ErrorMessage = "La segunda dirección adicional debe contener al menos una letra y solo puede incluir texto, números y signos básicos.")]
    public string? OtraDireccion2 { get; set; }

    [NotMapped]
    [StringLength(50, ErrorMessage = "La descripción no puede superar los 50 caracteres.")]
    [RegularExpression(
        @"^(?=.*[A-Za-zÁÉÍÓÚáéíóúÑñ])[A-Za-zÁÉÍÓÚáéíóúÑñ0-9 .,'()/#°º\-]+$",
        ErrorMessage = "La descripción debe contener al menos una letra y solo puede incluir texto, números y signos básicos.")]
    public string? DescripcionOtraDireccion2 { get; set; }
    public virtual ICollection<ContactoProveedor> ContactosAdicionales
    { get; set; } = new List<ContactoProveedor>();
    public virtual ICollection<CuentasPorPagar> CuentasPorPagars { get; set; } = new List<CuentasPorPagar>();

    // ================================================================
    // INFORMACIÓN FINANCIERA PARA CONSULTAS Y REPORTE No se almacena en la base de datos.
    

    [NotMapped]
    public int TotalCuentasValidas { get; set; }

    [NotMapped]
    public int TotalCuentasPendientes { get; set; }

    [NotMapped]
    public int TotalCuentasVencidas { get; set; }

    [NotMapped]
    public decimal SaldoTotalPendiente { get; set; }

    [NotMapped]
    public decimal SaldoTotalVencido { get; set; }

    [NotMapped]
    public string SituacionFinanciera { get; set; } =
        "Sin movimientos";

    public virtual ICollection<Factura> Facturas { get; set; } = new List<Factura>();

    public virtual ICollection<Gasto> Gastos { get; set; } = new List<Gasto>();
}
