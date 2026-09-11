using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace SistemaGestionFinanciera.Models;

public partial class Proyecto
{
    public int IdProyecto { get; set; }

    [Required(ErrorMessage = "El nombre es obligatorio.")]
    [StringLength(
        150,
        MinimumLength = 3,
        ErrorMessage = "El nombre debe contener entre 3 y 150 caracteres.")]
    [RegularExpression(
        @"^(?=.*[A-Za-zÁÉÍÓÚáéíóúÑñ])[A-Za-zÁÉÍÓÚáéíóúÑñ0-9\s.,()#/\-]+$",
        ErrorMessage = "El nombre debe contener al menos una letra y solo puede incluir letras, números, espacios y puntuación básica.")]
    [Display(Name = "Nombre")]
    public string Nombre { get; set; } = null!;


    [Required(ErrorMessage = "La descripción es obligatoria.")]
    [StringLength(
        500,
        MinimumLength = 5,
        ErrorMessage = "La descripción debe contener entre 5 y 500 caracteres.")]
    [RegularExpression(
        @"^(?=.*[A-Za-zÁÉÍÓÚáéíóúÑñ])[A-Za-zÁÉÍÓÚáéíóúÑñ0-9\s.,:;()#/\-]+$",
        ErrorMessage = "La descripción debe contener al menos una letra y solo puede incluir letras, números, espacios y puntuación básica.")]
    [Display(Name = "Descripción")]
    public string Descripcion { get; set; } = null!;


    [Required(ErrorMessage = "La fecha de inicio es obligatoria.")]
    [Display(Name = "Fecha de inicio")]
    public DateOnly? FechaInicio { get; set; }


    [Display(Name = "Fecha de fin")]
    public DateOnly? FechaFin { get; set; }


    [Required(ErrorMessage = "El estado es obligatorio.")]
    [StringLength(30)]
    public string Estado { get; set; } = null!;


    [Required(ErrorMessage = "El presupuesto asignado es obligatorio.")]
    [Range(
        typeof(decimal),
        "0.01",
        "999999999999999.99",
        ErrorMessage = "El presupuesto asignado debe ser mayor que cero.")]
    [Display(Name = "Presupuesto asignado")]
    public decimal? PresupuestoAsignado { get; set; }


    [NotMapped]
    [StringLength(
        300,
        ErrorMessage = "El motivo no puede superar los 300 caracteres.")]
    [RegularExpression(
        @"^(?=.*[A-Za-zÁÉÍÓÚáéíóúÑñ])[A-Za-zÁÉÍÓÚáéíóúÑñ0-9\s.,:;()#/\-]+$",
        ErrorMessage = "El motivo debe contener al menos una letra y solo puede incluir texto y puntuación básica.")]
    [Display(Name = "Motivo del ajuste presupuestario")]
    public string? MotivoAjustePresupuesto { get; set; }


    public bool Activo { get; set; } = true;

public virtual ICollection<CuentasPorCobrar> CuentasPorCobrars { get; set; } = new List<CuentasPorCobrar>();

    public virtual ICollection<CuentasPorPagar> CuentasPorPagars { get; set; } = new List<CuentasPorPagar>();

    public virtual ICollection<Gasto> Gastos { get; set; } = new List<Gasto>();

    public virtual ICollection<Ingreso> Ingresos { get; set; } = new List<Ingreso>();

    public virtual ICollection<MovimientosContable> MovimientosContables { get; set; } = new List<MovimientosContable>();

    public virtual ICollection<PresupuestosMensuale> PresupuestosMensuales { get; set; } = new List<PresupuestosMensuale>();
}
