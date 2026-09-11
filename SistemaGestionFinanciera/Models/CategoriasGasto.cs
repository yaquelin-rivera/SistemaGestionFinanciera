using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace SistemaGestionFinanciera.Models;

public partial class CategoriasGasto
{
    public int IdCategoriaGasto { get; set; }

    [Required(ErrorMessage = "El nombre es obligatorio.")]
    [StringLength(100, MinimumLength = 3,
    ErrorMessage = "El nombre debe tener entre 3 y 100 caracteres.")]
    [RegularExpression(@"^[A-Za-zÁÉÍÓÚáéíóúÑñ\s]+$",
    ErrorMessage = "El nombre solo puede contener letras y espacios.")]
    public string Nombre { get; set; } = null!;

    [StringLength(250,
        ErrorMessage = "La descripción no puede superar los 250 caracteres.")]
    public string? Descripcion { get; set; }

    [Required(ErrorMessage = "El límite mensual es obligatorio.")]
    [Range(typeof(decimal), "0", "999999999",
     ErrorMessage = "El límite mensual debe ser mayor o igual a 0.")]
    [Display(Name = "Límite mensual")]
    public decimal LimiteMensual { get; set; }

    public bool Activo { get; set; }
    public int? CuentaContableId { get; set; }

    public virtual CuentasContable? CuentaContable { get; set; }
    public virtual ICollection<Gasto> Gastos { get; set; } = new List<Gasto>();

    public virtual ICollection<PresupuestoDetalle> PresupuestoDetalles { get; set; } = new List<PresupuestoDetalle>();
}
