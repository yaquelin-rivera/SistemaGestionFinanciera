using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;


namespace SistemaGestionFinanciera.Models;

public partial class CategoriasIngreso
{
    public int IdCategoriaIngreso { get; set; }


    [Required(ErrorMessage = "El nombre es obligatorio.")]
    [StringLength(100, MinimumLength = 3,
     ErrorMessage = "El nombre debe tener entre 3 y 100 caracteres.")]
    [RegularExpression(@"^[A-Za-zÁÉÍÓÚáéíóúÑñ\s]+$",
     ErrorMessage = "El nombre solo puede contener letras y espacios.")]
    public string Nombre { get; set; } = null!;

    [StringLength(250,
   ErrorMessage = "La descripción no puede superar los 250 caracteres.")]
    public string? Descripcion { get; set; }

    public bool Activo { get; set; }
    public int? CuentaContableId { get; set; }

    public virtual CuentasContable? CuentaContable { get; set; }
    public virtual ICollection<Ingreso> Ingresos { get; set; } = new List<Ingreso>();

    public virtual ICollection<PresupuestoDetalle> PresupuestoDetalles { get; set; } = new List<PresupuestoDetalle>();
}
