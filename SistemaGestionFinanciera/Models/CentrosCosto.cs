using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace SistemaGestionFinanciera.Models;

public partial class CentrosCosto
{
    public int IdCentroCosto { get; set; }

    [Required(ErrorMessage = "El código es obligatorio.")]
    [StringLength(50, ErrorMessage = "El código no puede superar los 50 caracteres.")]
    [RegularExpression(@"^(?![\W_]+$)[A-Za-z0-9ÁÉÍÓÚáéíóúÑñ\s\-]+$",
        ErrorMessage = "El código no puede contener solo símbolos.")]
    public string? Codigo { get; set; }

    [Required(ErrorMessage = "El nombre es obligatorio.")]
    [StringLength(150, MinimumLength = 3, ErrorMessage = "El nombre debe tener entre 3 y 150 caracteres.")]
    [RegularExpression(@"^(?!\s+$)(?![\W_]+$)[A-Za-zÁÉÍÓÚáéíóúÑñ0-9\s\-]+$",
        ErrorMessage = "El nombre no puede estar vacío ni contener solo símbolos.")]
    public string Nombre { get; set; } = null!;

    [StringLength(300, ErrorMessage = "La descripción no puede superar los 300 caracteres.")]
    [RegularExpression(@"^$|^(?!\s+$)(?![\W_]+$)[A-Za-zÁÉÍÓÚáéíóúÑñ0-9\s.,;:()\-]+$",
        ErrorMessage = "La descripción no puede contener solo espacios o solo símbolos.")]
    public string? Descripcion { get; set; }

    public bool Activo { get; set; } = true;

    public virtual ICollection<MovimientosContable> MovimientosContables { get; set; } = new List<MovimientosContable>();
}
