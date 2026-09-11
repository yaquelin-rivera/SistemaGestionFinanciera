using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace SistemaGestionFinanciera.Models;

public partial class Role
{
    public int IdRol { get; set; }

    [Required(ErrorMessage = "El nombre del rol es obligatorio.")]
    [StringLength(
     100,
     MinimumLength = 3,
     ErrorMessage = "El nombre del rol debe contener entre 3 y 100 caracteres.")]
    [RegularExpression(
     @"^[A-Za-zÁÉÍÓÚáéíóúÑñ0-9\s-]+$",
     ErrorMessage = "El nombre solo puede contener letras, números, espacios y guiones.")]
    [Display(Name = "Nombre")]
    public string Nombre { get; set; } = string.Empty;


    [StringLength(
        250,
        ErrorMessage = "La descripción no puede superar los 250 caracteres.")]
    [RegularExpression(
        @"^[A-Za-zÁÉÍÓÚáéíóúÑñ0-9\s.,;:()/-]*$",
        ErrorMessage = "La descripción contiene caracteres no permitidos.")]
    [Display(Name = "Descripción")]
    public string? Descripcion { get; set; }

    public bool Activo { get; set; }

    public DateTime? FechaCreacion { get; set; }

    public virtual ICollection<RolesPermiso> RolesPermisos { get; set; } = new List<RolesPermiso>();

    public virtual ICollection<Usuario> Usuarios { get; set; } = new List<Usuario>();
}
