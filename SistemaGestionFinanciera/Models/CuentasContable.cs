using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace SistemaGestionFinanciera.Models;

public partial class CuentasContable
{
    public int IdCuentaContable { get; set; }

    [Required(ErrorMessage = "El código es obligatorio.")]
    [StringLength(50, ErrorMessage = "El código no puede superar los 50 caracteres.")]
    [RegularExpression(@"^(?!\s+$)(?![\W_]+$)[A-Za-z0-9ÁÉÍÓÚáéíóúÑñ\s\-]+$",
        ErrorMessage = "El código no puede estar vacío ni contener solo símbolos.")]
    public string Codigo { get; set; } = null!;

    [Required(ErrorMessage = "El nombre es obligatorio.")]
    [StringLength(150, MinimumLength = 3, ErrorMessage = "El nombre debe tener entre 3 y 150 caracteres.")]
    [RegularExpression(@"^(?!\s+$)(?![\W_]+$)[A-Za-zÁÉÍÓÚáéíóúÑñ0-9\s\-]+$",
        ErrorMessage = "El nombre no puede estar vacío ni contener solo símbolos.")]
    public string Nombre { get; set; } = null!;

    [Required(ErrorMessage = "El tipo de cuenta es obligatorio.")]
    public string TipoCuenta { get; set; } = null!;

    [Required(ErrorMessage = "La naturaleza es obligatoria.")]
    public string Naturaleza { get; set; } = null!;

    [Required(ErrorMessage = "El saldo inicial es obligatorio.")]
    [Range(0, 999999999999.99, ErrorMessage = "El saldo inicial no puede ser negativo.")]
    public decimal SaldoInicial { get; set; }

    [Required(ErrorMessage = "El saldo actual es obligatorio.")]
    [Range(0, 999999999999.99, ErrorMessage = "El saldo actual no puede ser negativo.")]
    public decimal SaldoActual { get; set; }

    public bool Activo { get; set; } = true;

    public DateTime? FechaCreacion { get; set; } = DateTime.Now;

    public virtual ICollection<MovimientosContable> MovimientosContables { get; set; } = new List<MovimientosContable>();
}