using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace SistemaGestionFinanciera.Models;

public partial class MovimientosContable
{
    public int IdMovimientoContable { get; set; }

    [Required(ErrorMessage = "La fecha del movimiento es obligatoria.")]
    public DateOnly Fecha { get; set; }

    [Required(ErrorMessage = "La cuenta contable es obligatoria.")]
    public int CuentaContableId { get; set; }

    public int? ProyectoId { get; set; }

    public int? CentroCostoId { get; set; }

    [Required(ErrorMessage = "El tipo de movimiento es obligatorio.")]
    [StringLength(20, ErrorMessage = "El tipo de movimiento no puede superar los 20 caracteres.")]
    public string TipoMovimiento { get; set; } = null!;

    [Required(ErrorMessage = "El monto es obligatorio.")]
    [Range(0.01, 999999999999.99, ErrorMessage = "El monto debe ser mayor que cero.")]
    public decimal Monto { get; set; }

    [Range(0, 999999999999.99, ErrorMessage = "El debe no puede ser negativo.")]
    public decimal Debe { get; set; }

    [Range(0, 999999999999.99, ErrorMessage = "El haber no puede ser negativo.")]
    public decimal Haber { get; set; }

    [StringLength(500, ErrorMessage = "La descripción no puede superar los 500 caracteres.")]
    public string? Descripcion { get; set; }

    [StringLength(100, ErrorMessage = "La referencia no puede superar los 100 caracteres.")]
    public string? Referencia { get; set; }

    [StringLength(80, ErrorMessage = "El módulo de origen no puede superar los 80 caracteres.")]
    public string? OrigenModulo { get; set; }

    public int? OrigenId { get; set; }

    public bool? Anulado { get; set; } = false;

    public DateTime? FechaCreacion { get; set; } = DateTime.Now;

    [Required(ErrorMessage = "El estado es obligatorio.")]
    [StringLength(30, ErrorMessage = "El estado no puede superar los 30 caracteres.")]
    public string Estado { get; set; } = "Registrado";

    public bool EsAutomatico { get; set; } = true;

    public int? UsuarioId { get; set; }

    public DateTime? FechaAnulacion { get; set; }

    [StringLength(500, ErrorMessage = "El motivo de anulación no puede superar los 500 caracteres.")]
    public string? MotivoAnulacion { get; set; }

    public int? MovimientoReversionId { get; set; }

    public virtual CentrosCosto? CentroCosto { get; set; }

    public virtual CuentasContable CuentaContable { get; set; } = null!;

    public virtual Proyecto? Proyecto { get; set; }

    public virtual Usuario? Usuario { get; set; }

    public virtual MovimientosContable? MovimientoReversion { get; set; }

    public virtual ICollection<MovimientosContable> MovimientosRevertidos { get; set; } = new List<MovimientosContable>();
}
