using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace SistemaGestionFinanciera.Models;

public partial class CuentasPorPagar
{
    public int IdCuentaPorPagar { get; set; }

    [Required(ErrorMessage = "Debe seleccionar el proveedor.")]
    public int ProveedorId { get; set; }

    public int? ProyectoId { get; set; }

    [Display(Name = "Período presupuestario")]
    public int? PresupuestoMensualId { get; set; }

    [Required(ErrorMessage = "Debe seleccionar la categoría de gasto.")]
    public int? CategoriaGastoId { get; set; }

    public int? CentroCostoId { get; set; }

    [Required(ErrorMessage = "Debe seleccionar el tipo de documento.")]
    [StringLength(50)]
    public string? TipoDocumento { get; set; }

    [StringLength(100)]
    public string? NumeroDocumento { get; set; }

    [StringLength(100)]
    public string? NumeroFactura { get; set; }

    [Required(ErrorMessage = "Debe ingresar el concepto.")]
    [StringLength(250, MinimumLength = 3,
        ErrorMessage = "El concepto debe tener entre 3 y 250 caracteres.")]
    [RegularExpression(
        @"^(?=.*[a-zA-ZáéíóúÁÉÍÓÚñÑ])[a-zA-ZáéíóúÁÉÍÓÚñÑ0-9\s.,;:/()#%\-]+$",
        ErrorMessage = "El concepto debe contener al menos una letra y solo caracteres permitidos.")]
    public string Concepto { get; set; } = null!;

    public DateOnly FechaEmision { get; set; }

    public DateOnly FechaVencimiento { get; set; }

    [Required(ErrorMessage = "Debe seleccionar los días de crédito.")]
    public int? DiasCredito { get; set; }

    [Range(0, 999999999, ErrorMessage = "El subtotal no puede ser negativo.")]
    public decimal? SubTotal { get; set; }

    [StringLength(50)]
    public string? TipoImpuesto { get; set; }

    [Range(0, 999999999, ErrorMessage = "El impuesto no puede ser negativo.")]
    public decimal? Impuesto { get; set; }

    [Range(0, 999999999, ErrorMessage = "El descuento no puede ser negativo.")]
    public decimal? Descuento { get; set; }

    [Required(ErrorMessage = "Debe ingresar el monto original.")]
    [Range(0.01, 999999999, ErrorMessage = "El monto original debe ser mayor a cero.")]
    public decimal MontoOriginal { get; set; }

    public decimal? MontoPagado { get; set; }

    [Range(0, 999999999, ErrorMessage = "El anticipo aplicado no puede ser negativo.")]
    public decimal? AnticipoAplicado { get; set; }

    public decimal SaldoPendiente { get; set; }

    public string Estado { get; set; } = null!;

    [Display(Name = "Origen")]
    [StringLength(20)]
    public string Origen { get; set; } = "Manual";

    [Display(Name = "Estado autorización")]
    [StringLength(30)]
    public string EstadoAutorizacion { get; set; } = "Pendiente";

    [StringLength(500)]
    [RegularExpression(
        @"^$|^(?=.*[a-zA-ZáéíóúÁÉÍÓÚñÑ])[a-zA-ZáéíóúÁÉÍÓÚñÑ0-9\s.,;:/()#%\-]*$",
        ErrorMessage = "La observación debe contener al menos una letra y solo caracteres permitidos.")]
    public string? Observacion { get; set; }

    [StringLength(500)]
    [RegularExpression(
        @"^$|^(?=.*[a-zA-ZáéíóúÁÉÍÓÚñÑ])[a-zA-ZáéíóúÁÉÍÓÚñÑ0-9\s.,;:/()#%\-]*$",
        ErrorMessage = "El motivo de anulación debe contener al menos una letra y solo caracteres permitidos.")]
    public string? MotivoAnulacion { get; set; }

    public bool Activo { get; set; }

    public virtual ICollection<PagosCuentaPorPagar> PagosCuentaPorPagars { get; set; } = new List<PagosCuentaPorPagar>();

    public virtual Proveedore Proveedor { get; set; } = null!;

    public virtual Proyecto? Proyecto { get; set; }

    [ForeignKey(nameof(PresupuestoMensualId))]
    public virtual PresupuestosMensuale? PresupuestoMensualRegistro { get; set; }
    public virtual CategoriasGasto? CategoriaGasto { get; set; }

    public virtual CentrosCosto? CentroCosto { get; set; }
}