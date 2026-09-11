using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace SistemaGestionFinanciera.Models;

public partial class Factura
{
    public int IdFactura { get; set; }

    [Required(ErrorMessage = "Debe seleccionar el tipo de factura.")]
    [Display(Name = "Tipo de factura")]
    public string TipoFactura { get; set; } = null!;

    [Required(ErrorMessage = "Debe seleccionar la condición de pago.")]
    [Display(Name = "Condición de pago")]
    public string CondicionPago { get; set; }

    [Display(Name = "Días de crédito")]
    public int? DiasCredito { get; set; }

    [Display(Name = "Monto pagado")]
    public decimal MontoPagado { get; set; }

    [Display(Name = "Número de factura")]
    public string? NumeroFactura { get; set; } = null!;

    [Display(Name = "Cliente")]
    public int? ClienteBeneficiarioId { get; set; }

    [Display(Name = "Proveedor")]
    public int? ProveedorId { get; set; }

    [Required(ErrorMessage = "La fecha de emisión es obligatoria.")]
    [Display(Name = "Fecha de emisión")]
    public DateOnly FechaEmision { get; set; }

    [Display(Name = "Fecha de vencimiento")]
    public DateOnly? FechaVencimiento { get; set; }

    [Required(ErrorMessage = "El subtotal es obligatorio.")]
    [Range(0.01, 999999999, ErrorMessage = "El subtotal debe ser mayor a cero.")]
    [Display(Name = "Subtotal")]
    public decimal SubTotal { get; set; }
    public int? CategoriaIngresoId { get; set; }
    public int? CategoriaGastoId { get; set; }
    public int? ProyectoId { get; set; }

    [Display(Name = "Período presupuestario")]
    public int? PresupuestoMensualId { get; set; }
    public int? CentroCostoId { get; set; }
    public int? ServicioCatalogoId { get; set; }

    public string? TipoIngreso { get; set; }

    [Range(0, 999999999, ErrorMessage = "El descuento no puede ser negativo.")]
    [Display(Name = "Descuento")]
    public decimal? DescuentoTotal { get; set; }

    [Display(Name = "Tipo de impuesto")]
    public string TipoImpuesto { get; set; } = "IVA 13%";

    [Display(Name = "Impuesto")]
    public decimal ImpuestoTotal { get; set; }

    [Display(Name = "Total")]
    public decimal Total { get; set; }

    [Display(Name = "Saldo pendiente")]
    public decimal SaldoPendiente { get; set; }

    [Display(Name = "Estado")]
    public string? Estado { get; set; } = null!;

    [StringLength(500, ErrorMessage = "La observación no puede superar los 500 caracteres.")]
    [RegularExpression(@"^(?=.*[A-Za-zÁÉÍÓÚáéíóúÑñ0-9])[A-Za-zÁÉÍÓÚáéíóúÑñ0-9\s\.,;:#\-\/\(\)°_]*$",
    ErrorMessage = "La observación debe contener texto válido y solo puede incluir letras, números, espacios y signos comunes como . , ; : # - / ( ) ° _.")]
    [Display(Name = "Observación")]
    public string? Observacion { get; set; }

    public virtual ClientesBeneficiario? ClienteBeneficiario { get; set; }

    public virtual ICollection<FacturaDetalle> FacturaDetalles { get; set; } = new List<FacturaDetalle>();

    public virtual Proveedore? Proveedor { get; set; }
    public virtual CategoriasIngreso? CategoriaIngreso { get; set; }

    public virtual CategoriasGasto? CategoriaGasto { get; set; }

    public virtual Proyecto? Proyecto { get; set; }

    [ForeignKey(nameof(PresupuestoMensualId))]
    public virtual PresupuestosMensuale? PresupuestoMensualRegistro { get; set; }

    public virtual CentrosCosto? CentroCosto { get; set; }

    [ForeignKey("ServicioCatalogoId")]
    public virtual FacturaDetalle? ServicioCatalogo { get; set; }
}