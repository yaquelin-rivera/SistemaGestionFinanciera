using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace SistemaGestionFinanciera.Models;

public partial class CuentasPorCobrar
{
    public int IdCuentaPorCobrar { get; set; }

    [Required(ErrorMessage = "Debe seleccionar el cliente o beneficiario.")]
    public int? ClienteBeneficiarioId { get; set; }
    public int? ProyectoId { get; set; }

    [Display(Name = "Período presupuestario")]
    public int? PresupuestoMensualId { get; set; }

    public int? FacturaId { get; set; }

    [Required(ErrorMessage = "Debe seleccionar el tipo de documento.")]
    [StringLength(50)]
    public string? TipoDocumento { get; set; }

    [Required(ErrorMessage = "Debe ingresar el concepto.")]
    [StringLength(250)]
    [RegularExpression(@"^[a-zA-ZáéíóúÁÉÍÓÚñÑ0-9\s.,;:/()#-]+$",
       ErrorMessage = "El concepto contiene caracteres no permitidos.")]
    public string Concepto { get; set; } = null!;


    [StringLength(100)]
    public string? NumeroDocumento { get; set; }

    public DateOnly FechaEmision { get; set; }

    public DateOnly FechaVencimiento { get; set; }

    [Required(ErrorMessage = "Debe ingresar el monto original de la cuenta.")]
    [Range(0.01, 999999999, ErrorMessage = "El monto original debe ser mayor a cero.")]
    public decimal? MontoOriginal { get; set; }

    public decimal? MontoPagado { get; set; }

    [Range(0, 999999999, ErrorMessage = "El anticipo aplicado no puede ser negativo.")]
    public decimal? AnticipoAplicado { get; set; }

    [Required(ErrorMessage = "Debe seleccionar los días de crédito.")]
    public int? DiasCredito { get; set; }
    public decimal SaldoPendiente { get; set; }

    public string Estado { get; set; } = null!;

    [Display(Name = "Origen")]
    [StringLength(20)]
    public string Origen { get; set; } = "Manual";

    [Display(Name = "Estado autorización")]
    [StringLength(30)]
    public string EstadoAutorizacion { get; set; } = "Pendiente";

    [StringLength(500)]
    [RegularExpression(@"^[a-zA-ZáéíóúÁÉÍÓÚñÑ0-9\s.,;:/()#-]*$",
        ErrorMessage = "La observación contiene caracteres no permitidos.")]
    public string? Observacion { get; set; }

    [Display(Name = "Motivo de anulación")]
    [StringLength(650)]
    [RegularExpression(
    @"^(?=(?:.*[a-zA-ZáéíóúÁÉÍÓÚñÑ]){3,})[a-zA-ZáéíóúÁÉÍÓÚñÑ0-9\s.,;:/()#\-]*$",
    ErrorMessage = "El motivo de anulación debe contener al menos tres letras y solo caracteres permitidos.")]
    public string? MotivoAnulacion { get; set; }
    public bool Activo { get; set; }

    public virtual ClientesBeneficiario ClienteBeneficiario { get; set; } = null!;

    public virtual Factura? Factura { get; set; }

    public virtual Proyecto? Proyecto { get; set; }

    public virtual PresupuestosMensuale? PresupuestoMensualRegistro { get; set; }
    public virtual ICollection<PagosCuentaPorCobrar> PagosCuentaPorCobrars { get; set; } = new List<PagosCuentaPorCobrar>();
}