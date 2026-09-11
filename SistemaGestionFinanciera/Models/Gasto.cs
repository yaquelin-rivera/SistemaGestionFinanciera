using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using Microsoft.AspNetCore.Mvc.ModelBinding.Validation;

namespace SistemaGestionFinanciera.Models;

public partial class Gasto
{
    public int IdGasto { get; set; }

    [Display(Name = "Fecha")]
    [Required(ErrorMessage = "La fecha del gasto es obligatoria.")]
    [DataType(DataType.Date)]
    public DateOnly Fecha { get; set; }

    [Display(Name = "Categoría de gasto")]
    [Required(ErrorMessage = "Debe seleccionar una categoría de gasto.")]
    public int CategoriaGastoId { get; set; }

    [Display(Name = "Proyecto")]
    public int? ProyectoId { get; set; }

    [Display(Name = "Período presupuestario")]
    public int? PresupuestoMensualId { get; set; }

    [Display(Name = "Proveedor")]
    public int? ProveedorId { get; set; }

    [Display(Name = "Concepto")]
    [Required(ErrorMessage = "Debe ingresar el concepto del gasto.")]
    [StringLength(250, MinimumLength = 3, ErrorMessage = "El concepto debe tener entre 3 y 250 caracteres.")]
    [RegularExpression(@"^(?=.*[A-Za-zÁÉÍÓÚáéíóúÑñ0-9]).+$",
     ErrorMessage = "El concepto no puede contener solo símbolos.")]
    public string Concepto { get; set; } = null!;

    [Display(Name = "Subtotal")]
    [Required(ErrorMessage = "Debe ingresar el subtotal.")]
    [Range(0.01, 999999999.99, ErrorMessage = "El subtotal debe ser mayor que 0.")]
    [Column(TypeName = "decimal(18,2)")]
    public decimal? Subtotal { get; set; }

    [Display(Name = "Impuesto")]
    [Range(0, 999999999.99, ErrorMessage = "El impuesto no puede ser negativo.")]
    [Column(TypeName = "decimal(18,2)")]
    public decimal? Impuesto { get; set; }

    [Display(Name = "Descuento")]
    [Range(0, 999999999.99, ErrorMessage = "El descuento no puede ser negativo.")]
    [Column(TypeName = "decimal(18,2)")]
    public decimal? Descuento { get; set; }

    [Display(Name = "Tipo de impuesto")]
    [Required(ErrorMessage = "Debe seleccionar el tipo de impuesto.")]
    [StringLength(20)]
    public string TipoImpuesto { get; set; } = "Sin impuesto";

    [Display(Name = "Monto total")]
    [Column(TypeName = "decimal(18,2)")]
    public decimal? MontoTotal { get; set; }

    [Display(Name = "Estado")]
    [StringLength(30)]
    public string Estado { get; set; } = "Registrado";

    [Display(Name = "Motivo del rechazo")]
    [StringLength(500, ErrorMessage = "El motivo del rechazo no puede superar los 500 caracteres.")]
    public string? MotivoRechazo { get; set; }

    [Display(Name = "Fecha de rechazo")]
    public DateTime? FechaRechazo { get; set; }

    [Display(Name = "Usuario que rechazó")]
    public int? UsuarioRechazoId { get; set; }

    [Display(Name = "Supera presupuesto")]
    public bool SuperaPresupuesto { get; set; }
    [Display(Name = "Motivo de anulación")]
    [StringLength(500, ErrorMessage = "El motivo de anulación no puede superar los 500 caracteres.")]
    public string? MotivoAnulacion { get; set; }

    [Display(Name = "Fecha de anulación")]
    public DateTime? FechaAnulacion { get; set; }

    [Display(Name = "Usuario que anuló")]
    public int? UsuarioAnulacionId { get; set; }

    public decimal? PresupuestoMensual { get; set; }
    public decimal? GastadoAcumulado { get; set; }
    public decimal? SaldoDisponible { get; set; }

    [Display(Name = "Observación")]
    [StringLength(500, ErrorMessage = "La observación no puede superar los 500 caracteres.")]
    [RegularExpression(@"^$|^(?=.*[A-Za-zÁÉÍÓÚáéíóúÑñ0-9]).+$",
    ErrorMessage = "La observación no puede contener solo símbolos.")]

    public string? Observacion { get; set; }
    [Display(Name = "Centro de costo")]
    public int? CentroCostoId { get; set; }

    [Display(Name = "Número de factura")]
    [StringLength(50, ErrorMessage = "El número de factura no puede superar los 50 caracteres.")]
    public string? NumeroFactura { get; set; }

    [Display(Name = "Activo")]
    public bool Activo { get; set; } = true;

    [ValidateNever]
    public virtual CategoriasGasto CategoriaGasto { get; set; } = null!;

    [ValidateNever]
    public virtual Proveedore? Proveedor { get; set; }

    [ValidateNever]
    public virtual Proyecto? Proyecto { get; set; }

    [ForeignKey(nameof(PresupuestoMensualId))]
    [ValidateNever]
    public virtual PresupuestosMensuale? PresupuestoMensualRegistro { get; set; }

    [ValidateNever]

    public virtual CentrosCosto? CentroCosto { get; set; }
}
