using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using Microsoft.AspNetCore.Mvc.ModelBinding.Validation;

namespace SistemaGestionFinanciera.Models;

public partial class Ingreso
{
    [Key]
public int IdIngreso { get; set; }
// intregracion
    [Display(Name = "Factura relacionada")]
    public int? FacturaId { get; set; }

    [Display(Name = "Fecha")]
[Required(ErrorMessage = "La fecha es obligatoria.")]
public DateOnly Fecha { get; set; }

[Display(Name = "Categoría de ingreso")]
[Required(ErrorMessage = "Debe seleccionar una categoría de ingreso.")]
public int CategoriaIngresoId { get; set; }

[Display(Name = "Proyecto")]
[Required(ErrorMessage = "Debe seleccionar un proyecto.")]
public int? ProyectoId { get; set; }
    [Display(Name = "Período presupuestario")]
    public int? PresupuestoMensualId { get; set; }

    [Display(Name = "Motivo pendiente presupuestario")]
    [StringLength(250, ErrorMessage = "El motivo pendiente presupuestario no puede superar los 250 caracteres.")]
    public string? MotivoPendientePresupuestario { get; set; }

    [Display(Name = "Cliente o beneficiario")]
[Required(ErrorMessage = "Debe seleccionar un cliente o beneficiario.")]
public int? ClienteBeneficiarioId { get; set; }

[Display(Name = "Fuente")]
[Required(ErrorMessage = "La fuente del ingreso es obligatoria.")]
[StringLength(80, MinimumLength = 3, ErrorMessage = "La fuente debe tener entre 3 y 80 caracteres.")]
[RegularExpression(@"^(?![\W_]+$)(?!\d+$)[A-Za-zÁÉÍÓÚáéíóúÑñ0-9\s.,-]+$",
    ErrorMessage = "La fuente no puede contener solo números, solo símbolos ni caracteres no permitidos.")]
public string? Fuente { get; set; }

[Display(Name = "Monto esperado")]
[Required(ErrorMessage = "El monto esperado es obligatorio.")]
[Range(1, 999999999.99, ErrorMessage = "El monto esperado debe ser mayor que cero.")]
public decimal? MontoEsperado { get; set; }

[Display(Name = "Monto real")]
[Required(ErrorMessage = "El monto real es obligatorio.")]
[Range(0, 999999999.99, ErrorMessage = "El monto real no puede ser negativo.")]
public decimal? MontoReal { get; set; }

[Display(Name = "Diferencia")]
public decimal? Diferencia { get; set; }

[Display(Name = "¿Es anticipo?")]
public bool EsAnticipo { get; set; }

[Display(Name = "Descripción")]
[Required(ErrorMessage = "La descripción es obligatoria.")]
[StringLength(500, MinimumLength = 10, ErrorMessage = "La descripción debe tener entre 10 y 500 caracteres.")]
[RegularExpression(@"^(?![\W_]+$)(?!\d+$)[A-Za-zÁÉÍÓÚáéíóúÑñ0-9\s.,;:()/-]+$",
    ErrorMessage = "La descripción no puede contener solo números, solo símbolos ni caracteres no permitidos.")]
public string? Descripcion { get; set; }

[Display(Name = "Estado")]
[StringLength(30, ErrorMessage = "El estado no puede superar los 30 caracteres.")]
public string? Estado { get; set; }

[Display(Name = "Tipo de ingreso")]
[Required(ErrorMessage = "Debe seleccionar un tipo de ingreso.")]
[StringLength(50, ErrorMessage = "El tipo de ingreso no puede superar los 50 caracteres.")]
public string? TipoIngreso { get; set; }
 [Display(Name = "Centro de costo")]
 public int? CentroCostoId { get; set; }

    [Display(Name = "Origen")]
    [StringLength(30, ErrorMessage = "El origen no puede superar los 30 caracteres.")]
    public string? Origen { get; set; } = "Manual";

    [Display(Name = "Observaciones")]
[StringLength(500, ErrorMessage = "Las observaciones no pueden superar los 500 caracteres.")]
public string? Observaciones { get; set; }

[Display(Name = "Comprobante")]
[StringLength(100, ErrorMessage = "El comprobante no puede superar los 100 caracteres.")]
public string? Comprobante { get; set; }
 [Display(Name = "Activo")]
public bool Activo { get; set; }
    [ValidateNever]
    public virtual CategoriasIngreso CategoriaIngreso { get; set; } = null!;

    [ValidateNever]
    public virtual ClientesBeneficiario? ClienteBeneficiario { get; set; }

    [ValidateNever]
    public virtual Proyecto? Proyecto { get; set; }

    [ValidateNever]
    public virtual PresupuestosMensuale? PresupuestoMensualRegistro { get; set; }

    [ValidateNever]
    public virtual CentrosCosto? CentroCosto { get; set; }

    [ValidateNever]
    public virtual Factura? Factura { get; set; }
  

}
