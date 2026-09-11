using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace SistemaGestionFinanciera.Models;

public partial class PagosCuentaPorCobrar
{
    public int IdPagoCuentaPorCobrar { get; set; }

    [Required(ErrorMessage = "Debe seleccionar una cuenta por cobrar.")]
    public int CuentaPorCobrarId { get; set; }

    public DateOnly FechaPago { get; set; }

    [Required(ErrorMessage = "Debe ingresar el monto del pago.")]
    [Range(0.01, 999999999, ErrorMessage = "El monto del pago debe ser mayor a cero.")]
    public decimal? Monto { get; set; }

    [Required(ErrorMessage = "Debe seleccionar el método de pago.")]
    [StringLength(80)]
    public string? MetodoPago { get; set; }

    [StringLength(100)]
    [RegularExpression(@"^[a-zA-ZáéíóúÁÉÍÓÚñÑ0-9\s.,;:/()#-]*$",
        ErrorMessage = "La referencia contiene caracteres no permitidos.")]
    public string? Referencia { get; set; }

    [StringLength(300)]
    [RegularExpression(@"^[a-zA-ZáéíóúÁÉÍÓÚñÑ0-9\s.,;:/()#-]*$",
        ErrorMessage = "La observación contiene caracteres no permitidos.")]
    public string? Observacion { get; set; }


    [StringLength(300)]
    [RegularExpression(@"^[a-zA-ZáéíóúÁÉÍÓÚñÑ0-9\s.,;:/()#-]*$",
    ErrorMessage = "El motivo de anulación contiene caracteres no permitidos.")]
    public string? MotivoAnulacion { get; set; }
    public bool Anulado { get; set; }

    public virtual CuentasPorCobrar CuentaPorCobrar { get; set; } = null!;
}