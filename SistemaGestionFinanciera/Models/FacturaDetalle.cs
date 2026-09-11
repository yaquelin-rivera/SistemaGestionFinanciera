using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace SistemaGestionFinanciera.Models
{
    public class FacturaDetalle
    {
        [Key]
        public int IdFacturaDetalle { get; set; }

        public int? FacturaId { get; set; }

        [ForeignKey("FacturaId")]
        public Factura? Factura { get; set; }

        [Required(ErrorMessage = "El nombre del servicio o producto es obligatorio.")]
        [StringLength(150)]
        [RegularExpression(@"^[a-zA-ZáéíóúÁÉÍÓÚñÑ0-9\s.,()-]+$",
         ErrorMessage = "Solo se permiten letras, números y signos básicos de puntuación.")]
        [Display(Name = "Servicio o producto")]
        public string Nombre { get; set; } = string.Empty;

        [StringLength(250)]
        [RegularExpression(@"^[a-zA-ZáéíóúÁÉÍÓÚñÑ0-9\s.,()-]*$",
         ErrorMessage = "La descripción contiene caracteres no permitidos.")]
        [Display(Name = "Descripción")]
        public string? Descripcion { get; set; }

        [Display(Name = "Precio unitario")]
        [Required(ErrorMessage = "El precio unitario es obligatorio.")]
        [Range(0.01, double.MaxValue, ErrorMessage = "El precio unitario debe ser mayor a cero.")]
        [RegularExpression(@"^\d+(\.\d{1,2})?$", ErrorMessage = "Ingrese un precio válido por favor use solo numeros y punto como separacion decimal, máximo dos decimales.")]
        public decimal PrecioUnitario { get; set; }
        public bool Activo { get; set; } = true;

        // Campos heredados de la tabla, no se usarán en el CRUD catálogo
        public int Cantidad { get; set; } = 1;
        public decimal Descuento { get; set; } = 0;
        public decimal Impuesto { get; set; } = 0;
        public decimal TotalLinea { get; set; }
    }
}