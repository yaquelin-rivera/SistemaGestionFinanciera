using System;
using System.ComponentModel.DataAnnotations;

namespace SistemaGestionFinanciera.Models
{
    public class CierreContable
    {
        [Key]
        public int IdCierreContable { get; set; }

        [Required]
        [Range(2000, 2100)]
        public int Anio { get; set; }

        [Required]
        [Range(1, 12)]
        public int Mes { get; set; }

        [Required]
        [StringLength(20)]
        public string Estado { get; set; } = "Cerrado";

        [Range(0, 999999999999.99)]
        public decimal TotalIngresos { get; set; }

        [Range(0, 999999999999.99)]
        public decimal TotalGastos { get; set; }

        public decimal ResultadoPeriodo { get; set; }

        [Range(0, 999999999999.99)]
        public decimal TotalDebe { get; set; }

        [Range(0, 999999999999.99)]
        public decimal TotalHaber { get; set; }

        [Range(0, int.MaxValue)]
        public int CantidadMovimientos { get; set; }

        public DateTime FechaCierre { get; set; } = DateTime.Now;

        [Required]
        public int UsuarioCierreId { get; set; }

        public DateTime? FechaReapertura { get; set; }

        public int? UsuarioReaperturaId { get; set; }

        [StringLength(500)]
        public string? MotivoReapertura { get; set; }

        // Usuario que realizó el cierre.
        public virtual Usuario UsuarioCierre { get; set; } = null!;

        // Usuario que realizó la última reapertura.
        public virtual Usuario? UsuarioReapertura { get; set; }
    }
}