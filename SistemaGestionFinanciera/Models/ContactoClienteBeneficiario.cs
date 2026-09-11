using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace SistemaGestionFinanciera.Models
{
    [Table("ContactosClientesBeneficiarios")]
    public class ContactoClienteBeneficiario
    {
        [Key]
        public int IdContactoCliente { get; set; }

        [Required]
        public int IdClienteBeneficiario { get; set; }

        [Required]
        [StringLength(20)]
        public string TipoContacto { get; set; } = null!;

        [Required]
        [StringLength(300)]
        public string Valor { get; set; } = null!;

        public bool Activo { get; set; }

        [StringLength(50)]
        public string? Descripcion { get; set; }

        [ForeignKey(nameof(IdClienteBeneficiario))]
        public virtual ClientesBeneficiario ClienteBeneficiario { get; set; } = null!;
    }
}