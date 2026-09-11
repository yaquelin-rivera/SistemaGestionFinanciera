using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace SistemaGestionFinanciera.Models
{
    [Table("ContactosProveedores")]
    public class ContactoProveedor
    {
        [Key]
        public int IdContactoProveedor { get; set; }

        [Required]
        public int IdProveedor { get; set; }

        [Required]
        [StringLength(20)]
        public string TipoContacto { get; set; } = null!;

        [Required]
        [StringLength(300)]
        public string Valor { get; set; } = null!;

        public bool Activo { get; set; }

        [StringLength(50)]
        public string? Descripcion { get; set; }

        [ForeignKey(nameof(IdProveedor))]
        public virtual Proveedore Proveedor { get; set; } = null!;
    }
}