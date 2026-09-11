using System;

namespace SistemaGestionFinanciera.Models.ViewModels
{
    public class MovimientoPresupuestoViewModel
    {
        public DateOnly Fecha { get; set; }

        public string Tipo { get; set; } = string.Empty;

        public string Categoria { get; set; } = string.Empty;

        public string Documento { get; set; } = string.Empty;

        public string Descripcion { get; set; } = string.Empty;

        public decimal Monto { get; set; }

        public string Origen { get; set; } = string.Empty;
    }
}