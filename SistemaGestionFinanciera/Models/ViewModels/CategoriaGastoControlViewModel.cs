using SistemaGestionFinanciera.Models;

namespace SistemaGestionFinanciera.Models.ViewModels
{
    public class CategoriaGastoControlViewModel
    {
        public CategoriasGasto Categoria { get; set; } = null!;

        public decimal GastoPeriodo { get; set; }

        public decimal PorcentajeConsumido { get; set; }

        public decimal Exceso { get; set; }

        public string EstadoLimite { get; set; } = "Sin consumo";
    }
}
