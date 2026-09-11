namespace SistemaGestionFinanciera.Models.ViewModels;

public class AlertaPresupuestoMenuViewModel
{
    public bool Mostrar { get; set; }

    public string Tipo { get; set; } =
        string.Empty;

    public int Cantidad { get; set; }

    public string Mensaje { get; set; } =
        string.Empty;
}
