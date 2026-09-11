namespace SistemaGestionFinanciera.Models.ViewModels
{
    public class DashboardViewModel
    {
        public string? Usuario { get; set; }
        public string? Rol { get; set; }

        public decimal IngresosMes { get; set; }
        public decimal GastosMes { get; set; }

        public decimal CuentasPorCobrarPendientes { get; set; }
        public decimal CuentasPorPagarPendientes { get; set; }

        public int FacturasMes { get; set; }
        public int MovimientosContablesMes { get; set; }

        public decimal PresupuestoIngresosPlanificado { get; set; }
        public decimal PresupuestoGastosPlanificado { get; set; }
        public decimal PresupuestoIngresosReal { get; set; }
        public decimal PresupuestoGastosReal { get; set; }

        public int CuentasPorCobrarVencidas { get; set; }
        public int CuentasPorPagarVencidas { get; set; }

        public int AuditoriasHoy { get; set; }

        public List<MovimientoDashboardItem> UltimosMovimientos { get; set; } = new();
        public List<AuditoriaDashboardItem> UltimasAuditorias { get; set; } = new();

        public List<string> Meses { get; set; } = new();

        public List<decimal> IngresosMensuales { get; set; } = new();

        public List<decimal> GastosMensuales { get; set; } = new();

        public int CuentasPorCobrarPendientesCantidad { get; set; }

        public int CuentasPorCobrarPagadasCantidad { get; set; }

        public int CuentasPorCobrarVencidasCantidad { get; set; }
    }

    public class MovimientoDashboardItem
    {
        public DateOnly Fecha { get; set; }
        public string? Modulo { get; set; }
        public string? Descripcion { get; set; }
        public decimal Debe { get; set; }
        public decimal Haber { get; set; }
    }

    public class AuditoriaDashboardItem
    {
        public DateTime Fecha { get; set; }
        public string? Modulo { get; set; }
        public string? Accion { get; set; }
        public string? Usuario { get; set; }
    }
}