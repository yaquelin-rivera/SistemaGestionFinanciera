using System;
using System.Collections.Generic;

namespace SistemaGestionFinanciera.Models;

public partial class PresupuestosMensuale
{
    public int IdPresupuestoMensual { get; set; }

    public int? ProyectoId { get; set; }

    public int Anio { get; set; }

    public int Mes { get; set; }

    public decimal? MontoIngresadoPlanificado { get; set; }

    public decimal? MontoGastoPlanificado { get; set; }

    public decimal? TotalIngresoReal { get; set; }

    public decimal? TotalGastoReal { get; set; }

    public decimal? DiferenciaIngreso { get; set; }

    public decimal? DiferenciaGasto { get; set; }

    public decimal? SaldoDisponible { get; set; }

    public string Estado { get; set; } = null!;

    public string EstadoAprobacion { get; set; } = "Borrador";

    public string? ObservacionRevision { get; set; }

    public DateTime? FechaCreacion { get; set; }

    public virtual ICollection<PresupuestoDetalle> PresupuestoDetalles { get; set; } = new List<PresupuestoDetalle>();

    public virtual Proyecto? Proyecto { get; set; }

    // Relación con gastos imputados a este presupuesto mensual
    public virtual ICollection<Gasto> Gastos { get; set; }
        = new List<Gasto>();

    // Relación con facturas imputadas a este presupuesto mensual
    public virtual ICollection<Factura> Facturas { get; set; }
        = new List<Factura>();

    // Relación con cuentas por pagar imputadas a este presupuesto mensual
    public virtual ICollection<CuentasPorPagar> CuentasPorPagars { get; set; }
        = new List<CuentasPorPagar>();
}
