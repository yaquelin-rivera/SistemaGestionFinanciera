using System;
using System.Collections.Generic;

namespace SistemaGestionFinanciera.Models;

public partial class PresupuestoDetalle
{
    public int IdPresupuestoDetalle { get; set; }

    public int PresupuestoMensualId { get; set; }

    public string Tipo { get; set; } = null!;

    public int? CategoriaIngresoId { get; set; }

    public int? CategoriaGastoId { get; set; }

    public decimal MontoPlanificado { get; set; }

    public decimal? MontoEjecutado { get; set; }

    public decimal? Diferencia { get; set; }

    public virtual CategoriasGasto? CategoriaGasto { get; set; }

    public virtual CategoriasIngreso? CategoriaIngreso { get; set; }

    public virtual PresupuestosMensuale PresupuestoMensual { get; set; } = null!;
}
