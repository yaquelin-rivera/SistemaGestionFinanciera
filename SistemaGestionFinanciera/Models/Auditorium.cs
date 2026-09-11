using System;
using System.Collections.Generic;

namespace SistemaGestionFinanciera.Models;

public partial class Auditorium
{
    public int IdAuditoria { get; set; }

    public int UsuarioId { get; set; }

    public string Tabla { get; set; } = null!;

    public int RegistroId { get; set; }

    public string Accion { get; set; } = null!;

    public string? Descripcion { get; set; }

    public DateTime Fecha { get; set; }

    public virtual Usuario Usuario { get; set; } = null!;
}
