using System;
using System.Collections.Generic;
using Microsoft.EntityFrameworkCore;
using SistemaGestionFinanciera.Models;

namespace SistemaGestionFinanciera.Data;

public partial class SistemaFinancieroContext : DbContext
{
    public SistemaFinancieroContext()
    {
    }

    public SistemaFinancieroContext(DbContextOptions<SistemaFinancieroContext> options)
        : base(options)
    {
    }

    public virtual DbSet<Auditorium> Auditoria { get; set; }

    public virtual DbSet<CategoriasGasto> CategoriasGastos { get; set; }

    public virtual DbSet<CategoriasIngreso> CategoriasIngresos { get; set; }

    public virtual DbSet<CentrosCosto> CentrosCostos { get; set; }

    public virtual DbSet<ClientesBeneficiario> ClientesBeneficiarios { get; set; }

    public virtual DbSet<ContactoClienteBeneficiario>
    ContactosClientesBeneficiarios
    { get; set; }

    public virtual DbSet<CuentasContable> CuentasContables { get; set; }
    public virtual DbSet<CierreContable> CierresContables { get; set; }
    public virtual DbSet<CuentasPorCobrar> CuentasPorCobrars { get; set; }

    public virtual DbSet<CuentasPorPagar> CuentasPorPagars { get; set; }

    public virtual DbSet<Factura> Facturas { get; set; }

    public virtual DbSet<FacturaDetalle> FacturaDetalles { get; set; }

    public virtual DbSet<Gasto> Gastos { get; set; }

    public virtual DbSet<Ingreso> Ingresos { get; set; }

    public virtual DbSet<MovimientosContable> MovimientosContables { get; set; }

    public virtual DbSet<PagosCuentaPorCobrar> PagosCuentaPorCobrars { get; set; }

    public virtual DbSet<PagosCuentaPorPagar> PagosCuentaPorPagars { get; set; }

    public virtual DbSet<Permiso> Permisos { get; set; }

    public virtual DbSet<PresupuestoDetalle> PresupuestoDetalles { get; set; }

    public virtual DbSet<PresupuestosMensuale> PresupuestosMensuales { get; set; }

    public virtual DbSet<Proveedore> Proveedores { get; set; }

    public virtual DbSet<ContactoProveedor> ContactosProveedores { get; set; }

    public virtual DbSet<Proyecto> Proyectos { get; set; }

    public virtual DbSet<Role> Roles { get; set; }

    public virtual DbSet<RolesPermiso> RolesPermisos { get; set; }

    public virtual DbSet<Usuario> Usuarios { get; set; }

    protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
#warning To protect potentially sensitive information in your connection string, you should move it out of source code. You can avoid scaffolding the connection string by using the Name= syntax to read it from configuration - see https://go.microsoft.com/fwlink/?linkid=2131148. For more guidance on storing connection strings, see https://go.microsoft.com/fwlink/?LinkId=723263.
        => optionsBuilder.UseSqlServer("Server=YAQUELIN-R\\SQLEXPRESS;Database=SistemaFinancieroUCAD;Integrated Security=True;TrustServerCertificate=True;");

    protected override void OnModelCreating(ModelBuilder modelBuilder)
    {
        modelBuilder.Entity<Auditorium>(entity =>
        {
            entity.HasKey(e => e.IdAuditoria).HasName("PK__Auditori__7FD13FA0B8BDF8DD");

            entity.HasIndex(e => e.UsuarioId, "IX_Auditoria_UsuarioId");

            entity.Property(e => e.Accion)
                .HasMaxLength(50)
                .IsUnicode(false);
            entity.Property(e => e.Descripcion)
                .HasMaxLength(500)
                .IsUnicode(false);
            entity.Property(e => e.Fecha)
                .HasDefaultValueSql("(getdate())")
                .HasColumnType("datetime");
            entity.Property(e => e.Tabla)
                .HasMaxLength(100)
                .IsUnicode(false);

            entity.HasOne(d => d.Usuario).WithMany(p => p.Auditoria)
                .HasForeignKey(d => d.UsuarioId)
                .OnDelete(DeleteBehavior.ClientSetNull)
                .HasConstraintName("FK_Auditoria_Usuarios");
        });

        modelBuilder.Entity<CategoriasGasto>(entity =>
        {
            entity.HasKey(e => e.IdCategoriaGasto).HasName("PK__Categori__59627481A218F5FE");

            entity.ToTable("CategoriasGasto");

            entity.Property(e => e.Activo).HasDefaultValue(true);
            entity.Property(e => e.Descripcion)
                .HasMaxLength(300)
                .IsUnicode(false);
            entity.Property(e => e.LimiteMensual).HasColumnType("decimal(18, 2)");
            entity.Property(e => e.Nombre)
                .HasMaxLength(120)
                .IsUnicode(false);
        });

        modelBuilder.Entity<CategoriasIngreso>(entity =>
        {
            entity.HasKey(e => e.IdCategoriaIngreso).HasName("PK__Categori__65B260F2B3BCC975");

            entity.ToTable("CategoriasIngreso");

            entity.Property(e => e.Activo).HasDefaultValue(true);
            entity.Property(e => e.Descripcion)
                .HasMaxLength(300)
                .IsUnicode(false);
            entity.Property(e => e.Nombre)
                .HasMaxLength(120)
                .IsUnicode(false);
        });

        modelBuilder.Entity<CentrosCosto>(entity =>
        {
            entity.HasKey(e => e.IdCentroCosto).HasName("PK__CentrosC__EE3651E84076DB7F");

            entity.ToTable("CentrosCosto");

            entity.Property(e => e.Activo).HasDefaultValue(true);
            entity.Property(e => e.Codigo)
                .HasMaxLength(50)
                .IsUnicode(false);
            entity.Property(e => e.Descripcion)
                .HasMaxLength(300)
                .IsUnicode(false);
            entity.Property(e => e.Nombre)
                .HasMaxLength(150)
                .IsUnicode(false);
        });

        modelBuilder.Entity<ClientesBeneficiario>(entity =>
        {
            entity.HasKey(e => e.IdClienteBeneficiario).HasName("PK__Clientes__BBE629D40194A73E");

            entity.Property(e => e.Activo).HasDefaultValue(true);
            entity.Property(e => e.Direccion)
                .HasMaxLength(300)
                .IsUnicode(false);
            entity.Property(e => e.Email)
                .HasMaxLength(150)
                .IsUnicode(false);
            entity.Property(e => e.Identificacion)
                .HasMaxLength(50)
                .IsUnicode(false);
            entity.Property(e => e.Nombre)
                .HasMaxLength(150)
                .IsUnicode(false);
            entity.Property(e => e.Telefono)
                .HasMaxLength(50)
                .IsUnicode(false);
            entity.Property(e => e.Tipo)
                .HasMaxLength(50)
                .IsUnicode(false);
        });

        modelBuilder.Entity<CuentasContable>(entity =>
        {
            entity.HasKey(e => e.IdCuentaContable).HasName("PK__CuentasC__458CB9B277560822");

            entity.Property(e => e.Activo).HasDefaultValue(true);
            entity.Property(e => e.Codigo)
                .HasMaxLength(50)
                .IsUnicode(false);
            entity.Property(e => e.FechaCreacion)
                .HasDefaultValueSql("(getdate())")
                .HasColumnType("datetime");
            entity.Property(e => e.Naturaleza)
                .HasMaxLength(20)
                .IsUnicode(false);
            entity.Property(e => e.Nombre)
                .HasMaxLength(150)
                .IsUnicode(false);
            entity.Property(e => e.SaldoActual).HasColumnType("decimal(18, 2)");
            entity.Property(e => e.SaldoInicial).HasColumnType("decimal(18, 2)");
            entity.Property(e => e.TipoCuenta)
                .HasMaxLength(50)
                .IsUnicode(false);
        });

        modelBuilder.Entity<CierreContable>(entity =>
        {
            entity.HasKey(e => e.IdCierreContable)
                .HasName("PK_CierresContables");

            entity.HasIndex(e => new { e.Anio, e.Mes })
                .IsUnique()
                .HasDatabaseName("UQ_CierresContables_AnioMes");

            entity.Property(e => e.Estado)
                .HasMaxLength(20)
                .IsUnicode(false)
                .HasDefaultValue("Cerrado");

            entity.Property(e => e.TotalIngresos)
                .HasColumnType("decimal(18, 2)")
                .HasDefaultValue(0m);

            entity.Property(e => e.TotalGastos)
                .HasColumnType("decimal(18, 2)")
                .HasDefaultValue(0m);

            entity.Property(e => e.ResultadoPeriodo)
                .HasColumnType("decimal(18, 2)")
                .HasDefaultValue(0m);

            entity.Property(e => e.TotalDebe)
                .HasColumnType("decimal(18, 2)")
                .HasDefaultValue(0m);

            entity.Property(e => e.TotalHaber)
                .HasColumnType("decimal(18, 2)")
                .HasDefaultValue(0m);

            entity.Property(e => e.CantidadMovimientos)
                .HasDefaultValue(0);

            entity.Property(e => e.FechaCierre)
                .HasColumnType("datetime")
                .HasDefaultValueSql("(getdate())");

            entity.Property(e => e.FechaReapertura)
                .HasColumnType("datetime");

            entity.Property(e => e.MotivoReapertura)
                .HasMaxLength(500)
                .IsUnicode(false);

            entity.HasOne(d => d.UsuarioCierre)
                .WithMany()
                .HasForeignKey(d => d.UsuarioCierreId)
                .OnDelete(DeleteBehavior.Restrict)
                .HasConstraintName("FK_CierresContables_UsuarioCierre");

            entity.HasOne(d => d.UsuarioReapertura)
                .WithMany()
                .HasForeignKey(d => d.UsuarioReaperturaId)
                .OnDelete(DeleteBehavior.Restrict)
                .HasConstraintName("FK_CierresContables_UsuarioReapertura");
        });

        modelBuilder.Entity<CuentasPorCobrar>(entity =>
        {
            entity.HasKey(e => e.IdCuentaPorCobrar).HasName("PK__CuentasP__17F592ACED90923B");

            entity.ToTable("CuentasPorCobrar");

            entity.HasIndex(e => e.ClienteBeneficiarioId, "IX_CuentasPorCobrar_ClienteBeneficiarioId");

            entity.HasIndex(e => e.ProyectoId, "IX_CuentasPorCobrar_ProyectoId");

            entity.HasIndex(e => e.FacturaId, "IX_CuentasPorCobrar_FacturaId");

            entity.Property(e => e.Activo).HasDefaultValue(true);
            entity.Property(e => e.AnticipoAplicado).HasColumnType("decimal(18, 2)");
            entity.Property(e => e.Concepto)
                .HasMaxLength(250)
                .IsUnicode(false);
            entity.Property(e => e.Estado)
                .HasMaxLength(30)
                .IsUnicode(false);
            entity.Property(e => e.MontoOriginal).HasColumnType("decimal(18, 2)");
            entity.Property(e => e.MontoPagado).HasColumnType("decimal(18, 2)");
            entity.Property(e => e.NumeroDocumento)
                .HasMaxLength(100)
                .IsUnicode(false);
            entity.Property(e => e.Observacion)
                .HasMaxLength(500)
                .IsUnicode(false);
            entity.Property(e => e.SaldoPendiente).HasColumnType("decimal(18, 2)");

            entity.Property(e => e.TipoDocumento)
                .HasMaxLength(50)
                .IsUnicode(false);

            // Guardamos los días que tiene el cliente para pagar
            entity.Property(e => e.DiasCredito);

            entity.HasOne(d => d.ClienteBeneficiario).WithMany(p => p.CuentasPorCobrars)
                .HasForeignKey(d => d.ClienteBeneficiarioId)
                .OnDelete(DeleteBehavior.ClientSetNull)
                .HasConstraintName("FK_CuentasPorCobrar_ClientesBeneficiarios");

            entity.HasOne(d => d.Proyecto).WithMany(p => p.CuentasPorCobrars)
                .HasForeignKey(d => d.ProyectoId)
                .HasConstraintName("FK_CuentasPorCobrar_Proyectos");

            entity.HasOne(d => d.PresupuestoMensualRegistro)
    .WithMany()
    .HasForeignKey(d => d.PresupuestoMensualId)
    .OnDelete(DeleteBehavior.Restrict)
    .HasConstraintName("FK_CuentasPorCobrar_PresupuestosMensuales");

            entity.HasIndex(
    e => e.PresupuestoMensualId,
    "IX_CuentasPorCobrar_PresupuestoMensualId");

            entity.HasOne(d => d.Factura)
                .WithMany()
                .HasForeignKey(d => d.FacturaId)
                .HasConstraintName("FK_CuentasPorCobrar_Facturas");
        });


        modelBuilder.Entity<CuentasPorPagar>(entity =>
        {
            entity.HasKey(e => e.IdCuentaPorPagar).HasName("PK__CuentasP__051712617811B11B");

            entity.ToTable("CuentasPorPagar");

            entity.HasIndex(e => e.ProveedorId, "IX_CuentasPorPagar_ProveedorId");

            entity.HasIndex(e => e.ProyectoId, "IX_CuentasPorPagar_ProyectoId");

            entity.Property(e => e.Activo).HasDefaultValue(true);
            entity.Property(e => e.AnticipoAplicado).HasColumnType("decimal(18, 2)");
            entity.Property(e => e.Concepto)
                .HasMaxLength(250)
                .IsUnicode(false);
            entity.Property(e => e.Descuento).HasColumnType("decimal(18, 2)");
            entity.Property(e => e.Estado)
                .HasMaxLength(30)
                .IsUnicode(false);
            entity.Property(e => e.Impuesto).HasColumnType("decimal(18, 2)");
            entity.Property(e => e.MontoOriginal).HasColumnType("decimal(18, 2)");
            entity.Property(e => e.MontoPagado).HasColumnType("decimal(18, 2)");
            entity.Property(e => e.TipoDocumento)
    .HasMaxLength(50)
             .IsUnicode(false);

            entity.Property(e => e.NumeroDocumento)
                .HasMaxLength(100)
                .IsUnicode(false);

            entity.Property(e => e.DiasCredito);
            entity.Property(e => e.NumeroFactura)
                .HasMaxLength(100)
                .IsUnicode(false);
            entity.Property(e => e.Observacion)
                .HasMaxLength(500)
                .IsUnicode(false);
            entity.Property(e => e.SaldoPendiente).HasColumnType("decimal(18, 2)");
            entity.Property(e => e.SubTotal).HasColumnType("decimal(18, 2)");

            entity.Property(e => e.CategoriaGastoId);

            entity.Property(e => e.CentroCostoId);

            entity.Property(e => e.TipoImpuesto)
                .HasMaxLength(50)
                 .IsUnicode(false);

            entity.Property(e => e.MotivoAnulacion)
                .HasMaxLength(500)
                 .IsUnicode(false);

            entity.HasOne(d => d.Proveedor).WithMany(p => p.CuentasPorPagars)
                .HasForeignKey(d => d.ProveedorId)
                .OnDelete(DeleteBehavior.ClientSetNull)
                .HasConstraintName("FK_CuentasPorPagar_Proveedores");

            entity.HasOne(d => d.Proyecto).WithMany(p => p.CuentasPorPagars)
                .HasForeignKey(d => d.ProyectoId)
                .HasConstraintName("FK_CuentasPorPagar_Proyectos");

            entity.HasOne(d => d.CategoriaGasto)
       .WithMany()
       .HasForeignKey(d => d.CategoriaGastoId)
       .HasConstraintName("FK_CuentasPorPagar_CategoriasGasto");

            entity.HasOne(d => d.CentroCosto)
                .WithMany()
                .HasForeignKey(d => d.CentroCostoId)
                .HasConstraintName("FK_CuentasPorPagar_CentrosCosto");

            entity.HasIndex(
    e => e.PresupuestoMensualId,
    "IX_CuentasPorPagar_PresupuestoMensualId");

            entity.HasOne(d => d.PresupuestoMensualRegistro)
                .WithMany(p => p.CuentasPorPagars)
                .HasForeignKey(d => d.PresupuestoMensualId)
                .OnDelete(DeleteBehavior.Restrict)
                .HasConstraintName(
                    "FK_CuentasPorPagar_PresupuestosMensuales");

        });

        modelBuilder.Entity<Factura>(entity =>
        {
            entity.HasKey(e => e.IdFactura).HasName("PK__Facturas__50E7BAF1C042CCAE");

            entity.HasIndex(e => e.ClienteBeneficiarioId, "IX_Facturas_ClienteBeneficiarioId");

            entity.HasIndex(e => e.ProveedorId, "IX_Facturas_ProveedorId");

            entity.Property(e => e.DescuentoTotal).HasColumnType("decimal(18, 2)");
            entity.Property(e => e.Estado)
                .HasMaxLength(30)
                .IsUnicode(false);
            entity.Property(e => e.ImpuestoTotal).HasColumnType("decimal(18, 2)");
            entity.Property(e => e.NumeroFactura)
                .HasMaxLength(100)
                .IsUnicode(false);
            entity.Property(e => e.Observacion)
                .HasMaxLength(500)
                .IsUnicode(false);
            entity.Property(e => e.SaldoPendiente).HasColumnType("decimal(18, 2)");
            entity.Property(e => e.SubTotal).HasColumnType("decimal(18, 2)");
            entity.Property(e => e.TipoFactura)
                .HasMaxLength(30)
                .IsUnicode(false);
            entity.Property(e => e.Total).HasColumnType("decimal(18, 2)");

            entity.HasOne(d => d.ClienteBeneficiario).WithMany(p => p.Facturas)
                .HasForeignKey(d => d.ClienteBeneficiarioId)
                .HasConstraintName("FK_Facturas_ClientesBeneficiarios");

            entity.HasOne(d => d.Proveedor).WithMany(p => p.Facturas)
                .HasForeignKey(d => d.ProveedorId)
                .HasConstraintName("FK_Facturas_Proveedores");

            entity.HasIndex(
    e => e.PresupuestoMensualId,
    "IX_Facturas_PresupuestoMensualId");

            entity.HasOne(d => d.PresupuestoMensualRegistro)
                .WithMany(p => p.Facturas)
                .HasForeignKey(d => d.PresupuestoMensualId)
                .OnDelete(DeleteBehavior.Restrict)
                .HasConstraintName(
                    "FK_Facturas_PresupuestosMensuales");
        });

        modelBuilder.Entity<FacturaDetalle>(entity =>
        {
            entity.HasKey(e => e.IdFacturaDetalle).HasName("PK__FacturaD__3D8E1AB80C80BEAF");

            entity.HasIndex(e => e.FacturaId, "IX_FacturaDetalles_FacturaId");

            entity.Property(e => e.Cantidad).HasColumnType("decimal(18, 2)");
            entity.Property(e => e.Descripcion)
                .HasMaxLength(250)
                .IsUnicode(false);
            entity.Property(e => e.Descuento).HasColumnType("decimal(18, 2)");
            entity.Property(e => e.Impuesto).HasColumnType("decimal(18, 2)");
            entity.Property(e => e.PrecioUnitario).HasColumnType("decimal(18, 2)");
            entity.Property(e => e.TotalLinea).HasColumnType("decimal(18, 2)");

            entity.HasOne(d => d.Factura).WithMany(p => p.FacturaDetalles)
                .HasForeignKey(d => d.FacturaId)
                .OnDelete(DeleteBehavior.ClientSetNull)
                .HasConstraintName("FK_FacturaDetalles_Facturas");
        });

        modelBuilder.Entity<Gasto>(entity =>
        {
            entity.HasKey(e => e.IdGasto).HasName("PK__Gastos__C630244D2FCC12DB");

            entity.HasIndex(e => e.CategoriaGastoId, "IX_Gastos_CategoriaGastoId");

            entity.HasIndex(e => e.ProveedorId, "IX_Gastos_ProveedorId");

            entity.HasIndex(e => e.ProyectoId, "IX_Gastos_ProyectoId");

            entity.Property(e => e.Activo).HasDefaultValue(true);
            entity.Property(e => e.Concepto)
                .HasMaxLength(250)
                .IsUnicode(false);
            entity.Property(e => e.Descuento).HasColumnType("decimal(18, 2)");
            entity.Property(e => e.Estado)
                .HasMaxLength(30)
                .IsUnicode(false);
            entity.Property(e => e.Impuesto).HasColumnType("decimal(18, 2)");
            entity.Property(e => e.MontoTotal).HasColumnType("decimal(18, 2)");
            entity.Property(e => e.Observacion)
                .HasMaxLength(500)
                .IsUnicode(false);
            entity.Property(e => e.Subtotal).HasColumnType("decimal(18, 2)");
            entity.Property(e => e.SuperaPresupuesto).HasDefaultValue(false);

            entity.HasOne(d => d.CategoriaGasto).WithMany(p => p.Gastos)
                .HasForeignKey(d => d.CategoriaGastoId)
                .OnDelete(DeleteBehavior.ClientSetNull)
                .HasConstraintName("FK_Gastos_CategoriasGasto");

            entity.HasOne(d => d.Proveedor).WithMany(p => p.Gastos)
                .HasForeignKey(d => d.ProveedorId)
                .HasConstraintName("FK_Gastos_Proveedores");

            entity.HasOne(d => d.Proyecto).WithMany(p => p.Gastos)
                .HasForeignKey(d => d.ProyectoId)
                .HasConstraintName("FK_Gastos_Proyectos");

            entity.HasOne(d => d.CentroCosto).WithMany()
                .HasForeignKey(d => d.CentroCostoId)
                .HasConstraintName("FK_Gastos_CentrosCosto");

            entity.HasIndex(
    e => e.PresupuestoMensualId,
    "IX_Gastos_PresupuestoMensualId");

            entity.HasOne(d => d.PresupuestoMensualRegistro)
                .WithMany(p => p.Gastos)
                .HasForeignKey(d => d.PresupuestoMensualId)
                .OnDelete(DeleteBehavior.Restrict)
                .HasConstraintName(
                    "FK_Gastos_PresupuestosMensuales");
        });

        modelBuilder.Entity<Ingreso>(entity =>
        {
            entity.HasKey(e => e.IdIngreso).HasName("PK__Ingresos__901EF2E3DAEDF58C");

            entity.HasIndex(e => e.CategoriaIngresoId, "IX_Ingresos_CategoriaIngresoId");

            entity.HasIndex(e => e.ClienteBeneficiarioId, "IX_Ingresos_ClienteBeneficiarioId");

            entity.HasIndex(e => e.ProyectoId, "IX_Ingresos_ProyectoId");

            entity.Property(e => e.Activo).HasDefaultValue(true);
            entity.Property(e => e.Descripcion)
                .HasMaxLength(500)
                .IsUnicode(false);
            entity.Property(e => e.Diferencia).HasColumnType("decimal(18, 2)");
            entity.Property(e => e.EsAnticipo).HasDefaultValue(false);
            entity.Property(e => e.Estado)
                .HasMaxLength(30)
                .IsUnicode(false);
            entity.Property(e => e.Fuente)
                .HasMaxLength(80)
                .IsUnicode(false);
            entity.Property(e => e.MontoEsperado).HasColumnType("decimal(18, 2)");
            entity.Property(e => e.MontoReal).HasColumnType("decimal(18, 2)");

            entity.HasOne(d => d.CategoriaIngreso).WithMany(p => p.Ingresos)
                .HasForeignKey(d => d.CategoriaIngresoId)
                .OnDelete(DeleteBehavior.ClientSetNull)
                .HasConstraintName("FK_Ingresos_CategoriasIngreso");

            entity.HasOne(d => d.ClienteBeneficiario).WithMany(p => p.Ingresos)
                .HasForeignKey(d => d.ClienteBeneficiarioId)
                .HasConstraintName("FK_Ingresos_ClientesBeneficiarios");

            entity.HasOne(d => d.Proyecto).WithMany(p => p.Ingresos)
                .HasForeignKey(d => d.ProyectoId)
                .HasConstraintName("FK_Ingresos_Proyectos");

            entity.HasOne(d => d.PresupuestoMensualRegistro)
    .WithMany()
    .HasForeignKey(d => d.PresupuestoMensualId)
    .OnDelete(DeleteBehavior.Restrict)
    .HasConstraintName("FK_Ingresos_PresupuestosMensuales");

            entity.HasIndex(
    e => e.PresupuestoMensualId,
    "IX_Ingresos_PresupuestoMensualId");

            entity.Property(e => e.MotivoPendientePresupuestario)
    .HasMaxLength(250)
    .IsUnicode(false);

            entity.Property(e => e.Origen)
    .HasMaxLength(30)
    .IsUnicode(false)
    .HasDefaultValue("Manual");

            // Relación entre ingreso y factura.
            // Permite identificar ingresos generados automáticamente desde facturación.
            entity.HasOne(d => d.Factura)
                .WithMany()
                .HasForeignKey(d => d.FacturaId)
                .HasConstraintName("FK_Ingresos_Facturas");
        });

        modelBuilder.Entity<MovimientosContable>(entity =>
        {
            entity.HasKey(e => e.IdMovimientoContable).HasName("PK__Movimien__1A1FB3AC247188D6");

            entity.HasIndex(e => e.CentroCostoId, "IX_MovimientosContables_CentroCostoId");
            entity.HasIndex(e => e.CuentaContableId, "IX_MovimientosContables_CuentaContableId");
            entity.HasIndex(e => e.ProyectoId, "IX_MovimientosContables_ProyectoId");
            entity.HasIndex(e => e.UsuarioId, "IX_MovimientosContables_UsuarioId");
            entity.HasIndex(e => e.MovimientoReversionId, "IX_MovimientosContables_MovimientoReversionId");

            entity.Property(e => e.Anulado).HasDefaultValue(false);

            entity.Property(e => e.Descripcion)
                .HasMaxLength(500)
                .IsUnicode(false);

            entity.Property(e => e.FechaCreacion)
                .HasDefaultValueSql("(getdate())")
                .HasColumnType("datetime");

            entity.Property(e => e.FechaAnulacion)
                .HasColumnType("datetime");

            entity.Property(e => e.Monto)
                .HasColumnType("decimal(18, 2)");

            entity.Property(e => e.Debe)
                .HasColumnType("decimal(18, 2)")
                .HasDefaultValue(0);

            entity.Property(e => e.Haber)
                .HasColumnType("decimal(18, 2)")
                .HasDefaultValue(0);

            entity.Property(e => e.OrigenModulo)
                .HasMaxLength(80)
                .IsUnicode(false);

            entity.Property(e => e.Referencia)
                .HasMaxLength(100)
                .IsUnicode(false);

            entity.Property(e => e.TipoMovimiento)
                .HasMaxLength(20)
                .IsUnicode(false);

            entity.Property(e => e.Estado)
                .HasMaxLength(30)
                .IsUnicode(false)
                .HasDefaultValue("Registrado");

            entity.Property(e => e.EsAutomatico)
                .HasDefaultValue(true);

            entity.Property(e => e.MotivoAnulacion)
                .HasMaxLength(500);

            entity.HasOne(d => d.CentroCosto)
                .WithMany(p => p.MovimientosContables)
                .HasForeignKey(d => d.CentroCostoId)
                .OnDelete(DeleteBehavior.ClientSetNull)
                .HasConstraintName("FK_MovimientosContables_CentrosCosto");

            entity.HasOne(d => d.CuentaContable)
                .WithMany(p => p.MovimientosContables)
                .HasForeignKey(d => d.CuentaContableId)
                .OnDelete(DeleteBehavior.ClientSetNull)
                .HasConstraintName("FK_MovimientosContables_CuentasContables");

            entity.HasOne(d => d.Proyecto)
                .WithMany(p => p.MovimientosContables)
                .HasForeignKey(d => d.ProyectoId)
                .OnDelete(DeleteBehavior.ClientSetNull)
                .HasConstraintName("FK_MovimientosContables_Proyectos");

            entity.HasOne(d => d.Usuario)
                .WithMany()
                .HasForeignKey(d => d.UsuarioId)
                .OnDelete(DeleteBehavior.ClientSetNull)
                .HasConstraintName("FK_MovimientosContables_Usuarios");

            entity.HasOne(d => d.MovimientoReversion)
                .WithMany(p => p.MovimientosRevertidos)
                .HasForeignKey(d => d.MovimientoReversionId)
                .OnDelete(DeleteBehavior.ClientSetNull)
                .HasConstraintName("FK_MovimientosContables_Reversion");
        });

        modelBuilder.Entity<PagosCuentaPorCobrar>(entity =>
        {
            entity.HasKey(e => e.IdPagoCuentaPorCobrar).HasName("PK__PagosCue__B4A116BFD26F01CE");

            entity.ToTable("PagosCuentaPorCobrar");

            entity.HasIndex(e => e.CuentaPorCobrarId, "IX_PagosCuentaPorCobrar_CuentaPorCobrarId");

            entity.Property(e => e.MetodoPago)
                .HasMaxLength(80)
                .IsUnicode(false);
            entity.Property(e => e.Monto).HasColumnType("decimal(18, 2)");
            entity.Property(e => e.Observacion)
                .HasMaxLength(300)
                .IsUnicode(false);
            entity.Property(e => e.Referencia)
                .HasMaxLength(100)
                .IsUnicode(false);

            entity.HasOne(d => d.CuentaPorCobrar).WithMany(p => p.PagosCuentaPorCobrars)
                .HasForeignKey(d => d.CuentaPorCobrarId)
                .OnDelete(DeleteBehavior.ClientSetNull)
                .HasConstraintName("FK_PagosCuentaPorCobrar_CuentasPorCobrar");
        });

        modelBuilder.Entity<PagosCuentaPorPagar>(entity =>
        {
            entity.HasKey(e => e.IdPagoCuentaPorPagar).HasName("PK__PagosCue__3878B28EEF07C973");

            entity.ToTable("PagosCuentaPorPagar");

            entity.HasIndex(e => e.CuentaPorPagarId, "IX_PagosCuentaPorPagar_CuentaPorPagarId");

            entity.Property(e => e.MetodoPago)
                .HasMaxLength(80)
                .IsUnicode(false);
            entity.Property(e => e.Monto).HasColumnType("decimal(18, 2)");
            entity.Property(e => e.Observacion)
                .HasMaxLength(300)
                .IsUnicode(false);
            entity.Property(e => e.Referencia)
                .HasMaxLength(100)
                .IsUnicode(false);

            entity.HasOne(d => d.CuentaPorPagar).WithMany(p => p.PagosCuentaPorPagars)
                .HasForeignKey(d => d.CuentaPorPagarId)
                .OnDelete(DeleteBehavior.ClientSetNull)
                .HasConstraintName("FK_PagosCuentaPorPagar_CuentasPorPagar");
        });

        modelBuilder.Entity<Permiso>(entity =>
        {
            entity.HasKey(e => e.IdPermiso).HasName("PK__Permisos__0D626EC851BAA1F9");

            entity.Property(e => e.Activo).HasDefaultValue(true);
            entity.Property(e => e.Descripcion)
                .HasMaxLength(250)
                .IsUnicode(false);
            entity.Property(e => e.FechaCreacion)
                .HasDefaultValueSql("(getdate())")
                .HasColumnType("datetime");
            entity.Property(e => e.Nombre)
                .HasMaxLength(100)
                .IsUnicode(false);
        });

        modelBuilder.Entity<PresupuestoDetalle>(entity =>
        {
            entity.HasKey(e => e.IdPresupuestoDetalle).HasName("PK__Presupue__032767A5EB14F516");

            entity.HasIndex(e => e.CategoriaGastoId, "IX_PresupuestoDetalles_CategoriaGastoId");

            entity.HasIndex(e => e.CategoriaIngresoId, "IX_PresupuestoDetalles_CategoriaIngresoId");

            entity.HasIndex(e => e.PresupuestoMensualId, "IX_PresupuestoDetalles_PresupuestoMensualId");

            entity.Property(e => e.Diferencia).HasColumnType("decimal(18, 2)");
            entity.Property(e => e.MontoEjecutado).HasColumnType("decimal(18, 2)");
            entity.Property(e => e.MontoPlanificado).HasColumnType("decimal(18, 2)");
            entity.Property(e => e.Tipo)
                .HasMaxLength(20)
                .IsUnicode(false);

            entity.HasOne(d => d.CategoriaGasto).WithMany(p => p.PresupuestoDetalles)
                .HasForeignKey(d => d.CategoriaGastoId)
                .HasConstraintName("FK_PresupuestoDetalles_CategoriasGasto");

            entity.HasOne(d => d.CategoriaIngreso).WithMany(p => p.PresupuestoDetalles)
                .HasForeignKey(d => d.CategoriaIngresoId)
                .HasConstraintName("FK_PresupuestoDetalles_CategoriasIngreso");

            entity.HasOne(d => d.PresupuestoMensual).WithMany(p => p.PresupuestoDetalles)
                .HasForeignKey(d => d.PresupuestoMensualId)
                .OnDelete(DeleteBehavior.ClientSetNull)
                .HasConstraintName("FK_PresupuestoDetalles_PresupuestosMensuales");
        });

        modelBuilder.Entity<PresupuestosMensuale>(entity =>
        {
            entity.HasKey(e => e.IdPresupuestoMensual).HasName("PK__Presupue__CDD0ECFF41243C76");

            entity.HasIndex(e => e.ProyectoId, "IX_PresupuestosMensuales_ProyectoId");

            entity.Property(e => e.DiferenciaGasto).HasColumnType("decimal(18, 2)");
            entity.Property(e => e.DiferenciaIngreso).HasColumnType("decimal(18, 2)");
            entity.Property(e => e.Estado)
                .HasMaxLength(30)
                .IsUnicode(false);

            entity.Property(e => e.EstadoAprobacion)
    .HasMaxLength(20)
    .HasDefaultValue("Borrador");

            entity.Property(e => e.ObservacionRevision)
                .HasMaxLength(500);

            entity.Property(e => e.FechaCreacion)
                .HasDefaultValueSql("(getdate())")
                .HasColumnType("datetime");
            entity.Property(e => e.MontoGastoPlanificado).HasColumnType("decimal(18, 2)");
            entity.Property(e => e.MontoIngresadoPlanificado).HasColumnType("decimal(18, 2)");
            entity.Property(e => e.SaldoDisponible).HasColumnType("decimal(18, 2)");
            entity.Property(e => e.TotalGastoReal).HasColumnType("decimal(18, 2)");
            entity.Property(e => e.TotalIngresoReal).HasColumnType("decimal(18, 2)");

            entity.HasOne(d => d.Proyecto).WithMany(p => p.PresupuestosMensuales)
                .HasForeignKey(d => d.ProyectoId)
                .HasConstraintName("FK_PresupuestosMensuales_Proyectos");
        });

        modelBuilder.Entity<Proveedore>(entity =>
        {
            entity.HasKey(e => e.IdProveedor).HasName("PK__Proveedo__E8B631AFEFF953F8");

            entity.Property(e => e.Activo).HasDefaultValue(true);
            entity.Property(e => e.Direccion)
                .HasMaxLength(300)
                .IsUnicode(false);
            entity.Property(e => e.Email)
                .HasMaxLength(150)
                .IsUnicode(false);
            entity.Property(e => e.Identificacion)
                .HasMaxLength(50)
                .IsUnicode(false);
            entity.Property(e => e.Nombre)
                .HasMaxLength(150)
                .IsUnicode(false);
            entity.Property(e => e.Telefono)
                .HasMaxLength(50)
                .IsUnicode(false);
        });
        modelBuilder.Entity<ContactoProveedor>(entity =>
        {
            entity.HasKey(e => e.IdContactoProveedor);

            entity.ToTable("ContactosProveedores");

            entity.Property(e => e.TipoContacto)
                .HasMaxLength(20)
                .IsUnicode(false);

            entity.Property(e => e.Valor)
                .HasMaxLength(300)
                .IsUnicode(false);

            entity.Property(e => e.Descripcion)
                .HasMaxLength(50)
                .IsUnicode(false);

            entity.Property(e => e.Activo)
                .HasDefaultValue(true);

            entity.HasOne(e => e.Proveedor)
                .WithMany(p => p.ContactosAdicionales)
                .HasForeignKey(e => e.IdProveedor)
                .OnDelete(DeleteBehavior.ClientSetNull);
        });

        modelBuilder.Entity<Proyecto>(entity =>
        {
            entity.HasKey(e => e.IdProyecto).HasName("PK__Proyecto__F4888673AC791DA5");

            entity.Property(e => e.Activo).HasDefaultValue(true);
            entity.Property(e => e.Descripcion)
                .HasMaxLength(500)
                .IsUnicode(false);
            entity.Property(e => e.Estado)
                .HasMaxLength(30)
                .IsUnicode(false);
            entity.Property(e => e.Nombre)
                .HasMaxLength(150)
                .IsUnicode(false);
            entity.Property(e => e.PresupuestoAsignado).HasColumnType("decimal(18, 2)");
        });

        modelBuilder.Entity<Role>(entity =>
        {
            entity.HasKey(e => e.IdRol).HasName("PK__Roles__2A49584CB55B1F9C");

            entity.Property(e => e.Activo).HasDefaultValue(true);
            entity.Property(e => e.Descripcion)
                .HasMaxLength(250)
                .IsUnicode(false);
            entity.Property(e => e.FechaCreacion)
                .HasDefaultValueSql("(getdate())")
                .HasColumnType("datetime");
            entity.Property(e => e.Nombre)
                .HasMaxLength(100)
                .IsUnicode(false);
        });

        modelBuilder.Entity<RolesPermiso>(entity =>
        {
            entity.HasKey(e => e.IdRolPermiso).HasName("PK__RolesPer__0CC73B1B299BE339");

            entity.HasIndex(e => e.PermisoId, "IX_RolesPermisos_PermisoId");

            entity.HasIndex(e => e.RolId, "IX_RolesPermisos_RolId");

            entity.HasOne(d => d.Permiso).WithMany(p => p.RolesPermisos)
                .HasForeignKey(d => d.PermisoId)
                .OnDelete(DeleteBehavior.ClientSetNull)
                .HasConstraintName("FK_RolesPermisos_Permisos");

            entity.HasOne(d => d.Rol).WithMany(p => p.RolesPermisos)
                .HasForeignKey(d => d.RolId)
                .OnDelete(DeleteBehavior.ClientSetNull)
                .HasConstraintName("FK_RolesPermisos_Roles");
        });

        modelBuilder.Entity<Usuario>(entity =>
        {
            entity.HasKey(e => e.IdUsuario).HasName("PK__Usuarios__5B65BF97EE5F26FB");

            entity.HasIndex(e => e.RolId, "IX_Usuarios_RolId");

            entity.Property(e => e.Activo).HasDefaultValue(true);
            entity.Property(e => e.Email)
                .HasMaxLength(150)
                .IsUnicode(false);
            entity.Property(e => e.FechaCreacion)
                .HasDefaultValueSql("(getdate())")
                .HasColumnType("datetime");
            entity.Property(e => e.Nombre)
                .HasMaxLength(150)
                .IsUnicode(false);
            entity.Property(e => e.PasswordHash)
                .HasMaxLength(150)
                .IsUnicode(false);
            entity.Property(e => e.UltimoAcceso).HasColumnType("datetime");

            entity.HasOne(d => d.Rol).WithMany(p => p.Usuarios)
                .HasForeignKey(d => d.RolId)
                .OnDelete(DeleteBehavior.ClientSetNull)
                .HasConstraintName("FK_Usuarios_Roles");
        });

        OnModelCreatingPartial(modelBuilder);
    }

    partial void OnModelCreatingPartial(ModelBuilder modelBuilder);
}
