using Microsoft.EntityFrameworkCore;
using SistemaGestionFinanciera.Data;
using SistemaGestionFinanciera.Models;

namespace SistemaGestionFinanciera.Services
{
    public class ContabilidadService
    {
        private readonly SistemaFinancieroContext _context;

        public ContabilidadService(SistemaFinancieroContext context)
        {
            _context = context;
        }
        // Obtiene la cuenta Caja General.
        private async Task<int> ObtenerCuentaCajaAsync()
        {
            var cuenta = await _context.CuentasContables
                .FirstOrDefaultAsync(c => c.Nombre == "Caja General" && c.Activo == true);

            if (cuenta == null)
                throw new Exception("No existe una cuenta contable activa llamada Caja General.");

            return cuenta.IdCuentaContable;
        }

        // Obtiene la cuenta Cuentas por Cobrar.
        private async Task<int> ObtenerCuentaCxCAsync()
        {
            var cuenta = await _context.CuentasContables
                .FirstOrDefaultAsync(c => c.Nombre == "Cuentas por Cobrar" && c.Activo == true);

            if (cuenta == null)
                throw new Exception("No existe una cuenta contable activa llamada Cuentas por Cobrar.");

            return cuenta.IdCuentaContable;
        }

        // Obtiene la cuenta Cuentas por Pagar.
        private async Task<int> ObtenerCuentaCxPAsync()
        {
            var cuenta = await _context.CuentasContables
                .FirstOrDefaultAsync(c => c.Nombre == "Cuentas por Pagar" && c.Activo == true);

            if (cuenta == null)
                throw new Exception("No existe una cuenta contable activa llamada Cuentas por Pagar.");

            return cuenta.IdCuentaContable;
        }


        //***************************** ASIENTO CONTABLE PARA INGRESOS *****************************
        // Registra el asiento contable de un ingreso manual pagado.
        // Genera exactamente 2 movimientos:
        // 1. Debe  -> Caja/Bancos
        // 2. Haber -> Cuenta de ingreso
        public async Task RegistrarIngresoAsync(
    Ingreso ingreso,
    int usuarioId)
        {
            decimal monto = ingreso.MontoReal ?? 0;

            // Si el ingreso no tiene monto real, no se contabiliza.
            if (monto <= 0)
            {
                return;
            }

            // Solo se contabilizan ingresos pagados.
            if (ingreso.Estado != "Pagado")
            {
                return;
            }

            int cuentaDebeId = await ObtenerCuentaCajaAsync();

            var categoria = await _context.CategoriasIngresos
                .FirstOrDefaultAsync(c => c.IdCategoriaIngreso == ingreso.CategoriaIngresoId);

            if (categoria == null)
                throw new Exception("No se encontró la categoría de ingreso.");

            if (categoria.CuentaContableId == null)
                throw new Exception("La categoría de ingreso no tiene una cuenta contable asignada.");

            int cuentaHaberId = categoria.CuentaContableId.Value;


            // Evita duplicar movimientos, pero no cuenta los movimientos de reversión.
            bool yaContabilizado = await _context.MovimientosContables
                .AnyAsync(m =>
                    m.OrigenModulo == "Ingresos" &&
                    m.OrigenId == ingreso.IdIngreso &&
                    m.Anulado == false &&
                    m.Estado != "Reversado");

            if (yaContabilizado)
            {
                return;
            }

            // Verifica que las cuentas contables existan y estén activas.
            bool cuentaDebeExiste = await _context.CuentasContables
                .AnyAsync(c => c.IdCuentaContable == cuentaDebeId && c.Activo == true);

            bool cuentaHaberExiste = await _context.CuentasContables
                .AnyAsync(c => c.IdCuentaContable == cuentaHaberId && c.Activo == true);

            if (!cuentaDebeExiste || !cuentaHaberExiste)
            {
                throw new Exception("No se encontraron las cuentas contables activas para registrar el ingreso.");
            }

            // Movimiento al Debe.
            var movimientoDebe = new MovimientosContable
            {
                Fecha = ingreso.Fecha,
                CuentaContableId = cuentaDebeId,
                ProyectoId = ingreso.ProyectoId,
                CentroCostoId = ingreso.CentroCostoId,
                TipoMovimiento = "Debe",
                Monto = monto,
                Debe = monto,
                Haber = 0,
                Descripcion = "Registro contable automático del ingreso: " + ingreso.Fuente,
                Referencia = ingreso.Fuente,
                OrigenModulo = "Ingresos",
                OrigenId = ingreso.IdIngreso,
                Anulado = false,
                FechaCreacion = DateTime.Now,
                Estado = "Registrado",
                EsAutomatico = true,
                UsuarioId = usuarioId
            };

            // Movimiento al Haber.
            var movimientoHaber = new MovimientosContable
            {
                Fecha = ingreso.Fecha,
                CuentaContableId = cuentaHaberId,
                ProyectoId = ingreso.ProyectoId,
                CentroCostoId = ingreso.CentroCostoId,
                TipoMovimiento = "Haber",
                Monto = monto,
                Debe = 0,
                Haber = monto,
                Descripcion = "Registro contable automático del ingreso: " + ingreso.Fuente,
                Referencia = ingreso.Fuente,
                OrigenModulo = "Ingresos",
                OrigenId = ingreso.IdIngreso,
                Anulado = false,
                FechaCreacion = DateTime.Now,
                Estado = "Registrado",
                EsAutomatico = true,
                UsuarioId = usuarioId
            };

            _context.MovimientosContables.Add(movimientoDebe);
            _context.MovimientosContables.Add(movimientoHaber);

            await _context.SaveChangesAsync();

            // Recalcula los saldos con base en los movimientos reales.
            await RecalcularSaldoCuentaAsync(cuentaDebeId);
            await RecalcularSaldoCuentaAsync(cuentaHaberId);
        }
        //   ***************** ASIENTO CONTABLE DE GASTO ******************************
        // Registra el asiento contable de un gasto aprobado.
        // Genera exactamente 2 movimientos:
        // 1. Debe  -> Cuenta de gasto
        // 2. Haber -> Caja/Bancos
        public async Task RegistrarGastoAsync(
      Gasto gasto,
      int usuarioId)
        {
            decimal monto = gasto.MontoTotal ?? 0;

            // Si el gasto no tiene monto, no se contabiliza.
            if (monto <= 0)
            {
                return;
            }

            // Solo se contabilizan gastos aprobados.
            if (gasto.Estado != "Aprobado")
            {
                return;
            }
            var categoria = await _context.CategoriasGastos
    .FirstOrDefaultAsync(c => c.IdCategoriaGasto == gasto.CategoriaGastoId);

            if (categoria == null)
                throw new Exception("No se encontró la categoría de gasto.");

            if (categoria.CuentaContableId == null)
                throw new Exception("La categoría de gasto no tiene una cuenta contable asignada.");

            int cuentaDebeId = categoria.CuentaContableId.Value;
            int cuentaHaberId = await ObtenerCuentaCajaAsync();
            // Evita duplicar movimientos, pero no cuenta los movimientos de reversión.
            bool yaContabilizado = await _context.MovimientosContables
                .AnyAsync(m =>
                    m.OrigenModulo == "Gastos" &&
                    m.OrigenId == gasto.IdGasto &&
                    m.Anulado == false &&
                    m.Estado != "Reversado");

            if (yaContabilizado)
            {
                return;
            }

            // Verifica que las cuentas contables existan y estén activas.
            bool cuentaDebeExiste = await _context.CuentasContables
                .AnyAsync(c => c.IdCuentaContable == cuentaDebeId && c.Activo == true);

            bool cuentaHaberExiste = await _context.CuentasContables
                .AnyAsync(c => c.IdCuentaContable == cuentaHaberId && c.Activo == true);

            if (!cuentaDebeExiste || !cuentaHaberExiste)
            {
                throw new Exception("No se encontraron las cuentas contables activas para registrar el gasto.");
            }

            // Movimiento al Debe.
            var movimientoDebe = new MovimientosContable
            {
                Fecha = gasto.Fecha,
                CuentaContableId = cuentaDebeId,
                ProyectoId = gasto.ProyectoId,
                CentroCostoId = gasto.CentroCostoId,
                TipoMovimiento = "Debe",
                Monto = monto,
                Debe = monto,
                Haber = 0,
                Descripcion = "Registro contable automático del gasto: " + gasto.Concepto,
                Referencia = gasto.NumeroFactura ?? gasto.Concepto,
                OrigenModulo = "Gastos",
                OrigenId = gasto.IdGasto,
                Anulado = false,
                FechaCreacion = DateTime.Now,
                Estado = "Registrado",
                EsAutomatico = true,
                UsuarioId = usuarioId
            };

            // Movimiento al Haber.
            var movimientoHaber = new MovimientosContable
            {
                Fecha = gasto.Fecha,
                CuentaContableId = cuentaHaberId,
                ProyectoId = gasto.ProyectoId,
                CentroCostoId = gasto.CentroCostoId,
                TipoMovimiento = "Haber",
                Monto = monto,
                Debe = 0,
                Haber = monto,
                Descripcion = "Registro contable automático del gasto: " + gasto.Concepto,
                Referencia = gasto.NumeroFactura ?? gasto.Concepto,
                OrigenModulo = "Gastos",
                OrigenId = gasto.IdGasto,
                Anulado = false,
                FechaCreacion = DateTime.Now,
                Estado = "Registrado",
                EsAutomatico = true,
                UsuarioId = usuarioId
            };

            _context.MovimientosContables.Add(movimientoDebe);
            _context.MovimientosContables.Add(movimientoHaber);

            await _context.SaveChangesAsync();

            // Recalcula los saldos con base en los movimientos reales.
            await RecalcularSaldoCuentaAsync(cuentaDebeId);
            await RecalcularSaldoCuentaAsync(cuentaHaberId);
        }
        //************** ASIENTO CONTABLE DE VENTA CONTADO **********************

        // Registra asiento contable para venta contado.
        // Debe: Caja/Bancos
        // Haber: Ingreso
        public async Task RegistrarVentaContadoAsync(Factura factura, int usuarioId)
        {
            if (factura.Total <= 0) return;

            if (factura.TipoFactura != "Venta" || factura.CondicionPago != "Contado") return;

            bool yaContabilizado = await _context.MovimientosContables
                .AnyAsync(m =>
                    m.OrigenModulo == "Facturas" &&
                    m.OrigenId == factura.IdFactura &&
                    m.Anulado == false &&
                    m.Estado != "Reversado");

            if (yaContabilizado) return;

            int cuentaCajaId = await ObtenerCuentaCajaAsync();

            var categoria = await _context.CategoriasIngresos
                .FirstOrDefaultAsync(c => c.IdCategoriaIngreso == factura.CategoriaIngresoId);

            if (categoria == null)
                throw new Exception("No se encontró la categoría de ingreso de la factura.");

            if (categoria.CuentaContableId == null)
                throw new Exception("La categoría de ingreso no tiene una cuenta contable asignada.");

            int cuentaIngresoId = categoria.CuentaContableId.Value;

            await CrearAsientoSimpleAsync(
                factura.FechaEmision,
                cuentaCajaId,
                cuentaIngresoId,
                factura.Total,
                factura.ProyectoId,
                factura.CentroCostoId,
                "Registro contable automático de venta contado: " + factura.NumeroFactura,
                factura.NumeroFactura,
                "Facturas",
                factura.IdFactura,
                usuarioId);
        }

//*********************** ASEINTO CONTABLE PARA VENTA CREDITO *********************************
        // Registra asiento contable para venta crédito.
        // Debe: Cuentas por Cobrar
        // Haber: Ingreso
        public async Task RegistrarVentaCreditoAsync(Factura factura, int usuarioId)
        {
            if (factura.Total <= 0) return;

            if (factura.TipoFactura != "Venta" ||
                (factura.CondicionPago != "Credito" && factura.CondicionPago != "Crédito"))
            {
                return;
            }

            bool yaContabilizado = await _context.MovimientosContables
                .AnyAsync(m =>
                    m.OrigenModulo == "Facturas" &&
                    m.OrigenId == factura.IdFactura &&
                    m.Anulado == false &&
                    m.Estado != "Reversado");

            if (yaContabilizado) return;

            int cuentaCxCId = await ObtenerCuentaCxCAsync();

            var categoria = await _context.CategoriasIngresos
                .FirstOrDefaultAsync(c => c.IdCategoriaIngreso == factura.CategoriaIngresoId);

            if (categoria == null)
                throw new Exception("No se encontró la categoría de ingreso de la factura.");

            if (categoria.CuentaContableId == null)
                throw new Exception("La categoría de ingreso no tiene una cuenta contable asignada.");

            int cuentaIngresoId = categoria.CuentaContableId.Value;

            await CrearAsientoSimpleAsync(
                factura.FechaEmision,
                cuentaCxCId,
                cuentaIngresoId,
                factura.Total,
                factura.ProyectoId,
                factura.CentroCostoId,
                "Registro contable automático de venta crédito: " + factura.NumeroFactura,
                factura.NumeroFactura,
                "Facturas",
                factura.IdFactura,
                usuarioId);
        }

//***************** ASIENTO CONTABLE PARA COMPRA CONTADO ********************************

        // Registra asiento contable para compra contado.
        // Debe: Gasto
        // Haber: Caja/Bancos
        public async Task RegistrarCompraContadoAsync(Factura factura, int usuarioId)
        {
            if (factura.Total <= 0) return;

            if (factura.TipoFactura != "Compra" || factura.CondicionPago != "Contado") return;

            bool yaContabilizado = await _context.MovimientosContables
                .AnyAsync(m =>
                    m.OrigenModulo == "Facturas" &&
                    m.OrigenId == factura.IdFactura &&
                    m.Anulado == false &&
                    m.Estado != "Reversado");

            if (yaContabilizado) return;

            int cuentaCajaId = await ObtenerCuentaCajaAsync();

            var categoria = await _context.CategoriasGastos
                .FirstOrDefaultAsync(c => c.IdCategoriaGasto == factura.CategoriaGastoId);

            if (categoria == null)
                throw new Exception("No se encontró la categoría de gasto de la factura.");

            if (categoria.CuentaContableId == null)
                throw new Exception("La categoría de gasto no tiene una cuenta contable asignada.");

            int cuentaGastoId = categoria.CuentaContableId.Value;

            await CrearAsientoSimpleAsync(
                factura.FechaEmision,
                cuentaGastoId,
                cuentaCajaId,
                factura.Total,
                factura.ProyectoId,
                factura.CentroCostoId,
                "Registro contable automático de compra contado: " + factura.NumeroFactura,
                factura.NumeroFactura,
                "Facturas",
                factura.IdFactura,
                usuarioId);
        }

// *************************** ASEINTO CONTABLE PARA COMRPA CREDITO ********************************

        // Registra asiento contable para compra crédito.
        // Debe: Gasto
        // Haber: Cuentas por Pagar
        public async Task RegistrarCompraCreditoAsync(Factura factura, int usuarioId)
        {
            if (factura.Total <= 0) return;

            if (factura.TipoFactura != "Compra" ||
                (factura.CondicionPago != "Credito" && factura.CondicionPago != "Crédito"))
            {
                return;
            }

            bool yaContabilizado = await _context.MovimientosContables
                .AnyAsync(m =>
                    m.OrigenModulo == "Facturas" &&
                    m.OrigenId == factura.IdFactura &&
                    m.Anulado == false &&
                    m.Estado != "Reversado");

            if (yaContabilizado) return;

            int cuentaCxPId = await ObtenerCuentaCxPAsync();

            var categoria = await _context.CategoriasGastos
                .FirstOrDefaultAsync(c => c.IdCategoriaGasto == factura.CategoriaGastoId);

            if (categoria == null)
                throw new Exception("No se encontró la categoría de gasto de la factura.");

            if (categoria.CuentaContableId == null)
                throw new Exception("La categoría de gasto no tiene una cuenta contable asignada.");

            int cuentaGastoId = categoria.CuentaContableId.Value;

            await CrearAsientoSimpleAsync(
                factura.FechaEmision,
                cuentaGastoId,
                cuentaCxPId,
                factura.Total,
                factura.ProyectoId,
                factura.CentroCostoId,
                "Registro contable automático de compra crédito: " + factura.NumeroFactura,
                factura.NumeroFactura,
                "Facturas",
                factura.IdFactura,
                usuarioId);
        }

        // Crea un asiento contable simple de dos líneas:
        // una línea al Debe y una línea al Haber.
        private async Task CrearAsientoSimpleAsync(
            DateOnly fecha,
            int cuentaDebeId,
            int cuentaHaberId,
            decimal monto,
            int? proyectoId,
            int? centroCostoId,
            string descripcion,
            string? referencia,
            string origenModulo,
            int origenId,
            int usuarioId)
        {
            bool cuentaDebeExiste = await _context.CuentasContables
                .AnyAsync(c => c.IdCuentaContable == cuentaDebeId && c.Activo == true);

            bool cuentaHaberExiste = await _context.CuentasContables
                .AnyAsync(c => c.IdCuentaContable == cuentaHaberId && c.Activo == true);

            if (!cuentaDebeExiste || !cuentaHaberExiste)
            {
                throw new Exception("No se encontraron las cuentas contables activas para registrar el asiento.");
            }

            var movimientoDebe = new MovimientosContable
            {
                Fecha = fecha,
                CuentaContableId = cuentaDebeId,
                ProyectoId = proyectoId,
                CentroCostoId = centroCostoId,
                TipoMovimiento = "Debe",
                Monto = monto,
                Debe = monto,
                Haber = 0,
                Descripcion = descripcion,
                Referencia = referencia,
                OrigenModulo = origenModulo,
                OrigenId = origenId,
                Anulado = false,
                FechaCreacion = DateTime.Now,
                Estado = "Registrado",
                EsAutomatico = true,
                UsuarioId = usuarioId
            };

            var movimientoHaber = new MovimientosContable
            {
                Fecha = fecha,
                CuentaContableId = cuentaHaberId,
                ProyectoId = proyectoId,
                CentroCostoId = centroCostoId,
                TipoMovimiento = "Haber",
                Monto = monto,
                Debe = 0,
                Haber = monto,
                Descripcion = descripcion,
                Referencia = referencia,
                OrigenModulo = origenModulo,
                OrigenId = origenId,
                Anulado = false,
                FechaCreacion = DateTime.Now,
                Estado = "Registrado",
                EsAutomatico = true,
                UsuarioId = usuarioId
            };

            _context.MovimientosContables.Add(movimientoDebe);
            _context.MovimientosContables.Add(movimientoHaber);

            await _context.SaveChangesAsync();

            await RecalcularSaldoCuentaAsync(cuentaDebeId);
            await RecalcularSaldoCuentaAsync(cuentaHaberId);
        }


        // Recalcula el saldo actual de una cuenta contable usando sus movimientos no anulados.
        // Esto evita depender del monto actual del ingreso, gasto, factura o pago.
        public async Task RecalcularSaldoCuentaAsync(int cuentaContableId)
        {
            var cuenta = await _context.CuentasContables
                .FirstOrDefaultAsync(c => c.IdCuentaContable == cuentaContableId);

            if (cuenta == null)
            {
                return;
            }

            decimal totalDebe = await _context.MovimientosContables
                .Where(m =>
    m.CuentaContableId == cuentaContableId &&
    m.Anulado == false &&
    m.Estado != "Reversado")
               .SumAsync(m => m.Debe);

            decimal totalHaber = await _context.MovimientosContables
                .Where(m =>
    m.CuentaContableId == cuentaContableId &&
    m.Anulado == false &&
    m.Estado != "Reversado")
                .SumAsync(m => m.Haber);
            // Las cuentas de naturaleza deudora aumentan por el Debe.
            if (cuenta.Naturaleza == "Deudora")
            {
                cuenta.SaldoActual = cuenta.SaldoInicial + totalDebe - totalHaber;
            }

            // Las cuentas de naturaleza acreedora aumentan por el Haber. 
            else if (cuenta.Naturaleza == "Acreedora")
            {
                cuenta.SaldoActual = cuenta.SaldoInicial + totalHaber - totalDebe;
            }

            await _context.SaveChangesAsync();
        }

        
        // Reversa los movimientos contables originales de una operación.
        // Usa los movimientos originales, no el monto actual del ingreso, gasto, factura o pago.
        public async Task ReversarMovimientosAsync(
        string origenModulo,
        int origenId,
        int usuarioId,
        string motivoAnulacion)
    {
            // Busca solo movimientos originales activos.
            // No debe reversar movimientos que ya son reversión.
            var movimientosOriginales = await _context.MovimientosContables
                .Where(m =>
                    m.OrigenModulo == origenModulo &&
                    m.OrigenId == origenId &&
                    m.Anulado == false &&
                    m.Estado != "Reversado")
                .ToListAsync();

            // Si no hay movimientos, no hay nada que reversar.
            if (!movimientosOriginales.Any())
        {
            return;
        }

        var cuentasAfectadas = movimientosOriginales
            .Select(m => m.CuentaContableId)
            .Distinct()
            .ToList();

        foreach (var movimiento in movimientosOriginales)
        {
            // Marca el movimiento original como anulado.
            movimiento.Anulado = true;
            movimiento.Estado = "Anulado";
            movimiento.FechaAnulacion = DateTime.Now;
            movimiento.MotivoAnulacion = motivoAnulacion;

            // Crea el movimiento inverso usando el movimiento original.
            var movimientoReversion = new MovimientosContable
            {
                Fecha = DateOnly.FromDateTime(DateTime.Today),
                CuentaContableId = movimiento.CuentaContableId,
                ProyectoId = movimiento.ProyectoId,
                CentroCostoId = movimiento.CentroCostoId,
                TipoMovimiento = movimiento.Debe > 0 ? "Haber" : "Debe",
                Monto = movimiento.Monto,
                Debe = movimiento.Haber,
                Haber = movimiento.Debe,
                Descripcion = "Reversión contable de: " + movimiento.Descripcion,
                Referencia = movimiento.Referencia,
                OrigenModulo = origenModulo,
                OrigenId = origenId,
                Anulado = false,
                FechaCreacion = DateTime.Now,
                Estado = "Reversado",
                EsAutomatico = true,
                UsuarioId = usuarioId,
                MovimientoReversionId = movimiento.IdMovimientoContable
            };

            _context.MovimientosContables.Add(movimientoReversion);
        }

        await _context.SaveChangesAsync();

        // Recalcula los saldos de las cuentas afectadas.
        foreach (var cuentaId in cuentasAfectadas)
        {
            await RecalcularSaldoCuentaAsync(cuentaId);
        }     
      }

   }
}