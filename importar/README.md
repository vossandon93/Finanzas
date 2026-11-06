# Cómo se cargan las cartolas

```
Vicente  ──►  Claude  ──►  normalizar.py  ──►  CSV  ──►  Drive "Finanzas – importar"  ──►  Apps Script  ──►  planilla
(cartola)                  (formato único)                (cada hora o "Importar ahora")
```

1. Vicente le pasa a Claude las cartolas (Excel de preferencia; PDF también sirve).
2. Claude corre `normalizar.py` y revisa con Vicente el resumen (cuántos movimientos, cuánto entra y sale).
3. Claude deja el CSV en la carpeta de Drive **Finanzas – importar**.
4. El script lo carga **sin duplicar**: si el mismo archivo o un mes que se cruza llega dos veces, los movimientos repetidos se saltan.
5. El archivo se mueve a **Finanzas – importar / Procesadas** y queda anotado en la hoja **Importaciones**.

## Formato del CSV

Movimientos (`fecha,banco,producto,descripcion,monto,cuotas,tipo,centro,cuenta,nota`):

| Columna | Ejemplo | Nota |
|---|---|---|
| fecha | `2026-10-03` | también acepta `03/10/2026` |
| banco | `Santander` | |
| producto | `Tarjeta Santander`, `CMR Falabella`, `Cuenta corriente Santander` | si dice tarjeta/CMR/Visa se trata como tarjeta |
| descripcion | `LIDER EXPRESS` | lo que dice la cartola |
| monto | `-85000` | **negativo sale, positivo entra** |
| cuotas | `3/12` | opcional |
| tipo | `Gasto`, `Ingreso`, `Pago tarjeta`, `Traspaso`, `Ignorar` | opcional; si falta, lo decide el script |
| centro, cuenta, nota | | opcionales |

Estado de cuenta de tarjeta (`producto,fecha_estado,deuda_total,monto_a_pagar,fecha_pago,cupo_total`): una fila por tarjeta, para la pestaña Tarjetas.

## Para no contar doble

- El **pago de la tarjeta** desde la cuenta corriente es tipo `Pago tarjeta`: no es gasto (los gastos ya están en la cartola de la tarjeta).
- Los **traspasos** entre tus cuentas son tipo `Traspaso`.
- Las compras en cuotas se cuentan **cuota a cuota**, en el mes en que se cobran (como aparecen en el estado de cuenta).
