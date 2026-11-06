# Finanzas · Vicente

App personal para ver mis finanzas con los movimientos de mis bancos (Santander y Banco Falabella): ingresos vs gastos, gasto por centro de costo y cuenta (los mismos de Casa App), presupuesto mensual, tarjetas y cuotas por pagar, y lo que falta categorizar.

- **App:** `index.html` (HTML + JS sin dependencias), publicada con GitHub Pages. Se entra con Google y después con Face ID.
- **Backend:** `apps-script/Code.gs`, un Google Apps Script pegado en la planilla **Finanzas**.
- **Datos:** Google Sheets (hojas Movimientos, Tarjetas, Presupuesto, Reglas, Centros, Cuentas, Importaciones).
- **Cartolas:** se convierten con `importar/normalizar.py` y se cargan desde la carpeta de Drive **Finanzas – importar** (ver [importar/README.md](importar/README.md)). Nunca se suben a GitHub.

## Pestañas

| Pestaña | Qué muestra |
|---|---|
| 📊 Resumen | Ahorro del mes, ingresos, gastos, pagos de tarjeta, 12 meses en barras y los mayores gastos |
| 🍩 Categorías | Dona por centro de costo (mes o año); cada centro se abre por cuenta y movimiento |
| 🎯 Presupuesto | Tope mensual por centro (y por cuenta), con barra y la rayita del ritmo del mes |
| 💳 Tarjetas | Deuda, monto a pagar y vencimiento de cada tarjeta, y cuotas por pagar los próximos 6 meses |
| 🔎 Revisar | Movimientos sin categoría y buscador. Al categorizar se puede crear una regla para todos los iguales |

Instalación: ver [SETUP.md](SETUP.md).
